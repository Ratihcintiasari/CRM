#!/usr/bin/env python3
"""Grafik 'Close Lost Unit' per minggu, stacked per Status Reason.

Dijalankan otomatis oleh GitHub Actions setiap file di folder data/ berubah.
Hasil: docs/index.html, docs/close_lost_unit.png, docs/close_lost_unit.xlsx
(ditampilkan sebagai halaman web docs/index.html lewat GitHub Pages).

Jalankan manual (opsional):  python generate_chart.py
"""
import calendar
import re
import sys
import textwrap
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D
from matplotlib.patches import Patch

# =====================================================================
# KONFIGURASI  (ubah bagian ini saja)
# =====================================================================
FILE_INPUT = "data/close_lost_unit.xlsx"   # kalau tidak ada, dipakai file .xlsx/.csv terbaru di folder data/
SHEET = 0
TAHUN = 2026
BULAN = 9
MODE_MINGGU = "senin"              # "senin"   = minggu Senin-Minggu
                                   # "tanggal" = W1: tgl 1-7, W2: 8-14, W3: 15-21, W4: 22-28, W5: 29+
SATUAN_BAGI = 1e9                  # IDR Bio
SEMBUNYIKAN_KOSONG = True          # kategori bernilai 0 di semua minggu tidak ditampilkan

# Nama kolom di file (cocok jika nama kolom MENGANDUNG teks ini, urut prioritas)
KOLOM_TANGGAL = ["actual close date", "est. close date"]   # kalau kosong, pakai yang berikutnya
KOLOM_ALASAN = ["status reason"]
KOLOM_NILAI = ["est. revenue"]

# Close Lost Unspecified
TAMPILKAN_UNSPECIFIED = True       # selalu tampil di grafik & tabel walau nilainya 0
KATA_UNSPECIFIED = "unspecified"
NAMA_UNSPECIFIED = "Close Lost Unspecified"
NILAI_MANUAL_UNSPECIFIED = {}      # input manual {minggu: IDR Bio}, contoh {2: 0.75, 4: 1.1}

FILE_PNG = "docs/close_lost_unit.png"
FILE_XLSX = "docs/close_lost_unit.xlsx"
FILE_HTML = "docs/index.html"
# =====================================================================

BLN_SINGKAT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
BLN_PENUH = ["January", "February", "March", "April", "May", "June", "July",
             "August", "September", "October", "November", "December"]
NAMA_BULAN = BLN_SINGKAT[BULAN - 1]
JUDUL = f"CLOSE LOST UNIT {BLN_PENUH[BULAN - 1].upper()} {TAHUN} - PER WEEK"

WARNA_TETAP = {"Product Spec": "#1B6B8F", "Lead Time": "#ED7D31",
               "Price": "#1E6B2A", "Others": "#7030A0", NAMA_UNSPECIFIED: "#B03A2E"}
PALET = ["#E6B800", "#C0392B", "#2E86C1", "#7F8C8D", "#16A085",
         "#D35400", "#8E6E53", "#2C3E50", "#AF7AC5", "#58D68D"]


# ---------- fungsi bantu ----------
def cari_kolom(df, daftar, wajib=True):
    for teks in daftar:
        for c in df.columns:
            if teks in str(c).strip().lower():
                return c
    if wajib:
        sys.exit(f"ERROR: kolom {daftar} tidak ditemukan.\nKolom di file: {list(df.columns)}")
    return None


def parse_rupiah(v):
    if pd.isna(v):
        return np.nan
    if isinstance(v, (int, float, np.number)):
        return float(v)
    s = re.sub(r"[^\d.,]", "", str(v))
    if not s:
        return np.nan
    if s.count(".") > 1:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    return float(s)


def nama_kategori(label):
    baku = {"price": "Price", "product spec": "Product Spec",
            "lead time": "Lead Time", "others": "Others", "other": "Others"}
    l = re.sub(r"\s+", " ", str(label)).strip()
    return baku.get(l.lower(), l)


def fmt_angka(x):
    if x == 0 or pd.isna(x):
        return "-"
    return f"{x + 1e-9:,.1f}".replace(",", "_").replace(".", ",").replace("_", ".")


def fmt_rp(bio):
    return "Rp" + f"{bio * SATUAN_BAGI:,.0f}".replace(",", ".")


# ---------- minggu ----------
HARI_TERAKHIR = calendar.monthrange(TAHUN, BULAN)[1]
OFFSET = calendar.weekday(TAHUN, BULAN, 1) if MODE_MINGGU == "senin" else 0
JUMLAH_MINGGU = (HARI_TERAKHIR + OFFSET - 1) // 7 + 1


def nomor_minggu(hari):
    return (hari - 1 + OFFSET) // 7 + 1


def label_minggu(w):
    awal = max(1, (w - 1) * 7 - OFFSET + 1)
    akhir = min(HARI_TERAKHIR, w * 7 - OFFSET)
    return f"W{w} ({awal}-{akhir} {NAMA_BULAN})"


# ---------- baca & olah ----------
def temukan_file():
    p = Path(FILE_INPUT)
    if p.exists():
        return p
    kandidat = [f for f in Path("data").glob("*") if f.suffix.lower() in (".xlsx", ".xls", ".xlsm", ".csv")]
    if not kandidat:
        sys.exit(f"ERROR: file '{FILE_INPUT}' tidak ditemukan dan folder data/ kosong.")
    terbaru = max(kandidat, key=lambda f: f.stat().st_mtime)
    print(f"Info: '{FILE_INPUT}' tidak ada, memakai file terbaru: {terbaru}")
    return terbaru


def baca():
    p = temukan_file()
    df = (pd.read_excel(p, sheet_name=SHEET) if p.suffix.lower() in (".xlsx", ".xls", ".xlsm")
          else pd.read_csv(p, sep=None, engine="python", encoding="utf-8-sig"))
    print("File:", p, "| jumlah baris:", len(df))

    kol_alasan = cari_kolom(df, KOLOM_ALASAN)
    kol_nilai = cari_kolom(df, KOLOM_NILAI)

    tgl = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    dipakai = []
    for teks in KOLOM_TANGGAL:
        c = cari_kolom(df, [teks], wajib=False)
        if c is not None:
            tgl = tgl.fillna(pd.to_datetime(df[c], errors="coerce", dayfirst=True))
            dipakai.append(c)
    if not dipakai:
        sys.exit(f"ERROR: kolom tanggal {KOLOM_TANGGAL} tidak ditemukan.\nKolom di file: {list(df.columns)}")
    print(f"Kolom alasan : {kol_alasan!r}\nKolom nilai  : {kol_nilai!r}\nKolom tanggal: {dipakai}")

    out = pd.DataFrame({
        "kategori": df[kol_alasan].map(lambda v: nama_kategori(v) if pd.notna(v) else "(Tanpa Reason)"),
        "nilai": df[kol_nilai].map(parse_rupiah) / SATUAN_BAGI,
        "tgl": tgl,
        "manual": False,
    })

    tanpa = out[out["tgl"].isna()]
    if len(tanpa):
        print(f"PERINGATAN: {len(tanpa)} baris tanpa tanggal tidak dihitung "
              f"(nilai {tanpa['nilai'].sum():,.3f}).")
    ok = (out["tgl"].dt.year == TAHUN) & (out["tgl"].dt.month == BULAN)
    luar = out[out["tgl"].notna() & ~ok]
    if len(luar):
        print(f"Info: {len(luar)} baris di luar {NAMA_BULAN} {TAHUN} diabaikan (nilai {luar['nilai'].sum():,.3f}).")
    out = out[ok].copy()
    out["nilai"] = out["nilai"].fillna(0.0)
    out["minggu"] = out["tgl"].dt.day.map(nomor_minggu)

    if NILAI_MANUAL_UNSPECIFIED:
        tambahan = pd.DataFrame([
            {"kategori": NAMA_UNSPECIFIED, "nilai": float(v), "tgl": pd.NaT,
             "manual": True, "minggu": int(w)} for w, v in NILAI_MANUAL_UNSPECIFIED.items()])
        out = pd.concat([out, tambahan], ignore_index=True)
    return out


def buat_keterangan(df):
    u = df[df["kategori"].str.lower().str.contains(KATA_UNSPECIFIED, na=False)]
    if u.empty:
        return "Keterangan: tidak ada opportunity dengan Status Reason Close Lost Unspecified."
    per = u.groupby("minggu").agg(jml=("manual", lambda s: int((~s).sum())),
                                  nilai=("nilai", "sum"), manual=("manual", "any"))
    rincian = "; ".join(
        f"{label_minggu(w)}: {int(r.jml)} opp ({fmt_rp(r.nilai)})" + (" [input manual]" if r.manual else "")
        for w, r in per.iterrows())
    total = u["nilai"].sum()
    akhir = (" Bernilai Rp0, sehingga tidak tampil di batang/tabel." if abs(total) < 1e-12
             else " Nilainya sudah termasuk di grafik.")
    return (f"Keterangan: {int((~u['manual']).sum())} opportunity dengan Status Reason "
            f"Close Lost Unspecified (total {fmt_rp(total)}). Rincian: {rincian}.{akhir}")


def hitung(df):
    pv = (df.groupby(["minggu", "kategori"])["nilai"].sum().unstack(fill_value=0)
            .reindex(index=range(1, JUMLAH_MINGGU + 1), fill_value=0))
    if TAMPILKAN_UNSPECIFIED and not any(KATA_UNSPECIFIED in str(k).lower() for k in pv.columns):
        pv[NAMA_UNSPECIFIED] = 0.0
    urut = pv.sum().sort_values(ascending=False).index.tolist()
    pv = pv[urut]

    def tampil(k):
        if TAMPILKAN_UNSPECIFIED and KATA_UNSPECIFIED in str(k).lower():
            return True
        return not (SEMBUNYIKAN_KOSONG and abs(pv[k].sum()) < 1e-12)

    kat = [k for k in urut if tampil(k)]
    pv = pv[kat]
    pv["Total"] = pv[kat].sum(axis=1)
    pv.index = [label_minggu(w) for w in pv.index]

    warna, i = dict(WARNA_TETAP), 0
    for k in kat:
        if k not in warna:
            warna[k] = PALET[i % len(PALET)]
            i += 1
    return pv, kat, warna


# ---------- gambar ----------
def gambar(pv, kat, warna, keterangan, data_as_of):
    n = len(pv)
    x = np.arange(n)
    nb = len(kat) + 1
    ket = textwrap.fill(keterangan, 135)
    h_ket = 0.22 * (ket.count("\n") + 1)
    H = 1.0 + 3.6 + 0.6 + 0.3 * (nb + 1) + 1.0 + 0.3 + h_ket
    fig = plt.figure(figsize=(11, H))

    def pos(bawah_in, tinggi_in):
        return [0.11, bawah_in / H, 0.78, tinggi_in / H]

    y_tab = 0.9 + 0.3 + h_ket
    h_tab = 0.3 * (nb + 1)
    ax_tab = fig.add_axes(pos(y_tab, h_tab))
    ax_tab.axis("off")
    ax = fig.add_axes(pos(y_tab + h_tab + 0.6, 3.6))

    bawah = np.zeros(n)
    for a in kat:
        ax.bar(x, pv[a], 0.45, bottom=bawah, color=warna[a], zorder=2)
        bawah += pv[a].values

    ax.plot(x, pv["Total"], color="black", ls="--", marker="o", ms=4, lw=1.1, zorder=4)
    for xi, v in zip(x, pv["Total"]):
        ax.annotate(fmt_angka(v), (xi, v), textcoords="offset points",
                    xytext=(0, 7), ha="center", fontsize=9, fontweight="bold", color="#333")

    ymax = max(pv["Total"].max(), 1) * 1.2
    ax.set_xlim(-0.5, n - 0.5)
    ax.set_ylim(0, ymax)
    ax.set_xticks(x)
    ax.set_xticklabels(pv.index, fontsize=9, color="#555")
    ax.yaxis.set_major_formatter(lambda v, _: fmt_angka(v))
    ax.tick_params(axis="y", labelsize=8, colors="#555", length=0)
    ax.tick_params(axis="x", length=0)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color("#bbbbbb")

    ax2 = ax.twinx()
    ax2.set_ylim(0, ymax)
    ax2.set_yticks(ax.get_yticks())
    ax2.set_ylim(0, ymax)
    ax2.yaxis.set_major_formatter(lambda v, _: fmt_angka(v))
    ax2.tick_params(axis="y", labelsize=8, colors="#555", length=0)
    for s in ("top", "right", "left", "bottom"):
        ax2.spines[s].set_visible(False)

    fig.text(0.5, 1 - 0.45 / H, JUDUL, ha="center", fontsize=15, fontweight="bold", color="#444")
    fig.text(0.5, 1 - 0.75 / H, "(IDR Bio)", ha="center", fontsize=10, color="#666")

    baris = kat[::-1] + ["Total"]
    isi = [[fmt_angka(v) if (r == "Total" or v != 0) else "" for v in pv[r]] for r in baris]
    tbl = ax_tab.table(cellText=isi, rowLabels=baris, colLabels=list(pv.index),
                       cellLoc="center", loc="center", bbox=[0, 0, 1, 1])
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8.5)
    for (r, c), cell in tbl.get_celld().items():
        cell.set_edgecolor("#cccccc")
        cell.set_linewidth(0.6)
        if c == -1 and r >= 1 and baris[r - 1] in warna:
            cell.set_facecolor(warna[baris[r - 1]])
            cell.get_text().set_color("white")
            cell.get_text().set_fontweight("bold")

    fig.text(0.11, (y_tab - 0.12) / H, ket, ha="left", va="top", fontsize=8,
             color="#B03A2E", style="italic")

    handel = [Patch(color=warna[a], label=a) for a in kat] + \
             [Line2D([0], [0], color="black", ls="--", marker="o", ms=4, label="Total")]
    fig.legend(handles=handel, loc="lower center", ncol=min(len(handel), 5), frameon=False,
               fontsize=9, bbox_to_anchor=(0.5, 0.2 / H))
    fig.text(0.02, 0.1 / H, data_as_of, fontsize=8, style="italic")

    Path(FILE_PNG).parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(FILE_PNG, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__JUDUL__</title>
<style>
  :root { --bg:#f4f6f8; --card:#fff; --teks:#2b2f33; --sub:#6b7280; --garis:#e5e7eb; --aksen:#1B6B8F; }
  @media (prefers-color-scheme: dark) {
    :root { --bg:#14181c; --card:#1e2429; --teks:#e8eaed; --sub:#9aa3ad; --garis:#2f373e; }
  }
  * { box-sizing: border-box; }
  body { margin:0; padding:24px 16px 40px; background:var(--bg); color:var(--teks);
         font-family: system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif; }
  .wrap { max-width: 1100px; margin: 0 auto; }
  h1 { font-size: 1.25rem; margin: 0 0 4px; }
  .sub { color: var(--sub); font-size: .85rem; margin-bottom: 16px; }
  .kartu { display:flex; gap:12px; flex-wrap:wrap; margin-bottom:16px; }
  .kpi { background:var(--card); border:1px solid var(--garis); border-radius:10px;
         padding:12px 16px; min-width:170px; }
  .kpi .l { font-size:.75rem; color:var(--sub); }
  .kpi .v { font-size:1.4rem; font-weight:700; }
  .grafik { background:#fff; border:1px solid var(--garis); border-radius:10px; padding:8px; }
  .grafik img { width:100%; height:auto; display:block; }
  .unduh { margin-top:16px; display:flex; gap:10px; flex-wrap:wrap; }
  .unduh a { background:var(--aksen); color:#fff; text-decoration:none; padding:8px 14px;
             border-radius:8px; font-size:.85rem; }
  .unduh a:hover { opacity:.9; }
</style>
</head>
<body>
<div class="wrap">
  <h1>__JUDUL__</h1>
  <div class="sub">Diperbarui otomatis: __WAKTU__ &middot; __ASOF__</div>
  <div class="kartu">
    <div class="kpi"><div class="l">Total (IDR Bio)</div><div class="v">__TOTAL__</div></div>
    <div class="kpi"><div class="l">Jumlah opportunity</div><div class="v">__OPP__</div></div>
    <div class="kpi"><div class="l">Minggu terbesar</div><div class="v">__TOP__</div></div>
  </div>
  <div class="grafik"><img src="close_lost_unit.png?v=__STAMP__" alt="__JUDUL__"></div>
  <div class="unduh">
    <a href="close_lost_unit.png?v=__STAMP__" download>Unduh grafik (PNG)</a>
    <a href="close_lost_unit.xlsx?v=__STAMP__" download>Unduh tabel (Excel)</a>
  </div>
</div>
</body>
</html>
"""


def tulis_html(pv, df, data_as_of):
    sekarang = datetime.now()
    tot = pv["Total"]
    minggu_top = tot.idxmax().split(" (")[0] + f" ({fmt_angka(tot.max())})" if tot.max() > 0 else "-"
    isi = (HTML_TEMPLATE
           .replace("__JUDUL__", JUDUL)
           .replace("__WAKTU__", sekarang.strftime("%d %b %Y %H:%M"))
           .replace("__ASOF__", data_as_of)
           .replace("__TOTAL__", fmt_angka(tot.sum()))
           .replace("__OPP__", str(int((~df["manual"]).sum())))
           .replace("__TOP__", minggu_top)
           .replace("__STAMP__", sekarang.strftime("%Y%m%d%H%M%S")))
    Path(FILE_HTML).parent.mkdir(parents=True, exist_ok=True)
    Path(FILE_HTML).write_text(isi, encoding="utf-8")


def main():
    df = baca()
    if df.empty:
        sys.exit(f"ERROR: tidak ada data {NAMA_BULAN} {TAHUN} yang bisa diolah.")
    print("\n=== Total per kategori (IDR Bio) - cocokkan dengan pivot ===")
    print(df.groupby("kategori")["nilai"].sum().round(3).sort_values(ascending=False).to_string())
    print(f"TOTAL: {df['nilai'].sum():,.3f}\n")

    keterangan = buat_keterangan(df)
    print(keterangan, "\n")

    pv, kat, warna = hitung(df)
    data_as_of = "Data as of " + datetime.now().strftime("%b-%d %Y")
    gambar(pv, kat, warna, keterangan, data_as_of)
    Path(FILE_XLSX).parent.mkdir(parents=True, exist_ok=True)
    pv.round(1).T.to_excel(FILE_XLSX)
    tulis_html(pv, df, data_as_of)
    print(pv.round(1).T.to_string())
    print(f"\nHalaman: {FILE_HTML}\nGrafik : {FILE_PNG}\nTabel  : {FILE_XLSX}")


if __name__ == "__main__":
    main()
