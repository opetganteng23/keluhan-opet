"""Merapikan istilah hasil ekspansi yang dipilih peneliti menjadi seed word baru + tabel untuk paper.

Masukan : Downloads/kandidat_ekspansi_seed_word.xlsx (kolom "Pilih" = Ya)
Keluaran: Downloads/tabel_seed_word_ekspansi.xlsx
  - Seed Awal       : 32 seed word awal
  - Seed Ekspansi   : seed word baru final (setelah dirapikan), kategori, status
  - Jejak Perapian  : setiap istilah terpilih -> hasilnya + aturan yang diterapkan

Aturan perapian (urut):
  1. Perbaiki istilah terpotong akibat penghapusan stopword (mis. "arus lalu" -> "arus lalu lintas").
  2. Gabungkan varian/duplikat ke satu bentuk (mis. "ugal", "ugalan" -> "ugal ugalan").
  3. Tambahkan "jakarta" bila istilah tidak memuat penanda Jakarta (agar pencarian tetap di wilayah studi).
  4. Buang yang sama persis (himpunan kata sama) dengan seed word awal atau seed ekspansi lain.
  5. Tandai istilah yang ruang lingkupnya perlu dikonfirmasi dosen, atau dibuang atas keputusan peneliti.

Pemakaian:  python persiapan_data/rapikan_seed_ekspansi.py
"""

import logging
import re
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("seed")

MASUKAN = Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "seed_word" / "kandidat_ekspansi_seed_word.xlsx"
KELUARAN = Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "seed_word" / "tabel_seed_word_ekspansi.xlsx"
FILE_CRAWLER = Path(__file__).resolve().parent.parent.parent / "crawling" / "crawl_per_bulan.py"
TEBAL = Font(bold=True)

PERBAIKI = {"arus lalu": "arus lalu lintas", "lampu lalu": "lampu lalu lintas"}
GABUNG = {
    "ugal": "ugal ugalan", "ugalan": "ugal ugalan",
    "penumpangnya": "penumpang", "penumpang lain": "penumpang", "seorang penumpang": "penumpang",
    "pejalan": "pejalan kaki", "pengendara motor": "pengendara",
    "kejebak macet": "kejebak", "macet macet": "macet parah",
    "motornya diangkut": "diangkut", "perhubungan": "dinas perhubungan",
    "transjakarta koridor": "koridor",
}
PENANDA_JAKARTA = {"jakarta", "dki", "transjakarta", "jaklingko", "mikrotrans", "metrotrans", "busway", "tj"}
KONFIRMASI: dict[str, str] = {}  # semua sudah diputuskan peneliti
# Keputusan peneliti (13-10-2026): ruang lingkup = keluhan sehari-hari warganet tentang perkotaan Jakarta
DIBUANG = {"pelecehan jakarta": "Dibuang peneliti: ranah keamanan/kepolisian, bukan kondisi transportasi kota",
           "tol jakarta": "Dibuang peneliti: jalan tol/antarkota, bukan jalan perkotaan",
           "kecelakaan jakarta": "Dibuang peneliti: cenderung berita/kepolisian, bukan keluhan sehari-hari"}
SEED_AWAL_KATEGORI = ["Angkutan Umum"] * 10 + ["Lahan Parkir"] * 6 + ["Manajemen Lalu Lintas"] * 12 + ["Dishub (umum)"] * 4


def seed_awal() -> list[str]:
    """32 seed word awal dari crawl_per_bulan.py."""
    blok = re.search(r"KATA_KUNCI = \[(.*?)\]", FILE_CRAWLER.read_text(encoding="utf-8"), re.S).group(1)
    return re.findall(r'"([^"]+)"', blok)


def kata(istilah: str) -> frozenset[str]:
    """Himpunan kata (pencarian X tidak peduli urutan kata)."""
    return frozenset(re.findall(r"[a-z@:]+", istilah.lower()))


def rapikan(istilah: str) -> tuple[str, list[str]]:
    """Terapkan aturan 1-3; kembalikan istilah final + daftar aturan yang dipakai."""
    aturan, hasil = [], istilah.strip().lower()
    if hasil in PERBAIKI:
        hasil = PERBAIKI[hasil]
        aturan.append("1. diperbaiki (terpotong)")
    if hasil in GABUNG:
        hasil = GABUNG[hasil]
        aturan.append("2. digabung (varian)")
    if not kata(hasil) & PENANDA_JAKARTA:
        hasil = f"{hasil} jakarta"
        aturan.append("3. + 'jakarta'")
    return hasil, aturan


def main() -> None:
    """Baca pilihan peneliti, rapikan, tulis tabel."""
    awal = seed_awal()
    set_awal = {kata(s): s for s in awal}
    buku_masuk = load_workbook(MASUKAN, read_only=True)
    jejak, final = [], {}
    for sheet in buku_masuk.worksheets[1:]:
        for b in list(sheet.iter_rows(values_only=True))[1:]:
            if str(b[5] or "").strip().lower() != "ya":
                continue
            hasil, aturan = rapikan(b[1])
            k = kata(hasil)
            if k in set_awal:
                jejak.append([sheet.title, b[1], "-", "4. dibuang: sama dengan seed awal '%s'" % set_awal[k]])
                continue
            if k in {kata(x) for x in final}:
                jejak.append([sheet.title, b[1], hasil, "; ".join(aturan + ["4. digabung: sudah ada"])])
                continue
            final[hasil] = sheet.title
            jejak.append([sheet.title, b[1], hasil, "; ".join(aturan) or "tetap"])

    buku = Workbook()
    s1 = buku.active
    s1.title = "Seed Awal"
    s1.append(["No", "Seed Word", "Kategori"])
    for i, (s, k) in enumerate(zip(awal, SEED_AWAL_KATEGORI), 1):
        s1.append([i, s, k])

    s2 = buku.create_sheet("Seed Ekspansi")
    s2.append(["No", "Seed Word Ekspansi", "Kategori", "Status", "Catatan"])
    urut = sorted(final.items(), key=lambda x: (x[1], x[0]))
    for i, (s, k) in enumerate(urut, 1):
        status = "Dibuang" if s in DIBUANG else "Konfirmasi dosen" if s in KONFIRMASI else "Siap dipakai"
        s2.append([i, s, k, status, DIBUANG.get(s) or KONFIRMASI.get(s, "")])

    s3 = buku.create_sheet("Jejak Perapian")
    s3.append(["Kategori", "Istilah Dipilih", "Seed Word Final", "Aturan"])
    for b in jejak:
        s3.append(b)

    s4 = buku.create_sheet("Aturan")
    for t in (__doc__ or "").strip().splitlines():
        s4.append([t])
    for sheet, lebar in ((s1, [5, 30, 24]), (s2, [5, 32, 24, 18, 40]), (s3, [24, 26, 32, 60]), (s4, [110])):
        for sel in sheet[1]:
            sel.font = TEBAL
        for j, w in enumerate(lebar):
            sheet.column_dimensions[chr(65 + j)].width = w
    buku.save(KELUARAN)

    per_kat = {}
    for _, k in final.items():
        per_kat[k] = per_kat.get(k, 0) + 1
    logger.info("Dipilih %d istilah -> %d seed word ekspansi final %s", len(jejak), len(final), per_kat)
    for k in sorted(set(final.values())):
        logger.info("%s: %s", k, ", ".join(s for s, kk in urut if kk == k))
    logger.info("File: %s", KELUARAN)


if __name__ == "__main__":
    main()
