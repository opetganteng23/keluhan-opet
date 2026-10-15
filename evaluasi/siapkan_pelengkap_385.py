"""Lembar pelengkap 385 sampel uji untuk definisi versi 2/3 (keluhan, sebab, akibat, emosi).

Kategori & Sentimen diisi dari lembar_label_pelabel1.xlsx (label manual sebelumnya) dan BOLEH
direvisi jika menurut Pedoman versi 2 labelnya berbeda (tulis alasannya di Catatan).

Hasil: Downloads/evaluasi_akurasi/lembar_label_v2_pelabel1.xlsx
Pemakaian:  python evaluasi/siapkan_pelengkap_385.py
"""

import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from siapkan_kalibrasi import EMOSI, KATEGORI, PEDOMAN, SENTIMEN

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("pelengkap")

FOLDER = Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "evaluasi_akurasi"
TEBAL = Font(bold=True)
TERISI = PatternFill("solid", fgColor="E8F1FB")   # biru muda: sudah terisi, boleh direvisi
ISIAN = PatternFill("solid", fgColor="FFF4CC")    # kuning: perlu diisi


def main() -> None:
    """Salin label lama, tambahkan kolom keluhan/sebab/akibat/emosi."""
    lama = list(load_workbook(FOLDER / "lembar_label_pelabel1.xlsx", read_only=True)["Pelabelan"]
                .iter_rows(values_only=True))[1:]
    buku = Workbook()
    sheet = buku.active
    sheet.title = "Pelabelan"
    sheet.append(["No", "ID Sampel", "Teks", "Kategori", "Sentimen", "Keluhan", "Sebab", "Akibat", "Emosi", "Catatan"])
    for sel in sheet[1]:
        sel.font = TEBAL
    for i, b in enumerate(lama, 1):
        sheet.append([b[0], b[1], b[2], b[3], b[4], "", "", "", "", b[5] or ""])
        sheet.cell(row=i + 1, column=3).alignment = Alignment(wrap_text=True, vertical="top")
        for kolom in (4, 5):
            sheet.cell(row=i + 1, column=kolom).fill = TERISI
        for kolom in range(6, 10):
            sheet.cell(row=i + 1, column=kolom).fill = ISIAN
    akhir = len(lama) + 1
    for kolom, pilihan in (("D", KATEGORI), ("E", SENTIMEN), ("F", ["Ya", "Tidak"]), ("I", EMOSI)):
        dv = DataValidation(type="list", formula1='"' + ",".join(pilihan) + '"', allow_blank=True)
        sheet.add_data_validation(dv)
        dv.add(f"{kolom}2:{kolom}{akhir}")
    for h, w in zip("ABCDEFGHIJ", [5, 10, 70, 22, 11, 9, 28, 28, 15, 22]):
        sheet.column_dimensions[h].width = w
    sheet.freeze_panes = "D2"

    pedoman = buku.create_sheet("Pedoman")
    for b in PEDOMAN + [[], ["KOLOM BIRU", "Kategori & Sentimen sudah terisi dari label sebelumnya. Revisi hanya jika "
                                         "menurut pedoman ini berbeda (kategori = TOPIK tweet; pengumuman/berita "
                                         "tetap masuk topiknya). Tulis alasan di Catatan."],
                        ["KOLOM KUNING", "Isi Keluhan, Sebab, Akibat, Emosi."]]:
        pedoman.append(b)
    for b in pedoman.iter_rows():
        if b[0].value and (str(b[0].value).isupper() or str(b[0].value).startswith("Pedoman")):
            b[0].font = TEBAL
        for sel in b:
            sel.alignment = Alignment(wrap_text=True, vertical="top")
    pedoman.column_dimensions["A"].width = 24
    pedoman.column_dimensions["B"].width = 110

    tujuan = FOLDER / "lembar_label_v2_pelabel1.xlsx"
    buku.save(tujuan)
    logger.info("Lembar pelengkap: %s (%d baris)", tujuan, len(lama))


if __name__ == "__main__":
    main()
