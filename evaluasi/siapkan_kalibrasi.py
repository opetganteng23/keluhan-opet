"""Lembar kalibrasi prompt (dev set): 100 tweet untuk dilabeli manual dengan definisi keluhan versi 2.

- Diambil dari 15.000 tweet berlabel LLM versi 1 (progres_llm_15000.jsonl), TIDAK termasuk 385 sampel uji.
- Berstrata menurut label LLM versi 1 (hanya untuk komposisi, tidak ditampilkan):
  30 Angkutan Umum, 30 Manajemen Lalu Lintas, 20 Lahan Parkir, 20 Tidak Relevan.
- Kolom isian: Kategori, Sentimen, Keluhan, Sebab, Akibat, Emosi (+ Catatan).

Hasil: Downloads/kalibrasi/lembar_kalibrasi_100.xlsx  (+ kunci_kalibrasi.json berisi id tweet)
Pemakaian:  python evaluasi/siapkan_kalibrasi.py
"""

import json
import logging
import random
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("kalibrasi")

FILE_LLM = Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "analisis_llm" / "progres_llm_15000.jsonl"
FILE_UJI = Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "evaluasi_akurasi" / "hasil_engine_sampel.xlsx"
FOLDER = Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "kalibrasi"
BENIH = 2027
KOMPOSISI = {"Angkutan Umum": 30, "Manajemen Lalu Lintas": 30, "Lahan Parkir": 20, "Tidak Relevan": 20}
KATEGORI = ["Angkutan Umum", "Lahan Parkir", "Manajemen Lalu Lintas", "Tidak Relevan"]
SENTIMEN = ["Positif", "Negatif", "Netral"]
EMOSI = ["Lelah", "Frustasi", "Kesal", "Kecewa", "Marah", "Takut/Khawatir", "Tidak Ada"]
TEBAL = Font(bold=True)
ISIAN = PatternFill("solid", fgColor="FFF4CC")

PEDOMAN = [
    ["Pedoman Pelabelan – Keluhan Transportasi Perkotaan Jakarta (versi 2)"],
    [],
    ["Ruang lingkup", "Keluhan sehari-hari warganet tentang kondisi transportasi perkotaan JAKARTA: "
                      "jalan & lalu lintas, angkutan umum, dan parkir. Bukan jalan tol/antarkota, bukan kota lain."],
    ["Cara mengisi", "Isi sel kuning. Kategori, Sentimen, Keluhan, Emosi: pilih dari dropdown. "
                     "Sebab & Akibat: ketik singkat (maks. ±8 kata), atau 'tidak disebutkan'."],
    ["Penting", "Nilai hanya dari isi teks. Jangan melihat hasil sistem. Ragu? Pilih yang paling sesuai + tulis di Catatan."],
    [],
    ["KATEGORI", ""],
    ["Manajemen Lalu Lintas", "Kemacetan, jalan rusak/berlubang, galian, rambu, lampu lalu lintas, marka, rekayasa "
                              "lalu lintas, ganjil genap, trotoar, lampu jalan/PJU, JPO, zebra cross, perilaku "
                              "pengendara di jalan kota, petugas Dishub di jalan."],
    ["Lahan Parkir", "Parkir liar, juru parkir/tukang parkir, tarif/pungutan parkir, ketersediaan lahan, diderek/diangkut."],
    ["Angkutan Umum", "TransJakarta, JakLingko/Mikrotrans, angkot, bus kota, ojek, halte di Jakarta: armada, "
                      "sopir/petugas, waktu tunggu, kepadatan, kenyamanan, keamanan, tarif, rute."],
    ["Tidak Relevan", "Review aplikasi, iklan, candaan, berita/pengumuman tanpa pendapat warga, cerita pribadi yang "
                      "hanya MENYEBUT transportasi tanpa menilai kondisinya, jalan tol/antarkota, kota lain, "
                      "kecelakaan yang hanya diberitakan, topik lain (banjir, sampah)."],
    [],
    ["KELUHAN", "Ya = penulis MENGELUHKAN kondisi pada salah satu dari 3 topik. Tidak = bertanya, memberi info, "
                "memuji, atau Tidak Relevan (Tidak Relevan selalu Keluhan = Tidak)."],
    ["SEBAB", "Penyebab masalah yang dikeluhkan, sesuai teks. Tidak ada di teks -> 'tidak disebutkan'."],
    ["AKIBAT", "Kondisi buruk/dampak yang dikeluhkan, sesuai teks. Tidak ada -> 'tidak disebutkan'."],
    ["Pola", "AKIBAT terjadi karena SEBAB."],
    ["Contoh 1", "'Jalanan macet nih gara-gara galian' -> Sebab: galian jalan | Akibat: jalanan macet"],
    ["Contoh 2", "'Lampu merah ini macet banget, lampunya kelamaan' -> Sebab: durasi lampu merah terlalu lama | "
                 "Akibat: macet parah di lampu merah"],
    ["Contoh 3", "'Parkir di sini banyak jukirnya, matok harga gak bener' -> Sebab: banyak jukir liar | "
                 "Akibat: tarif parkir tidak wajar"],
    ["Contoh 4", "'Udah sejam kejebak macet, telat ngantor lagi, capek' -> Sebab: tidak disebutkan | "
                 "Akibat: terjebak macet sejam, telat kerja | Emosi: Lelah"],
    [],
    ["SENTIMEN", "Sikap penulis terhadap kondisi yang dibahas. Sarkasme ('hebat, nunggu sejam') = Negatif."],
    ["EMOSI", "Perasaan yang paling tampak: Lelah, Frustasi, Kesal, Kecewa, Marah, Takut/Khawatir, Tidak Ada."],
]


def teks_uji() -> set[str]:
    """Teks 385 sampel uji (dikecualikan)."""
    sheet = load_workbook(FILE_UJI, read_only=True)["Label Engine"]
    return {b[1] for b in list(sheet.iter_rows(values_only=True))[1:]}


def ambil_sampel() -> list[dict]:
    """Sampel berstrata menurut label LLM versi 1, tanpa sampel uji."""
    uji = teks_uji()
    per_kat: dict[str, list[dict]] = {k: [] for k in KOMPOSISI}
    with FILE_LLM.open(encoding="utf-8") as f:
        for d in map(json.loads, f):
            if d["Metode"] == "LLM" and d["teks_bersih"].strip() and d["teks_bersih"] not in uji:
                per_kat[d["Kategori"]].append(d)
    random.seed(BENIH)
    sampel = [d for k, n in KOMPOSISI.items() for d in random.sample(per_kat[k], n)]
    random.shuffle(sampel)  # campur supaya urutan tidak memberi petunjuk kategori
    return sampel


def buat_lembar(sampel: list[dict]) -> Path:
    """Lembar pelabelan dengan dropdown + sheet pedoman."""
    buku = Workbook()
    sheet = buku.active
    sheet.title = "Kalibrasi"
    sheet.append(["No", "ID", "Teks", "Kategori", "Sentimen", "Keluhan", "Sebab", "Akibat", "Emosi", "Catatan"])
    for sel in sheet[1]:
        sel.font = TEBAL
    for i, d in enumerate(sampel, 1):
        sheet.append([i, f"K{i:03d}", d["teks_bersih"], "", "", "", "", "", "", ""])
        sheet.cell(row=i + 1, column=3).alignment = Alignment(wrap_text=True, vertical="top")
        for kolom in range(4, 10):
            sheet.cell(row=i + 1, column=kolom).fill = ISIAN
    akhir = len(sampel) + 1
    for kolom, pilihan in (("D", KATEGORI), ("E", SENTIMEN), ("F", ["Ya", "Tidak"]), ("I", EMOSI)):
        dv = DataValidation(type="list", formula1='"' + ",".join(pilihan) + '"', allow_blank=True)
        dv.error, dv.errorTitle = "Pilih dari daftar", "Isian tidak valid"
        sheet.add_data_validation(dv)
        dv.add(f"{kolom}2:{kolom}{akhir}")
    for h, w in zip("ABCDEFGHIJ", [5, 7, 70, 22, 11, 9, 28, 28, 15, 22]):
        sheet.column_dimensions[h].width = w
    sheet.freeze_panes = "D2"

    pedoman = buku.create_sheet("Pedoman")
    for b in PEDOMAN:
        pedoman.append(b)
    for b in pedoman.iter_rows():
        if b[0].value and (b[0].value.isupper() or b[0].value.startswith("Pedoman")):
            b[0].font = TEBAL
        for sel in b:
            sel.alignment = Alignment(wrap_text=True, vertical="top")
    pedoman.column_dimensions["A"].width = 24
    pedoman.column_dimensions["B"].width = 110

    FOLDER.mkdir(parents=True, exist_ok=True)
    tujuan = FOLDER / "lembar_kalibrasi_100.xlsx"
    buku.save(tujuan)
    kunci = [{"ID": f"K{i:03d}", "id_tweet": d["id"], "teks_bersih": d["teks_bersih"]} for i, d in enumerate(sampel, 1)]
    (FOLDER / "kunci_kalibrasi.json").write_text(json.dumps(kunci, ensure_ascii=False, indent=1), encoding="utf-8")
    return tujuan


if __name__ == "__main__":
    logger.info("Lembar kalibrasi: %s", buat_lembar(ambil_sampel()))
