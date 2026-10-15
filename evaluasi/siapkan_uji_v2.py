"""Evaluasi final (definisi v2/v3): jalankan LLM v3.1 (Batch API + pasca-proses) dan Kata Kunci
pada 385 sampel uji, tulis Downloads/evaluasi_akurasi/hasil_engine_sampel_v2.xlsx.

Naive Bayes & SVM menambahkan kolomnya sendiri lewat:
    EVAL_VERSI=v2 python metode_3_tfidf_naive_bayes/naive_bayes.py
    EVAL_VERSI=v2 python metode_4_tfidf_svm/svm.py

Pemakaian:  python evaluasi/siapkan_uji_v2.py
"""

import json
import logging
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from metode_1_llm import analisis_batch as ab
from metode_1_llm.pasca_proses import koreksi
from siapkan_sampel import bukti_kata_kunci

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("uji_v2")

FOLDER = Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "evaluasi_akurasi"
NAMA_BATCH = "uji_385_v3"
TEBAL = Font(bold=True)


def sampel_uji() -> list[dict]:
    """385 sampel uji (ID Sampel + teks bersih) dari lembar pelabelan v2."""
    baris = list(load_workbook(FOLDER / "lembar_label_v2_pelabel1.xlsx", read_only=True)["Pelabelan"]
                 .iter_rows(values_only=True))[1:]
    return [{"id": b[1], "ID Sampel": b[1], "teks_bersih": b[2] or ""} for b in baris]


def label_llm(sampel: list[dict]) -> dict[str, dict]:
    """Kirim batch (jika belum), tunggu, ambil, terapkan pasca-proses R1/R2."""
    if not ab.path_status(NAMA_BATCH).exists():
        ab.kirim(NAMA_BATCH, sampel, "385 sampel uji (evaluasi final, prompt v3.1)")
    while ab.cek(NAMA_BATCH) != "ended":
        time.sleep(30)
    hasil = ab.ambil(NAMA_BATCH)
    for b in hasil:
        b["Koreksi"] = koreksi(b)
    return {b["ID Sampel"]: b for b in hasil}


def main() -> None:
    """Tulis label LLM + Kata Kunci (+ bukti kata kunci) untuk 385 sampel uji."""
    sampel = sampel_uji()
    llm = label_llm(sampel)

    buku = Workbook()
    sheet = buku.active
    sheet.title = "Label Engine"
    sheet.append(["ID Sampel", "Teks",
                  "Kategori LLM", "Sentimen LLM", "Keluhan LLM", "Sebab LLM", "Akibat LLM", "Emosi LLM",
                  "Kategori Kata Kunci", "Sentimen Kata Kunci", "Keluhan Kata Kunci"])
    bukti = buku.create_sheet("Bukti Kata Kunci")
    bukti.append(["ID Sampel", "Teks", "Cocok: Angkutan Umum", "Cocok: Lahan Parkir", "Cocok: Manajemen Lalu Lintas",
                  "Alasan Kategori", "Kata Positif", "Kata Negatif"])
    for s in sampel:
        l = llm[s["ID Sampel"]]
        kk = bukti_kata_kunci(s["teks_bersih"])
        # Baseline kata kunci tidak mengenal "keluhan": didefinisikan keluhan = sentimen Negatif
        sheet.append([s["ID Sampel"], s["teks_bersih"], l["Kategori"], l["Sentimen"], l["Keluhan"], l["Sebab"],
                      l["Akibat"], l["Emosi"], kk["kategori"], kk["sentimen"],
                      "Ya" if kk["sentimen"] == "Negatif" else "Tidak"])
        bukti.append([s["ID Sampel"], s["teks_bersih"], kk["cocok"]["Angkutan Umum"], kk["cocok"]["Lahan Parkir"],
                      kk["cocok"]["Manajemen Lalu Lintas"], kk["alasan_kategori"], kk["kata_positif"],
                      kk["kata_negatif"]])
    for sh, lebar in ((sheet, [11, 60, 22, 11, 9, 28, 28, 14, 22, 11, 9]), (bukti, [11, 60, 22, 22, 22, 30, 18, 18])):
        for sel in sh[1]:
            sel.font = TEBAL
        for j, w in enumerate(lebar):
            sh.column_dimensions[chr(65 + j)].width = w
        sh.freeze_panes = "C2"
    konf = buku.create_sheet("Konfigurasi")
    for b in [["LLM", "claude-haiku-4-5, prompt v3.1, temperature 0, Batch API, pasca-proses R1/R2"],
              ["Kata Kunci", "classifier.py + sentiment.py; Keluhan = Ya jika sentimen Negatif (baseline)"],
              ["Batch", json.loads(ab.path_status(NAMA_BATCH).read_text(encoding="utf-8"))["batch_id"]]]:
        konf.append(b)
    konf.column_dimensions["B"].width = 90
    tujuan = FOLDER / "hasil_engine_sampel_v2.xlsx"
    buku.save(tujuan)
    logger.info("Label engine (LLM + Kata Kunci): %s", tujuan)


if __name__ == "__main__":
    main()
