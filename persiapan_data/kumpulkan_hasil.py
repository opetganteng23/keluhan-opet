"""Kumpulkan semua hasil penelitian ke satu folder rapi: Downloads/Hasil_Penelitian/.

File disalin dari _file_kerja (dan kode) sehingga skrip tetap berjalan. Folder bernomor 01-09 dibuat ulang
setiap kali dijalankan (isinya hanya salinan); _file_kerja tidak disentuh.
Pemakaian:  python persiapan_data/kumpulkan_hasil.py
"""

import logging
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("kumpulkan")

TUJUAN = Path.home() / "Downloads" / "Hasil_Penelitian"
KERJA = TUJUAN / "_file_kerja"          # file kerja yang dibaca/ditulis skrip
EV = KERJA / "evaluasi_akurasi"
AN = KERJA / "analisis_llm"
KA = KERJA / "kalibrasi"
DS = KERJA / "data_siap_analisis"
SW = KERJA / "seed_word"
KODE = Path(__file__).resolve().parent.parent

STRUKTUR = {
    "01_Dataset": {
        "dataset_final_10000_tweet.xlsx": AN / "hasil_v3_10000_final.xlsx",
        "data_bersih_19016_tweet.xlsx": DS / "data_siap_analisis_20261013_2127.xlsx",
        "tabel_seed_word_awal.xlsx": SW / "tabel_seed_word.xlsx",
        "tabel_seed_word_ekspansi.xlsx": SW / "tabel_seed_word_ekspansi.xlsx",
        "kandidat_ekspansi_seed_word_terpilih.xlsx": SW / "kandidat_ekspansi_seed_word.xlsx",
    },
    "02_Metode_1_LLM": {
        "hasil_LLM_10000_tweet_sebelum_koreksi.xlsx": AN / "hasil_v3_10000.xlsx",
        "hasil_LLM_385_sampel_uji.xlsx": AN / "hasil_uji_385_v3.xlsx",
        "hasil_LLM_615_sampel_uji.xlsx": AN / "hasil_uji_615_v3.xlsx",
        "kode_llm_analyzer.py": KODE / "metode_1_llm" / "llm_analyzer.py",
        "kode_analisis_batch.py": KODE / "metode_1_llm" / "analisis_batch.py",
        "kode_pasca_proses.py": KODE / "metode_1_llm" / "pasca_proses.py",
    },
    "03_Metode_2_Kata_Kunci": {
        "kode_classifier.py": KODE / "metode_2_kata_kunci" / "classifier.py",
        "kode_sentiment.py": KODE / "metode_2_kata_kunci" / "sentiment.py",
    },
    "04_Metode_3_TFIDF_Naive_Bayes": {
        "bukti_cara_kerja_NB.xlsx": EV / "bukti_NB_v3.xlsx",
        "kode_naive_bayes.py": KODE / "metode_3_tfidf_naive_bayes" / "naive_bayes.py",
        "kode_tfidf_umum.py": KODE / "persiapan_data" / "tfidf_umum.py",
    },
    "05_Metode_4_TFIDF_SVM": {
        "bukti_cara_kerja_SVM.xlsx": EV / "bukti_SVM_v3.xlsx",
        "kode_svm.py": KODE / "metode_4_tfidf_svm" / "svm.py",
        "kode_tfidf_umum.py": KODE / "persiapan_data" / "tfidf_umum.py",
    },
    "06_Metode_5_IndoBERT": {
        "riwayat_pelatihan_per_epoch.xlsx": EV / "indobert" / "riwayat_pelatihan_indobert.xlsx",
        "grafik_loss_akurasi_per_epoch.png": EV / "indobert" / "8_epoch_indobert.png",
        "kode_indobert.py": KODE / "metode_5_indobert" / "indobert.py",
    },
    "07_Evaluasi_Akurasi": {
        "laporan_evaluasi_5_metode.xlsx": EV / "laporan_evaluasi_v3.xlsx",
        "label_semua_metode_1000_sampel.xlsx": EV / "hasil_engine_sampel_v3.xlsx",
        "label_manual_1000_sampel.xlsx": EV / "lembar_label_v3_1000.xlsx",
    },
    "08_Kalibrasi_Prompt": {
        "label_manual_100_kalibrasi.xlsx": KA / "lembar_kalibrasi_100.xlsx",
        "hasil_kalibrasi_1_prompt_v2.xlsx": KA / "hasil_kalibrasi_v2.xlsx",
        "hasil_kalibrasi_2_prompt_v3.xlsx": KA / "hasil_kalibrasi_v3.xlsx",
        "hasil_kalibrasi_3_prompt_v3_1.xlsx": KA / "hasil_kalibrasi_v3_1.xlsx",
        "hasil_kalibrasi_4_prompt_v3_1_final.xlsx": KA / "hasil_kalibrasi_v3_1_t0.xlsx",
        "hasil_kalibrasi_5_uji_ulang_konsistensi.xlsx": KA / "hasil_kalibrasi_v3_1_t0_ulang.xlsx",
        "hasil_kalibrasi_6_uji_batch.xlsx": AN / "hasil_uji_kalibrasi.xlsx",
    },
    "09_Arsip": {
        "evaluasi_v1_laporan.xlsx": EV / "laporan_evaluasi.xlsx",
        "evaluasi_v1_label_manual_385.xlsx": EV / "lembar_label_pelabel1.xlsx",
        "evaluasi_v1_hasil_LLM_15000_tweet.xlsx": AN / "hasil_llm_15000.xlsx",
        "evaluasi_v2_laporan_4_metode_385_sampel.xlsx": EV / "laporan_evaluasi_v2.xlsx",
        "label_615_pelabelan_cepat_TIDAK_DIPAKAI.xlsx": EV / "lembar_label_v2_tambahan615_ARSIP_pelabelan_cepat.xlsx",
    },
}
FOLDER_GRAFIK = {"07_Evaluasi_Akurasi/grafik": EV / "grafik_v3", "09_Arsip/grafik_evaluasi_v1": EV / "grafik",
                 "09_Arsip/grafik_evaluasi_v2_385": EV / "grafik_v2"}

KETERANGAN = {
    "01_Dataset": "Dataset penelitian. FILE UTAMA: dataset_final_10000_tweet.xlsx (kategori, sentimen, keluhan, "
                  "sebab, akibat, emosi; kolom Koreksi = aturan pasca-proses R1/R2).",
    "02_Metode_1_LLM": "Metode utama: LLM claude-haiku-4-5, prompt v3.1, temperature 0, Batch API.",
    "03_Metode_2_Kata_Kunci": "Baseline leksikon (kamus kata). Bukti kata cocok per sampel: folder 07, "
                              "label_semua_metode_1000_sampel.xlsx sheet 'Bukti Kata Kunci'.",
    "04_Metode_3_TFIDF_Naive_Bayes": "TF-IDF + Multinomial Naive Bayes, dilatih dari label LLM (tanpa 1.000 sampel uji).",
    "05_Metode_4_TFIDF_SVM": "TF-IDF + SVM linear (class_weight balanced), dilatih dari label LLM (tanpa 1.000 sampel uji).",
    "06_Metode_5_IndoBERT": "Fine-tuning IndoBERT (indobert-base-p1), 4 epoch, model epoch terbaik menurut macro-F1 "
                            "validasi; data latih sama dengan NB/SVM.",
    "07_Evaluasi_Akurasi": "Evaluasi FINAL 5 metode vs label manual (1.000 sampel). Grafik PNG 300 dpi di subfolder "
                           "grafik; laporan_evaluasi.html = halaman laporan lengkap dengan rumus.",
    "08_Kalibrasi_Prompt": "Penyetelan prompt LLM pada 100 tweet kalibrasi (terpisah dari sampel uji), urut 1-6.",
    "09_Arsip": "Riwayat: evaluasi v1 (definisi lama), evaluasi v2 (385 sampel, 4 metode), dan label 615 versi cepat "
                "yang diganti pelabelan ulang. Tidak dipakai untuk hasil final.",
}
SINGKAT_ENGINE = {"02_Metode_1_LLM": "LLM", "03_Metode_2_Kata_Kunci": "Kata Kunci",
                  "04_Metode_3_TFIDF_Naive_Bayes": "NB", "05_Metode_4_TFIDF_SVM": "SVM",
                  "06_Metode_5_IndoBERT": "IndoBERT"}


def prediksi_per_metode(folder: str, akhiran: str) -> None:
    """Prediksi satu metode pada 1.000 sampel uji, berdampingan dengan label manual."""
    engine = list(load_workbook(EV / "hasil_engine_sampel_v3.xlsx", read_only=True)["Label Engine"]
                  .iter_rows(values_only=True))
    gold = {b[1]: b for b in list(load_workbook(EV / "lembar_label_v3_1000.xlsx", read_only=True)["Pelabelan"]
                                   .iter_rows(values_only=True))[1:]}
    pos = {j: i for i, j in enumerate(engine[0])}
    buku = Workbook()
    sheet = buku.active
    sheet.title = "Prediksi 1000 sampel"
    sheet.append(["ID Sampel", "Teks", "Kategori Manual", f"Kategori {akhiran}", "Sentimen Manual",
                  f"Sentimen {akhiran}", "Keluhan Manual", f"Keluhan {akhiran}"])
    for b in engine[1:]:
        g = gold[b[0]]
        sheet.append([b[0], b[1], g[3], b[pos[f"Kategori {akhiran}"]], g[4], b[pos[f"Sentimen {akhiran}"]],
                      g[5], b[pos[f"Keluhan {akhiran}"]]])
    for sel in sheet[1]:
        sel.font = Font(bold=True)
    for h, w in zip("ABCDEFGH", [11, 60, 20, 20, 12, 12, 10, 10]):
        sheet.column_dimensions[h].width = w
    buku.save(TUJUAN / folder / f"prediksi_{akhiran.replace(' ', '_')}_1000_sampel.xlsx")


def main() -> None:
    """Buat ulang folder 01-09 dari _file_kerja + kode, lengkap dengan BACA_SAYA per folder."""
    for lama in TUJUAN.iterdir() if TUJUAN.exists() else []:
        if lama.is_dir() and lama.name[:2].isdigit() and lama.name != "_file_kerja":
            shutil.rmtree(lama)
    hilang = []
    for folder, isi in STRUKTUR.items():
        (TUJUAN / folder).mkdir(parents=True, exist_ok=True)
        for nama, sumber in isi.items():
            if sumber.exists():
                shutil.copy2(sumber, TUJUAN / folder / nama)
            else:
                hilang.append(str(sumber))
        (TUJUAN / folder / "BACA_SAYA.txt").write_text(KETERANGAN[folder] + "\n", encoding="utf-8")
        if folder in SINGKAT_ENGINE:
            prediksi_per_metode(folder, SINGKAT_ENGINE[folder])
    for folder, sumber in FOLDER_GRAFIK.items():
        if sumber.exists():
            shutil.copytree(sumber, TUJUAN / folder, dirs_exist_ok=True)
    daftar = "\n".join(f"{f}\n    {KETERANGAN[f]}" for f in STRUKTUR)
    (KERJA / "BACA_SAYA.txt").write_text(
        "File kerja yang dibaca/ditulis skrip. Jangan diganti nama/dipindah. Untuk dilihat: pakai folder 01-09.\n",
        encoding="utf-8")
    (TUJUAN / "00_BACA_SAYA.txt").write_text(
        f"HASIL PENELITIAN - Keluhan Transportasi Perkotaan Jakarta\nDiperbarui: {datetime.now():%d-%m-%Y %H:%M}\n\n"
        f"{daftar}\n\nKode lengkap: {KODE}\n", encoding="utf-8")
    if hilang:
        logger.warning("File sumber tidak ditemukan (dilewati): %s", hilang)
    logger.info("Selesai: %s", TUJUAN)


if __name__ == "__main__":
    main()
