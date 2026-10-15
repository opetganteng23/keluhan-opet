"""Metode 4: TF-IDF + Support Vector Machine (linear).

Pemakaian (dari folder keluhanku):  python metode_4_tfidf_svm/svm.py
- Melatih model kategori & sentimen dari label LLM (tanpa 385 sampel uji)
- Validasi silang 5-fold, menyimpan model ke model/, bukti ke Downloads/evaluasi_akurasi/bukti_SVM.xlsx
- Menulis prediksi 385 sampel uji ke hasil_engine_sampel.xlsx (kolom 'Kategori SVM', 'Sentimen SVM')
"""

import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # agar bisa dijalankan langsung

from sklearn.svm import LinearSVC

from persiapan_data.tfidf_umum import BENIH, jalankan

logging.basicConfig(level=logging.INFO, format="%(message)s")


def buat_klasifier() -> LinearSVC:
    """SVM linear, C = 1 (baku); class_weight 'balanced' agar kelas kecil (Lahan Parkir) tidak tenggelam."""
    return LinearSVC(C=1.0, class_weight="balanced", random_state=BENIH)


if __name__ == "__main__":
    jalankan("TF-IDF + SVM", "SVM", buat_klasifier, Path(__file__).resolve().parent)
