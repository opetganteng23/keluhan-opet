"""Metode 3: TF-IDF + Multinomial Naive Bayes.

Pemakaian (dari folder keluhanku):  python metode_3_tfidf_naive_bayes/naive_bayes.py
- Melatih model kategori & sentimen dari label LLM (tanpa 385 sampel uji)
- Validasi silang 5-fold, menyimpan model ke model/, bukti ke Downloads/evaluasi_akurasi/bukti_NB.xlsx
- Menulis prediksi 385 sampel uji ke hasil_engine_sampel.xlsx (kolom 'Kategori NB', 'Sentimen NB')
"""

import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # agar bisa dijalankan langsung

from sklearn.naive_bayes import MultinomialNB

from persiapan_data.tfidf_umum import jalankan

logging.basicConfig(level=logging.INFO, format="%(message)s")


def buat_klasifier() -> MultinomialNB:
    """Multinomial Naive Bayes dengan smoothing Laplace (alpha = 1, nilai baku)."""
    return MultinomialNB(alpha=1.0)


if __name__ == "__main__":
    jalankan("TF-IDF + Naive Bayes", "NB", buat_klasifier, Path(__file__).resolve().parent)
