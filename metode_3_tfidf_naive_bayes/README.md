# Metode 3 – TF-IDF + Naive Bayes

Teks diubah menjadi vektor TF-IDF (unigram + bigram), lalu diklasifikasi dengan Multinomial Naive Bayes (alpha = 1).
Dua model: kategori dan sentimen.

- `naive_bayes.py` – jalankan: `python metode_3_tfidf_naive_bayes/naive_bayes.py` (dari folder induk)
- `model/` – model hasil pelatihan (`kategori.joblib`, `sentimen.joblib`)
- Data latih: label LLM dari 15.000 tweet, **tanpa** 385 sampel uji.
- Bukti cara kerja: `Downloads\evaluasi_akurasi\bukti_NB.xlsx` (validasi silang + istilah berbobot tertinggi per kelas).
- Bagian bersama TF-IDF: `persiapan_data/tfidf_umum.py`.
