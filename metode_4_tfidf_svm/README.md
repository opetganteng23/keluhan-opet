# Metode 4 – TF-IDF + SVM

Teks diubah menjadi vektor TF-IDF (unigram + bigram), lalu diklasifikasi dengan SVM linear (C = 1, class_weight = balanced).
Dua model: kategori dan sentimen.

- `svm.py` – jalankan: `python metode_4_tfidf_svm/svm.py` (dari folder induk)
- `model/` – model hasil pelatihan (`kategori.joblib`, `sentimen.joblib`)
- Data latih: label LLM dari 15.000 tweet, **tanpa** 385 sampel uji.
- Bukti cara kerja: `Downloads\evaluasi_akurasi\bukti_SVM.xlsx` (validasi silang + istilah berbobot tertinggi per kelas).
- Bagian bersama TF-IDF: `persiapan_data/tfidf_umum.py`.
