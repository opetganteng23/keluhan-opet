# Metode 5 – IndoBERT (fine-tuning)

Model `indobenchmark/indobert-base-p1` dilatih ulang (4 epoch) untuk kategori, sentimen, dan keluhan.

- `indobert.py latih` – latih 3 model, catat loss & akurasi per epoch (grafik `8_epoch_indobert.png`).
- `indobert.py prediksi` – prediksi sampel uji.
- `model/` – model hasil pelatihan (besar, tidak di-commit).
- Data latih: label LLM dataset 10.000 tweet, **tanpa** 1.000 sampel uji.
