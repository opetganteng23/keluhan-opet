# Klasifikasi Keluhan Transportasi

Aplikasi web analisis keluhan transportasi publik. Input: CSV mentah. Output: Excel dengan kolom kategori + sentimen.

## Stack

- **Backend:** Python + FastAPI
- **Analisis utama:** Claude API (`claude-haiku-4-5`) via SDK `anthropic` — `llm_analyzer.py`
- **Analisis cadangan/pembanding:** kata kunci — `classifier.py` + `sentiment.py`
- **Config:** `.env` (`LLM_API_KEY`), dibaca `python-dotenv`
- **Export:** openpyxl
- **Frontend:** HTML/CSS/JS vanilla (satu file, tidak ada framework)
- **Package manager:** pip + venv

## Struktur Project

```
keluhanku/
├── CLAUDE.md, README.md, start.bat, stop.bat, requirements*.txt, .env
├── metode_1_llm/              # llm_analyzer.py — metode utama
├── metode_2_kata_kunci/       # classifier.py, sentiment.py — baseline leksikon
├── metode_3_tfidf_naive_bayes/  # (belum dibuat)
├── metode_4_tfidf_svm/          # (belum dibuat)
├── persiapan_data/            # pembersih.py, seleksi.py, ekspor_siap_analisis.py
├── aplikasi_web/              # main.py, exporter.py, static/, uploads/, klasifikasi_universal.html
├── evaluasi/                  # siapkan_sampel.py, hitung_evaluasi.py
├── dokumentasi/               # TECHNICAL_SPEC.md, PROGRESS.md
└── tests/
```

Import antar-folder memakai nama paket, mis. `from metode_2_kata_kunci.classifier import ...` (jalankan dari folder induk).
Istilah penelitian: 32 kata pencarian crawling disebut **seed word** (paper 1, jurnal *Data in Brief*).

## Commands

```bash
# Setup
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt

# Jalankan server
uvicorn aplikasi_web.main:app --reload --port 8000

# Frontend
# Buka http://localhost:8000 di browser

# Test (LLM ditiru, tanpa API key)
pip install -r requirements-dev.txt
pytest
```

## Kategori Klasifikasi

Tiga kategori wajib + satu kategori penyaring — tidak boleh ditambah atau dikurangi:

1. **Angkutan Umum** — bus, angkot, transjakarta, halte, terminal, penumpang, sopir, rute, tarif, ojek
2. **Lahan Parkir** — parkir, jukir, retribusi parkir, gedung parkir, parkir liar, tarif parkir
3. **Manajemen Lalu Lintas** — macet, kemacetan, rambu, lampu merah, trotoar, dishub, ganjil genap

4. **Tidak Relevan** — hanya dari mode LLM: review aplikasi, iklan, berita, topik di luar Dishub

Mode kata kunci: kalau skor semua kategori nol → default ke **Manajemen Lalu Lintas**.

## Output Excel

Wajib dua sheet:

- **Sheet 1 "Hasil Klasifikasi":** semua kolom asli dari CSV + kolom tambahan: `Isi Keluhan`, `Teks Bersih`, `Kategori`, `Sentimen`, `Metode`
- **Sheet 2 "Ringkasan":** tabel jumlah dan persentase per kategori, sentimen, metode + tabel Seleksi Data
- **Sheet 3 "Data Disaring"** (export semua data, jika ada): baris yang tidak dianalisis + `Alasan Disaring`

Alur data: seleksi → pembersihan teks → analisis (LLM / kata kunci). Teks asli tetap di `Isi Keluhan`.

## Aturan Coding

- Bahasa komentar dan variabel: **Indonesia** (sesuai domain)
- Semua fungsi wajib ada docstring singkat
- Jangan hardcode path — pakai `pathlib.Path`
- Encoding CSV: coba UTF-8 dulu, fallback ke `latin-1`
- Hapus file upload setelah diproses (jangan tumpuk di server)
- Jangan pakai `print()` untuk logging — pakai `logging` standar Python

## API Endpoints

```
POST /upload        # Upload CSV, return preview + list kolom
GET  /status-llm    # Cek apakah API key LLM sudah diisi
POST /classify      # Proses klasifikasi (metode: llm | kata_kunci), return JSON hasil
GET  /export/{id}   # Download file Excel hasil
```

## Yang Tidak Boleh Dilakukan

- Jangan install model besar (IndoBERT dll) tanpa konfirmasi — berat, butuh GPU
- Jangan ubah tiga nama kategori
- Jangan simpan data pengguna permanen — semua temporary (teks keluhan dikirim ke Anthropic hanya di mode LLM)
- Jangan commit `.env` / jangan hardcode API key
- Jangan diam-diam campur metode: error akun/kredit LLM harus menghentikan proses
- Jangan pakai pandas untuk export Excel — pakai `openpyxl` langsung
- Jangan tambah autentikasi kecuali diminta
