# Aplikasi Web

Upload CSV → pilih kolom → pilih metode (LLM / Kata Kunci) → download Excel.
Jalankan dari folder induk dengan `start.bat` (atau `uvicorn aplikasi_web.main:app --port 8000`).

- `main.py` – server FastAPI · `exporter.py` – Excel · `static/index.html` – tampilan
- `klasifikasi_universal.html` – versi mandiri tanpa server · `uploads/` – file sementara
