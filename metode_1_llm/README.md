# Metode 1 – LLM (metode utama)

Model bahasa besar (`claude-haiku-4-5`) membaca makna setiap teks dan memilih kategori + sentimen dari daftar tetap (jawaban dikunci dengan skema JSON).

- `llm_analyzer.py` – prompt penelitian, skema jawaban, pengiriman 20 teks per request, penanganan error.
- Bukti cara kerja: teks yang dikirim + jawaban mentah model per request (lihat `evaluasi/`).
- Butuh `LLM_API_KEY` di file `.env` (folder induk).
