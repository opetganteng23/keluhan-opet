# Progres Project

Terakhir diperbarui: 10 Oktober 2026

Status: ✅ Sudah · ⏳ Belum · ⚠️ Perlu perhatian

## Tahap 0 – Perapian & Perbaikan Dasar

| # | Pekerjaan | Status | Catatan |
|---|---|---|---|
| 0.1 | File lama/duplikat dipindah ke `_arsip/` | ✅ | Tidak ada yang dihapus |
| 0.2 | Pencocokan kata utuh + frasa terpanjang dulu | ✅ | |
| 0.3 | Daftar kata negatif dibersihkan | ✅ | |
| 0.4 | Excel dihapus otomatis setelah download | ✅ | Teruji otomatis |
| 0.5 | Sesi kedaluwarsa setelah 1 jam | ✅ | |
| 0.6 | `.env`, `.env.example`, `.gitignore` | ✅ | |

## Tahap 1 – Mode Analisis LLM

| # | Pekerjaan | Status | Catatan |
|---|---|---|---|
| 1.1 | `llm_analyzer.py` (Haiku 4.5, prompt penelitian) | ✅ | 20 teks/request, 4 paralel |
| 1.2 | Kategori **Tidak Relevan** | ✅ | |
| 1.3 | Pilihan metode, `GET /status-llm`, kolom **Metode** | ✅ | |
| 1.4 | Error akun → berhenti; gangguan sementara → fallback | ✅ | Teruji otomatis |
| 1.5 | Frontend: pilihan metode & peringatan key | ✅ | ⚠️ Belum dicoba di browser |
| 1.6 | Cakupan ditambah: JakLingko/Mikrotrans, lampu jalan/PJU, JPO | ✅ | Di kata kunci dan prompt LLM |

## Tahap 1b – Perbaikan Lanjutan

| # | Pekerjaan | Status | Catatan |
|---|---|---|---|
| 1b.1 | `klasifikasi_universal.html` diperbarui | ✅ | Hanya preset transportasi; mode LLM; parser CSV diperbaiki; versi lama di `_arsip/`. Logika teruji dengan Node. ⚠️ Belum dicoba di browser |
| 1b.2 | `crawling/crawling.py` → `claude-haiku-4-5` | ✅ | ±5× lebih murah dari Opus 5 |
| 1b.3 | Test otomatis (pytest) | ✅ | 65 test, semua lolos |
| 1b.4 | Batas upload 50 MB | ✅ | |
| 1b.5 | **Bug keamanan:** nama file upload dipakai sebagai path | ✅ | Diganti nama acak |
| 1b.6 | **Bug:** judul sheet Ringkasan tertimpa header | ✅ | Ditemukan & diperbaiki saat menambah tabel Metode |
| 1b.7 | Ringkasan Excel memuat tabel per Metode + Total | ✅ | |
| 1b.8 | README, TECHNICAL_SPEC (v1.1), CLAUDE.md | ✅ | |

## Tahap 1c – Persiapan Data

| # | Pekerjaan | Status | Catatan |
|---|---|---|---|
| 1c.1 | Pembersihan teks (emoji, link, mention, RT, HTML) + kolom **Teks Bersih** | ✅ | `pembersih.py` |
| 1c.2 | Seleksi data + sheet **Data Disaring** + tabel Seleksi Data | ✅ | `seleksi.py`; bisa dimatikan |
| 1c.3 | Uji seleksi pada `tweets_raw.csv` | ✅ | 2.706 → **1.410** dianalisis |
| 1c.4 | Excel lama `klasifikasi_keluhan_transportasi.xlsx` dinyatakan rusak | ✅ | 55% baris kolomnya bergeser; jangan dipakai |
| 1c.5 | Pembersihan & seleksi di `klasifikasi_universal.html` | ⏳ | Belum |
| 1c.6 | Tinjau daftar `AKUN_RESMI_MEDIA` | ⏳ | **Kamu**: pastikan sesuai pertimbangan penelitian |

## Tahap 2 – Uji Coba LLM (butuh kredit)

| # | Pekerjaan | Status | Catatan |
|---|---|---|---|
| 2.1 | Isi `LLM_API_KEY` di `.env` | ✅ | Terdeteksi oleh aplikasi |
| 2.2 | Uji `test.csv`, lalu 50 tweet: cek kualitas label | ⏳ | |
| 2.3 | Ukur biaya nyata, lalu proyeksikan ke 1.410 tweet | ⏳ | |
| 2.4 | Coba kedua aplikasi di browser | ⏳ | |

## Tahap 3 – Analisis Data Penuh

| # | Pekerjaan | Status | Catatan |
|---|---|---|---|
| 3.1 | Jalankan 1.410 tweet (hasil seleksi) dengan LLM | ⏳ | Sumber: `tweets_raw.csv` |
| 3.2 | Jalankan juga dengan kata kunci (pembanding) | ⏳ | |
| 3.3 | Validasi manual 100–200 sampel | ⏳ | Untuk bab hasil skripsi |

## Tahap 4 – Clustering Sub-topik

| # | Pekerjaan | Status | Catatan |
|---|---|---|---|
| 4.1 | Rancang & implementasi | ⏳ | Sebaiknya setelah Tahap 3, supaya kategorinya sudah valid |

## Perlu Perhatian

| # | Hal | Catatan |
|---|---|---|
| P.1 | Mode LLM belum pernah dijalankan dengan key asli | Prompt, skema tool, dan biaya baru terbukti di Tahap 2 |
| P.2 | Mode LLM di aplikasi utama belum punya progress bar | Untuk 2.390 baris perlu menunggu beberapa menit; jangan tutup halaman |
| P.3 | Library terbaru memberi peringatan *deprecation* di Python 3.14 | Hanya peringatan; aplikasi berjalan normal |
