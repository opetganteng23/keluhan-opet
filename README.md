# Klasifikasi Keluhan Transportasi

Aplikasi untuk **menganalisis keluhan warga Jakarta tentang transportasi secara otomatis**, sebagai bagian dari penelitian tugas akhir tentang kinerja Dinas Perhubungan (Dishub) DKI Jakarta di lapangan.

Ribuan keluhan (misalnya hasil crawling Twitter/X dari folder `crawling/`) dalam bentuk CSV dianalisis, lalu setiap keluhan diberi:

| Label | Pertanyaan yang dijawab | Nilai |
|---|---|---|
| **Kategori** | Keluhan ini soal apa? | Angkutan Umum / Lahan Parkir / Manajemen Lalu Lintas / Tidak Relevan |
| **Sentimen** | Nadanya bagaimana? | Positif / Negatif / Netral |
| **Metode** | Siapa yang memberi label? | LLM / Kata Kunci / Kata Kunci (fallback) / Kosong |

Hasil akhirnya berupa file Excel. Dari file itu langsung terlihat masalah mana yang paling banyak dikeluhkan dan seberapa negatif nadanya.

Detail teknis lengkap ada di [TECHNICAL_SPEC.md](TECHNICAL_SPEC.md). Status pekerjaan ada di [PROGRESS.md](PROGRESS.md).

## Ruang Lingkup

| Kategori | Cakupan |
|---|---|
| **Angkutan Umum** | TransJakarta, JakLingko/Mikrotrans, angkot, ojek, halte, sopir, tarif, rute |
| **Lahan Parkir** | Parkir liar, jukir, retribusi/tarif parkir, lahan parkir |
| **Manajemen Lalu Lintas** | Kemacetan, rambu, lampu lalu lintas, ganjil genap, petugas Dishub, dan fasilitas jalan (trotoar, lampu jalan/PJU, JPO, zebra cross) |
| **Tidak Relevan** | Review aplikasi, iklan, berita, banjir/sampah, dan topik di luar Dishub. *Hanya dari mode LLM.* |

## Persiapan Data Otomatis

Sebelum dianalisis, data melewati dua tahap. Keduanya berlaku untuk metode LLM maupun kata kunci.

**1. Seleksi data** (bisa dimatikan lewat kotak centang di halaman). Data yang disaring tidak dihapus, tetapi dipindah ke sheet **Data Disaring** bersama alasannya.

| Disaring | Alasan |
|---|---|
| Teks kosong setelah dibersihkan | Isinya hanya emoji, link, atau mention |
| Akun resmi/media | Pengumuman/berita, bukan suara warga. Daftar akunnya ada di `AKUN_RESMI_MEDIA` di [seleksi.py](seleksi.py) dan bisa diedit |
| Duplikat | Teks yang sama persis (setelah dibersihkan); hanya yang pertama dipertahankan |

**2. Pembersihan teks** ([pembersih.py](pembersih.py)): emoji, link, mention, awalan `RT`, dan kode HTML dibuang; `#macet` → `macet`; spasi dirapikan. Hasilnya ada di kolom **Teks Bersih**, sedangkan teks asli tetap di **Isi Keluhan**.

Contoh pada `tweets_raw.csv`: 2.706 data awal → 1.268 akun resmi/media, 9 duplikat, 19 kosong → **1.410 dianalisis**.

## Dua Metode Analisis

| | **AI / LLM** (utama) | **Kata Kunci** (cadangan/pembanding) |
|---|---|---|
| Cara kerja | Model `claude-haiku-4-5` membaca makna kalimat | Mencocokkan daftar kata |
| Paham konteks & sarkasme | Ya | Tidak |
| Bisa memilah "Tidak Relevan" | Ya | Tidak |
| Biaya | Berbayar (perlu API key) | Gratis |
| Kecepatan | Hitungan menit untuk ribuan baris | Hitungan detik |
| Data keluar komputer | Ya, teks dikirim ke penyedia API LLM | Tidak |

Untuk skripsi, kedua metode bisa dijalankan pada data yang sama, lalu hasilnya dibandingkan.

## Persiapan

1. Isi API key di file **`.env`**:
   ```
   LLM_API_KEY=...
   ```
   File `.env` tidak ikut ter-upload ke Git. Tanpa key, aplikasi tetap bisa dipakai dengan mode Kata Kunci.
2. Klik dua kali **`start.bat`**. Library akan terpasang dan server menyala, lalu browser membuka http://localhost:8000.

## Alur Pakai

```
CSV → Upload → Pilih kolom keluhan → Pilih metode → Jalankan → Lihat ringkasan → Download Excel
```

1. Upload file CSV. Untuk mencoba, pakai `test.csv` dengan kolom `isi`.
2. Pilih kolom yang berisi teks keluhan.
3. Pilih metode: **AI / LLM** atau **Kata Kunci**.
4. Klik **Jalankan Analisis**.
5. Download Excel, bisa untuk semua data atau per kategori.
6. Selesai? Klik **`stop.bat`**.

## Isi File Excel

- **Sheet "Hasil Klasifikasi":** data yang lolos seleksi: kolom asli CSV + `Isi Keluhan`, `Teks Bersih`, `Kategori`, `Sentimen`, `Metode`.
- **Sheet "Ringkasan":** jumlah dan persentase per Kategori, per Sentimen, dan per Metode, total data, plus tabel **Seleksi Data**.
- **Sheet "Data Disaring":** baris yang tidak dianalisis + alasannya (hanya pada export semua data).

## Kalau API LLM Bermasalah

| Masalah | Yang terjadi |
|---|---|
| API key salah / kredit habis | Proses **dihentikan** dengan pesan jelas. Tidak ada data yang dicampur metode lain. |
| Gangguan jaringan / server sibuk | Dicoba ulang otomatis. Kalau tetap gagal, kelompok itu memakai kata kunci dan ditandai **"Kata Kunci (fallback)"** di kolom Metode. |

## Cara Kerja Mode Kata Kunci

- Kategori dengan jumlah kata kunci terbanyak yang menang. Kalau tidak ada yang cocok → **Manajemen Lalu Lintas** (default). Kalau seri → kategori yang lebih dulu di tabel Ruang Lingkup.
- Sentimen: jumlah kata positif dibandingkan dengan kata negatif. Kalau sama banyak → **Netral**.
- Hanya kata utuh yang dihitung, jadi "busuk" ≠ "bus". Frasa panjang dicocokkan lebih dulu, jadi "tarif parkir" ≠ "tarif" dan "tidak nyaman" ≠ "nyaman".
- Daftar kata bisa diedit di [classifier.py](classifier.py) dan [sentiment.py](sentiment.py).

## Versi Mandiri: `klasifikasi_universal.html`

Aplikasi satu file yang dibuka langsung di browser, **tanpa server Python**. Kategorinya sama, dan kata kuncinya bisa ditambah atau dihapus lewat tampilan. Mode LLM juga tersedia: key diketik di halaman dan tidak disimpan. Butuh internet untuk library Excel dan untuk LLM.

## Evaluasi Akurasi (untuk penelitian)

Mengukur seberapa sering label engine sama dengan label manual manusia (*gold standard*).

1. `python evaluasi/siapkan_sampel.py` → ambil sampel acak 385 tweet (Slovin, e = 5%), jalankan kedua engine, lalu buat di `Downloads\evaluasi_akurasi\`:
   - `lembar_label_pelabel1.xlsx` dan `lembar_label_pelabel2.xlsx`: diisi manual (dropdown, plus sheet **Pedoman**), **tanpa** label sistem.
   - `hasil_engine_sampel.xlsx`: label kedua engine + **bukti cara kerja** (kata yang cocok, skor, prompt, jawaban mentah LLM per request). Jangan dibuka sebelum selesai melabeli.
2. Isi lembar pelabelan. Pelabel 2 opsional, untuk Cohen's Kappa.
3. `python evaluasi/hitung_evaluasi.py` → `laporan_evaluasi.xlsx`: akurasi, precision, recall, F1, macro-F1, confusion matrix (heatmap), grafik perbandingan, Cohen's Kappa, dan daftar kesalahan beserta buktinya.

## Struktur Project

```
keluhanku├── metode_1_llm\                 LLM (metode utama)
├── metode_2_kata_kunci\          Leksikon / kata kunci (baseline)
├── metode_3_tfidf_naive_bayes\   TF-IDF + Naive Bayes (belum dibuat)
├── metode_4_tfidf_svm\           TF-IDF + SVM (belum dibuat)
├── persiapan_data\               Seleksi, pembersihan teks, ekspor Excel siap analisis
├── evaluasi\                     Sampel 385, lembar pelabelan, hitung akurasi
├── aplikasi_web\                 Aplikasi upload CSV (FastAPI) + versi mandiri HTML
├── dokumentasi\                  TECHNICAL_SPEC.md, PROGRESS.md
├── tests\                        Test otomatis (pytest)
└── start.bat · stop.bat · requirements.txt · .env
```

Setiap folder punya README.md singkat. Jalankan semua perintah dari folder `keluhanku`, misalnya `python persiapan_data/ekspor_siap_analisis.py`.

## Menjalankan Test

```bash
venv\Scripts\activate
pip install -r requirements-dev.txt
pytest
```

Test tidak memanggil API LLM sungguhan (ditiru), jadi tidak butuh API key dan tidak ada biaya.

## Privasi Data

- CSV upload dihapus segera setelah dibaca. Excel dihapus setelah didownload. Sesi di memori dibersihkan setelah 1 jam.
- Ukuran upload maksimal 50 MB.
- **Mode LLM mengirim teks keluhan ke penyedia API LLM** untuk dianalisis. Mode Kata Kunci sepenuhnya lokal.

## Setup Manual (tanpa start.bat)

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
uvicorn aplikasi_web.main:app --reload --port 8000
```
