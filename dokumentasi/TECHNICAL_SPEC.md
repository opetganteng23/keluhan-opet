# Master Technical Specification
## Klasifikasi Keluhan Transportasi

| | |
|---|---|
| Versi dokumen | 1.1 |
| Tanggal | 10 Oktober 2026 |
| Status | Sesuai kode saat ini (as-built) |
| Pemilik | Rizky Fadhillah |

**Perubahan dari 1.0:** mode analisis LLM, kategori Tidak Relevan, kolom Metode, endpoint `/status-llm`, batas ukuran upload, test otomatis, perbaikan sheet Ringkasan, dan versi mandiri yang diperbarui.

---

## 1. Ringkasan

### 1.1 Tujuan
Mengklasifikasikan keluhan warga Jakarta tentang transportasi secara otomatis, untuk penelitian tugas akhir tentang **kinerja nyata Dishub DKI Jakarta di lapangan**. Input berupa CSV mentah (biasanya hasil crawling Twitter/X). Output berupa Excel berisi label kategori, sentimen, dan metode untuk setiap keluhan, ditambah ringkasan statistiknya.

### 1.2 Posisi dalam Pipeline Penelitian
```
crawling/ (twscrape)            keluhanku/ (aplikasi ini)
──────────────────────          ──────────────────────────────────────
crawl_raw_only.py ──▶ tweets_raw.csv ──▶ upload ──▶ analisis ──▶ Excel
```
Folder `crawling/` juga punya analisis Claude sendiri (`crawling.py`, kini memakai `claude-haiku-4-5`). Aplikasi ini adalah tempat analisis utama karena punya antarmuka, export, dan dua metode yang bisa dibandingkan.

### 1.3 Ruang Lingkup
| Termasuk | Tidak termasuk |
|---|---|
| Upload CSV, pemilihan kolom keluhan | Crawling (ada di folder `crawling/`) |
| 3 kategori transportasi + Tidak Relevan | Topik di luar Dishub |
| Sentimen 3 kelas | Model lokal (IndoBERT, dll.) |
| Dua metode: LLM & kata kunci | Autentikasi, multi-user, database |
| Export Excel (semua / per kategori) | Penyimpanan riwayat analisis |
| Test otomatis | Clustering sub-topik (rencana, §12) |

---

## 2. Arsitektur

```
┌─────────────────────┐   HTTP (JSON/multipart)   ┌───────────────────────────────┐   HTTPS   ┌──────────────┐
│ Browser             │ ────────────────────────▶ │ FastAPI (main.py)             │ ────────▶ │ Anthropic API│
│ static/index.html   │ ◀──────────────────────── │  ├─ llm_analyzer.py (Claude)  │ ◀──────── │ claude-haiku │
└─────────────────────┘        JSON / .xlsx       │  ├─ classifier.py  (kata)     │           └──────────────┘
                                                  │  ├─ sentiment.py   (kata)     │   (hanya mode LLM)
                                                  │  ├─ exporter.py    (Excel)    │
                                                  │  └─ SESI (dict di memori)     │
                                                  └──────────────┬────────────────┘
                                                                 ▼
                                                            uploads/ (sementara)
```

### 2.1 Stack Teknologi

| Lapisan | Teknologi | Versi |
|---|---|---|
| Bahasa | Python | 3.14 (venv) |
| Web framework | FastAPI | 0.115.0 |
| Server ASGI | Uvicorn[standard] | 0.30.6 |
| Upload multipart | python-multipart | 0.0.9 |
| Export Excel | openpyxl | 3.1.5 |
| LLM | anthropic (SDK Python) | 1.13.0 terpasang |
| Config | python-dotenv | terbaru |
| Test | pytest + httpx | `requirements-dev.txt` |
| Frontend | HTML/CSS/JS vanilla, satu file | — |

Pandas sengaja **tidak** dipakai.

### 2.2 Struktur File

| File | Tanggung jawab |
|---|---|
| `main.py` | Endpoint, sesi, pembacaan CSV, batas upload, middleware no-cache |
| `llm_analyzer.py` | Prompt penelitian, skema tool, pengelompokan & paralelisme, fallback |
| `classifier.py` | Kamus kategori, fungsi `hitung_kemunculan` (dipakai bersama) |
| `sentiment.py` | Leksikon positif/negatif |
| `exporter.py` | Workbook dua sheet |
| `static/index.html` | Antarmuka |
| `klasifikasi_universal.html` | Versi mandiri tanpa server (§10) |
| `tests/` | Test otomatis (§11) |
| `.env` | `LLM_API_KEY` (tidak di-commit) |

---

## 3. Alur Proses

```
[1] POST /upload   → batas 50 MB → simpan (nama acak) → parse → hapus file → sesi → preview
[2] POST /classify → metode?
        llm        → kelompok 20 teks → maks 4 paralel → LLM → label
        kata_kunci → classifier + sentiment → label
     → simpan hasil di sesi → JSON hasil + ringkasan
[3] GET /export/{id} → (filter kategori) → tulis .xlsx → kirim → hapus file
```

---

## 4. Spesifikasi API

Base URL: `http://localhost:8000`

### 4.1 `GET /status-llm`
Memberi tahu apakah mode LLM bisa dipakai.
```json
{ "tersedia": true }
```
`true` jika `LLM_API_KEY` diawali `sk-ant-` dan bukan placeholder dari `.env.example`.

### 4.2 `POST /upload`
**Request:** `multipart/form-data`, field `file` (`.csv`, maks 50 MB).

**Response 200:**
```json
{ "id": "<id sesi>", "kolom": ["no","isi","lokasi"], "preview": [ ... maks 5 baris ... ], "jumlah_baris": 2390 }
```

| Kode | Kondisi |
|---|---|
| 400 | Bukan `.csv`, atau tidak ada baris data |
| 413 | File > 50 MB |

### 4.3 `POST /classify`
**Request:**
```json
{ "id": "<id sesi>", "kolom_keluhan": "isi", "metode": "llm" }
```
`metode`: `"llm"` atau `"kata_kunci"` (default `"kata_kunci"`).

**Response 200:**
```json
{
  "id": "<id sesi>",
  "hasil": [{ "...kolom asli...": "", "Isi Keluhan": "...", "Kategori": "Lahan Parkir", "Sentimen": "Negatif", "Metode": "LLM" }],
  "ringkasan_kategori": { "Lahan Parkir": 1 },
  "ringkasan_sentimen": { "Negatif": 1 },
  "metode": "llm",
  "jumlah_fallback": 0
}
```

| Kode | Kondisi |
|---|---|
| 404 | Sesi tidak ada / kedaluwarsa |
| 400 | Kolom tidak ada, metode tidak dikenal, atau mode LLM tanpa API key |
| 502 | API LLM menolak (key salah, kredit habis, permintaan tidak valid). Pesan dari Anthropic diteruskan. |

### 4.4 `GET /export/{id_sesi}?kategori=...`
`kategori` opsional: salah satu dari 3 kategori atau `Tidak Relevan`.

| Kode | Kondisi |
|---|---|
| 200 | File `.xlsx` (`hasil_klasifikasi_keluhan.xlsx` atau `hasil_<kategori>.xlsx`) |
| 404 | Sesi tidak ada / belum di-classify |
| 400 | Kategori tidak dikenal |

### 4.5 `GET /`
Menyajikan `static/index.html`. Semua respons memakai `Cache-Control: no-store`.

---

## 5. Logika Analisis

### 5.1 Kategori

| Kategori | Cakupan | Sumber |
|---|---|---|
| **Angkutan Umum** | TransJakarta, JakLingko/Mikrotrans, angkot, ojek, halte, sopir, tarif, rute | Kedua metode |
| **Lahan Parkir** | Parkir liar, jukir, retribusi, lahan parkir | Kedua metode |
| **Manajemen Lalu Lintas** | Macet, rambu, lampu lalu lintas, ganjil genap, petugas Dishub, fasilitas jalan (trotoar, PJU, JPO, zebra cross, marka) | Kedua metode |
| **Tidak Relevan** | Review aplikasi, iklan, berita/pengumuman, topik di luar Dishub | Hanya LLM |

### 5.2 Metode LLM (`llm_analyzer.py`)

| Parameter | Nilai | Alasan |
|---|---|---|
| Model | `claude-haiku-4-5` | Termurah; cukup untuk klasifikasi tertutup |
| Teks per request | 20 (`UKURAN_KELOMPOK`) | Prompt sistem hanya dibayar sekali per 20 teks |
| Request paralel | 4 (`MAKS_PARALEL`) | Cepat tanpa mudah terkena rate limit |
| Panjang teks maks | 1.000 karakter | Mencegah teks sangat panjang memboroskan token |
| Retry SDK | 4 kali | Untuk 429/5xx/jaringan |
| Format output | Tool `kirim_analisis` dengan `strict: true` dan `tool_choice` paksa | Jawaban selalu sesuai skema (enum kategori/sentimen) |

**Prompt sistem** diadaptasi dari `crawling/crawling.py`. Isinya: konteks penelitian Dishub, definisi 3 kategori, aturan tegas kapan memilih `tidak_relevan`, contoh relevan/tidak relevan, dan instruksi menangkap sarkasme.

**Penanganan error:**

| Situasi | Perilaku |
|---|---|
| `AuthenticationError`, `PermissionDeniedError`, `BadRequestError` (termasuk kredit habis) | **Seluruh proses dihentikan** → HTTP 502 |
| Error lain setelah retry habis (jaringan, 5xx, 429) | Kelompok itu dianalisis dengan kata kunci → `Metode = "Kata Kunci (fallback)"` |
| LLM melewatkan nomor tertentu | Teks itu memakai kata kunci → fallback |
| Teks kosong | Tidak dikirim → `Tidak Relevan / Netral / "Kosong"` |

Alasan menghentikan proses saat ada masalah akun: data penelitian tidak boleh diam-diam tercampur dua metode. Ini pernah terjadi di `crawling/`, ketika 84 permintaan gagal karena kredit habis.

### 5.3 Metode Kata Kunci

**Algoritma pencocokan (`hitung_kemunculan`):**
1. Ubah teks ke huruf kecil.
2. Urutkan kata kunci dari frasa terpanjang.
3. Cari kecocokan kata utuh `(?<!\w)kata(?!\w)`.
4. Setiap kecocokan menambah skor dan diganti spasi, supaya tidak terhitung lagi oleh kata yang lebih pendek.

**Kategori:** skor tertinggi menang. Semua skor 0 → Manajemen Lalu Lintas. Seri → urutan Angkutan Umum, Lahan Parkir, Manajemen Lalu Lintas.

**Sentimen:** positif > negatif → Positif; negatif > positif → Negatif; sama → Netral. Kata "keluhan" sengaja tidak masuk daftar negatif.

---

## 6. Spesifikasi Data

### 6.1 Input CSV
| Aspek | Ketentuan |
|---|---|
| Format | CSV ber-header, delimiter koma |
| Encoding | `utf-8-sig`, fallback `latin-1` |
| Ukuran | Maks 50 MB |

### 6.2 Output Excel
**Sheet "Hasil Klasifikasi":** kolom asli + `Isi Keluhan`, `Kategori`, `Sentimen`, `Metode`.

**Sheet "Ringkasan":** tiga tabel berurutan (per Kategori, per Sentimen, per Metode), masing-masing berisi Nilai, Jumlah, dan Persentase (teks 1 desimal), lalu baris **Total Data**. Kalau export difilter, ringkasan dihitung dari data yang sudah difilter.

### 6.3 Sesi di Memori
```python
SESI[id] = {"kolom": [...], "baris": [...], "dibuat": float, "hasil": [...]}
```

---

## 7. Keamanan & Privasi

| Aspek | Ketentuan |
|---|---|
| CSV upload | Disimpan dengan nama acak (nama file pengguna **tidak** dipakai sebagai path), dihapus segera setelah dibaca |
| Excel export | Nama acak, dihapus setelah terkirim |
| Sesi | Dihapus jika > 1 jam (dicek saat upload baru) atau saat server restart |
| API key | Hanya di `.env` (masuk `.gitignore`); tidak pernah dikirim ke browser |
| Data ke pihak ketiga | **Mode LLM mengirim teks ke penyedia API LLM.** Mode kata kunci sepenuhnya lokal. |
| Jaringan | Server listen di `127.0.0.1` saja |
| Autentikasi | Tidak ada (sesuai aturan project) |

---

## 8. Kebutuhan Non-Fungsional

| Aspek | Ketentuan |
|---|---|
| Platform | Windows (`.bat`); kode lintas-platform |
| Kinerja kata kunci | 2.390 baris dalam hitungan detik |
| Kinerja LLM | ≈ 120 request untuk 2.390 baris, 4 paralel. **Waktu & biaya nyata belum diukur** (PROGRESS 2.2–2.3). |
| Logging | `logging` standar; jumlah fallback dicatat per proses |
| Gaya kode | Indonesia, docstring di setiap fungsi, `pathlib.Path` |

---

## 9. Batasan & Risiko

| # | Batasan | Mitigasi |
|---|---|---|
| 1 | Akurasi LLM belum divalidasi pada data nyata | Validasi manual 100–200 sampel (§11.3) |
| 2 | Permintaan `/classify` mode LLM bisa berjalan beberapa menit tanpa indikator progres | Pesan "jangan tutup halaman"; progres bertahap bisa ditambah nanti |
| 3 | Mode kata kunci: kosakata terbatas, tanpa stemming/slang, negasi terbatas | Pakai mode LLM sebagai utama |
| 4 | Mode kata kunci: default Manajemen Lalu Lintas bisa menggelembungkan kategori itu | Bandingkan dengan LLM yang punya Tidak Relevan |
| 5 | Sesi di memori, satu proses | Cukup untuk pemakaian lokal |
| 6 | Hasil LLM bisa sedikit berbeda jika dijalankan ulang | Simpan Excel hasil sebagai data final penelitian |

---

## 10. Versi Mandiri: `klasifikasi_universal.html`

Satu file HTML yang berjalan di browser tanpa server.

| Aspek | Ketentuan |
|---|---|
| Kategori | Preset transportasi (sama dengan aplikasi utama); kata kunci bisa diedit; tombol "Kembalikan ke Default" |
| Metode | Kata kunci (offline) atau Claude (`claude-haiku-4-5`, dipanggil langsung dari browser) |
| API key | Diketik di halaman, **tidak disimpan** |
| Parser CSV | Mendukung teks berkutip berisi enter, koma, dan `""`; delimiter `,` atau `;` otomatis |
| Encoding | UTF-8; jika muncul karakter rusak, dibaca ulang sebagai Latin-1 |
| Tanpa kata kunci cocok | `Tidak Terklasifikasi` (tidak dipaksa ke kategori pertama) |
| Excel | SheetJS 0.18.5 (CDN). Kolom asli + Isi Teks, Kategori, Sentimen, Metode; Ringkasan per kategori, sentimen, dan metode |
| Error akun LLM | Proses dihentikan; gangguan sementara → fallback kata kunci per kelompok |

Versi sebelum perubahan disimpan di `_arsip/klasifikasi_universal_lama.html`.

---

## 11. Pengujian

### 11.1 Test Otomatis
`pytest` (43 test, semua lolos). LLM ditiru, jadi test tidak butuh key dan tidak ada biaya.

| File | Cakupan |
|---|---|
| `tests/test_kata_kunci.py` | Kategori & sentimen, kata utuh, frasa terpanjang, duplikat, default |
| `tests/test_llm_analyzer.py` | Pemetaan label, teks kosong, nomor terlewat, fallback gangguan sementara, penghentian saat error akun, pengelompokan 20 teks, deteksi placeholder key |
| `tests/test_api.py` | Upload (latin-1, ditolak, terlalu besar), classify (dua metode, tanpa key, input salah), export (2 sheet, kolom Metode, struktur Ringkasan, file terhapus, filter kategori) |

Menjalankan: `pip install -r requirements-dev.txt` lalu `pytest`.

### 11.2 Belum Diuji
- Mode LLM dengan API key sungguhan
- Antarmuka di browser (kedua aplikasi)

### 11.3 Validasi Akurasi (untuk penelitian)
1. Ambil sampel acak 100–200 baris, lalu beri label manual tanpa melihat label sistem.
2. Hitung akurasi dan confusion matrix untuk LLM dan kata kunci.
3. Laporkan keduanya sebagai perbandingan metode.

---

## 12. Rencana Pengembangan

| Prioritas | Item |
|---|---|
| Tinggi | Uji 50 tweet dengan key asli: kualitas + biaya |
| Tinggi | Validasi akurasi (§11.3) |
| Sedang | Clustering sub-topik per kategori dengan LLM |
| Sedang | Indikator progres untuk mode LLM di aplikasi utama |
| Rendah | Batch API (diskon 50%) jika data jauh lebih besar |

---

## 13. Aturan Project (ringkas, lihat CLAUDE.md)

- Kategori: 3 kategori wajib + Tidak Relevan, tidak boleh diubah
- Excel memakai `openpyxl` langsung, bukan pandas
- Data pengguna tidak disimpan permanen
- API key hanya di `.env`
- Error akun LLM harus menghentikan proses, tidak boleh diam-diam mencampur metode
- Tanpa autentikasi kecuali diminta; tanpa model besar lokal tanpa konfirmasi
