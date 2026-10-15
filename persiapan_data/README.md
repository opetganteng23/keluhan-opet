# Persiapan Data

- `seleksi.py` – buang akun resmi/media (`AKUN_RESMI_MEDIA`), duplikat, teks kosong.
- `pembersih.py` – buang emoji, link, mention, awalan RT, kode HTML.
- `ekspor_siap_analisis.py` – buat Excel data siap analisis dari `..\..\crawling\tweets_raw.csv`.
  `python persiapan_data/ekspor_siap_analisis.py` (Excel) · `... --hitung` (jumlah saja)
