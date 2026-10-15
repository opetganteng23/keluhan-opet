"""Test pra-pemrosesan teks (penetralan sebelum analisis)."""

import pytest

from persiapan_data.pembersih import bersihkan_teks


@pytest.mark.parametrize("asli, bersih", [
    ("Hebat banget TJ, nunggu 1 jam 🙃", "Hebat banget TJ, nunggu 1 jam"),
    ("Macet 😡😡😡 parah", "Macet parah"),
    ("Keluarga 👨‍👩‍👧 naik bus 👍🏽", "Keluarga naik bus"),
    ("Bendera 🇮🇩 merah putih", "Bendera merah putih"),
    ("@dishubjakarta tolong cek halte ini", "tolong cek halte ini"),
    ("RT @akun: parkir liar lagi", "parkir liar lagi"),
    ("Lihat https://t.co/abc123 dan www.contoh.com ya", "Lihat dan ya"),
    ("#macet #JakLingko pagi ini", "macet JakLingko pagi ini"),
    ("Bus &amp; halte &quot;kotor&quot;", 'Bus & halte "kotor"'),
    ("baris satu\n\nbaris   dua", "baris satu baris dua"),
    ("email@contoh.com tetap", "email@contoh.com tetap"),
    ("Tarif Rp3.500, jam 06.00 – 07.30", "Tarif Rp3.500, jam 06.00 – 07.30"),
])
def test_bersihkan_teks(asli, bersih):
    """Emoji, link, mention, RT, dan kode HTML dibuang; isi kalimat utuh."""
    assert bersihkan_teks(asli) == bersih


@pytest.mark.parametrize("asli", ["", "   ", "🙃😡", "@dishubjakarta https://t.co/x"])
def test_hanya_noise_jadi_kosong(asli):
    """Teks yang isinya hanya emoji/mention/link menjadi kosong."""
    assert bersihkan_teks(asli) == ""
