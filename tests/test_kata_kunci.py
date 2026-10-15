"""Test klasifikasi kategori dan sentimen berbasis kata kunci."""

import pytest

from metode_2_kata_kunci.classifier import KATEGORI_DEFAULT, hitung_kemunculan, klasifikasi_keluhan
from metode_2_kata_kunci.sentiment import analisis_sentimen


@pytest.mark.parametrize("teks, kategori", [
    ("Bus transjakarta selalu telat", "Angkutan Umum"),
    ("Naik JakLingko hari ini", "Angkutan Umum"),
    ("Parkir liar di depan halte, tarif parkir mahal", "Lahan Parkir"),
    ("Jalan macet karena lampu merah rusak", "Manajemen Lalu Lintas"),
    ("Lampu jalan mati, trotoar gelap", "Manajemen Lalu Lintas"),
])
def test_kategori(teks, kategori):
    """Kategori ditentukan oleh kata kunci yang paling banyak muncul."""
    assert klasifikasi_keluhan(teks) == kategori


@pytest.mark.parametrize("teks", ["", "   ", "Busuk banget pelayanannya"])
def test_kategori_default(teks):
    """Teks kosong atau tanpa kata kunci jatuh ke kategori default."""
    assert klasifikasi_keluhan(teks) == KATEGORI_DEFAULT


def test_hanya_kata_utuh():
    """'busuk' tidak boleh terbaca sebagai 'bus'."""
    assert hitung_kemunculan("busuk", {"a": ["bus"]}) == {"a": 0}


def test_frasa_terpanjang_dulu():
    """'tarif parkir' dihitung sekali sebagai frasa, bukan ditambah 'tarif'."""
    skor = hitung_kemunculan("tarif parkir naik", {"angkutan": ["tarif"], "parkir": ["tarif parkir"]})
    assert skor == {"angkutan": 0, "parkir": 1}


def test_duplikat_kata_kunci_tidak_dobel():
    """Kata kunci yang ditulis dua kali tetap dihitung sekali."""
    assert hitung_kemunculan("goblok", {"neg": ["goblok", "goblok"]}) == {"neg": 1}


@pytest.mark.parametrize("teks, sentimen", [
    ("Sopir ugal-ugalan dan telat", "Negatif"),
    ("Pelayanan ramah dan tepat waktu", "Positif"),
    ("Halte tidak nyaman", "Negatif"),
    ("Saya mau menyampaikan keluhan", "Netral"),
    ("", "Netral"),
])
def test_sentimen(teks, sentimen):
    """Sentimen dari selisih jumlah kata positif dan negatif."""
    assert analisis_sentimen(teks) == sentimen
