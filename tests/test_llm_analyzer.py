"""Test analisis LLM dengan klien tiruan (tanpa API key, tanpa biaya)."""

import asyncio
from types import SimpleNamespace

import anthropic
import pytest

from metode_1_llm import llm_analyzer


def _galat(kelas):
    """Membuat objek error SDK tanpa respons HTTP sungguhan."""
    galat = kelas.__new__(kelas)
    galat.message = "galat uji"
    return galat


class KlienPalsu:
    """Pengganti AsyncAnthropic; perilaku diatur lewat fungsi `jawab`."""

    def __init__(self, jawab):
        self.jawab = jawab
        self.messages = SimpleNamespace(create=self._create)
        self.jumlah_panggilan = 0

    async def _create(self, **kwargs):
        self.jumlah_panggilan += 1
        return self.jawab(kwargs)

    async def __aenter__(self):
        return self

    async def __aexit__(self, *args):
        return False


def _pasang(monkeypatch, jawab):
    """Mengganti AsyncAnthropic di llm_analyzer dengan KlienPalsu."""
    klien = KlienPalsu(jawab)
    monkeypatch.setattr(llm_analyzer.anthropic, "AsyncAnthropic", lambda **kw: klien)
    return klien


def _respons_tool(hasil):
    """Membuat respons tiruan berisi satu blok tool_use kirim_analisis."""
    blok = SimpleNamespace(type="tool_use", name="kirim_analisis", input={"hasil": hasil})
    return SimpleNamespace(content=[blok])


def _item(nomor, kategori, sentimen, keluhan=False, sebab="tidak disebutkan", akibat="tidak disebutkan",
          emosi="tidak_ada"):
    """Satu jawaban tiruan sesuai skema tool versi 2."""
    return {"nomor": nomor, "kategori": kategori, "sentimen": sentimen, "keluhan": keluhan,
            "sebab": sebab, "akibat": akibat, "emosi": emosi}


def test_label_dipetakan_ke_nama_tampilan(monkeypatch):
    """Kode kategori/sentimen/emosi LLM diubah ke nama tampilan; sebab-akibat diteruskan."""
    _pasang(monkeypatch, lambda kw: _respons_tool([
        _item(1, "parkir", "negatif", True, "banyak jukir liar", "tarif parkir tidak wajar", "kesal"),
        _item(2, "tidak_relevan", "netral"),
    ]))
    hasil = asyncio.run(llm_analyzer.analisis_dengan_llm(["parkir liar", "promo diskon"]))
    assert hasil == [
        {"Kategori": "Lahan Parkir", "Sentimen": "Negatif", "Keluhan": "Ya", "Sebab": "banyak jukir liar",
         "Akibat": "tarif parkir tidak wajar", "Emosi": "Kesal", "Keluhan Lengkap": "Ya", "Metode": "LLM"},
        {"Kategori": "Tidak Relevan", "Sentimen": "Netral", "Keluhan": "Tidak", "Sebab": "tidak disebutkan",
         "Akibat": "tidak disebutkan", "Emosi": "Tidak Ada", "Keluhan Lengkap": "Tidak", "Metode": "LLM"},
    ]


def test_keluhan_tanpa_sebab_tidak_lengkap(monkeypatch):
    """Keluhan tanpa sebab tetap keluhan, tapi Keluhan Lengkap = Tidak."""
    _pasang(monkeypatch, lambda kw: _respons_tool([
        _item(1, "manajemen_lalu_lintas", "negatif", True, "tidak disebutkan", "telat kerja", "lelah"),
    ]))
    h = asyncio.run(llm_analyzer.analisis_dengan_llm(["kejebak macet sejam, telat ngantor"]))[0]
    assert (h["Keluhan"], h["Keluhan Lengkap"], h["Emosi"]) == ("Ya", "Tidak", "Lelah")


def test_tidak_relevan_tidak_boleh_keluhan(monkeypatch):
    """Walau LLM menjawab keluhan=true, kategori Tidak Relevan dipaksa bukan keluhan."""
    _pasang(monkeypatch, lambda kw: _respons_tool([
        _item(1, "tidak_relevan", "negatif", True, "banjir", "rumah kebanjiran", "kesal"),
    ]))
    h = asyncio.run(llm_analyzer.analisis_dengan_llm(["banjir lagi"]))[0]
    assert (h["Keluhan"], h["Sebab"], h["Keluhan Lengkap"]) == ("Tidak", "tidak disebutkan", "Tidak")


def test_teks_kosong_tidak_dikirim(monkeypatch):
    """Teks kosong langsung diberi label tanpa memanggil API."""
    klien = _pasang(monkeypatch, lambda kw: pytest.fail("API tidak boleh dipanggil"))
    hasil = asyncio.run(llm_analyzer.analisis_dengan_llm(["", "  "]))
    assert all(h["Metode"] == "Kosong" for h in hasil)
    assert klien.jumlah_panggilan == 0


def test_nomor_terlewat_pakai_kata_kunci(monkeypatch):
    """Teks yang tidak dijawab LLM dianalisis ulang dengan kata kunci."""
    _pasang(monkeypatch, lambda kw: _respons_tool([
        _item(1, "angkutan_umum", "negatif"),
    ]))
    hasil = asyncio.run(llm_analyzer.analisis_dengan_llm(["bus telat", "parkir liar"]))
    assert hasil[0]["Metode"] == "LLM"
    assert hasil[1] == {"Kategori": "Lahan Parkir", "Sentimen": "Negatif", "Metode": "Kata Kunci (fallback)",
                        "Keluhan": "", "Sebab": "", "Akibat": "", "Emosi": "", "Keluhan Lengkap": ""}


def test_gangguan_sementara_pakai_kata_kunci(monkeypatch):
    """Error jaringan/server: kelompok itu memakai kata kunci, proses tetap jalan."""
    def jawab(kw):
        raise _galat(anthropic.APIConnectionError)
    _pasang(monkeypatch, jawab)
    hasil = asyncio.run(llm_analyzer.analisis_dengan_llm(["bus telat"]))
    assert hasil[0]["Metode"] == "Kata Kunci (fallback)"


@pytest.mark.parametrize("kelas", [
    anthropic.AuthenticationError, anthropic.PermissionDeniedError, anthropic.BadRequestError,
])
def test_masalah_akun_menghentikan_proses(monkeypatch, kelas):
    """Key salah / kredit habis tidak boleh diam-diam diganti kata kunci."""
    def jawab(kw):
        raise _galat(kelas)
    _pasang(monkeypatch, jawab)
    with pytest.raises(kelas):
        asyncio.run(llm_analyzer.analisis_dengan_llm(["bus telat"]))


def test_dikirim_per_kelompok(monkeypatch):
    """45 teks dikirim dalam 3 request (20 + 20 + 5)."""
    def jawab(kw):
        jumlah = kw["messages"][0]["content"].count("\n\n") + 1
        return _respons_tool([
            _item(n, "angkutan_umum", "netral") for n in range(1, jumlah + 1)
        ])
    klien = _pasang(monkeypatch, jawab)
    hasil = asyncio.run(llm_analyzer.analisis_dengan_llm([f"bus {i}" for i in range(45)]))
    assert klien.jumlah_panggilan == 3
    assert all(h["Metode"] == "LLM" for h in hasil)


@pytest.mark.parametrize("kunci, tersedia", [
    ("", False), ("sk-ant-isi-key-anda-di-sini", False), ("sk-ant-api03-abc", True),
])
def test_api_key_tersedia(monkeypatch, kunci, tersedia):
    """Placeholder di .env tidak dianggap sebagai key yang valid."""
    monkeypatch.setenv("LLM_API_KEY", kunci)
    assert llm_analyzer.api_key_tersedia() is tersedia
