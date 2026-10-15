"""Test endpoint FastAPI: upload, classify, export, dan kebersihan file sementara."""

import io

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from aplikasi_web import main

CSV_CONTOH = (
    "no,isi\n"
    "1,Bus transjakarta selalu telat\n"
    "2,Parkir liar dan jukir nakal\n"
    "3,Jalan macet parah\n"
).encode("utf-8")


@pytest.fixture
def klien():
    """Klien uji dengan sesi bersih."""
    main.SESI.clear()
    return TestClient(main.app)


def _upload(klien, isi=CSV_CONTOH, nama="data.csv"):
    """Upload CSV dan kembalikan JSON respons."""
    return klien.post("/upload", files={"file": (nama, io.BytesIO(isi), "text/csv")})


def test_upload_mengembalikan_kolom_dan_preview(klien):
    """Upload CSV valid mengembalikan id sesi, kolom, dan preview."""
    data = _upload(klien).json()
    assert data["kolom"] == ["no", "isi"]
    assert data["jumlah_baris"] == 3
    assert len(data["preview"]) == 3


def test_upload_latin1(klien):
    """CSV ber-encoding latin-1 tetap terbaca."""
    data = _upload(klien, "no,isi\n1,caf\xe9 macet\n".encode("latin-1")).json()
    assert data["jumlah_baris"] == 1


@pytest.mark.parametrize("isi, nama", [(CSV_CONTOH, "data.txt"), (b"no,isi\n", "kosong.csv")])
def test_upload_ditolak(klien, isi, nama):
    """File bukan .csv atau tanpa baris data ditolak."""
    assert _upload(klien, isi, nama).status_code == 400


def test_upload_terlalu_besar(klien, monkeypatch):
    """File melebihi batas ukuran ditolak dengan 413."""
    monkeypatch.setattr(main, "UKURAN_MAKS_UPLOAD", 10)
    assert _upload(klien).status_code == 413


def test_classify_kata_kunci(klien):
    """Mode kata kunci memberi label dan ringkasan."""
    id_sesi = _upload(klien).json()["id"]
    data = klien.post("/classify", json={"id": id_sesi, "kolom_keluhan": "isi"}).json()
    assert [b["Kategori"] for b in data["hasil"]] == ["Angkutan Umum", "Lahan Parkir", "Manajemen Lalu Lintas"]
    assert all(b["Metode"] == "Kata Kunci" for b in data["hasil"])
    assert data["jumlah_fallback"] == 0


def test_classify_llm(klien, monkeypatch):
    """Mode llm memanggil analisis_dengan_llm (ditiru)."""
    async def palsu(daftar):
        return [{"Kategori": "Tidak Relevan", "Sentimen": "Netral", "Metode": "LLM"} for _ in daftar]
    monkeypatch.setattr(main, "api_key_tersedia", lambda: True)
    monkeypatch.setattr(main, "analisis_dengan_llm", palsu)
    id_sesi = _upload(klien).json()["id"]
    data = klien.post("/classify", json={"id": id_sesi, "kolom_keluhan": "isi", "metode": "llm"}).json()
    assert data["ringkasan_kategori"] == {"Tidak Relevan": 3}


def test_llm_menerima_teks_bersih(klien, monkeypatch):
    """Teks dinetralkan sebelum dikirim ke LLM; teks asli tetap di Isi Keluhan."""
    diterima = []
    async def palsu(daftar):
        diterima.extend(daftar)
        return [{"Kategori": "Angkutan Umum", "Sentimen": "Negatif", "Metode": "LLM"} for _ in daftar]
    monkeypatch.setattr(main, "api_key_tersedia", lambda: True)
    monkeypatch.setattr(main, "analisis_dengan_llm", palsu)
    csv_emoji = "isi\n@dishubjakarta bus telat 😡 https://t.co/x\n".encode("utf-8")
    id_sesi = _upload(klien, csv_emoji).json()["id"]
    data = klien.post("/classify", json={"id": id_sesi, "kolom_keluhan": "isi", "metode": "llm"}).json()
    assert diterima == ["bus telat"]
    assert data["hasil"][0]["Isi Keluhan"] == "@dishubjakarta bus telat 😡 https://t.co/x"
    assert data["hasil"][0]["Teks Bersih"] == "bus telat"


def test_classify_llm_tanpa_key(klien, monkeypatch):
    """Mode llm tanpa API key ditolak dengan pesan jelas."""
    monkeypatch.setattr(main, "api_key_tersedia", lambda: False)
    id_sesi = _upload(klien).json()["id"]
    res = klien.post("/classify", json={"id": id_sesi, "kolom_keluhan": "isi", "metode": "llm"})
    assert res.status_code == 400
    assert "LLM_API_KEY" in res.json()["detail"]


@pytest.mark.parametrize("isi_json, kode", [
    ({"id": "tidak-ada", "kolom_keluhan": "isi"}, 404),
    ({"kolom_keluhan": "salah"}, 400),
    ({"kolom_keluhan": "isi", "metode": "ngawur"}, 400),
])
def test_classify_input_salah(klien, isi_json, kode):
    """Sesi tidak ada, kolom salah, atau metode tidak dikenal ditolak."""
    id_sesi = _upload(klien).json()["id"]
    isi_json = {"id": id_sesi, **isi_json}
    assert klien.post("/classify", json=isi_json).status_code == kode


def test_export_excel_dan_file_dihapus(klien):
    """Excel punya 2 sheet + kolom Metode, dan file sementara terhapus setelah dikirim."""
    id_sesi = _upload(klien).json()["id"]
    klien.post("/classify", json={"id": id_sesi, "kolom_keluhan": "isi"})
    res = klien.get(f"/export/{id_sesi}")
    assert res.status_code == 200

    buku = load_workbook(io.BytesIO(res.content))
    assert buku.sheetnames == ["Hasil Klasifikasi", "Ringkasan"]
    header = [sel.value for sel in buku["Hasil Klasifikasi"][1]]
    assert header == ["no", "isi", "Isi Keluhan", "Teks Bersih", "Kategori", "Sentimen", "Metode"]
    assert not list(main.DIREKTORI_UPLOAD.glob("hasil_*.xlsx"))

    ringkasan = [[sel.value for sel in baris] for baris in buku["Ringkasan"].iter_rows()]
    assert ringkasan[0][0] == "Ringkasan per Kategori"
    assert ringkasan[1] == ["Kategori", "Jumlah", "Persentase"]
    judul = [b[0] for b in ringkasan]
    assert "Ringkasan per Sentimen" in judul and "Ringkasan per Metode" in judul
    assert ["Total Data", 3] in [b[:2] for b in ringkasan]
    assert "Seleksi Data" in judul


@pytest.mark.parametrize("kategori, kode", [("Lahan Parkir", 200), ("Tidak Relevan", 200), ("Ngawur", 400)])
def test_export_per_kategori(klien, kategori, kode):
    """Filter export menerima 3 kategori + Tidak Relevan saja."""
    id_sesi = _upload(klien).json()["id"]
    klien.post("/classify", json={"id": id_sesi, "kolom_keluhan": "isi"})
    assert klien.get(f"/export/{id_sesi}", params={"kategori": kategori}).status_code == kode


def test_export_sebelum_classify(klien):
    """Export sebelum klasifikasi ditolak."""
    id_sesi = _upload(klien).json()["id"]
    assert klien.get(f"/export/{id_sesi}").status_code == 404
