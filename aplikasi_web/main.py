"""Entry point FastAPI untuk aplikasi klasifikasi keluhan transportasi publik."""

import csv
import io
import logging
import time
import uuid
from pathlib import Path

import anthropic

from fastapi import BackgroundTasks, FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from metode_2_kata_kunci.classifier import KATEGORI_KEYWORDS, klasifikasi_keluhan
from persiapan_data.pembersih import bersihkan_teks
from persiapan_data.seleksi import cari_kolom_akun, seleksi_data
from aplikasi_web.exporter import buat_file_excel
from metode_1_llm.llm_analyzer import KATEGORI_TIDAK_RELEVAN, analisis_dengan_llm, api_key_tersedia
from metode_2_kata_kunci.sentiment import analisis_sentimen

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("klasifikasi_keluhan")

app = FastAPI(title="Klasifikasi Keluhan Transportasi")


@app.middleware("http")
async def cegah_cache_browser(request, call_next):
    """Mencegah browser menyimpan cache static file (index.html/JS) yang basi."""
    response = await call_next(request)
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
    return response

DIREKTORI_UPLOAD = Path(__file__).parent / "uploads"
DIREKTORI_UPLOAD.mkdir(exist_ok=True)

# Penyimpanan sesi sementara di memori (id -> data hasil parsing/klasifikasi)
SESI: dict[str, dict] = {}
UMUR_SESI_DETIK = 60 * 60
UKURAN_MAKS_UPLOAD = 50 * 1024 * 1024  # 50 MB


def _bersihkan_sesi_kedaluwarsa() -> None:
    """Menghapus sesi di memori yang lebih tua dari UMUR_SESI_DETIK."""
    batas = time.time() - UMUR_SESI_DETIK
    for id_sesi in [i for i, s in SESI.items() if s["dibuat"] < batas]:
        del SESI[id_sesi]
        logger.info("Sesi kedaluwarsa dihapus, id_sesi=%s", id_sesi)


class PermintaanKlasifikasi(BaseModel):
    """Payload permintaan klasifikasi: id sesi upload dan nama kolom isi keluhan."""

    id: str
    kolom_keluhan: str
    metode: str = "kata_kunci"  # "kata_kunci" atau "llm"
    seleksi: bool = True  # buang kosong, akun resmi/media, dan duplikat sebelum analisis


def _baca_csv(isi_bytes: bytes) -> tuple[list[str], list[dict]]:
    """Membaca isi CSV dari bytes, coba UTF-8 dulu lalu fallback ke latin-1."""
    for enkoding in ("utf-8-sig", "latin-1"):
        try:
            teks = isi_bytes.decode(enkoding)
            pembaca = csv.DictReader(io.StringIO(teks))
            baris = list(pembaca)
            kolom = pembaca.fieldnames or []
            return list(kolom), baris
        except UnicodeDecodeError:
            continue
    raise HTTPException(status_code=400, detail="Gagal membaca encoding file CSV")


@app.get("/status-llm")
async def status_llm():
    """Memberi tahu frontend apakah mode LLM bisa dipakai (API key sudah diisi)."""
    return {"tersedia": api_key_tersedia()}


@app.post("/upload")
async def upload_csv(file: UploadFile = File(...)):
    """Menerima upload CSV, mem-parsing isinya, lalu mengembalikan preview dan daftar kolom."""
    if not file.filename or not file.filename.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="File harus berformat .csv")

    isi_bytes = await file.read(UKURAN_MAKS_UPLOAD + 1)
    if len(isi_bytes) > UKURAN_MAKS_UPLOAD:
        raise HTTPException(status_code=413, detail="File terlalu besar (maksimal 50 MB)")

    # Nama acak saja: nama file dari pengguna tidak dipakai sebagai path
    path_sementara = DIREKTORI_UPLOAD / f"{uuid.uuid4().hex}.csv"
    path_sementara.write_bytes(isi_bytes)

    try:
        kolom, baris = _baca_csv(isi_bytes)
    finally:
        path_sementara.unlink(missing_ok=True)

    if not baris:
        raise HTTPException(status_code=400, detail="File CSV kosong atau tidak valid")

    _bersihkan_sesi_kedaluwarsa()
    id_sesi = uuid.uuid4().hex
    SESI[id_sesi] = {"kolom": kolom, "baris": baris, "dibuat": time.time()}

    logger.info("Upload berhasil, id_sesi=%s, jumlah_baris=%d", id_sesi, len(baris))

    return {
        "id": id_sesi,
        "kolom": kolom,
        "preview": baris[:5],
        "jumlah_baris": len(baris),
    }


@app.post("/classify")
async def classify(permintaan: PermintaanKlasifikasi):
    """Memproses klasifikasi kategori dan sentimen untuk seluruh baris pada suatu sesi."""
    sesi = SESI.get(permintaan.id)
    if sesi is None:
        raise HTTPException(status_code=404, detail="Sesi tidak ditemukan, silakan upload ulang")

    if permintaan.kolom_keluhan not in sesi["kolom"]:
        raise HTTPException(status_code=400, detail="Kolom isi keluhan tidak valid")

    if permintaan.metode not in ("kata_kunci", "llm"):
        raise HTTPException(status_code=400, detail="Metode harus 'kata_kunci' atau 'llm'")
    if permintaan.metode == "llm" and not api_key_tersedia():
        raise HTTPException(status_code=400, detail="LLM_API_KEY belum diisi di file .env")

    # Seleksi data: buang kosong / akun resmi-media / duplikat (disimpan dengan alasannya)
    semua_baris = sesi["baris"]
    if permintaan.seleksi:
        kolom_akun = cari_kolom_akun(sesi["kolom"])
        indeks_lolos, alasan_disaring = seleksi_data(semua_baris, permintaan.kolom_keluhan, kolom_akun)
    else:
        indeks_lolos, alasan_disaring = list(range(len(semua_baris))), {}
    baris_lolos = [semua_baris[i] for i in indeks_lolos]
    sesi["disaring"] = [
        {**semua_baris[i], "Alasan Disaring": alasan} for i, alasan in sorted(alasan_disaring.items())
    ]
    ringkasan_seleksi = {"Data awal": len(semua_baris)}
    for alasan in alasan_disaring.values():
        ringkasan_seleksi[alasan] = ringkasan_seleksi.get(alasan, 0) + 1
    ringkasan_seleksi["Dianalisis"] = len(baris_lolos)
    sesi["ringkasan_seleksi"] = ringkasan_seleksi

    daftar_teks = [baris.get(permintaan.kolom_keluhan, "") or "" for baris in baris_lolos]
    # Pra-pemrosesan: kedua metode menganalisis teks yang sudah dinetralkan
    daftar_bersih = [bersihkan_teks(teks) for teks in daftar_teks]
    if permintaan.metode == "llm":
        try:
            label = await analisis_dengan_llm(daftar_bersih)
        except anthropic.APIStatusError as galat:
            logger.error("Analisis LLM dihentikan: %s", galat)
            pesan = galat.body.get("error", {}).get("message", str(galat)) if isinstance(galat.body, dict) else str(galat)
            raise HTTPException(status_code=502, detail=f"API LLM menolak permintaan: {pesan}")
    else:
        label = [
            {
                "Kategori": klasifikasi_keluhan(t),
                "Sentimen": analisis_sentimen(t),
                "Metode": "Kata Kunci" if t else "Kosong",
            }
            for t in daftar_bersih
        ]

    hasil: list[dict] = []
    for baris, teks_keluhan, teks_bersih, item in zip(baris_lolos, daftar_teks, daftar_bersih, label):
        baris_hasil = dict(baris)
        baris_hasil["Isi Keluhan"] = teks_keluhan
        baris_hasil["Teks Bersih"] = teks_bersih
        baris_hasil.update(item)
        hasil.append(baris_hasil)

    sesi["hasil"] = hasil

    jumlah_kategori: dict[str, int] = {}
    jumlah_sentimen: dict[str, int] = {}
    for baris in hasil:
        jumlah_kategori[baris["Kategori"]] = jumlah_kategori.get(baris["Kategori"], 0) + 1
        jumlah_sentimen[baris["Sentimen"]] = jumlah_sentimen.get(baris["Sentimen"], 0) + 1

    jumlah_fallback = sum(1 for b in hasil if b["Metode"] == "Kata Kunci (fallback)")
    logger.info(
        "Klasifikasi selesai, id_sesi=%s, metode=%s, jumlah_baris=%d, fallback=%d",
        permintaan.id, permintaan.metode, len(hasil), jumlah_fallback,
    )

    return {
        "id": permintaan.id,
        "hasil": hasil,
        "ringkasan_kategori": jumlah_kategori,
        "ringkasan_sentimen": jumlah_sentimen,
        "metode": permintaan.metode,
        "jumlah_fallback": jumlah_fallback,
        "ringkasan_seleksi": ringkasan_seleksi,
    }


@app.get("/export/{id_sesi}")
async def export_excel(id_sesi: str, latar: BackgroundTasks, kategori: str | None = None):
    """Membuat dan mengirim file Excel hasil klasifikasi, opsional difilter per kategori."""
    sesi = SESI.get(id_sesi)
    if sesi is None or "hasil" not in sesi:
        raise HTTPException(status_code=404, detail="Hasil klasifikasi belum tersedia untuk sesi ini")

    baris_hasil = sesi["hasil"]
    nama_file = "hasil_klasifikasi_keluhan.xlsx"
    # Data disaring & ringkasan seleksi hanya disertakan pada export semua data
    disaring = sesi.get("disaring", [])
    ringkasan_seleksi = sesi.get("ringkasan_seleksi", {})

    if kategori:
        if kategori not in KATEGORI_KEYWORDS and kategori != KATEGORI_TIDAK_RELEVAN:
            raise HTTPException(status_code=400, detail="Kategori tidak dikenal")
        baris_hasil = [baris for baris in baris_hasil if baris.get("Kategori") == kategori]
        nama_file = f"hasil_{kategori.lower().replace(' ', '_')}.xlsx"
        disaring, ringkasan_seleksi = [], {}

    path_excel = DIREKTORI_UPLOAD / f"hasil_{uuid.uuid4().hex}.xlsx"
    buat_file_excel(sesi["kolom"], baris_hasil, path_excel, disaring, ringkasan_seleksi)
    # Hapus file setelah terkirim, supaya tidak menumpuk di server
    latar.add_task(path_excel.unlink, missing_ok=True)

    logger.info("Export Excel selesai, id_sesi=%s, kategori=%s", id_sesi, kategori or "semua")

    return FileResponse(
        path=path_excel,
        filename=nama_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


app.mount("/", StaticFiles(directory=Path(__file__).parent / "static", html=True), name="static")
