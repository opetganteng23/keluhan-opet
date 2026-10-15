"""Seleksi data sebelum analisis: buang teks kosong, akun resmi/media, dan duplikat.

Data yang disaring tidak dihapus, tetapi dikembalikan bersama alasannya supaya
bisa dilaporkan (sheet "Data Disaring" di Excel).
"""

from persiapan_data.pembersih import bersihkan_teks

# Akun institusi & media: isinya pengumuman/berita, bukan keluhan warga.
# Silakan tambah/kurangi sesuai pertimbangan penelitian (huruf besar/kecil diabaikan).
AKUN_RESMI_MEDIA: set[str] = {
    # Institusi pemerintah & operator transportasi
    "dishubjakarta", "pt_transjakarta", "ptjasamarga", "tmcpoldametro",
    "dkijakarta", "jakartagoid", "beritajakarta", "mrtjakarta", "tfjakarta",
    "p_kelapagading", "senkomsoroja", "rute_t11_t12",
    # Media berita
    "kompascom", "hariankompas", "kompasotomotif", "radioelshinta", "media_akurat",
    "akuratco", "mediaindonesia", "metro_tv", "kumparan", "sindonews", "antaranews",
    "pikiran_rakyat", "beritasatu", "inilahdotcom", "sonorafm92", "kontannews",
    "detikcom", "idntimes", "tvonenews", "officialinews_", "cnnindonesia",
    "tempodotco", "liputan6dotcom", "okezonenews", "republikaonline",
}

ALASAN_KOSONG = "Kosong setelah dibersihkan"
ALASAN_AKUN = "Akun resmi/media"
ALASAN_DUPLIKAT = "Duplikat"

KANDIDAT_KOLOM_AKUN = ("user", "username", "akun", "screen_name", "author")


def cari_kolom_akun(kolom: list[str]) -> str | None:
    """Mencari kolom nama akun di CSV (mis. 'user'); None jika tidak ada."""
    peta = {k.lower().strip(): k for k in kolom}
    for kandidat in KANDIDAT_KOLOM_AKUN:
        if kandidat in peta:
            return peta[kandidat]
    return None


def seleksi_data(
    baris: list[dict], kolom_keluhan: str, kolom_akun: str | None
) -> tuple[list[int], dict[int, str]]:
    """Mengembalikan (indeks baris yang lolos, {indeks baris disaring: alasan}).

    Urutan pemeriksaan: kosong -> akun resmi/media -> duplikat (yang pertama dipertahankan).
    Duplikat dibandingkan dari teks bersih huruf kecil.
    """
    lolos: list[int] = []
    disaring: dict[int, str] = {}
    sudah_ada: set[str] = set()

    for indeks, data in enumerate(baris):
        bersih = bersihkan_teks(data.get(kolom_keluhan, "") or "").lower()
        akun = (data.get(kolom_akun, "") or "").strip().lstrip("@").lower() if kolom_akun else ""

        if not bersih:
            disaring[indeks] = ALASAN_KOSONG
        elif akun in AKUN_RESMI_MEDIA:
            disaring[indeks] = ALASAN_AKUN
        elif bersih in sudah_ada:
            disaring[indeks] = ALASAN_DUPLIKAT
        else:
            sudah_ada.add(bersih)
            lolos.append(indeks)
    return lolos, disaring
