"""Analisis kategori dan sentimen keluhan memakai API LLM (penyedia: Anthropic).

Keluhan dikirim per kelompok (UKURAN_KELOMPOK teks per request) supaya hemat
biaya, dan beberapa kelompok diproses paralel. Kelompok yang gagal karena
gangguan sementara (jaringan, server sibuk) dianalisis ulang dengan metode kata
kunci. Masalah akun (key salah, kredit habis) menghentikan proses.
"""

import asyncio
import json
import logging
import os
from pathlib import Path

import anthropic
from dotenv import load_dotenv

from metode_2_kata_kunci.classifier import klasifikasi_keluhan
from metode_2_kata_kunci.sentiment import analisis_sentimen

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

logger = logging.getLogger("klasifikasi_keluhan")

MODEL = "claude-haiku-4-5"
UKURAN_KELOMPOK = 20
MAKS_PARALEL = 4
MAKS_KARAKTER_TEKS = 1000

KATEGORI_TIDAK_RELEVAN = "Tidak Relevan"

# Kode kategori dari LLM -> nama kategori yang ditampilkan
PETA_KATEGORI: dict[str, str] = {
    "angkutan_umum": "Angkutan Umum",
    "parkir": "Lahan Parkir",
    "manajemen_lalu_lintas": "Manajemen Lalu Lintas",
    "tidak_relevan": KATEGORI_TIDAK_RELEVAN,
}
PETA_SENTIMEN: dict[str, str] = {"positif": "Positif", "negatif": "Negatif", "netral": "Netral"}

# Versi 1 (dipakai untuk analisis 15.000 tweet pertama, 13-10-2026); diarsipkan
PROMPT_SISTEM_V1 = """\
Kamu adalah asisten yang menganalisis teks warga Jakarta (tweet maupun ulasan \
aplikasi) untuk penelitian tugas akhir tentang kinerja NYATA Dinas Perhubungan \
(Dishub) DKI Jakarta di LAPANGAN. Batasan masalah penelitian ini hanya mencakup \
3 topik pelayanan nyata:

1. manajemen_lalu_lintas - kemacetan, rambu, lampu lalu lintas, rekayasa jalan, \
ganjil genap, kinerja petugas Dishub di jalan, serta fasilitas jalan umum: \
trotoar, lampu penerangan jalan (PJU), JPO, zebra cross, marka.
2. parkir - perparkiran liar, juru parkir, tarif/retribusi parkir, ketersediaan \
lahan parkir.
3. angkutan_umum - kualitas layanan TransJakarta/JakLingko/Mikrotrans/angkot/ojek/halte \
di lapangan: \
kondisi armada, perilaku sopir/petugas, ketepatan waktu, kepadatan penumpang, \
keamanan/kenyamanan perjalanan, kebijakan tarif/rute.

Tandai "tidak_relevan" walau teks menyebut kata kunci di atas, jika:
- Isinya review APLIKASI (bug, crash, login, tampilan, fitur pelacak bus, top up \
saldo). Itu soal software, bukan pelayanan di lapangan.
- Pujian/keluhan generik tanpa pengalaman nyata di jalan/transportasi \
("aplikasinya bagus", "makasih infonya").
- Topik di luar 3 kategori (banjir, sampah, berita umum, iklan, candaan, obrolan \
lain), termasuk layanan pemerintah DKI lain di luar Dishub.
- Pengumuman/informasi resmi tanpa keluhan atau pendapat warga.

Contoh RELEVAN (angkutan_umum): "supir busnya galak banget", "bus telat 1 jam", \
"halte kotor dan rusak".
Contoh TIDAK RELEVAN meski menyebut TransJakarta: "aplikasinya keren buat cek \
posisi bus", "app wajib buat naik TJ".

Untuk setiap teks tentukan:
- kategori: "manajemen_lalu_lintas", "parkir", "angkutan_umum", atau "tidak_relevan".
- sentimen: sikap warga terhadap pelayanan yang dibahas: "positif", "negatif", \
atau "netral". Jika kategori "tidak_relevan", isi "netral".
Tangkap sarkasme sesuai maksud sebenarnya. Nilai harus murni berdasarkan isi \
teks, jangan menebak di luar konteks.

Kamu akan menerima beberapa teks bernomor. Panggil tool kirim_analisis satu \
kali dengan tepat satu hasil untuk SETIAP nomor.
"""

PETA_EMOSI: dict[str, str] = {
    "lelah": "Lelah", "frustasi": "Frustasi", "kesal": "Kesal", "kecewa": "Kecewa",
    "marah": "Marah", "takut": "Takut/Khawatir", "tidak_ada": "Tidak Ada",
}
TIDAK_DISEBUTKAN = "tidak disebutkan"

# Versi 2 (13-10-2026): ruang lingkup = keluhan sehari-hari warganet tentang perkotaan Jakarta,
# ditambah anotasi keluhan, sebab, akibat, dan emosi.
# Versi 2 (kalibrasi 100 tweet: kategori 71%, keluhan 87%); diarsipkan
PROMPT_SISTEM_V2 = """\
Kamu menganalisis tweet warganet untuk penelitian tentang KELUHAN SEHARI-HARI warga \
mengenai kondisi transportasi perkotaan JAKARTA. Ruang lingkupnya 3 topik:

1. manajemen_lalu_lintas - kondisi jalan dan lalu lintas kota Jakarta: kemacetan, \
jalan rusak/berlubang, galian, rambu, lampu lalu lintas, marka, rekayasa lalu lintas, \
ganjil genap, trotoar, lampu penerangan jalan (PJU), JPO, zebra cross, perilaku \
pengendara di jalan kota, petugas Dishub di jalan.
2. parkir - parkir liar, juru parkir (jukir/tukang parkir), tarif/pungutan parkir, \
ketersediaan lahan parkir, kendaraan diderek/diangkut.
3. angkutan_umum - layanan TransJakarta/JakLingko/Mikrotrans/angkot/bus kota/ojek \
dan halte di Jakarta: armada, sopir/petugas, waktu tunggu, kepadatan, kenyamanan, \
keamanan, tarif, rute.

Tandai "tidak_relevan" jika: review aplikasi (bug, login, top up, fitur), iklan, \
candaan, berita/pengumuman tanpa pengalaman atau pendapat warga, cerita pribadi yang \
hanya menyebut transportasi tanpa menilai kondisinya, jalan tol/antarkota, kecelakaan \
yang hanya diberitakan, serta topik di luar 3 topik di atas (banjir, sampah, dll).

Untuk setiap teks tentukan:
- kategori: "manajemen_lalu_lintas", "parkir", "angkutan_umum", atau "tidak_relevan".
- sentimen: sikap penulis terhadap kondisi yang dibahas: "positif", "negatif", atau \
"netral". Tangkap sarkasme sesuai maksud sebenarnya ("hebat, nunggu sejam" = negatif).
- keluhan: true jika penulis MENGELUHKAN kondisi pada salah satu dari 3 topik \
(bukan sekadar bertanya, memberi info, atau memuji); false jika tidak.
- sebab: penyebab masalah yang dikeluhkan, ringkas (maks. 8 kata), sesuai isi teks.
- akibat: kondisi buruk/dampak yang dikeluhkan, ringkas (maks. 8 kata), sesuai isi teks.
- emosi: perasaan yang paling tampak: "lelah", "frustasi", "kesal", "kecewa", \
"marah", "takut", atau "tidak_ada".

Pola keluhan: AKIBAT terjadi karena SEBAB. Contoh:
- "Jalanan macet nih gara-gara galian" -> sebab: galian jalan; akibat: jalanan macet.
- "Lampu merah ini macet banget, lampunya kelamaan" -> sebab: durasi lampu merah \
terlalu lama; akibat: macet parah di lampu merah.
- "Parkir di sini banyak jukirnya, matok harga gak bener" -> sebab: banyak jukir \
liar; akibat: tarif parkir tidak wajar.
- "Udah sejam kejebak macet, telat ngantor lagi, capek" -> sebab: "tidak disebutkan"; \
akibat: terjebak macet sejam, telat kerja; emosi: lelah.
Jika sebab atau akibat tidak ada di teks, isi "tidak disebutkan" (jangan mengarang). \
Jika keluhan false, isi sebab dan akibat "tidak disebutkan" dan emosi "tidak_ada". \
Jika kategori "tidak_relevan", keluhan harus false dan sentimen "netral".

Kamu akan menerima beberapa teks bernomor. Panggil tool kirim_analisis satu \
kali dengan tepat satu hasil untuk SETIAP nomor.
"""

# Versi 3.1 (14-10-2026): + aturan tidak_relevan (sebutan sambil lalu, ensiklopedis, kecelakaan) & topik = sumber masalah.
# Versi 3 (14-10-2026): kategori = TOPIK (terpisah dari keluhan); keluhan termasuk tersirat/sarkas;
# panduan emosi diperjelas. Disusun dari pola kalibrasi, contoh di bawah dibuat baru (bukan tweet kalibrasi).
PROMPT_SISTEM = """\
Kamu menganalisis tweet warganet untuk penelitian tentang keluhan sehari-hari warga \
mengenai transportasi perkotaan JAKARTA. Setiap tweet dinilai dalam DUA lapis terpisah:
(1) TOPIK-nya (kategori), lalu (2) apakah penulis MENGELUH (keluhan).

LAPIS 1 - KATEGORI = topik tweet, APA PUN jenis tweetnya (keluhan, pujian, pertanyaan, \
info, pengumuman, berita, opini, candaan tentang topik itu):
- manajemen_lalu_lintas: kemacetan dan lalu lintas di Jakarta, kondisi jalan (rusak, \
berlubang, galian), rambu, lampu lalu lintas, marka, rekayasa lalu lintas, ganjil genap, \
trotoar, lampu jalan/PJU, JPO, zebra cross, perilaku pengendara, petugas Dishub di jalan.
- parkir: parkir liar, juru/tukang parkir, tarif atau pungutan parkir, lahan parkir, \
kendaraan diderek/diangkut, termasuk cerita atau hitung-hitungan tentang bisnis parkir.
- angkutan_umum: TransJakarta, JakLingko/Mikrotrans, angkot, bus kota, ojek, halte di \
Jakarta: layanan, armada, sopir/petugas, rute, tarif, jadwal, penutupan/perubahan halte \
atau rute, kebijakan, kepadatan, kenyamanan, keamanan.
- tidak_relevan: HANYA jika tweet tidak membahas salah satu topik di atas di Jakarta, \
misalnya kota lain, topik lain (banjir, kos, belanja, hiburan), atau kata transportasi \
hanya sambil lalu tanpa dibahas (mis. "kosku dekat halte" saat membahas harga kos), \nseruan/candaan singkat yang hanya menyebut nama moda ("gas naik busway!"), penjelasan \numum/ensiklopedis tentang sejarah suatu moda, dan berita kecelakaan/tabrakan/kriminal.
Pengumuman, berita layanan, info, dan opini TETAP masuk kategori topiknya, bukan tidak_relevan.
Jika beberapa topik muncul, pilih topik SUMBER MASALAHNYA (mis. bus tidak bisa berhenti \nkarena mobil parkir sembarangan -> parkir; bus telat karena jalan macet -> manajemen_lalu_lintas \nbila yang dikeluhkan macetnya, angkutan_umum bila yang dikeluhkan layanan busnya).

LAPIS 2 - KELUHAN = true jika penulis menyatakan ketidakpuasan, beban, atau penderitaan \
terhadap kondisi pada topik itu, termasuk yang TERSIRAT, pasrah, bercanda, atau sarkas, \
misalnya: "otw berangkat perang lawan macet lagi", "ilmu sabar level dewa: tiap hari \
lewat jalan ini", "mantap banget busnya, berdiri sambil goyang dua jam". Berita atau \
pengumuman tentang kejadian merugikan yang disertai kecaman juga dihitung keluhan.
keluhan = false untuk info/pengumuman netral, pertanyaan biasa, pujian, saran tanpa \
keluhan, dan semua tweet tidak_relevan.

Untuk setiap teks tentukan:
- kategori: "manajemen_lalu_lintas", "parkir", "angkutan_umum", atau "tidak_relevan".
- sentimen: sikap penulis terhadap topik: "positif", "negatif", atau "netral". Keluhan \
(termasuk sarkas) = negatif; info/pengumuman tanpa penilaian = netral.
- keluhan: true/false sesuai LAPIS 2.
- sebab: penyebab masalah yang dikeluhkan, ringkas (maks. 8 kata), sesuai isi teks.
- akibat: kondisi buruk/dampak yang dikeluhkan, ringkas (maks. 8 kata), sesuai isi teks.
- emosi (hanya jika keluhan; selain itu "tidak_ada"):
  "kesal" = jengkel/sebal secara umum (PILIHAN UTAMA untuk keluhan biasa);
  "lelah" = capek atau pasrah menghadapi beban rutin (macet harian, perjalanan panjang);
  "frustasi" = putus asa, merasa tidak ada jalan keluar;
  "marah" = kata kasar, makian, nada sangat keras;
  "kecewa" = harapan tidak terpenuhi, terutama terhadap kebijakan/pemerintah/operator;
  "takut" = khawatir akan keselamatan, keamanan, atau kerugian.

Pola keluhan: AKIBAT terjadi karena SEBAB. Contoh:
- "Jalanan macet nih gara-gara galian" -> sebab: galian jalan; akibat: jalanan macet.
- "Lampu merah ini macet banget, lampunya kelamaan" -> sebab: durasi lampu merah \
terlalu lama; akibat: macet parah di lampu merah.
- "Parkir di sini banyak jukirnya, matok harga gak bener" -> sebab: banyak jukir liar; \
akibat: tarif parkir tidak wajar.
- "Udah sejam kejebak macet, telat ngantor lagi, capek" -> sebab: "tidak disebutkan"; \
akibat: terjebak macet sejam, telat kerja; emosi: lelah.
Jika sebab atau akibat tidak ada di teks, isi "tidak disebutkan" (jangan mengarang). \
Jika keluhan false, isi sebab dan akibat "tidak disebutkan".

Kamu akan menerima beberapa teks bernomor. Panggil tool kirim_analisis satu \
kali dengan tepat satu hasil untuk SETIAP nomor.
"""

TOOL_ANALISIS = {
    "name": "kirim_analisis",
    "description": "Kirim hasil analisis untuk semua teks bernomor.",
    "input_schema": {
        "type": "object",
        "properties": {
            "hasil": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "nomor": {"type": "integer"},
                        "kategori": {"type": "string", "enum": list(PETA_KATEGORI)},
                        "sentimen": {"type": "string", "enum": list(PETA_SENTIMEN)},
                        "keluhan": {"type": "boolean"},
                        "sebab": {"type": "string"},
                        "akibat": {"type": "string"},
                        "emosi": {"type": "string", "enum": list(PETA_EMOSI)},
                    },
                    "required": ["nomor", "kategori", "sentimen", "keluhan", "sebab", "akibat", "emosi"],
                    "additionalProperties": False,
                },
            }
        },
        "required": ["hasil"],
        "additionalProperties": False,
    },
    "strict": True,
}


# Kolom keluhan untuk baris yang tidak dianalisis LLM (fallback / teks kosong)
KOLOM_KELUHAN_KOSONG = {"Keluhan": "", "Sebab": "", "Akibat": "", "Emosi": "", "Keluhan Lengkap": ""}


def api_key_tersedia() -> bool:
    """Mengecek apakah LLM_API_KEY sudah diisi (bukan placeholder)."""
    kunci = os.environ.get("LLM_API_KEY", "")
    return kunci.startswith("sk-ant-") and "isi-key" not in kunci


def _analisis_kata_kunci(teks: str) -> dict:
    """Fallback: analisis satu teks dengan metode kata kunci."""
    return {
        "Kategori": klasifikasi_keluhan(teks),
        "Sentimen": analisis_sentimen(teks),
        "Metode": "Kata Kunci (fallback)",
        **KOLOM_KELUHAN_KOSONG,
    }


def susun_pesan(daftar_teks: list[str]) -> str:
    """Teks bernomor yang dikirim ke LLM untuk satu kelompok."""
    return "\n\n".join(
        f"[{nomor}] {teks[:MAKS_KARAKTER_TEKS]}" for nomor, teks in enumerate(daftar_teks, start=1)
    )


def parameter_request(daftar_teks: list[str]) -> dict:
    """Parameter Messages API untuk satu kelompok (sama persis untuk mode langsung & Batch API)."""
    return {
        "model": MODEL,
        "max_tokens": 4096,
        "temperature": 0,  # deterministik: hasil bisa diulang (penting untuk penelitian)
        "system": PROMPT_SISTEM,
        "tools": [TOOL_ANALISIS],
        "tool_choice": {"type": "tool", "name": "kirim_analisis"},
        "messages": [{"role": "user", "content": susun_pesan(daftar_teks)}],
    }


def baca_jawaban(blok_konten, daftar_teks: list[str]) -> tuple[list[dict], dict[int, dict]]:
    """Ubah blok konten jawaban LLM menjadi label per teks; nomor yang terlewat -> kata kunci."""
    hasil_per_nomor: dict[int, dict] = {}
    for blok in blok_konten:
        tipe = blok["type"] if isinstance(blok, dict) else blok.type
        nama = blok.get("name") if isinstance(blok, dict) else getattr(blok, "name", None)
        if tipe == "tool_use" and nama == "kirim_analisis":
            masukan = blok["input"] if isinstance(blok, dict) else blok.input
            for item in masukan.get("hasil", []):
                hasil_per_nomor[item["nomor"]] = item

    keluaran: list[dict] = []
    for nomor, teks in enumerate(daftar_teks, start=1):
        item = hasil_per_nomor.get(nomor)
        if item is None:
            logger.warning("LLM melewatkan teks nomor %d, pakai kata kunci", nomor)
            keluaran.append(_analisis_kata_kunci(teks))
            continue
        sebab = (item["sebab"] or TIDAK_DISEBUTKAN).strip()
        akibat = (item["akibat"] or TIDAK_DISEBUTKAN).strip()
        keluhan = bool(item["keluhan"]) and item["kategori"] != "tidak_relevan"
        lengkap = keluhan and TIDAK_DISEBUTKAN not in (sebab.lower(), akibat.lower())
        keluaran.append({
            "Kategori": PETA_KATEGORI[item["kategori"]],
            "Sentimen": PETA_SENTIMEN[item["sentimen"]],
            "Keluhan": "Ya" if keluhan else "Tidak",
            "Sebab": sebab if keluhan else TIDAK_DISEBUTKAN,
            "Akibat": akibat if keluhan else TIDAK_DISEBUTKAN,
            "Emosi": PETA_EMOSI[item["emosi"]],
            "Keluhan Lengkap": "Ya" if lengkap else "Tidak",
            "Metode": "LLM",
        })
    return keluaran, hasil_per_nomor


async def _analisis_kelompok(
    klien: anthropic.AsyncAnthropic,
    daftar_teks: list[str],
    semafor: asyncio.Semaphore,
    jejak: list[dict] | None = None,
) -> list[dict]:
    """Mengirim satu kelompok teks ke LLM dan mengembalikan hasil per teks."""
    param = parameter_request(daftar_teks)
    isi_pesan = param["messages"][0]["content"]
    suhu = param.pop("temperature")  # SDK 1.x: temperature dikirim lewat extra_body
    async with semafor:
        try:
            respons = await klien.messages.create(**param, extra_body={"temperature": suhu})
        except (anthropic.AuthenticationError, anthropic.PermissionDeniedError, anthropic.BadRequestError):
            # Masalah akun/kredit/permintaan: hentikan, jangan diam-diam campur metode
            raise
        except anthropic.APIError as galat:
            logger.warning("LLM gagal untuk satu kelompok, pakai kata kunci: %s", galat)
            return [_analisis_kata_kunci(teks) for teks in daftar_teks]

    keluaran, hasil_per_nomor = baca_jawaban(respons.content, daftar_teks)

    if jejak is not None:
        # Bukti cara kerja: apa yang dikirim dan apa jawaban mentah model
        jejak.append({
            "id_request": getattr(respons, "_request_id", "") or respons.id,
            "model": respons.model,
            "token_input": respons.usage.input_tokens,
            "token_output": respons.usage.output_tokens,
            "pesan_dikirim": isi_pesan,
            "jawaban_mentah": json.dumps(
                [hasil_per_nomor[n] for n in sorted(hasil_per_nomor)], ensure_ascii=False
            ),
        })

    return keluaran


async def analisis_dengan_llm(daftar_teks: list[str], jejak: list[dict] | None = None) -> list[dict]:
    """Menganalisis seluruh teks dengan LLM; urutan hasil sama dengan input.

    Teks kosong langsung diberi label Tidak Relevan / Netral tanpa memanggil API.
    Jika `jejak` diberikan, setiap request dicatat ke dalamnya (bukti cara kerja).
    """
    hasil: list[dict | None] = [None] * len(daftar_teks)
    indeks_berisi = [i for i, teks in enumerate(daftar_teks) if teks.strip()]
    for i, teks in enumerate(daftar_teks):
        if not teks.strip():
            hasil[i] = {"Kategori": KATEGORI_TIDAK_RELEVAN, "Sentimen": "Netral", "Metode": "Kosong",
                        **KOLOM_KELUHAN_KOSONG}

    kelompok_indeks = [
        indeks_berisi[awal:awal + UKURAN_KELOMPOK]
        for awal in range(0, len(indeks_berisi), UKURAN_KELOMPOK)
    ]
    semafor = asyncio.Semaphore(MAKS_PARALEL)
    async with anthropic.AsyncAnthropic(api_key=os.environ.get("LLM_API_KEY"), max_retries=4) as klien:
        semua_hasil = await asyncio.gather(*(
            _analisis_kelompok(klien, [daftar_teks[i] for i in kelompok], semafor, jejak)
            for kelompok in kelompok_indeks
        ))

    for kelompok, hasil_kelompok in zip(kelompok_indeks, semua_hasil):
        for i, item in zip(kelompok, hasil_kelompok):
            hasil[i] = item
    return hasil
