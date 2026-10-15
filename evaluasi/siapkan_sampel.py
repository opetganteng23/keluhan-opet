"""Langkah 1 evaluasi akurasi: ambil sampel acak, jalankan kedua engine, buat lembar pelabelan.

Menghasilkan (di Downloads/evaluasi_akurasi/):
- lembar_label_pelabel1.xlsx, lembar_label_pelabel2.xlsx : untuk dilabeli manual (TANPA label sistem)
- hasil_engine_sampel.xlsx : label kedua engine + bukti cara kerjanya
  (JANGAN dibuka sebelum pelabelan manual selesai, supaya penilaian tidak terpengaruh)

Pemakaian:  python evaluasi/siapkan_sampel.py [jumlah_sampel]   (default 385 = Slovin e=5%)
"""

import asyncio
import json
import logging
import os
import math
import random
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from metode_2_kata_kunci.classifier import KATEGORI_DEFAULT, KATEGORI_KEYWORDS, kata_cocok
from persiapan_data.ekspor_siap_analisis import _aman, muat_data
from metode_1_llm.llm_analyzer import MODEL, PROMPT_SISTEM, TOOL_ANALISIS, UKURAN_KELOMPOK, analisis_dengan_llm
from persiapan_data.pembersih import bersihkan_teks
from metode_2_kata_kunci.sentiment import KATA_NEGATIF, KATA_POSITIF

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("evaluasi")

FOLDER = Path(os.environ.get("EVALUASI_FOLDER", Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "evaluasi_akurasi"))
BENIH_ACAK = 2026  # supaya sampel bisa diulang persis sama
LABEL_KATEGORI = ["Angkutan Umum", "Lahan Parkir", "Manajemen Lalu Lintas", "Tidak Relevan"]
LABEL_SENTIMEN = ["Positif", "Negatif", "Netral"]
TEBAL = Font(bold=True)
ISIAN = PatternFill("solid", fgColor="FFF4CC")

PEDOMAN = [
    ["Pedoman Pelabelan"],
    [],
    ["Cara mengisi", "Baca kolom Teks, lalu pilih Kategori dan Sentimen dari dropdown (sel kuning). "
                     "Isi berdasarkan isi teks saja, jangan menebak di luar teks."],
    ["Penting", "Kerjakan sendiri, jangan berdiskusi dengan pelabel lain sampai keduanya selesai. "
                "Jangan membuka hasil_engine_sampel.xlsx sebelum selesai."],
    [],
    ["KATEGORI", ""],
    ["Angkutan Umum", "Pengalaman/pendapat tentang layanan TransJakarta, JakLingko/Mikrotrans, angkot, ojek, "
                      "halte, terminal: armada, sopir/petugas, ketepatan waktu, kepadatan, kenyamanan, tarif, rute."],
    ["Lahan Parkir", "Parkir liar, juru parkir (jukir), tarif/retribusi parkir, ketersediaan lahan parkir, derek."],
    ["Manajemen Lalu Lintas", "Kemacetan, rambu, lampu lalu lintas, rekayasa lalu lintas, ganjil genap, petugas "
                              "Dishub di jalan, fasilitas jalan: trotoar, lampu jalan/PJU, JPO, zebra cross, marka."],
    ["Tidak Relevan", "Review aplikasi (bug, login, top up), iklan, candaan, cerita pribadi tanpa keluhan "
                      "layanan, berita/pengumuman tanpa pendapat warga, topik di luar Dishub (banjir, sampah, dll)."],
    [],
    ["SENTIMEN", "Sikap penulis terhadap layanan yang dibahas"],
    ["Positif", "Memuji, puas, berterima kasih atas layanan."],
    ["Negatif", "Mengeluh, kecewa, marah, mengkritik. Sarkasme ('hebat, nunggu 1 jam') = Negatif."],
    ["Netral", "Bertanya, memberi informasi, tidak jelas sikapnya. Jika kategori Tidak Relevan, pilih Netral."],
    [],
    ["Ragu-ragu?", "Pilih yang paling sesuai, lalu tulis alasannya di kolom Catatan."],
]


def ukuran_slovin(populasi: int, galat: float = 0.05) -> int:
    """Ukuran sampel rumus Slovin: n = N / (1 + N e^2), dibulatkan ke atas."""
    return math.ceil(populasi / (1 + populasi * galat ** 2))


def bukti_kata_kunci(teks: str) -> dict:
    """Menjalankan engine kata kunci sekaligus mencatat kata yang cocok (bukti cara kerja)."""
    per_kategori = kata_cocok(teks, KATEGORI_KEYWORDS)
    skor = {k: len(v) for k, v in per_kategori.items()}
    if not teks or max(skor.values()) == 0:
        kategori, alasan_k = KATEGORI_DEFAULT, "Tidak ada kata kunci cocok -> kategori default"
    else:
        kategori = max(skor, key=skor.get)
        alasan_k = f"Skor tertinggi ({skor[kategori]})"
    sent = kata_cocok(teks, {"positif": KATA_POSITIF, "negatif": KATA_NEGATIF})
    pos, neg = len(sent["positif"]), len(sent["negatif"])
    sentimen = "Netral" if pos == neg else ("Positif" if pos > neg else "Negatif")
    return {
        "kategori": kategori,
        "sentimen": sentimen,
        "cocok": {k: ", ".join(v) for k, v in per_kategori.items()},
        "skor": skor,
        "alasan_kategori": alasan_k,
        "kata_positif": ", ".join(sent["positif"]),
        "kata_negatif": ", ".join(sent["negatif"]),
        "alasan_sentimen": f"positif {pos} vs negatif {neg}",
    }


def _lebar(sheet, lebar: list[int]) -> None:
    """Mengatur lebar kolom A, B, C, ... sesuai daftar."""
    for i, w in enumerate(lebar):
        sheet.column_dimensions[chr(65 + i)].width = w


def buat_lembar_label(sampel: list[dict], pelabel: int) -> Path:
    """Lembar pelabelan manual: teks + dropdown kategori/sentimen, tanpa label sistem."""
    buku = Workbook()
    sheet = buku.active
    sheet.title = "Pelabelan"
    sheet.append(["No", "ID Sampel", "Teks", "Kategori", "Sentimen", "Catatan"])
    for sel in sheet[1]:
        sel.font = TEBAL
    for i, s in enumerate(sampel, 1):
        sheet.append([i, s["id_sampel"], s["teks_bersih"], "", "", ""])
        sheet.cell(row=i + 1, column=3).alignment = Alignment(wrap_text=True, vertical="top")
        for kolom in (4, 5):
            sheet.cell(row=i + 1, column=kolom).fill = ISIAN
    akhir = len(sampel) + 1
    dv_k = DataValidation(type="list", formula1='"' + ",".join(LABEL_KATEGORI) + '"', allow_blank=True)
    dv_s = DataValidation(type="list", formula1='"' + ",".join(LABEL_SENTIMEN) + '"', allow_blank=True)
    for dv, kolom in ((dv_k, "D"), (dv_s, "E")):
        dv.error, dv.errorTitle = "Pilih dari daftar", "Isian tidak valid"
        sheet.add_data_validation(dv)
        dv.add(f"{kolom}2:{kolom}{akhir}")
    _lebar(sheet, [6, 11, 80, 24, 12, 30])
    sheet.freeze_panes = "C2"

    pedoman = buku.create_sheet("Pedoman")
    for baris in PEDOMAN:
        pedoman.append(baris)
    for baris in pedoman.iter_rows():
        if baris[0].value in ("Pedoman Pelabelan", "KATEGORI", "SENTIMEN"):
            baris[0].font = TEBAL
        for sel in baris:
            sel.alignment = Alignment(wrap_text=True, vertical="top")
    _lebar(pedoman, [24, 100])

    tujuan = FOLDER / f"lembar_label_pelabel{pelabel}.xlsx"
    buku.save(tujuan)
    return tujuan


def buat_hasil_engine(sampel: list[dict], jejak: list[dict], info: dict) -> Path:
    """Label kedua engine + bukti cara kerja + konfigurasi, untuk dibandingkan dengan label manual."""
    buku = Workbook()
    ringkas = buku.active
    ringkas.title = "Label Engine"
    ringkas.append(["ID Sampel", "Teks", "Kategori LLM", "Sentimen LLM", "Metode LLM",
                    "Kategori Kata Kunci", "Sentimen Kata Kunci"])
    for s in sampel:
        ringkas.append([s["id_sampel"], s["teks_bersih"], s["llm"]["Kategori"], s["llm"]["Sentimen"],
                        s["llm"]["Metode"], s["kk"]["kategori"], s["kk"]["sentimen"]])
    _lebar(ringkas, [11, 70, 22, 13, 10, 22, 13])

    kk = buku.create_sheet("Bukti Kata Kunci")
    kategori = list(KATEGORI_KEYWORDS)
    kk.append(["ID Sampel", "Teks"] + [f"Cocok: {k}" for k in kategori] + [f"Skor: {k}" for k in kategori]
              + ["Kategori", "Alasan Kategori", "Kata Positif", "Kata Negatif", "Sentimen", "Alasan Sentimen"])
    for s in sampel:
        b = s["kk"]
        kk.append([s["id_sampel"], s["teks_bersih"]] + [b["cocok"][k] for k in kategori]
                  + [b["skor"][k] for k in kategori]
                  + [b["kategori"], b["alasan_kategori"], b["kata_positif"], b["kata_negatif"],
                     b["sentimen"], b["alasan_sentimen"]])
    _lebar(kk, [11, 60, 22, 22, 22, 9, 9, 9, 22, 30, 18, 18, 10, 18])

    jl = buku.create_sheet("Bukti LLM (per request)")
    jl.append(["Request ke", "ID Request", "Model", "Token Input", "Token Output",
               "Teks yang Dikirim", "Jawaban Mentah Model (JSON)"])
    for i, j in enumerate(jejak, 1):
        jl.append([i, j["id_request"], j["model"], j["token_input"], j["token_output"],
                   _aman(j["pesan_dikirim"]), _aman(j["jawaban_mentah"])])
        for kolom in (6, 7):
            jl.cell(row=i + 1, column=kolom).alignment = Alignment(wrap_text=True, vertical="top")
    _lebar(jl, [10, 30, 18, 12, 12, 80, 80])

    konf = buku.create_sheet("Konfigurasi")
    baris_konf = [["Konfigurasi Eksperimen"], []] + [[k, v] for k, v in info.items()] + [
        [], ["Engine 1: LLM"], ["Model", MODEL], ["Teks per request", UKURAN_KELOMPOK],
        ["Format jawaban", "Tool 'kirim_analisis' (strict JSON schema): nomor, kategori, sentimen"],
        ["Skema tool", json.dumps(TOOL_ANALISIS["input_schema"], ensure_ascii=False)],
        ["Prompt sistem", PROMPT_SISTEM],
        [], ["Engine 2: Kata Kunci"],
        ["Cara kerja", "Hitung kata kunci (kata utuh, frasa terpanjang dulu) per kategori; skor tertinggi menang; "
                       "semua nol -> Manajemen Lalu Lintas; seri -> urutan kamus. Sentimen: jumlah kata positif vs negatif."],
    ] + [[f"Kamus: {k}", ", ".join(v)] for k, v in KATEGORI_KEYWORDS.items()] + [
        ["Kamus: positif", ", ".join(KATA_POSITIF)], ["Kamus: negatif", ", ".join(KATA_NEGATIF)],
    ]
    for b in baris_konf:
        konf.append(b)
    for b in konf.iter_rows():
        if b[0].value in ("Konfigurasi Eksperimen", "Engine 1: LLM", "Engine 2: Kata Kunci"):
            b[0].font = TEBAL
        for sel in b:
            sel.alignment = Alignment(wrap_text=True, vertical="top")
    _lebar(konf, [24, 110])

    for sheet in buku.worksheets[:3]:
        for sel in sheet[1]:
            sel.font = TEBAL
        sheet.freeze_panes = "B2"
    tujuan = FOLDER / "hasil_engine_sampel.xlsx"
    buku.save(tujuan)
    return tujuan


def main() -> None:
    """Ambil sampel, jalankan kedua engine, simpan lembar pelabelan & hasil engine."""
    FOLDER.mkdir(parents=True, exist_ok=True)
    baris, lolos, _ = muat_data()
    n = int(sys.argv[1]) if len(sys.argv) > 1 else ukuran_slovin(len(lolos))
    random.seed(BENIH_ACAK)
    indeks = sorted(random.sample(lolos, n))
    sampel = [{
        "id_sampel": f"S{nomor:03d}",
        "id_tweet": baris[i]["id"],
        "teks_bersih": _aman(bersihkan_teks(baris[i]["text"])),
    } for nomor, i in enumerate(indeks, 1)]
    logger.info("Populasi %d tweet bersih -> sampel acak %d (benih %d)", len(lolos), n, BENIH_ACAK)

    for s in sampel:
        s["kk"] = bukti_kata_kunci(s["teks_bersih"])
    jejak: list[dict] = []
    label_llm = asyncio.run(analisis_dengan_llm([s["teks_bersih"] for s in sampel], jejak))
    for s, label in zip(sampel, label_llm):
        s["llm"] = label

    info = {
        "Tanggal": datetime.now().strftime("%d-%m-%Y %H:%M"),
        "Populasi (tweet bersih)": len(lolos),
        "Ukuran sampel": n,
        "Rumus": "Slovin, e = 5%" if len(sys.argv) <= 1 else "ditentukan manual",
        "Teknik sampling": f"Acak sederhana (random.sample), benih {BENIH_ACAK}",
        "Total token": f"input {sum(j['token_input'] for j in jejak)}, output {sum(j['token_output'] for j in jejak)}",
    }
    hasil = buat_hasil_engine(sampel, jejak, info)
    for p in (1, 2):
        logger.info("Lembar pelabelan: %s", buat_lembar_label(sampel, p))
    logger.info("Hasil engine (jangan dibuka sebelum melabeli): %s", hasil)
    logger.info("Token LLM: %s", info["Total token"])


if __name__ == "__main__":
    main()
