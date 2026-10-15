"""Analisis LLM lewat Message Batches API (diskon 50%), dengan prompt & aturan yang sama persis.

Perintah (dari folder keluhanku):
  python metode_1_llm/analisis_batch.py kirim 10000      -> sampel acak 10.000 tweet bersih, kirim batch
  python metode_1_llm/analisis_batch.py uji-kalibrasi    -> batch kecil 100 tweet kalibrasi (cek kesamaan dg mode langsung)
  python metode_1_llm/analisis_batch.py cek <nama>       -> status batch
  python metode_1_llm/analisis_batch.py ambil <nama>     -> unduh hasil, tulis JSONL + Excel

Status batch disimpan di Downloads/analisis_llm/batch_<nama>.json, jadi bisa dicek/diambil kapan saja.
"""

import json
import logging
import os
import random
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # agar bisa dijalankan langsung

import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font

from metode_1_llm.llm_analyzer import MODEL, UKURAN_KELOMPOK, baca_jawaban, parameter_request
from persiapan_data.ekspor_siap_analisis import _aman, muat_data
from persiapan_data.pembersih import bersihkan_teks

logging.basicConfig(level=logging.INFO, format="%(message)s")
for nama in ("httpx", "httpx2", "anthropic"):
    logging.getLogger(nama).setLevel(logging.WARNING)
logger = logging.getLogger("batch")

FOLDER = Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "analisis_llm"
BENIH = 2026
HARGA_INPUT, HARGA_OUTPUT = 0.5, 2.5  # USD per 1 juta token, harga Batch API (50% dari harga normal)
TEBAL = Font(bold=True)
KATEGORI = ["Angkutan Umum", "Lahan Parkir", "Manajemen Lalu Lintas", "Tidak Relevan"]
SENTIMEN = ["Negatif", "Netral", "Positif"]
KOLOM_LABEL = ["Kategori", "Sentimen", "Keluhan", "Sebab", "Akibat", "Emosi", "Keluhan Lengkap", "Metode", "Koreksi"]


def klien() -> anthropic.Anthropic:
    """Klien Anthropic memakai LLM_API_KEY dari .env (dimuat oleh llm_analyzer)."""
    return anthropic.Anthropic(api_key=os.environ.get("LLM_API_KEY"), max_retries=4)


def path_status(nama: str) -> Path:
    """File status batch."""
    return FOLDER / f"batch_{nama}.json"


def kirim(nama: str, sampel: list[dict], keterangan: str) -> None:
    """Kelompokkan per 20 teks, kirim sebagai satu batch, simpan status + data sampel."""
    kelompok = [sampel[i:i + UKURAN_KELOMPOK] for i in range(0, len(sampel), UKURAN_KELOMPOK)]
    permintaan = [{"custom_id": f"k{j:05d}", "params": parameter_request([d["teks_bersih"] for d in grup])}
                  for j, grup in enumerate(kelompok)]
    batch = klien().messages.batches.create(requests=permintaan)
    FOLDER.mkdir(parents=True, exist_ok=True)
    path_status(nama).write_text(json.dumps({
        "batch_id": batch.id, "dibuat": datetime.now().isoformat(timespec="seconds"), "model": MODEL,
        "keterangan": keterangan, "jumlah_tweet": len(sampel), "jumlah_request": len(permintaan),
        "kelompok": {f"k{j:05d}": grup for j, grup in enumerate(kelompok)},
    }, ensure_ascii=False), encoding="utf-8")
    logger.info("Batch terkirim: %s | %d tweet dalam %d request | status: %s",
                batch.id, len(sampel), len(permintaan), batch.processing_status)


def cek(nama: str) -> str:
    """Tampilkan status batch; kembalikan processing_status."""
    status = json.loads(path_status(nama).read_text(encoding="utf-8"))
    b = klien().messages.batches.retrieve(status["batch_id"])
    c = b.request_counts
    logger.info("%s | %s | selesai %d, proses %d, gagal %d, kedaluwarsa %d, batal %d",
                b.id, b.processing_status, c.succeeded, c.processing, c.errored, c.expired, c.canceled)
    return b.processing_status


def ambil(nama: str) -> list[dict]:
    """Unduh hasil batch -> label per tweet (urutan sesuai sampel); tulis JSONL + Excel."""
    status = json.loads(path_status(nama).read_text(encoding="utf-8"))
    hasil_per_id: dict[str, list[dict]] = {}
    tok_in = tok_out = 0
    gagal = []
    for entri in klien().messages.batches.results(status["batch_id"]):
        grup = status["kelompok"][entri.custom_id]
        if entri.result.type != "succeeded":
            gagal.append(entri.custom_id)
            continue
        pesan = entri.result.message
        tok_in += pesan.usage.input_tokens
        tok_out += pesan.usage.output_tokens
        label, _ = baca_jawaban(pesan.content, [d["teks_bersih"] for d in grup])
        hasil_per_id[entri.custom_id] = label
    if gagal:
        logger.warning("%d request tidak berhasil: %s (bisa dikirim ulang)", len(gagal), gagal[:10])

    baris = []
    for kid, grup in status["kelompok"].items():
        for d, label in zip(grup, hasil_per_id.get(kid, [])):
            baris.append({**d, **label})
    biaya = tok_in / 1e6 * HARGA_INPUT + tok_out / 1e6 * HARGA_OUTPUT
    with (FOLDER / f"hasil_{nama}.jsonl").open("w", encoding="utf-8") as f:
        for b in baris:
            f.write(json.dumps(b, ensure_ascii=False) + "\n")
    tujuan = simpan_excel(nama, baris, status, tok_in, tok_out, biaya, len(gagal))
    logger.info("Diambil %d tweet | token in %d, out %d | biaya ≈ $%.2f | Excel: %s",
                len(baris), tok_in, tok_out, biaya, tujuan)
    return baris


def simpan_excel(nama, baris, status, tok_in, tok_out, biaya, n_gagal) -> Path:
    """Excel: hasil per tweet, ringkasan (kategori x sentimen, keluhan, emosi), info."""
    buku = Workbook()
    sheet = buku.active
    sheet.title = "Hasil"
    sheet.append(["id", "date", "user", "text", "Teks Bersih", "seed_query"] + KOLOM_LABEL)
    for d in baris:
        sheet.append([str(d.get("id", "")), d.get("date", ""), d.get("user", ""), _aman(d.get("text", "")),
                      _aman(d["teks_bersih"]), d.get("query", "")] + [d.get(k, "") for k in KOLOM_LABEL])
    for sel in sheet[1]:
        sel.font = TEBAL
    for h, w in zip("ABCDEFGHIJKLMN", [21, 26, 16, 50, 50, 26, 22, 10, 9, 28, 28, 14, 10, 8]):
        sheet.column_dimensions[h].width = w
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    ring = buku.create_sheet("Ringkasan")
    n = len(baris) or 1
    silang = Counter((d["Kategori"], d["Sentimen"]) for d in baris)
    ring.append(["Kategori x Sentimen"])
    ring.append(["Kategori"] + SENTIMEN + ["Total", "%"])
    for k in KATEGORI:
        v = [silang[(k, s)] for s in SENTIMEN]
        ring.append([k] + v + [sum(v), round(sum(v) / n * 100, 1)])
    ring.append([])
    ring.append(["Keluhan per kategori"])
    ring.append(["Kategori", "Keluhan", "Keluhan Lengkap (ada sebab & akibat)"])
    for k in KATEGORI[:3]:
        ring.append([k, sum(d["Kategori"] == k and d["Keluhan"] == "Ya" for d in baris),
                     sum(d["Kategori"] == k and d["Keluhan Lengkap"] == "Ya" for d in baris)])
    ring.append(["Total", sum(d["Keluhan"] == "Ya" for d in baris), sum(d["Keluhan Lengkap"] == "Ya" for d in baris)])
    ring.append([])
    ring.append(["Emosi (hanya keluhan)"])
    ring.append(["Emosi", "Jumlah"])
    for e, j in Counter(d["Emosi"] for d in baris if d["Keluhan"] == "Ya").most_common():
        ring.append([e, j])
    for b in ring.iter_rows():
        if b[0].value in ("Kategori x Sentimen", "Keluhan per kategori", "Emosi (hanya keluhan)", "Kategori", "Emosi"):
            for sel in b:
                sel.font = TEBAL
    ring.column_dimensions["A"].width = 26
    ring.column_dimensions["C"].width = 34

    info = buku.create_sheet("Info")
    for b in [["Batch ID", status["batch_id"]], ["Dikirim", status["dibuat"]], ["Model", status["model"]],
              ["Keterangan", status["keterangan"]], ["Jumlah tweet", len(baris)],
              ["Request gagal", n_gagal],
              ["Token", f"input {tok_in}, output {tok_out}" if tok_in else "lihat file hasil batch asli"],
              ["Biaya (Batch API, 50%)", f"${biaya:.2f}" if biaya else "lihat file hasil batch asli"],
              ["Prompt", "versi 3.1 (metode_1_llm/llm_analyzer.py, PROMPT_SISTEM), temperature 0"]]:
        info.append(b)
    info.column_dimensions["A"].width = 24
    info.column_dimensions["B"].width = 80
    tujuan = FOLDER / f"hasil_{nama}.xlsx"
    buku.save(tujuan)
    return tujuan


def sampel_acak(n: int) -> list[dict]:
    """n tweet acak dari data bersih (benih tetap)."""
    baris, lolos, _ = muat_data()
    random.seed(BENIH)
    pilih = random.sample(lolos, min(n, len(lolos)))
    return [{"id": str(baris[i]["id"]), "date": baris[i]["date"], "user": baris[i]["user"], "text": baris[i]["text"],
             "query": baris[i]["query"], "teks_bersih": bersihkan_teks(baris[i]["text"])} for i in pilih]


def sampel_kalibrasi() -> list[dict]:
    """100 tweet kalibrasi (urutan sama dengan lembar kalibrasi)."""
    kunci = json.loads((Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "kalibrasi" / "kunci_kalibrasi.json").read_text(encoding="utf-8"))
    return [{"id": k["id_tweet"], "teks_bersih": k["teks_bersih"], "ID": k["ID"]} for k in kunci]


def main() -> None:
    """Jalankan perintah dari argumen."""
    perintah = sys.argv[1] if len(sys.argv) > 1 else ""
    if perintah == "kirim":
        n = int(sys.argv[2]) if len(sys.argv) > 2 else 10000
        kirim(f"v3_{n}", sampel_acak(n), f"{n} tweet acak dari data bersih final (benih {BENIH})")
    elif perintah == "uji-kalibrasi":
        kirim("uji_kalibrasi", sampel_kalibrasi(), "100 tweet kalibrasi (uji kesamaan dengan mode langsung)")
    elif perintah == "cek":
        cek(sys.argv[2])
    elif perintah == "tunggu":
        while cek(sys.argv[2]) != "ended":
            time.sleep(60)
    elif perintah == "ambil":
        ambil(sys.argv[2])
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
