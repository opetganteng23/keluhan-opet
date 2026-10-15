"""Kalibrasi prompt: jalankan LLM (prompt aktif) pada 100 tweet kalibrasi, bandingkan dengan label manual.

Metrik: kecocokan Kategori, Sentimen, Keluhan, Emosi (+ precision/recall Keluhan),
dan kecocokan ADA/TIDAK-nya Sebab & Akibat. Isi Sebab/Akibat ditampilkan berdampingan untuk dinilai manual.

Hasil: Downloads/kalibrasi/hasil_kalibrasi_<versi>.xlsx
Pemakaian:  python evaluasi/evaluasi_kalibrasi.py [label_versi]   (default "v2")
"""

import asyncio
import logging
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from metode_1_llm.llm_analyzer import TIDAK_DISEBUTKAN, analisis_dengan_llm

logging.basicConfig(level=logging.INFO, format="%(message)s")
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpx2").setLevel(logging.WARNING)
logger = logging.getLogger("kalibrasi")

FOLDER = Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "kalibrasi"
HARGA_INPUT, HARGA_OUTPUT = 1.0, 5.0
TEBAL = Font(bold=True)
SALAH = PatternFill("solid", fgColor="FBE3E3")
NORMAL_EMOSI = {"khawatir": "Takut/Khawatir", "takut": "Takut/Khawatir"}


def norm_emosi(e: str) -> str:
    """Samakan penulisan emosi (mis. 'Khawatir' -> 'Takut/Khawatir')."""
    e = (e or "").strip()
    return NORMAL_EMOSI.get(e.lower(), e)


def ada(teks: str) -> bool:
    """True jika sebab/akibat diisi (bukan 'tidak disebutkan')."""
    return bool(teks) and teks.strip().lower() != TIDAK_DISEBUTKAN


def baca_manual() -> list[dict]:
    """Label manual dari lembar kalibrasi."""
    baris = list(load_workbook(FOLDER / "lembar_kalibrasi_100.xlsx", read_only=True)["Kalibrasi"]
                 .iter_rows(values_only=True))[1:]
    return [{"ID": b[1], "Teks": b[2], "Kategori": b[3], "Sentimen": b[4], "Keluhan": b[5],
             "Sebab": str(b[6] or "").strip(), "Akibat": str(b[7] or "").strip(),
             "Emosi": norm_emosi(b[8]), "Catatan": b[9] or ""} for b in baris if b[1]]


def main() -> None:
    """Jalankan, bandingkan, tulis laporan."""
    versi = sys.argv[1] if len(sys.argv) > 1 else "v2"
    manual = baca_manual()
    jejak: list[dict] = []
    llm = asyncio.run(analisis_dengan_llm([m["Teks"] for m in manual], jejak))
    n = len(manual)

    def cocok(kunci, f=lambda x: x):
        return sum(f(m[kunci]) == f(l[kunci]) for m, l in zip(manual, llm))

    hasil = {
        "Kategori": cocok("Kategori"),
        "Sentimen": cocok("Sentimen"),
        "Keluhan": cocok("Keluhan"),
        "Emosi": cocok("Emosi", lambda e: norm_emosi(e)),
        "Sebab ada/tidak": cocok("Sebab", ada),
        "Akibat ada/tidak": cocok("Akibat", ada),
    }
    tp = sum(m["Keluhan"] == "Ya" and l["Keluhan"] == "Ya" for m, l in zip(manual, llm))
    pred = sum(l["Keluhan"] == "Ya" for l in llm)
    akt = sum(m["Keluhan"] == "Ya" for m in manual)
    biaya = sum(j["token_input"] for j in jejak) / 1e6 * HARGA_INPUT + \
        sum(j["token_output"] for j in jejak) / 1e6 * HARGA_OUTPUT

    buku = Workbook()
    ring = buku.active
    ring.title = "Ringkasan"
    ring.append([f"Kalibrasi prompt {versi} (n = {n})"])
    ring.append(["Aspek", "Cocok", "Persen"])
    for k, v in hasil.items():
        ring.append([k, v, round(v / n * 100, 1)])
    ring.append([])
    ring.append(["Keluhan: precision", f"{tp}/{pred}", round(tp / pred * 100, 1) if pred else 0])
    ring.append(["Keluhan: recall", f"{tp}/{akt}", round(tp / akt * 100, 1) if akt else 0])
    ring.append(["Biaya", f"${biaya:.4f}"])
    for b in (1, 2):
        for sel in ring[b]:
            sel.font = TEBAL
    ring.column_dimensions["A"].width = 24

    detail = buku.create_sheet("Perbandingan")
    kolom = ["Kategori", "Sentimen", "Keluhan", "Emosi"]
    detail.append(["ID", "Teks"] + [f"{k} {s}" for k in kolom for s in ("Manual", "LLM")]
                  + ["Sebab Manual", "Sebab LLM", "Akibat Manual", "Akibat LLM", "Catatan"])
    for sel in detail[1]:
        sel.font = TEBAL
    for m, l in zip(manual, llm):
        detail.append([m["ID"], m["Teks"]] + [x for k in kolom for x in (m[k], l[k])]
                      + [m["Sebab"], l["Sebab"], m["Akibat"], l["Akibat"], m["Catatan"]])
        r = detail.max_row
        for j, k in enumerate(kolom):
            if norm_emosi(str(m[k])) != norm_emosi(str(l[k])):
                for c in (3 + 2 * j, 4 + 2 * j):
                    detail.cell(row=r, column=c).fill = SALAH
        for c in (2, 11, 12, 13, 14):
            detail.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")
    for h, w in zip("ABCDEFGHIJKLMNO", [6, 50, 18, 18, 9, 9, 8, 8, 13, 13, 24, 24, 24, 24, 20]):
        detail.column_dimensions[h].width = w
    detail.freeze_panes = "C2"
    detail.auto_filter.ref = detail.dimensions

    tujuan = FOLDER / f"hasil_kalibrasi_{versi}.xlsx"
    buku.save(tujuan)
    for k, v in hasil.items():
        logger.info("%-17s %3d/%d  (%.0f%%)", k, v, n, v / n * 100)
    logger.info("Keluhan precision %d/%d, recall %d/%d | biaya $%.4f", tp, pred, tp, akt, biaya)
    salah_kat = Counter((m["Kategori"], l["Kategori"]) for m, l in zip(manual, llm) if m["Kategori"] != l["Kategori"])
    salah_kel = Counter((m["Keluhan"], l["Keluhan"]) for m, l in zip(manual, llm) if m["Keluhan"] != l["Keluhan"])
    logger.info("Salah kategori (manual -> LLM): %s", dict(salah_kat))
    logger.info("Salah keluhan  (manual -> LLM): %s", dict(salah_kel))
    logger.info("Laporan: %s", tujuan)


if __name__ == "__main__":
    main()
