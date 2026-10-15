"""Pasca-proses hasil LLM: perbaiki "sebab" yang sebenarnya bukan penyebab (aturan transparan).

Aturan:
  R1. Sebab hanya berupa "macet/kemacetan (parah/panjang) (di) Jakarta" (atau "Jakarta macet")
      -> itu kondisi yang dikeluhkan (akibat), bukan penyebab -> sebab = "tidak disebutkan".
  R2. Sebab sama persis dengan akibat -> sebab = "tidak disebutkan".
Setelah aturan diterapkan, "Keluhan Lengkap" dihitung ulang (Ya hanya jika sebab & akibat ada).
Kolom "Koreksi" mencatat aturan yang diterapkan per baris; data asli tidak ditimpa (file baru).

Pemakaian:  python metode_1_llm/pasca_proses.py v3_10000
  -> Downloads/analisis_llm/hasil_v3_10000_final.jsonl + .xlsx
"""

import json
import logging
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # agar bisa dijalankan langsung

from metode_1_llm.analisis_batch import FOLDER, path_status, simpan_excel
from metode_1_llm.llm_analyzer import TIDAK_DISEBUTKAN

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("pasca")

POLA_MACET_SAJA = re.compile(
    r"^(?:(?:ke)?macet(?:an)?(?:\s+(?:parah|panjang|total|berat))?(?:\s+(?:di\s+)?(?:kota\s+)?jakarta)?"
    r"|jakarta\s+(?:sangat\s+)?macet(?:\s+parah)?)$"
)


def koreksi(baris: dict) -> str:
    """Terapkan R1/R2 pada satu baris (ubah di tempat); kembalikan kode aturan yang dipakai."""
    if baris.get("Keluhan") != "Ya":
        return ""
    sebab = baris["Sebab"].strip().lower().rstrip(".")
    akibat = baris["Akibat"].strip().lower().rstrip(".")
    aturan = ""
    if sebab != TIDAK_DISEBUTKAN and POLA_MACET_SAJA.match(sebab):
        aturan = "R1"
    elif sebab != TIDAK_DISEBUTKAN and sebab == akibat:
        aturan = "R2"
    if aturan:
        baris["Sebab"] = TIDAK_DISEBUTKAN
    ada_sebab = baris["Sebab"].strip().lower() != TIDAK_DISEBUTKAN
    ada_akibat = baris["Akibat"].strip().lower() != TIDAK_DISEBUTKAN
    baris["Keluhan Lengkap"] = "Ya" if ada_sebab and ada_akibat else "Tidak"
    return aturan


def main() -> None:
    """Baca hasil batch, koreksi, tulis versi final."""
    nama = sys.argv[1] if len(sys.argv) > 1 else "v3_10000"
    baris = [json.loads(x) for x in (FOLDER / f"hasil_{nama}.jsonl").open(encoding="utf-8")]
    lengkap_awal = sum(b["Keluhan Lengkap"] == "Ya" for b in baris)
    contoh = {}
    for b in baris:
        sebab_asli = b["Sebab"]
        b["Koreksi"] = koreksi(b)
        if b["Koreksi"]:
            contoh.setdefault(b["Koreksi"], []).append(sebab_asli)
    with (FOLDER / f"hasil_{nama}_final.jsonl").open("w", encoding="utf-8") as f:
        for b in baris:
            f.write(json.dumps(b, ensure_ascii=False) + "\n")
    status = json.loads(path_status(nama).read_text(encoding="utf-8"))
    status["keterangan"] += " | pasca-proses R1/R2 (sebab bukan penyebab -> 'tidak disebutkan')"
    tujuan = simpan_excel(f"{nama}_final", baris, status, 0, 0, 0.0, 0)
    for kode, daftar in contoh.items():
        logger.info("%s: %d baris dikoreksi, contoh sebab asli: %s", kode, len(daftar), sorted(set(daftar))[:8])
    logger.info("Keluhan Lengkap: %d -> %d | file: %s",
                lengkap_awal, sum(b["Keluhan Lengkap"] == "Ya" for b in baris), tujuan)


if __name__ == "__main__":
    main()
