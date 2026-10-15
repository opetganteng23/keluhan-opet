"""Langkah 2 evaluasi akurasi: bandingkan label engine dengan label manual, hitung metrik + visual.

Membaca dari Downloads/evaluasi_akurasi/:
- hasil_engine_sampel.xlsx (dibuat siapkan_sampel.py)
- lembar_label_pelabel1.xlsx (WAJIB, sudah diisi) -> dipakai sebagai kunci jawaban (gold standard)
- lembar_label_pelabel2.xlsx (opsional, sudah diisi) -> untuk Cohen's Kappa antar-pelabel

Menghasilkan laporan_evaluasi.xlsx: ringkasan + grafik, metrik per kelas, confusion matrix
(heatmap), kesepakatan pelabel, dan daftar kesalahan beserta bukti cara kerja engine.

Pemakaian:  python evaluasi/hitung_evaluasi.py
"""

import logging
import os
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("evaluasi")

FOLDER = Path(os.environ.get("EVALUASI_FOLDER", Path.home() / "Downloads" / "Hasil_Penelitian" / "_file_kerja" / "evaluasi_akurasi"))
# EVAL_VERSI=v2 -> evaluasi final (definisi keluhan v2/v3): + tugas Keluhan, file *_v2
VERSI = os.environ.get("EVAL_VERSI", "v1")
AKHIRAN = f"_{VERSI}" if VERSI in ("v2", "v3") else ""
FILE_ENGINE = f"hasil_engine_sampel{AKHIRAN}.xlsx"
FILE_GOLD = {"v3": "lembar_label_v3_1000.xlsx", "v2": "lembar_label_v2_pelabel1.xlsx"}.get(VERSI, "lembar_label_pelabel1.xlsx")
FILE_GOLD2 = "lembar_label_v2_pelabel2.xlsx" if VERSI in ("v2", "v3") else "lembar_label_pelabel2.xlsx"
FILE_LAPORAN = f"laporan_evaluasi{AKHIRAN}.xlsx"
LABEL = {
    "Kategori": ["Angkutan Umum", "Lahan Parkir", "Manajemen Lalu Lintas", "Tidak Relevan"],
    "Sentimen": ["Positif", "Negatif", "Netral"],
}
if VERSI in ("v2", "v3"):
    LABEL["Keluhan"] = ["Ya", "Tidak"]
# Engine -> akhiran nama kolom di sheet "Label Engine" (kolom "<Tugas> <akhiran>"); engine dievaluasi
# hanya jika kolom semua tugas tersedia
KOLOM_ENGINE = {"LLM": "LLM", "IndoBERT": "IndoBERT", "SVM": "SVM", "Naive Bayes": "NB", "Kata Kunci": "Kata Kunci"}
# Urutan warna tetap per engine (tervalidasi aman buta warna; ada label angka di grafik)
WARNA_ENGINE = {"LLM": "2A78D6", "Kata Kunci": "EB6834", "Naive Bayes": "1BAF7A", "SVM": "EDA100", "IndoBERT": "E87BA4"}
ENGINE: list[str] = []  # diisi saat membaca hasil_engine_sampel.xlsx
TEBAL = Font(bold=True)


# ---------------------------------------------------------------- metrik


def matriks_kebingungan(benar: list[str], tebakan: list[str], label: list[str]) -> list[list[int]]:
    """Confusion matrix: baris = label manual (benar), kolom = label engine (tebakan)."""
    posisi = {nama: i for i, nama in enumerate(label)}
    matriks = [[0] * len(label) for _ in label]
    for b, t in zip(benar, tebakan):
        matriks[posisi[b]][posisi[t]] += 1
    return matriks


def metrik_per_kelas(matriks: list[list[int]], label: list[str]) -> list[dict]:
    """Precision, recall, F1, dan support untuk setiap kelas."""
    hasil = []
    for i, nama in enumerate(label):
        tp = matriks[i][i]
        prediksi = sum(baris[i] for baris in matriks)
        aktual = sum(matriks[i])
        p = tp / prediksi if prediksi else 0.0
        r = tp / aktual if aktual else 0.0
        f1 = 2 * p * r / (p + r) if p + r else 0.0
        hasil.append({"kelas": nama, "precision": p, "recall": r, "f1": f1, "support": aktual})
    return hasil


def akurasi(matriks: list[list[int]]) -> float:
    """Proporsi tebakan yang benar (diagonal / total)."""
    total = sum(map(sum, matriks))
    return sum(matriks[i][i] for i in range(len(matriks))) / total if total else 0.0


def macro_f1(per_kelas: list[dict]) -> float:
    """Rata-rata F1 kelas yang muncul di data (support > 0)."""
    ada = [k["f1"] for k in per_kelas if k["support"] > 0]
    return sum(ada) / len(ada) if ada else 0.0


def cohen_kappa(a: list[str], b: list[str], label: list[str]) -> float:
    """Cohen's Kappa: kesepakatan dua pelabel, dikoreksi peluang kebetulan."""
    n = len(a)
    if n == 0:
        return 0.0
    po = sum(x == y for x, y in zip(a, b)) / n
    pe = sum((a.count(k) / n) * (b.count(k) / n) for k in label)
    return (po - pe) / (1 - pe) if pe < 1 else 1.0


def tafsir_kappa(k: float) -> str:
    """Interpretasi Landis & Koch (1977)."""
    for batas, arti in ((0.8, "Hampir sempurna"), (0.6, "Substansial"), (0.4, "Sedang"),
                        (0.2, "Cukup"), (0.0, "Rendah")):
        if k > batas:
            return arti
    return "Buruk (di bawah kebetulan)"


# ---------------------------------------------------------------- baca data


def baca_lembar(path: Path) -> dict[str, dict]:
    """Membaca lembar pelabelan (format v1 atau v2) -> {ID Sampel: {Kategori, Sentimen, [Keluhan...], Catatan}}."""
    semua = list(load_workbook(path, read_only=True)["Pelabelan"].iter_rows(values_only=True))
    pos = {j: i for i, j in enumerate(semua[0])}
    hasil = {}
    for baris in semua[1:]:
        if not baris[1]:
            continue
        d = {k: str(baris[pos[k]] or "").strip() for k in ("Kategori", "Sentimen", "Keluhan", "Sebab", "Akibat", "Emosi")
             if k in pos}
        if "Emosi" in d:
            d["Emosi"] = {"khawatir": "Takut/Khawatir", "takut": "Takut/Khawatir", "": "Tidak Ada"}.get(
                d["Emosi"].lower(), d["Emosi"])
        d["Catatan"] = baris[pos["Catatan"]] or ""
        hasil[baris[1]] = d
    return hasil


def baca_engine(path: Path) -> tuple[dict[str, dict], dict[str, dict]]:
    """Membaca label engine dan bukti kata kunci per ID Sampel."""
    buku = load_workbook(path, read_only=True)
    label, bukti = {}, {}
    baris = list(buku["Label Engine"].iter_rows(values_only=True))
    posisi = {j: i for i, j in enumerate(baris[0])}
    ENGINE[:] = [n for n, akh in KOLOM_ENGINE.items() if all(f"{t} {akh}" in posisi for t in LABEL)]
    for b in baris[1:]:
        label[b[0]] = {"Teks": b[1]}
        for n in ENGINE:
            label[b[0]][n] = {t: b[posisi[f"{t} {KOLOM_ENGINE[n]}"]] for t in LABEL}
        for ekstra in ("Sebab LLM", "Akibat LLM", "Emosi LLM"):
            if ekstra in posisi:
                label[b[0]][ekstra] = b[posisi[ekstra]] or ""
    sheet = buku["Bukti Kata Kunci"]
    judul = [c for c in next(sheet.iter_rows(values_only=True))]
    for b in list(sheet.iter_rows(values_only=True))[1:]:
        baris = dict(zip(judul, b))
        cocok = "; ".join(f"{j[7:]}: {baris[j]}" for j in judul if j.startswith("Cocok: ") and baris[j])
        bukti[b[0]] = (f"Kata cocok -> {cocok or '(tidak ada)'} | {baris['Alasan Kategori']} | "
                       f"positif: {baris['Kata Positif'] or '-'}, negatif: {baris['Kata Negatif'] or '-'}")
    return label, bukti


def validasi(gold: dict[str, dict], nama: str) -> None:
    """Hentikan jika ada baris yang belum diisi atau isiannya tidak valid."""
    salah = [i for i, g in gold.items() if any(g.get(t) not in LABEL[t] for t in LABEL)]
    if salah:
        logger.error("%s: %d baris belum diisi / tidak valid, contoh: %s", nama, len(salah), ", ".join(salah[:10]))
        sys.exit(1)


# ---------------------------------------------------------------- tulis laporan


def _judul(sheet, baris: int, teks: str) -> None:
    """Menulis judul bagian bercetak tebal."""
    sheet.cell(row=baris, column=1, value=teks).font = TEBAL


def tulis_matriks(sheet, baris: int, kolom: int, matriks, label, judul: str) -> int:
    """Menulis confusion matrix + heatmap (satu warna biru, terang->gelap). Mengembalikan baris berikutnya."""
    sheet.cell(row=baris, column=kolom, value=judul).font = TEBAL
    sheet.cell(row=baris + 1, column=kolom, value="Manual \\ Engine").font = TEBAL
    for j, nama in enumerate(label):
        sheet.cell(row=baris + 1, column=kolom + 1 + j, value=nama).font = TEBAL
    for i, nama in enumerate(label):
        sheet.cell(row=baris + 2 + i, column=kolom, value=nama).font = TEBAL
        for j in range(len(label)):
            sheet.cell(row=baris + 2 + i, column=kolom + 1 + j, value=matriks[i][j]).alignment = \
                Alignment(horizontal="center")
    awal = sheet.cell(row=baris + 2, column=kolom + 1).coordinate
    akhir = sheet.cell(row=baris + 1 + len(label), column=kolom + len(label)).coordinate
    sheet.conditional_formatting.add(f"{awal}:{akhir}", ColorScaleRule(
        start_type="min", start_color="FCFCFB", end_type="max", end_color="1C5CAB"))
    return baris + len(label) + 3


def _rapikan_grafik(grafik: BarChart) -> None:
    """Label angka saja (tanpa nama seri/kategori) di ujung batang; sumbu X & Y ditampilkan."""
    grafik.dataLabels = DataLabelList()
    grafik.dataLabels.showVal = True
    for atribut in ("showSerName", "showCatName", "showLegendKey", "showPercent"):
        setattr(grafik.dataLabels, atribut, False)
    grafik.dataLabels.position = "outEnd"
    grafik.x_axis.delete = False
    grafik.y_axis.delete = False
    grafik.y_axis.numFmt = "0"
    grafik.legend.overlay = False


def grafik_batang(sheet, ref_baris: tuple[int, int], judul: str, posisi: str) -> None:
    """Grafik batang berkelompok LLM vs Kata Kunci (satu sumbu, 0-100%)."""
    grafik = BarChart()
    grafik.type = "col"
    grafik.title = judul
    grafik.y_axis.scaling.min, grafik.y_axis.scaling.max = 0, 100
    grafik.y_axis.majorGridlines = None
    grafik.gapWidth, grafik.overlap = 80, -10
    awal, akhir = ref_baris
    data = Reference(sheet, min_col=2, max_col=1 + len(ENGINE), min_row=awal, max_row=akhir)
    kategori = Reference(sheet, min_col=1, min_row=awal + 1, max_row=akhir)
    grafik.add_data(data, titles_from_data=True)
    grafik.set_categories(kategori)
    for seri, nama in zip(grafik.series, ENGINE):
        seri.graphicalProperties.solidFill = WARNA_ENGINE[nama]
        seri.graphicalProperties.line.solidFill = WARNA_ENGINE[nama]
    _rapikan_grafik(grafik)
    grafik.legend.position = "r"
    grafik.width, grafik.height = 22, 10
    sheet.add_chart(grafik, posisi)


def tulis_evaluasi_anotasi(buku: Workbook, engine: dict, gold: dict, ids: list[str]) -> None:
    """Sheet khusus LLM: emosi & ada/tidaknya sebab-akibat pada keluhan yang disepakati (manual & LLM = Ya)."""
    def ada(t: str) -> bool:
        return bool(t) and str(t).strip().lower() != "tidak disebutkan"

    sama = [i for i in ids if gold[i]["Keluhan"] == "Ya" and engine[i]["LLM"]["Keluhan"] == "Ya"]
    sheet = buku.create_sheet("LLM Sebab-Akibat-Emosi")
    _judul(sheet, 1, f"Anotasi keluhan LLM vs manual, pada {len(sama)} tweet yang sama-sama dinilai keluhan")
    sheet.append(["Aspek", "Cocok", "Persen"])
    for sel in sheet[2]:
        sel.font = TEBAL
    n = len(sama) or 1
    for nama, f in [("Emosi", lambda i: gold[i]["Emosi"] == engine[i]["Emosi LLM"]),
                    ("Sebab ada/tidak", lambda i: ada(gold[i]["Sebab"]) == ada(engine[i]["Sebab LLM"])),
                    ("Akibat ada/tidak", lambda i: ada(gold[i]["Akibat"]) == ada(engine[i]["Akibat LLM"]))]:
        cocok = sum(f(i) for i in sama)
        sheet.append([nama, f"{cocok}/{len(sama)}", round(cocok / n * 100, 1)])
    sheet.append([])
    sheet.append(["ID", "Teks", "Sebab Manual", "Sebab LLM", "Akibat Manual", "Akibat LLM", "Emosi Manual", "Emosi LLM"])
    for sel in sheet[sheet.max_row]:
        sel.font = TEBAL
    for i in sama:
        sheet.append([i, engine[i]["Teks"], gold[i]["Sebab"], engine[i]["Sebab LLM"], gold[i]["Akibat"],
                      engine[i]["Akibat LLM"], gold[i]["Emosi"], engine[i]["Emosi LLM"]])
    for kol, w in zip("ABCDEFGH", [16, 55, 28, 28, 28, 28, 14, 14]):
        sheet.column_dimensions[kol].width = w


def main() -> None:
    """Hitung metrik semua engine x tugas, lalu tulis laporan_evaluasi.xlsx."""
    engine, bukti = baca_engine(FOLDER / FILE_ENGINE)
    gold = baca_lembar(FOLDER / FILE_GOLD)
    validasi(gold, "Pelabel 1")
    path2 = FOLDER / FILE_GOLD2
    pelabel2 = baca_lembar(path2) if path2.exists() else {}
    ada_pelabel2 = bool(pelabel2) and all(all(g.get(t) in LABEL[t] for t in LABEL) for g in pelabel2.values())
    ids = sorted(gold)

    buku = Workbook()
    ring = buku.active
    ring.title = "Ringkasan"
    _judul(ring, 1, f"Evaluasi Akurasi Engine (n = {len(ids)} sampel, kunci jawaban = Pelabel 1)")
    hasil: dict[tuple[str, str], dict] = {}
    for tugas, label in LABEL.items():
        benar = [gold[i][tugas] for i in ids]
        for nama in ENGINE:
            tebakan = [engine[i][nama][tugas] for i in ids]
            m = matriks_kebingungan(benar, tebakan, label)
            pk = metrik_per_kelas(m, label)
            hasil[(tugas, nama)] = {"matriks": m, "per_kelas": pk, "akurasi": akurasi(m), "macro_f1": macro_f1(pk)}

    # Tabel ringkasan + grafik
    ring.append([])
    ring.append(["Metrik"] + ENGINE)
    for sel in ring[3]:
        sel.font = TEBAL
    for tugas in LABEL:
        ring.append([f"Akurasi {tugas}"] + [round(hasil[(tugas, n)]["akurasi"] * 100, 1) for n in ENGINE])
        ring.append([f"Macro-F1 {tugas}"] + [round(hasil[(tugas, n)]["macro_f1"] * 100, 1) for n in ENGINE])
    grafik_batang(ring, (3, 3 + 2 * len(LABEL)), "Perbandingan Akurasi & Macro-F1 (%)", "H2")

    baris = 6 + 2 * len(LABEL)
    _judul(ring, baris, "Cara membaca")
    for teks in [
        "Akurasi = % label engine yang sama dengan label manual.",
        "Macro-F1 = rata-rata F1 semua kelas; lebih adil untuk kelas kecil (mis. Lahan Parkir).",
        "Engine Kata Kunci tidak punya kelas 'Tidak Relevan', sehingga semua tweet tidak relevan pasti salah kategorinya.",
        "IndoBERT, Naive Bayes & SVM dilatih dari label LLM (tanpa sampel uji); keterbatasan ini perlu disebut di metodologi.",
    ]:
        baris += 1
        ring.cell(row=baris, column=1, value=teks)
    baris += 2
    _judul(ring, baris, "Kesepakatan antar-pelabel (Cohen's Kappa)")
    if ada_pelabel2:
        bersama = [i for i in ids if i in pelabel2]
        for tugas, label in LABEL.items():
            k = cohen_kappa([gold[i][tugas] for i in bersama], [pelabel2[i][tugas] for i in bersama], label)
            baris += 1
            ring.cell(row=baris, column=1, value=f"Kappa {tugas}")
            ring.cell(row=baris, column=2, value=round(k, 3))
            ring.cell(row=baris, column=3, value=tafsir_kappa(k))
    else:
        baris += 1
        ring.cell(row=baris, column=1, value="Belum ada lembar Pelabel 2 yang terisi lengkap -> Kappa tidak dihitung.")
    ring.column_dimensions["A"].width = 30
    for kol in "BCDE":
        ring.column_dimensions[kol].width = 14

    # Detail per tugas: metrik per kelas + confusion matrix
    for tugas, label in LABEL.items():
        sheet = buku.create_sheet(f"Detail {tugas}")
        baris = 1
        for nama in ENGINE:
            h = hasil[(tugas, nama)]
            _judul(sheet, baris, f"{nama}: akurasi {h['akurasi'] * 100:.1f}%, macro-F1 {h['macro_f1'] * 100:.1f}%")
            sheet.append(["Kelas", "Precision (%)", "Recall (%)", "F1 (%)", "Jumlah (manual)"])
            for sel in sheet[baris + 1]:
                sel.font = TEBAL
            for k in h["per_kelas"]:
                sheet.append([k["kelas"], round(k["precision"] * 100, 1), round(k["recall"] * 100, 1),
                              round(k["f1"] * 100, 1), k["support"]])
            baris = tulis_matriks(sheet, baris + len(label) + 3, 1, h["matriks"], label,
                                  f"Confusion matrix {tugas} - {nama}")
            baris += 1
        sheet.column_dimensions["A"].width = 24
        for kol in "BCDEF":
            sheet.column_dimensions[kol].width = 22

        # Grafik F1 per kelas, LLM vs Kata Kunci
        sheet.cell(row=1, column=8, value="Kelas").font = TEBAL
        for j, nama in enumerate(ENGINE):
            sheet.cell(row=1, column=9 + j, value=nama).font = TEBAL
        for i, nama_kelas in enumerate(label):
            sheet.cell(row=2 + i, column=8, value=nama_kelas)
            for j, nama in enumerate(ENGINE):
                sheet.cell(row=2 + i, column=9 + j, value=round(hasil[(tugas, nama)]["per_kelas"][i]["f1"] * 100, 1))
        grafik = BarChart()
        grafik.type, grafik.title = "col", f"F1 per Kelas - {tugas} (%)"
        grafik.y_axis.scaling.min, grafik.y_axis.scaling.max = 0, 100
        grafik.y_axis.majorGridlines = None
        grafik.add_data(Reference(sheet, min_col=9, max_col=8 + len(ENGINE), min_row=1, max_row=1 + len(label)),
                        titles_from_data=True)
        grafik.set_categories(Reference(sheet, min_col=8, min_row=2, max_row=1 + len(label)))
        for seri, nama in zip(grafik.series, ENGINE):
            seri.graphicalProperties.solidFill = WARNA_ENGINE[nama]
            seri.graphicalProperties.line.solidFill = WARNA_ENGINE[nama]
        _rapikan_grafik(grafik)
        grafik.legend.position = "r"
        grafik.width, grafik.height = 22, 10
        sheet.add_chart(grafik, "H8")

    # Daftar kesalahan + bukti cara kerja
    salah = buku.create_sheet("Kesalahan & Bukti")
    salah.append(["ID Sampel", "Teks", "Manual Kategori", "Manual Sentimen"]
                 + [f"{n} {t}" for n in ENGINE for t in LABEL] + ["Bukti Kata Kunci", "Catatan Pelabel"])
    for sel in salah[1]:
        sel.font = TEBAL
    for i in ids:
        e, g = engine[i], gold[i]
        if any(e[n][t] != g[t] for n in ENGINE for t in LABEL):
            salah.append([i, e["Teks"], g["Kategori"], g["Sentimen"]]
                         + [e[n][t] for n in ENGINE for t in LABEL] + [bukti.get(i, ""), g["Catatan"]])
    kol_bukti = 5 + 2 * len(ENGINE)
    for b in salah.iter_rows(min_row=2):
        b[1].alignment = b[kol_bukti - 1].alignment = Alignment(wrap_text=True, vertical="top")
    lebar = [11, 60, 20, 12] + [20, 12] * len(ENGINE) + [60, 25]
    for j, w in enumerate(lebar, 1):
        salah.column_dimensions[salah.cell(row=1, column=j).column_letter].width = w
    salah.freeze_panes = "C2"
    salah.auto_filter.ref = salah.dimensions

    if ada_pelabel2:
        beda = buku.create_sheet("Beda Pelabel")
        beda.append(["ID Sampel", "Teks", "P1 Kategori", "P2 Kategori", "P1 Sentimen", "P2 Sentimen"])
        for sel in beda[1]:
            sel.font = TEBAL
        for i in ids:
            if i in pelabel2 and any(gold[i][t] != pelabel2[i][t] for t in LABEL):
                beda.append([i, engine[i]["Teks"], gold[i]["Kategori"], pelabel2[i]["Kategori"],
                             gold[i]["Sentimen"], pelabel2[i]["Sentimen"]])
        beda.column_dimensions["B"].width = 70

    if VERSI in ("v2", "v3"):
        tulis_evaluasi_anotasi(buku, engine, gold, ids)
    tujuan = FOLDER / FILE_LAPORAN
    buku.save(tujuan)
    for tugas in LABEL:
        logger.info("%s: %s", tugas, " | ".join(f"{n} {hasil[(tugas, n)]['akurasi'] * 100:.1f}%" for n in ENGINE))
    logger.info("Laporan: %s", tujuan)


if __name__ == "__main__":
    main()
