"""Grafik evaluasi untuk paper (PNG 300 dpi): perbandingan metode, heatmap confusion matrix, F1 per kelas.

Membaca file yang sama dengan hitung_evaluasi.py (hasil_engine_sampel.xlsx + lembar_label_pelabel1.xlsx)
dan menyimpan gambar ke Downloads/evaluasi_akurasi/grafik/.

Pemakaian:  python evaluasi/buat_grafik.py
"""

import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import matplotlib

matplotlib.use("Agg")  # tanpa jendela, langsung ke file
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import LinearSegmentedColormap

import hitung_evaluasi as he

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("grafik")

FOLDER_GRAFIK = he.FOLDER / (f"grafik_{he.VERSI}" if he.VERSI in ("v2", "v3") else "grafik")
DPI = 300
TINTA, TINTA_2, GARIS = "#0b0b0b", "#52514e", "#d9d8d4"
WARNA = {n: "#" + w.lower() for n, w in he.WARNA_ENGINE.items()}
# Heatmap: satu warna (biru), terang -> gelap
PETA_BIRU = LinearSegmentedColormap.from_list("biru", ["#fcfcfb", "#cde2fb", "#86b6ef", "#2a78d6", "#184f95"])
SINGKAT = {"Angkutan Umum": "Angkutan\nUmum", "Lahan Parkir": "Lahan\nParkir",
           "Manajemen Lalu Lintas": "Manajemen\nLalu Lintas", "Tidak Relevan": "Tidak\nRelevan"}

plt.rcParams.update({
    "font.size": 10, "axes.edgecolor": GARIS, "axes.labelcolor": TINTA_2, "xtick.color": TINTA_2,
    "ytick.color": TINTA_2, "axes.titlecolor": TINTA, "axes.titleweight": "bold", "axes.titlesize": 11,
    "figure.facecolor": "white", "axes.facecolor": "white", "savefig.bbox": "tight",
})


def hitung_semua() -> tuple[dict, int]:
    """Metrik semua engine x tugas (memakai fungsi di hitung_evaluasi.py)."""
    engine, _ = he.baca_engine(he.FOLDER / he.FILE_ENGINE)
    gold = he.baca_lembar(he.FOLDER / he.FILE_GOLD)
    he.validasi(gold, "Pelabel 1")
    ids = sorted(gold)
    hasil = {}
    for tugas, label in he.LABEL.items():
        benar = [gold[i][tugas] for i in ids]
        for nama in he.ENGINE:
            m = he.matriks_kebingungan(benar, [engine[i][nama][tugas] for i in ids], label)
            pk = he.metrik_per_kelas(m, label)
            hasil[(tugas, nama)] = {"matriks": m, "per_kelas": pk, "akurasi": he.akurasi(m),
                                    "macro_f1": he.macro_f1(pk)}
    return hasil, len(ids)


def _gaya_sumbu(ax) -> None:
    """Sumbu minimalis: tanpa bingkai atas/kanan, grid horizontal tipis di belakang."""
    for sisi in ("top", "right", "left"):
        ax.spines[sisi].set_visible(False)
    ax.yaxis.grid(True, color=GARIS, linewidth=0.6)
    ax.set_axisbelow(True)
    ax.tick_params(axis="y", length=0)


def batang_berkelompok(ax, kelompok: list[str], nilai: dict[str, list[float]]) -> None:
    """Batang berkelompok per engine dengan label angka di ujung batang."""
    n = len(nilai)
    lebar = 0.8 / n
    x = np.arange(len(kelompok))
    for j, (nama, angka) in enumerate(nilai.items()):
        posisi = x - 0.4 + lebar * (j + 0.5)
        batang = ax.bar(posisi, angka, width=lebar - 0.02, color=WARNA[nama], label=nama,
                        edgecolor="white", linewidth=1)
        for b, v in zip(batang, angka):
            ax.text(b.get_x() + b.get_width() / 2, v + 1.2, f"{v:.1f}", ha="center", va="bottom",
                    fontsize=7.5, color=TINTA)
    ax.set_xticks(x, kelompok)
    ax.set_ylim(0, 105)
    ax.set_ylabel("Persen (%)")
    _gaya_sumbu(ax)


def grafik_perbandingan(hasil: dict, n: int) -> Path:
    """Akurasi & macro-F1 per tugas, keempat metode."""
    kelompok = [f"{m}\n{t}" for t in he.LABEL for m in ("Akurasi", "Macro-F1")]
    nilai = {nama: [hasil[(t, nama)][k] * 100 for t in he.LABEL for k in ("akurasi", "macro_f1")]
             for nama in he.ENGINE}
    fig, ax = plt.subplots(figsize=(4.5 + 1.5 * len(kelompok), 4.6))
    batang_berkelompok(ax, kelompok, nilai)
    ax.set_title(f"Perbandingan kinerja metode terhadap label manual (n = {n})", loc="left")
    ax.legend(ncol=len(nilai), frameon=False, loc="upper left", bbox_to_anchor=(0, -0.14))
    return simpan(fig, "1_perbandingan_metode.png")


def grafik_f1_per_kelas(hasil: dict, tugas: str) -> Path:
    """F1 per kelas untuk satu tugas, keempat metode."""
    label = he.LABEL[tugas]
    nilai = {nama: [k["f1"] * 100 for k in hasil[(tugas, nama)]["per_kelas"]] for nama in he.ENGINE}
    fig, ax = plt.subplots(figsize=(9, 4.6))
    batang_berkelompok(ax, [SINGKAT.get(k, k) for k in label], nilai)
    ax.set_title(f"F1-score per kelas – {tugas}", loc="left")
    ax.legend(ncol=len(nilai), frameon=False, loc="upper left", bbox_to_anchor=(0, -0.16))
    nomor = 2 + 2 * list(he.LABEL).index(tugas)
    return simpan(fig, f"{nomor}_f1_per_kelas_{tugas.lower()}.png")


def grafik_confusion(hasil: dict, tugas: str) -> Path:
    """Heatmap confusion matrix 2x2 (satu panel per metode); warna = % per baris, angka = jumlah + %."""
    label = he.LABEL[tugas]
    nama_singkat = [SINGKAT.get(k, k) for k in label]
    kolom = 3 if len(he.ENGINE) > 4 else 2
    baris = int(np.ceil(len(he.ENGINE) / kolom))
    ukuran = 4.3 if len(label) == 4 else 3.6
    fig, axs = plt.subplots(baris, kolom, figsize=(ukuran * kolom + 1, ukuran * baris), squeeze=False)
    for ax, nama in zip(axs.flat, he.ENGINE):
        m = np.array(hasil[(tugas, nama)]["matriks"], dtype=float)
        persen = m / np.maximum(m.sum(axis=1, keepdims=True), 1) * 100
        ax.imshow(persen, cmap=PETA_BIRU, vmin=0, vmax=100)
        for i in range(len(label)):
            for j in range(len(label)):
                gelap = persen[i, j] > 55
                ax.text(j, i, f"{int(m[i, j])}\n({persen[i, j]:.0f}%)", ha="center", va="center", fontsize=8,
                        color="white" if gelap else TINTA, fontweight="bold" if i == j else "normal")
        ax.set_xticks(range(len(label)), nama_singkat, fontsize=8)
        ax.set_yticks(range(len(label)), nama_singkat, fontsize=8)
        ax.set_xlabel("Label metode")
        ax.set_ylabel("Label manual")
        h = hasil[(tugas, nama)]
        ax.set_title(f"{nama}  (akurasi {h['akurasi'] * 100:.1f}%, macro-F1 {h['macro_f1'] * 100:.1f}%)",
                     fontsize=9.5, loc="left")
        for sisi in ax.spines.values():
            sisi.set_visible(False)
        ax.set_xticks(np.arange(-0.5, len(label)), minor=True)
        ax.set_yticks(np.arange(-0.5, len(label)), minor=True)
        ax.grid(which="minor", color="white", linewidth=2)
        ax.tick_params(which="both", length=0)
    for ax in list(axs.flat)[len(he.ENGINE):]:
        ax.axis("off")
    fig.suptitle(f"Confusion matrix – {tugas} (warna = % per baris label manual)", x=0.02, ha="left",
                 fontweight="bold", color=TINTA)
    fig.tight_layout()
    nomor = 3 + 2 * list(he.LABEL).index(tugas)
    return simpan(fig, f"{nomor}_confusion_matrix_{tugas.lower()}.png")


def simpan(fig, nama: str) -> Path:
    """Simpan PNG 300 dpi lalu tutup gambar."""
    FOLDER_GRAFIK.mkdir(parents=True, exist_ok=True)
    tujuan = FOLDER_GRAFIK / nama
    fig.savefig(tujuan, dpi=DPI)
    plt.close(fig)
    return tujuan


def main() -> None:
    """Buat semua grafik evaluasi."""
    hasil, n = hitung_semua()
    berkas = [grafik_perbandingan(hasil, n)]
    for tugas in he.LABEL:
        berkas.append(grafik_f1_per_kelas(hasil, tugas))
        berkas.append(grafik_confusion(hasil, tugas))
    for b in sorted(berkas):
        logger.info("Grafik: %s", b)


if __name__ == "__main__":
    main()
