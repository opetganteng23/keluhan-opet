"""Test rumus metrik evaluasi dengan contoh yang hasilnya sudah diketahui."""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "evaluasi"))

from hitung_evaluasi import akurasi, cohen_kappa, macro_f1, matriks_kebingungan, metrik_per_kelas  # noqa: E402

LABEL = ["A", "B"]
BENAR = ["A", "A", "A", "B", "B"]
TEBAK = ["A", "A", "B", "B", "A"]


def test_matriks_kebingungan():
    """Baris = benar, kolom = tebakan."""
    assert matriks_kebingungan(BENAR, TEBAK, LABEL) == [[2, 1], [1, 1]]


def test_akurasi():
    """3 dari 5 benar = 60%."""
    assert akurasi(matriks_kebingungan(BENAR, TEBAK, LABEL)) == pytest.approx(0.6)


def test_precision_recall_f1():
    """Kelas A: P=2/3, R=2/3, F1=2/3. Kelas B: P=1/2, R=1/2, F1=1/2."""
    pk = metrik_per_kelas(matriks_kebingungan(BENAR, TEBAK, LABEL), LABEL)
    assert pk[0]["precision"] == pytest.approx(2 / 3) and pk[0]["recall"] == pytest.approx(2 / 3)
    assert pk[1]["f1"] == pytest.approx(0.5)
    assert macro_f1(pk) == pytest.approx((2 / 3 + 0.5) / 2)


def test_kelas_tanpa_data_tidak_dihitung_di_macro_f1():
    """Kelas yang tidak pernah muncul di label manual tidak menurunkan macro-F1."""
    pk = metrik_per_kelas(matriks_kebingungan(["A", "A"], ["A", "A"], ["A", "B"]), ["A", "B"])
    assert macro_f1(pk) == pytest.approx(1.0)


def test_cohen_kappa():
    """Contoh buku teks: po=0.7, pe=0.5 -> kappa 0.4; sepakat penuh -> 1."""
    a = ["Y"] * 5 + ["N"] * 5
    b = ["Y"] * 4 + ["N"] + ["Y"] * 2 + ["N"] * 3
    assert cohen_kappa(a, b, ["Y", "N"]) == pytest.approx(0.4)
    assert cohen_kappa(a, a, ["Y", "N"]) == pytest.approx(1.0)
