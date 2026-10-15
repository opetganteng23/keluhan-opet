"""Halaman laporan evaluasi (HTML): tabel hasil, grafik, heatmap, F1 per kelas, dan rumus + contoh hitungan.

Semua angka dibaca dari laporan_evaluasi_4_metode.xlsx (tidak diketik manual).
Hasil: Downloads/Hasil_Penelitian/06_Evaluasi_Akurasi/laporan_evaluasi.html (grafik dari subfolder grafik/).
Pemakaian:  python evaluasi/buat_laporan_html.py
"""

import html
import logging
from pathlib import Path

from openpyxl import load_workbook

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("laporan_html")

FOLDER = Path.home() / "Downloads" / "Hasil_Penelitian" / "07_Evaluasi_Akurasi"
SUMBER = FOLDER / "laporan_evaluasi_5_metode.xlsx"
RIWAYAT_EPOCH = Path.home() / "Downloads" / "Hasil_Penelitian" / "06_Metode_5_IndoBERT" / "riwayat_pelatihan_per_epoch.xlsx"
TUGAS = ["Kategori", "Sentimen", "Keluhan"]
ENGINE = ["LLM", "IndoBERT", "SVM", "Naive Bayes", "Kata Kunci"]
WARNA = {"LLM": "#2a78d6", "Kata Kunci": "#eb6834", "Naive Bayes": "#1baf7a", "SVM": "#eda100", "IndoBERT": "#e87ba4"}


def baca() -> dict:
    """Ringkasan, detail per kelas + confusion matrix per tugas/engine, dan anotasi LLM."""
    buku = load_workbook(SUMBER, read_only=True)
    ring = {}
    for r in buku["Ringkasan"].iter_rows(min_row=4, max_row=3 + 2 * len(TUGAS), values_only=True):
        ring[r[0]] = dict(zip(ENGINE, r[1:1 + len(ENGINE)]))
    detail = {}
    for tugas in TUGAS:
        baris = [r for r in buku[f"Detail {tugas}"].iter_rows(max_col=7, values_only=True)]
        i, d = 0, {}
        while i < len(baris):
            judul = str(baris[i][0] or "")
            nama = judul.split(":")[0]
            if nama in ENGINE and "akurasi" in judul:
                kelas, j = [], i + 2
                while j < len(baris) and baris[j][0] and isinstance(baris[j][1], (int, float)):
                    kelas.append(baris[j][:5])
                    j += 1
                while j < len(baris) and str(baris[j][0] or "") != "Manual \\ Engine":
                    j += 1
                mulai = j + 1  # baris pertama matriks setelah header "Manual \ Engine"
                matriks = [list(baris[k][1:1 + len(kelas)]) for k in range(mulai, mulai + len(kelas))]
                d[nama] = {"judul": judul, "kelas": kelas, "matriks": matriks}
                i = mulai + len(kelas)
            else:
                i += 1
        detail[tugas] = d
    anot = [r for r in buku["LLM Sebab-Akibat-Emosi"].iter_rows(min_row=1, max_row=5, values_only=True)]
    epoch = [r for r in load_workbook(RIWAYAT_EPOCH, read_only=True)["Riwayat Epoch"].iter_rows(values_only=True)]
    return {"ring": ring, "detail": detail, "anotasi": anot, "epoch": epoch}


def tabel_ringkasan(ring: dict) -> str:
    """Tabel akurasi & macro-F1 semua engine; nilai tertinggi per baris ditebalkan."""
    kepala = "".join(f'<th><span class="dot" style="background:{WARNA[e]}"></span>{e}</th>' for e in ENGINE)
    isi = []
    for metrik, nilai in ring.items():
        top = max(nilai.values())
        sel = "".join(f'<td class="num{" best" if v == top else ""}">{v:.1f}</td>' for v in nilai.values())
        isi.append(f"<tr><th scope='row'>{html.escape(metrik)} (%)</th>{sel}</tr>")
    return f'<div class="tbl"><table><thead><tr><th>Metrik</th>{kepala}</tr></thead><tbody>{"".join(isi)}</tbody></table></div>'


def tabel_kelas(det: dict) -> str:
    """Precision/recall/F1 per kelas untuk satu tugas, semua engine berdampingan."""
    kelas = [k[0] for k in det["LLM"]["kelas"]]
    kepala = "".join(f'<th colspan="3"><span class="dot" style="background:{WARNA[e]}"></span>{e}</th>' for e in ENGINE)
    sub = "".join("<th>P</th><th>R</th><th>F1</th>" for _ in ENGINE)
    isi = []
    for i, k in enumerate(kelas):
        sel = ""
        for e in ENGINE:
            _, p, r, f1, n = det[e]["kelas"][i]
            sel += f'<td class="num">{p:.1f}</td><td class="num">{r:.1f}</td><td class="num f1">{f1:.1f}</td>'
        n = det["LLM"]["kelas"][i][4]
        isi.append(f"<tr><th scope='row'>{html.escape(k)} <span class='n'>n={n}</span></th>{sel}</tr>")
    return (f'<div class="tbl"><table class="wide"><thead><tr><th rowspan="2">Kelas</th>{kepala}</tr>'
            f'<tr>{sub}</tr></thead><tbody>{"".join(isi)}</tbody></table></div>')


def main() -> None:
    """Tulis halaman HTML."""
    d = baca()
    llm_kel = d["detail"]["Keluhan"]["LLM"]
    (tp, fn), (fp, tn) = llm_kel["matriks"]
    n = tp + fn + fp + tn
    p_ya, r_ya = tp / (tp + fp), tp / (tp + fn)
    f1_ya = 2 * p_ya * r_ya / (p_ya + r_ya)
    p_t, r_t = tn / (tn + fn), tn / (tn + fp)
    f1_t = 2 * p_t * r_t / (p_t + r_t)
    akurasi = (tp + tn) / n
    anot = d["anotasi"]
    ep = d["epoch"]
    kol = {k: i for i, k in enumerate(ep[0])}
    baris_ep = "".join(
        f"<tr{' class=\'best\'' if r[kol['terbaik']] else ''}><th scope='row'>{r[kol['tugas']]}</th><td class='num'>{r[kol['epoch']]}</td>"
        f"<td class='num'>{r[kol['loss_latih']]:.3f}</td><td class='num'>{r[kol['loss_validasi']]:.3f}</td>"
        f"<td class='num'>{r[kol['akurasi_validasi']]*100:.1f}</td><td class='num'>{r[kol['macro_f1_validasi']]*100:.1f}</td></tr>"
        for r in ep[1:])
    tabel_epoch = ('<div class="tbl"><table><thead><tr><th>Tugas</th><th class="num">Epoch</th><th class="num">Loss latih</th>'
                   '<th class="num">Loss validasi</th><th class="num">Akurasi val (%)</th><th class="num">Macro-F1 val (%)</th>'
                   f'</tr></thead><tbody>{baris_ep}</tbody></table></div>')

    seksi_tugas = []
    nomor = {"Kategori": (2, 3), "Sentimen": (4, 5), "Keluhan": (6, 7)}
    keterangan = {
        "Kategori": "Empat kelas: Angkutan Umum, Lahan Parkir, Manajemen Lalu Lintas, Tidak Relevan.",
        "Sentimen": "Tiga kelas: Positif, Negatif, Netral.",
        "Keluhan": "Dua kelas: Ya (tweet berisi keluhan) dan Tidak.",
    }
    for t in TUGAS:
        f1_png, cm_png = nomor[t]
        seksi_tugas.append(f"""
<section id="{t.lower()}">
  <h2>{t}</h2>
  <p class="lead">{keterangan[t]} P = precision, R = recall, dalam persen.</p>
  {tabel_kelas(d["detail"][t])}
  <figure><img src="grafik/{f1_png}_f1_per_kelas_{t.lower()}.png" alt="Grafik batang F1 per kelas untuk tugas {t}, empat metode" loading="lazy">
    <figcaption>F1-score per kelas, tugas {t}.</figcaption></figure>
  <figure><img src="grafik/{cm_png}_confusion_matrix_{t.lower()}.png" alt="Heatmap confusion matrix tugas {t} untuk empat metode" loading="lazy">
    <figcaption>Confusion matrix (heatmap). Baris = label manual, kolom = label metode, warna = persentase per baris. Diagonal = tebakan benar.</figcaption></figure>
</section>""")

    halaman = f"""<title>Evaluasi Klasifikasi Keluhan</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,600;8..60,700&family=Public+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>
/* Laporan riset: satu kolom baca, tabel & figur selebar konten, rumus sebagai blok. */
:root {{
  --bg:#f7f8fa; --surface:#ffffff; --fg:#14181f; --muted:#5a6270; --line:#dde1e7; --accent:#2a78d6; --best:#e7f0fc;
  --display:"Source Serif 4", Georgia, serif; --body:"Public Sans", system-ui, sans-serif; --mono:"IBM Plex Mono", ui-monospace, monospace;
}}
@media (prefers-color-scheme: dark) {{ :root:not([data-theme="light"]) {{
  --bg:#12151a; --surface:#1a1e25; --fg:#e8eaee; --muted:#a0a8b5; --line:#2d333d; --accent:#5b9bea; --best:#1f3150; color-scheme:dark }} }}
:root[data-theme="dark"] {{ --bg:#12151a; --surface:#1a1e25; --fg:#e8eaee; --muted:#a0a8b5; --line:#2d333d; --accent:#5b9bea; --best:#1f3150; color-scheme:dark }}
body {{ background:var(--bg); color:var(--fg); font:15px/1.6 var(--body); }}
main {{ max-width:960px; margin:0 auto; padding-inline:16px; padding-block:32px 64px; display:grid; gap:40px; }}
h1,h2,h3 {{ font-family:var(--display); text-wrap:balance; line-height:1.2; margin:0; }}
h1 {{ font-size:clamp(1.7rem,4vw,2.3rem); }} h2 {{ font-size:1.45rem; padding-bottom:6px; border-bottom:1px solid var(--line); }}
h3 {{ font-size:1.1rem; }}
section {{ display:grid; gap:16px; min-width:0; }}
p {{ margin:0; max-width:68ch; }} .lead {{ color:var(--muted); }}
.meta {{ color:var(--muted); font-size:.9rem; }}
nav {{ display:flex; flex-wrap:wrap; gap:8px 16px; font-size:.9rem; }}
nav a {{ color:var(--accent); text-decoration:none; }} nav a:hover, nav a:focus-visible {{ text-decoration:underline; }}
.tbl {{ overflow-x:auto; border:1px solid var(--line); border-radius:6px; background:var(--surface); }}
table {{ border-collapse:collapse; width:100%; font-size:.9rem; }}
table.wide {{ min-width:880px; }}
th,td {{ padding:7px 10px; border-bottom:1px solid var(--line); text-align:left; white-space:nowrap; }}
thead th {{ font-weight:600; font-size:.8rem; letter-spacing:.02em; color:var(--muted); }}
tbody tr:last-child th, tbody tr:last-child td {{ border-bottom:0; }}
.num {{ text-align:right; font-variant-numeric:tabular-nums; font-family:var(--mono); }}
.best {{ background:var(--best); font-weight:600; }} .f1 {{ font-weight:600; }}
.n {{ color:var(--muted); font-weight:400; font-size:.8rem; }}
.dot {{ display:inline-block; width:9px; height:9px; border-radius:2px; margin-right:6px; vertical-align:baseline; }}
figure {{ margin:0; display:grid; gap:8px; }}
figure img {{ width:100%; height:auto; background:#fff; border:1px solid var(--line); border-radius:6px; }}
figcaption {{ color:var(--muted); font-size:.85rem; }}
.rumus {{ display:grid; gap:14px; }}
.r {{ background:var(--surface); border:1px solid var(--line); border-radius:6px; padding:14px 16px; display:grid; gap:6px; min-width:0; }}
.r .math {{ overflow-x:auto; }}
.r p {{ font-size:.92rem; }}
.calc {{ font-family:var(--mono); font-size:.85rem; color:var(--muted); overflow-x:auto; }}
.grid2 {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,300px),1fr)); gap:14px; }}
ul {{ margin:0; padding-left:20px; max-width:68ch; }}
</style>
<script>window.MathJax={{tex:{{inlineMath:[["\\\\(","\\\\)"]]}},svg:{{fontCache:"global"}}}};</script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/mathjax/3.2.2/es5/tex-svg.min.js"></script>
<main>
<header style="display:grid;gap:10px">
  <h1>Evaluasi Metode Klasifikasi Keluhan Transportasi Jakarta</h1>
  <p class="lead">Lima metode diuji pada <strong>{n} sampel uji</strong> dan dibandingkan dengan label manual peneliti. Tiga tugas: kategori, sentimen, dan keluhan.</p>
  <nav aria-label="Isi"><a href="#ringkasan">Ringkasan</a><a href="#kategori">Kategori</a><a href="#sentimen">Sentimen</a><a href="#keluhan">Keluhan</a><a href="#anotasi">Sebab-akibat-emosi</a><a href="#indobert">Epoch IndoBERT</a><a href="#rumus">Rumus</a><a href="#catatan">Catatan</a></nav>
</header>

<section id="ringkasan">
  <h2>Ringkasan</h2>
  <p>LLM memberi hasil tertinggi di ketiga tugas, diikuti IndoBERT. Sel bersorot = nilai tertinggi pada baris itu.</p>
  {tabel_ringkasan(d["ring"])}
  <figure><img src="grafik/1_perbandingan_metode.png" alt="Grafik batang akurasi dan macro-F1 empat metode untuk tugas kategori, sentimen, dan keluhan">
    <figcaption>Akurasi dan macro-F1 keempat metode terhadap label manual (n = {n}).</figcaption></figure>
</section>
{"".join(seksi_tugas)}
<section id="anotasi">
  <h2>Anotasi sebab, akibat, dan emosi (LLM)</h2>
  <p class="lead">Dihitung pada {anot[2][1].split('/')[1] if False else anot[3][1].split('/')[1]} tweet yang dinilai keluhan oleh peneliti dan LLM.</p>
  <div class="tbl"><table><thead><tr><th>Aspek</th><th>Cocok</th><th class="num">Persen</th></tr></thead><tbody>
    {"".join(f"<tr><th scope='row'>{html.escape(str(r[0]))}</th><td class='num'>{r[1]}</td><td class='num'>{r[2]:.1f}</td></tr>" for r in anot[2:5])}
  </tbody></table></div>
</section>

<section id="indobert">
  <h2>Pelatihan IndoBERT per epoch</h2>
  <p class="lead">Model dipilih dari epoch dengan macro-F1 validasi tertinggi (garis putus). Loss validasi yang naik setelah epoch 2 sampai 3 menandakan overfitting.</p>
  {tabel_epoch}
  <figure><img src="grafik/8_epoch_indobert.png" alt="Grafik loss dan akurasi latih serta validasi IndoBERT per epoch untuk tugas kategori, sentimen, dan keluhan" loading="lazy">
    <figcaption>Loss dan akurasi data latih dan validasi per epoch, untuk tiap tugas.</figcaption></figure>
</section>

<section id="rumus">
  <h2>Rumus</h2>
  <p class="lead">Contoh hitungan memakai confusion matrix tugas <strong>Keluhan</strong> metode <strong>LLM</strong>: TP = {tp}, FN = {fn}, FP = {fp}, TN = {tn} (kelas positif = "Ya").</p>
  <div class="rumus">
    <div class="r"><h3>Ukuran sampel (Slovin)</h3>
      <div class="math">\\[ n = \\frac{{N}}{{1 + N e^2}} \\qquad e = \\sqrt{{\\frac{{N - n}}{{N\\,n}}}} \\]</div>
      <p>N = populasi tweet bersih, e = batas kesalahan. Sampel awal 385 (e = 5% dari N = 10.079), lalu ditambah 615 sampel acak menjadi 1.000.</p>
      <div class="calc">e = √((19.016 − 1.000) / (19.016 × 1.000)) = 0,031 → margin of error ±3,1%</div></div>
    <div class="grid2">
      <div class="r"><h3>Akurasi</h3>
        <div class="math">\\[ \\text{{Akurasi}} = \\frac{{TP + TN}}{{TP + TN + FP + FN}} \\]</div>
        <div class="calc">({tp} + {tn}) / {n} = {akurasi*100:.1f}%</div></div>
      <div class="r"><h3>Precision</h3>
        <div class="math">\\[ \\text{{Precision}} = \\frac{{TP}}{{TP + FP}} \\]</div>
        <p>Dari yang ditebak "Ya", berapa yang benar.</p>
        <div class="calc">{tp} / ({tp} + {fp}) = {p_ya*100:.1f}%</div></div>
      <div class="r"><h3>Recall</h3>
        <div class="math">\\[ \\text{{Recall}} = \\frac{{TP}}{{TP + FN}} \\]</div>
        <p>Dari semua yang sebenarnya "Ya", berapa yang ditemukan.</p>
        <div class="calc">{tp} / ({tp} + {fn}) = {r_ya*100:.1f}%</div></div>
      <div class="r"><h3>F1-score</h3>
        <div class="math">\\[ F_1 = 2 \\times \\frac{{\\text{{Precision}} \\times \\text{{Recall}}}}{{\\text{{Precision}} + \\text{{Recall}}}} \\]</div>
        <div class="calc">2 × ({p_ya:.3f} × {r_ya:.3f}) / ({p_ya:.3f} + {r_ya:.3f}) = {f1_ya*100:.1f}%</div></div>
    </div>
    <div class="r"><h3>Macro-F1</h3>
      <div class="math">\\[ \\text{{Macro-}}F_1 = \\frac{{1}}{{K}} \\sum_{{k=1}}^{{K}} F_{{1,k}} \\]</div>
      <p>Rata-rata F1 semua kelas (K = jumlah kelas). Lebih adil untuk kelas kecil seperti Lahan Parkir karena tiap kelas berbobot sama.</p>
      <div class="calc">F1(Ya) = {f1_ya*100:.1f}%, F1(Tidak) = {f1_t*100:.1f}% → ({f1_ya*100:.1f} + {f1_t*100:.1f}) / 2 = {(f1_ya+f1_t)/2*100:.1f}%</div></div>
    <div class="r"><h3>Cohen's Kappa (kesepakatan antar-pelabel)</h3>
      <div class="math">\\[ \\kappa = \\frac{{p_o - p_e}}{{1 - p_e}} \\]</div>
      <p>p<sub>o</sub> = proporsi label yang sama antara dua pelabel; p<sub>e</sub> = kesepakatan yang diharapkan secara kebetulan. Interpretasi Landis &amp; Koch: &gt; 0,6 substansial, &gt; 0,8 hampir sempurna. Belum dilaporkan karena belum ada pelabel kedua yang independen.</p></div>
    <div class="r"><h3>Confusion matrix</h3>
      <p>Tabel baris = label manual, kolom = label metode. Sel diagonal adalah tebakan benar; sel di luar diagonal menunjukkan ke kelas mana metode keliru. Heatmap pada laporan diwarnai menurut persentase per baris.</p></div>
  </div>
</section>

<section id="catatan">
  <h2>Catatan dan keterbatasan</h2>
  <ul>
    <li>Label manual dibuat oleh satu peneliti tanpa melihat label metode (blind labeling).</li>
    <li>IndoBERT, Naive Bayes, dan SVM dilatih dari label LLM pada dataset 10.000 tweet (9.467 tweet setelah sampel uji dikeluarkan).</li>
    <li>615 sampel tambahan dilabeli ulang karena pelabelan pertama tidak konsisten dengan pedoman; label pertama diarsipkan dan tidak dipakai.</li>
    <li>Kata Kunci tidak mengenal kelas Tidak Relevan dan konsep keluhan; keluhannya didefinisikan sebagai sentimen Negatif.</li>
    <li>LLM cenderung menilai tweet sebagai keluhan: recall keluhan {r_ya*100:.1f}% tetapi precision {p_ya*100:.1f}%.</li>
    <li>Prompt LLM disetel pada 100 tweet kalibrasi yang terpisah dari sampel uji (versi 3.1, temperature 0).</li>
  </ul>
</section>
</main>
"""
    tujuan = FOLDER / "laporan_evaluasi.html"
    tujuan.write_text(halaman, encoding="utf-8")
    logger.info("Laporan HTML: %s", tujuan)


if __name__ == "__main__":
    main()
