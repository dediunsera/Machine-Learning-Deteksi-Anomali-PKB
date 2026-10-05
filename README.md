# Penelitian 4 — Deteksi Anomali Transaksi & Kualitas Data SAMSAT (PKB Banten)


[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/dediunsera/Machine-Learning-Deteksi-Anomali-PKB/blob/main/PKB_Anomali_SAMSAT_Colab.ipynb)
Penulis: **Ahmad Dedi Jubaedi** (korespondensi, Universitas Serang Raya), **Saleh Dwiyatno** (Universitas Serang Raya), **Rahmat** (AMIK Serang)
Jurnal sasaran: **INOVATIF – Jurnal Inovasi Teknologi Informasi dan Informatika** (Universitas Ibn Khaldun Bogor), template `Template Inovatif 2022 ok.doc`.

## Ringkasan
| Item | Isi |
|---|---|
| Data | 4 file `DATA PKB BULAN MEI - JUNI TAHUN 2021..2024.xlsx` (ISQL-WEB SAMSAT Banten), 193.899 transaksi, 24 atribut |
| Metode | Lapis 1: 11 aturan domain (R01–R11) · Lapis 2: Isolation Forest + LOF + Autoencoder (fusi maksimum persentil) · Lapis 3: skor hibrida + kode alasan |
| Tambahan | Deteksi pergeseran skema kode (JSD) & harmonisasi; kartu skor kualitas data 6 dimensi |
| Evaluasi | Latih 2021–2023, uji 2024; injeksi 8 tipe anomali × 100 × 10 ulangan; ROC-AUC, PR-AUC, P@k, Recall@5%, F1; Friedman + Wilcoxon-Holm; sensitivitas; ablasi |

## Temuan utama
- **Hibrida PR-AUC 0,973 ± 0,006; ROC-AUC 0,996** vs Aturan 0,817, AE 0,862, Ensemble ML 0,748, IF 0,178, LOF 0,133 (Friedman p = 1,4×10⁻⁹; hibrida menang 10/10 ulangan, Wilcoxon-Holm p = 0,0098).
- Aturan presisi 1,000 tetapi recall 0,809; gagal total pada anomali tanpa aturan (A8 BBNKB janggal, recall 0,01) yang ditangkap AE (0,99) → hibrida 0,83.
- Fitur domain: PR-AUC ensemble 0,54 → 0,76; AE 0,68 → 0,87; hibrida 0,92 → 0,97.
- **Pergeseran kode 2024**: kd_status 3/4 → 5/6 (JSD = 1,0) dan daftar ulang+mutasi 1|x → 2|x (JSD = 0,19); setelah harmonisasi JSD ≈ 0,001–0,007. Penggabungan makna kode 2|4 memicu lonjakan R08 (775 transaksi) pada 2024.
- Data riil: 10.103 transaksi (5,21%) melanggar aturan; R06 (tunggakan tanpa pokok berjalan) naik 1,34% → 2,92%; 7.051 transaksi bernilai nol; 2.068 transaksi ditandai ML saja. Daftar audit hibrida 5,65–7,64% transaksi per tahun.

## Struktur folder
```
project4/
├── README.md, requirements.txt
├── Template Inovatif 2022 ok.doc / .docx (hasil konversi LibreOffice, dipakai skrip naskah)
├── src/
│   ├── pkb_anomaly_pipeline.py   ← seluruh eksperimen (sel # %%), seed 42, gambar ID & EN
│   ├── build_manuscript.py       ← menyusun naskah DOCX (ID & EN) + suplemen langsung dari tabel hasil
│   ├── content_id.py / content_en.py / refs.py
├── notebooks/PKB_Anomali_SAMSAT_Colab.ipynb   ← cek silang di Google Colab
├── data_colab/  dataset input terpseudonimisasi untuk Colab (193.899 transaksi, 25 kolom; hasil identik dengan file xlsx mentah)
├── data_processed/pkb_2021_2024_scored_pseudonymised.csv.gz  (tanpa nopol/nama/alamat; skor & bendera per transaksi)
├── results/
│   ├── tables/  T01–T13 (tabel utama), S01–S08 (pendukung)
│   ├── figures/ (ID) & figures_en/ (EN): F01 flowchart … F16 pergeseran kode
│   ├── run_summary.json, run_log.txt
└── manuscript/
    ├── INOVATIF_Deteksi_Anomali_SAMSAT_PKB_Banten_ID.docx (+ .pdf pratinjau)
    ├── INOVATIF_SAMSAT_Anomaly_Detection_PKB_Banten_EN.docx (+ .pdf pratinjau)
    └── Suplemen_…_ID.docx / Supplementary_…_EN.docx (gambar & tabel tambahan)
```

## Daftar gambar
F01 flowchart metode · F02 arsitektur sistem · F03 prevalensi aturan · F04 kartu skor DQ · F05 nilai PKB per jenis (R09) · F06 matriks kode mohon×mutasi · F07 ROC & PR · F08 recall per tipe anomali · F09 komparasi metrik · F10 distribusi skor · F11 sensitivitas · F12 ablasi fitur · F13 temuan data riil · F14 proyeksi PCA · F15 waktu komputasi · F16 pergeseran kode

## Cara menjalankan
```
pip install -r requirements.txt
python src/pkb_anomaly_pipeline.py --data_dir "/path/PKB Banten" --out_dir "/path/project4"    # ±9 menit (CPU 2 inti); --fast untuk uji cepat
python src/build_manuscript.py --template "Template Inovatif 2022 ok.docx" --lang both
```
**Google Colab:** struktur Drive `MyDrive/Penelitian/INOVATIF_Deteksi_Anomali/` → `data/` (pkb_2021_2024_input_pseudonim.csv.gz + pkb_inventaris_mentah.csv, atau 4 xlsx mentah dengan `FROM_XLSX = True`), `src/` (semua .py), `project4/` (hasil; taruh juga `Template Inovatif 2022 ok.docx` di sini untuk menyusun naskah). Buka `notebooks/PKB_Anomali_SAMSAT_Colab.ipynb` → *Run all*.

## Catatan metodologis
- Kamus kode resmi tidak tersedia → semantik kode (R07, R08) diturunkan dari profil 2021–2023; perlu dikonfirmasi petugas SAMSAT/Bapenda.
- Ground truth tidak ada → evaluasi kuantitatif memakai injeksi anomali sintetis; temuan pada data riil adalah **kandidat** untuk diverifikasi, bukan kesalahan terbukti.
- Fusi ensemble memakai maksimum persentil; fusi rerata dilaporkan di analisis sensitivitas (PR-AUC 0,556 vs 0,790 pada seed 0) agar pilihan desain transparan.
- Transaksi bernilai nol dipisahkan sebagai kategori informatif (kemungkinan pengesahan tanpa pembayaran).

## Yang perlu dicek penulis sebelum submit
1. Naskah 11 halaman (batas template 4–10 halaman): bila editor ketat, pindahkan Gambar 2 (arsitektur) atau Tabel 7 ke suplemen.
2. Alamat afiliasi AMIK Serang (jalan & kode pos) belum dicantumkan; email Rahmat diambil dari biografi.docx (rahmat042@gmail.com).
3. 27 referensi jurnal 2021–2026 + 1 regulasi (UU 1/2022). DOI sudah dicantumkan; mohon cek ulang volume/halaman lewat Mendeley/Crossref sebelum submit.
4. Konfirmasi ke Bapenda: dugaan kebijakan relaksasi denda 2021–2022 (R02/R03) dan makna kode baru kd_status 5/6 serta kd_jen_mohon 2 pada 2024.
5. Izin penggunaan data: data diperoleh untuk eksperimen melalui petugas SAMSAT; pertimbangkan surat izin resmi sebelum publikasi.
