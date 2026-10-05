# -*- coding: utf-8 -*-
"""Isi naskah Bahasa Indonesia (INOVATIF). Semua angka diambil dari tabel hasil (N)."""
import os
import pandas as pd
from docx.enum.text import WD_ALIGN_PARAGRAPH

def build(b, N, ROOT):
    L = "id"; F = lambda n: os.path.join(ROOT, "results", "figures", n + ".png")
    num = lambda x, d=0: (f"{x:,.{d}f}").replace(",", "§").replace(".", ",").replace("§", ".")
    dec = lambda x, d=3: f"{x:.{d}f}".replace(".", ",")
    pct = lambda x, d=2: dec(x, d) + "%"
    S, agg, T12, T03, T04, DR, T08, T11, T10, cross = N["S"], N["agg"], N["T12"], N["T03"], N["T04"], N["DRIFT"], N["T08"], N["T11"], N["T10"], N["cross"]
    g = lambda meth, met: dec(agg.loc[meth, (met, "mean")])
    gs = lambda meth, met: dec(agg.loc[meth, (met, "mean")]) + " ± " + dec(agg.loc[meth, (met, "std")])
    d24 = DR[DR.tahun == 2024].set_index("field")
    t03 = T03.set_index("tahun"); t12 = T12.set_index("tahun"); t04 = T04.set_index("tahun")
    ab = lambda meth, fs: dec(T11.loc[meth, ("mean", fs)])
    n_tr, n_te = S["n_train_bersih"], S["n_uji_bersih"]
    tm = N["TM"]; w9 = N["T09"]
    sens = T10.set_index(["parameter", "nilai"])
    hyb_pct = [dec(t12.loc[y, "hibrida_pct"], 2) for y in [2021, 2022, 2023, 2024]]

    # ---------------- halaman judul
    b.front("Deteksi Anomali Hibrida Aturan–Pembelajaran Mesin pada Data Pajak Kendaraan SAMSAT",
            [("Ahmad Dedi Jubaedi", "1,*"), ("Saleh Dwiyatno", "1"), ("Rahmat", "2")],
            [("1", "Fakultas Ilmu Komputer, Universitas Serang Raya, Jl. Raya Cilegon Km. 5, Taktakan, Serang, Banten 42162, Indonesia"),
             ("2", "AMIK Serang, Serang, Banten, Indonesia")],
            "E-mail: dedi@unsera.ac.id (penulis korespondensi), salehdwiyatno@gmail.com, rahmat042@gmail.com")

    # ---------------- abstrak (Inggris wajib, abstrak Indonesia menyertai)
    b.para("Abstract", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    b.para(f"Motor vehicle tax (PKB) transactions recorded by SAMSAT offices underpin regional revenue reporting, yet their quality is rarely audited systematically. "
           f"**Problem:** inconsistencies such as penalties without arrears, mutation codes that contradict the application type and PKB values implausible for the vehicle type are hidden among hundreds of thousands of records, and no labelled ground truth exists. "
           f"**Solution:** we propose a three-layer hybrid detector that combines eleven domain rules with an ensemble of Isolation Forest, Local Outlier Factor (LOF) and an autoencoder, preceded by coding-scheme drift detection and harmonisation. "
           f"The method was applied to {S['n_transaksi']:,} ISQL-WEB SAMSAT transactions from Banten (May–June 2021–2024), trained on 2021–2023 and tested on 2024 with eight injected anomaly types over ten runs. "
           f"**Results:** the hybrid detector reached PR-AUC {agg.loc['Hibrida', ('PR_AUC', 'mean')]:.3f} and ROC-AUC {agg.loc['Hibrida', ('ROC_AUC', 'mean')]:.3f}, outperforming rules alone ({agg.loc['Aturan', ('PR_AUC', 'mean')]:.3f}), the autoencoder ({agg.loc['AE', ('PR_AUC', 'mean')]:.3f}), Isolation Forest ({agg.loc['IF', ('PR_AUC', 'mean')]:.3f}) and LOF ({agg.loc['LOF', ('PR_AUC', 'mean')]:.3f}) (Friedman p < 0.001). "
           f"Domain features raised the ensemble PR-AUC from {T11.loc['Ensemble ML', ('mean', 'dasar')]:.2f} to {T11.loc['Ensemble ML', ('mean', 'dasar+domain')]:.2f}. "
           f"On real data, {S['rule_flag_pct']:.1f}% of transactions violated at least one rule and a 2024 recoding of status and application codes was detected.", italic=False)
    b.para("**Keywords:** anomaly detection; data quality; Isolation Forest; autoencoder; motor vehicle tax")
    b.blank()
    b.para("Abstrak", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    b.para(f"Transaksi pajak kendaraan bermotor (PKB) yang dicatat SAMSAT menjadi dasar pelaporan pendapatan daerah, tetapi kualitasnya jarang diaudit secara sistematis. "
           f"**Masalah:** inkonsistensi seperti denda tanpa tunggakan, kode mutasi yang bertentangan dengan jenis permohonan, dan nilai PKB yang janggal untuk jenis kendaraannya tersembunyi di antara ratusan ribu baris, sementara label kebenaran tidak tersedia. "
           f"**Solusi:** penelitian ini mengusulkan detektor hibrida tiga lapis yang menggabungkan sebelas aturan domain dengan ensemble Isolation Forest, Local Outlier Factor (LOF), dan autoencoder, didahului deteksi serta harmonisasi pergeseran skema kode. "
           f"Metode diterapkan pada {num(S['n_transaksi'])} transaksi ISQL-WEB SAMSAT Banten (Mei–Juni 2021–2024), dilatih pada 2021–2023 dan diuji pada 2024 dengan delapan tipe anomali injeksi dalam sepuluh ulangan. "
           f"**Hasil:** detektor hibrida mencapai PR-AUC {g('Hibrida', 'PR_AUC')} dan ROC-AUC {g('Hibrida', 'ROC_AUC')}, lebih baik daripada aturan saja ({g('Aturan', 'PR_AUC')}), autoencoder ({g('AE', 'PR_AUC')}), Isolation Forest ({g('IF', 'PR_AUC')}), dan LOF ({g('LOF', 'PR_AUC')}) (Friedman p < 0,001). "
           f"Fitur berbasis domain menaikkan PR-AUC ensemble dari {dec(T11.loc['Ensemble ML', ('mean', 'dasar')], 2)} menjadi {dec(T11.loc['Ensemble ML', ('mean', 'dasar+domain')], 2)}. "
           f"Pada data riil, {dec(S['rule_flag_pct'], 1)}% transaksi melanggar sedikitnya satu aturan dan terdeteksi pengodean ulang kode status serta permohonan pada 2024.")
    b.para("**Kata kunci:** deteksi anomali; kualitas data; Isolation Forest; autoencoder; pajak kendaraan bermotor")

    # ---------------- 1. Pendahuluan
    b.h1("1. Pendahuluan")
    b.para("Pajak kendaraan bermotor (PKB) dan bea balik nama kendaraan bermotor (BBNKB) merupakan sumber utama pendapatan asli daerah provinsi di Indonesia, dan pengelolaannya diatur ulang melalui Undang-Undang Hubungan Keuangan Pusat dan Daerah [@uuhkpd]. "
           "Seluruh transaksi PKB dicatat oleh Sistem Administrasi Manunggal Satu Atap (SAMSAT); di Banten, data tersebut dihimpun aplikasi ISQL-WEB SAMSAT dan menjadi bahan pelaporan penerimaan, penagihan tunggakan, serta evaluasi kebijakan keringanan. "
           "Kajian PKB di Indonesia selama ini didominasi analisis kepatuhan wajib pajak dan layanan elektronik [@ambarwati; @ervina] serta peramalan realisasi penerimaan [@kurniasari], sedangkan mutu data transaksinya sendiri jarang diperiksa. "
           "Padahal, kualitas data merupakan prasyarat analitik dan pembelajaran mesin yang andal [@ehrlinger; @kapoor], dan studi pada data pemerintah terbuka menunjukkan bahwa nilai hilang serta anomali dapat mencapai porsi yang besar bila tidak dipantau [@karamanou].")
    b.para("Di bidang administrasi perpajakan, pembelajaran mesin telah dipakai untuk mendeteksi risiko pajak, memilih sasaran audit, dan memprediksi tunggakan [@zheng; @battaglini; @baghdasaryan; @belahouaoui; @yang; @khreis]. "
           "Karena label kecurangan atau kesalahan hampir selalu langka, pendekatan tak terawasi menjadi pilihan dominan [@savic; @hilal]. "
           "Isolation Forest mengisolasi titik langka melalui partisi acak [@hariri; @xu], Local Outlier Factor (LOF) membandingkan kerapatan lokal [@alghushairy], dan autoencoder menilai galat rekonstruksi terhadap pola normal [@pang; @ruff]. "
           "Namun, sistem audit di lingkungan pemerintahan umumnya masih bertumpu pada aturan yang ditulis petugas; aturan mudah dijelaskan tetapi hanya menangkap kesalahan yang telah diantisipasi, sebaliknya model tak terawasi dapat menemukan pola baru tetapi sulit dijelaskan [@westerski; @lixad].")
    b.para("**Pernyataan masalah.** Data transaksi PKB SAMSAT Banten memuat inkonsistensi yang tidak terdeteksi oleh validasi entri, misalnya denda tanpa tunggakan, kode mutasi yang tidak sesuai dengan jenis permohonan, BBNKB yang tidak sejalan dengan kode transaksi, dan nilai PKB yang janggal untuk jenis kendaraannya. "
           "Tiga kendala membuat masalah ini sulit: (i) tidak tersedia label kebenaran sehingga kinerja detektor tidak dapat diukur langsung; (ii) kamus kode resmi tidak tersedia bagi analis dan skema kode dapat berubah antartahun sehingga aturan statis menjadi usang; dan (iii) petugas membutuhkan daftar prioritas yang dapat dijelaskan, bukan sekadar skor kotak hitam.")
    b.para("**Penelitian terdahulu.** Tabel 1 merangkum penelitian yang paling dekat. Savić dkk. [@savic] menggabungkan klasterisasi dan pembelajaran representasi untuk risiko penghindaran pajak penghasilan, Westerski dkk. [@westerski] menerapkan deteksi anomali yang dapat dijelaskan pada pengadaan pemerintah, Karamanou dkk. [@karamanou] memakai Isolation Forest untuk menilai kualitas data pemerintah terbuka, sedangkan Carcillo dkk. [@carcillo] dan Bakumenko & Elragal [@bakumenko] menggabungkan skor tak terawasi dengan model terawasi pada data keuangan berlabel. "
           "Belum ada studi yang (a) memadukan aturan domain perpajakan daerah dengan ensemble IF–LOF–autoencoder pada tingkat transaksi SAMSAT, (b) mengevaluasinya secara kuantitatif tanpa label melalui injeksi anomali realistis [@steinbuss] dengan validasi temporal, dan (c) sekaligus mengukur kualitas data serta pergeseran skema kode antartahun [@bayram].")
    T1 = pd.DataFrame([
        ["Savić dkk. [@savic]", "PPh orang pribadi, Serbia", "Klasterisasi + autoencoder (HUNOD), model surrogate", "Sebagian", "Validasi internal", "Tidak"],
        ["Westerski dkk. [@westerski]", "Pengadaan pemerintah", "Deteksi anomali dapat dijelaskan", "Ya", "Umpan balik implementasi", "Tidak"],
        ["Karamanou dkk. [@karamanou]", "Data lalu lintas OGD, Yunani", "Isolation Forest + dekomposisi musiman", "Tidak", "Deskriptif", "Ya"],
        ["Bakumenko & Elragal [@bakumenko]", "Data jurnal keuangan", "ML terawasi & tak terawasi", "Tidak", "Label anomali", "Tidak"],
        ["Carcillo dkk. [@carcillo]", "Transaksi kartu kredit", "Skor outlier + klasifikasi terawasi", "Tidak", "Label fraud", "Tidak"],
        ["Kurniasari dkk. [@kurniasari]", "Realisasi BBNKB, Lampung", "Jaringan saraf tiruan (prediksi)", "Tidak", "Galat prediksi", "Tidak"],
        ["**Penelitian ini**", "**Transaksi PKB SAMSAT Banten 2021–2024**", "**11 aturan + IF/LOF/AE, fusi maksimum, skor hibrida**", "**Ya**", "**Injeksi 8 tipe × 10 ulangan, uji temporal**", "**Ya (6 dimensi + JSD kode)**"],
    ], columns=["Studi", "Domain/data", "Metode", "Aturan domain", "Evaluasi", "Kualitas data & drift"])
    b.table(T1, "Perbandingan dengan penelitian terdahulu", widths=[2.9, 2.9, 3.6, 1.5, 2.4, 2.1])
    b.para("**Solusi yang diusulkan dan kebaruan.** Penelitian ini mengusulkan kerangka hibrida tiga lapis untuk transaksi SAMSAT: lapis pertama adalah mesin aturan domain R01–R11 yang profil kodenya diturunkan dari data referensi; lapis kedua adalah ensemble Isolation Forest, LOF, dan autoencoder dengan fusi maksimum persentil; lapis ketiga menggabungkan keduanya menjadi daftar prioritas audit beserta kode alasan. "
           "Kebaruannya adalah: (1) katalog aturan kualitas data PKB yang dapat dioperasionalkan dan dipetakan ke enam dimensi kualitas data; (2) deteksi dan harmonisasi pergeseran skema kode berbasis Jensen–Shannon divergence sebelum aturan dan model dijalankan; (3) fitur berbasis domain (rasio denda, z-skor nilai per jenis, rasio BBNKB/PKB) yang terbukti meningkatkan detektor tak terawasi; dan (4) protokol evaluasi tanpa label menggunakan delapan tipe anomali injeksi, sepuluh ulangan, validasi temporal, dan uji Friedman–Wilcoxon.")

    # ---------------- 2. Metodologi
    b.h1("2. Metodologi")
    b.para(f"Tahapan penelitian ditunjukkan pada Gambar 1. Seluruh eksperimen dijalankan dengan Python (pandas, scikit-learn {S['versions']['sklearn']}) dengan benih acak 42, dan diperiksa ulang di Google Colab menggunakan notebook yang identik dengan skrip.")
    b.figure(F("F01_flowchart_metode"), "Diagram alir metode penelitian (tiga lapis deteksi dan evaluasi)", width=11.0)
    b.h2("2.1 Data dan pra-pemrosesan")
    t01 = N["T01"]
    b.para(f"Data berasal dari ekspor ISQL-WEB SAMSAT Banten periode Mei–Juni 2021, 2022, 2023, dan 2024 (empat berkas, {num(t01.baris_mentah.sum())} baris). Setelah membuang {num(t01.baris_header_footer.sum())} baris judul/pengulang header, diperoleh {num(S['n_transaksi'])} transaksi dengan 24 atribut: kode status, jenis permohonan (kd_jen_mohon), jenis mutasi (kd_jen_mutasi), lokasi, kabupaten/kota, jenis kendaraan (kd_jenis_kb), merek, tanggal bayar, pokok dan denda BBNKB, pokok dan denda PKB tahun berjalan, serta pokok dan denda tunggakan lima tahun. "
           "Nomor polisi hanya dipakai untuk memeriksa format dan duplikasi, kemudian diganti pseudonim SHA-256; nama dan alamat tidak pernah diolah. Data diperoleh melalui petugas SAMSAT untuk keperluan eksperimen dan dianalisis dalam bentuk terpseudonimisasi.")
    b.h2("2.2 Deteksi pergeseran skema kode")
    b.para("Distribusi setiap atribut kategorikal pada tahun *y* dibandingkan dengan tahun referensi lain menggunakan Jensen–Shannon divergence (JSD) berbasis log~2~ (nilai 0–1):")
    b.eq("JSD(P‖Q) = ½·KL(P‖M) + ½·KL(Q‖M),   M = ½(P + Q)", 1)
    b.para("Kode yang muncul atau hilang serentak dengan volume setara diperlakukan sebagai pengodean ulang dan diharmonisasi sebelum aturan dan model dijalankan (Subbab 3.1).")
    b.h2("2.3 Mesin aturan domain")
    b.para("Sebelas aturan (Tabel 2) disusun dari logika tarif PKB dan pola kode pada tahun referensi 2021–2023, lalu dipetakan ke dimensi konsistensi, akurasi, plausibilitas, dan keunikan [@ehrlinger]. "
           "Karena kamus kode resmi tidak tersedia, konsistensi kode (R07, R08) diturunkan dari data: pasangan kode permohonan–mutasi yang porsinya < 0,5% dalam kode permohonannya (atau n < 30) dianggap tidak lazim, sedangkan pasangan yang ≥ 95% (≤ 5%) disertai BBNKB menjadi acuan wajib (tanpa) BBNKB. "
           "Kejanggalan nilai PKB terhadap jenis kendaraan (R09) memakai z-skor robust:")
    b.eq("z~i~ = (log~10~ PKB~i~ − med~g~) / max(MAD~g~, 0,05),   |z~i~| > 3,5 ⇒ R09", 2)
    b.para("dengan *g* adalah kode jenis kendaraan, med~g~ median dan MAD~g~ median absolute deviation (berskala normal) dari data referensi. Skor aturan adalah jumlah terbobot pelanggaran (bobot 1–3, Tabel 2).")
    T2 = N["T02"].copy()
    T2 = T2[["ID", "Aturan", "Dimensi", "Bobot", "Logika"]]
    b.table(T2, "Katalog aturan domain kualitas data transaksi PKB", widths=[0.9, 3.6, 1.9, 1.0, 8.0], size=7.5)
    b.h2("2.4 Fitur dan detektor tak terawasi")
    b.para(f"Detektor dilatih hanya pada transaksi bersih (tanpa pelanggaran aturan dan bukan transaksi bernilai nol) tahun 2021–2023 (n = {num(n_tr)}; sampel latih 40.000) dan diuji pada 2024 sehingga tidak terjadi kebocoran waktu [@kapoor]. "
           "Fitur dasar terdiri atas logaritma nilai pokok/denda PKB, BBNKB dan tunggakan, jumlah tahun tunggakan, frekuensi kode, serta kode status. Fitur domain menambahkan rasio denda/pokok, rasio denda/tunggakan, rasio tunggakan/pokok, rasio BBNKB/PKB, z-skor nilai per jenis (persamaan 2), z-skor BBNKB/PKB per pasangan kode, frekuensi pasangan kode, dan selisih status BBNKB terhadap kelaziman kodenya. "
           "Ketiga detektor adalah Isolation Forest (300 pohon, ψ = 512) dengan skor")
    b.eq("s(x) = 2^−E[h(x)]/c(ψ)^", 3)
    b.para("LOF mode *novelty* (k = 35), dan autoencoder MLP 32–8–32 (ReLU, Adam, *early stopping*) dengan galat rekonstruksi")
    b.eq("e(x) = (1/d) Σ~j~ (x~j~ − x̂~j~)²", 4)
    b.para("Pemilihan model dangkal dan jaringan kecil mengikuti temuan bahwa data tabular tidak selalu memerlukan arsitektur dalam [@shwartz; @borisov].")
    b.h2("2.5 Fusi ensemble dan skor hibrida")
    b.para("Skor setiap detektor *m* dikalibrasi menjadi persentil F~m~ terhadap distribusi skor data latih, lalu digabung dengan fungsi maksimum (dengan rerata sebagai pemecah seri) agar sinyal kuat dari satu detektor tidak teredam detektor lain [@pang]. Skor hibrida menempatkan transaksi yang melanggar aturan pada prioritas pertama:")
    b.eq("S~ML~(x) = max~m~ F~m~(s~m~(x));      H(x) = 𝟙[R(x) > 0] + S~ML~(x)", 5)
    b.para("Transaksi ditandai ML bila S~ML~ melampaui persentil ke-99 data latih. Kode alasan diambil dari aturan yang dilanggar dan tiga fitur dengan galat rekonstruksi autoencoder terbesar [@lixad]. Gambar 2 memperlihatkan rancangan penerapannya pada sistem informasi SAMSAT.")
    b.figure(F("F02_arsitektur_sistem"), "Arsitektur penerapan detektor hibrida pada alur data SAMSAT", width=14.5)
    b.h2("2.6 Desain evaluasi")
    b.para(f"Karena label kebenaran tidak tersedia, kinerja diukur dengan injeksi anomali terkendali yang meniru kesalahan nyata [@steinbuss]. Pada setiap ulangan, 20.000 transaksi bersih 2024 diambil acak dan 800 anomali (8 tipe × 100) disisipkan (kontaminasi 3,85%), sebanyak sepuluh ulangan (Tabel 3). "
           "Tipe A8 sengaja tidak memiliki aturan pasangan untuk menguji kemampuan menemukan pola yang tidak diantisipasi. Metrik yang dipakai adalah ROC-AUC, PR-AUC (*average precision*), presisi pada k = jumlah anomali (P@k), Recall pada anggaran audit 5%, serta presisi, recall, dan F1 pada ambang operasional. "
           "Perbedaan antarmetode diuji dengan Friedman dan Wilcoxon berpasangan terkoreksi Holm. Analisis sensitivitas mencakup k LOF, jumlah pohon IF, arsitektur AE, ukuran data latih, dan strategi fusi; ablasi membandingkan fitur dasar dengan fitur dasar + domain.")
    T3 = N["T06"][["Kode", "Tipe anomali", "Aturan terkait"]].copy()
    T3["Cara injeksi"] = ["denda tahun tunggakan ke-2 = 5–25% pokok, tanpa pokok tunggakan",
                          "kd_jen_mutasi diganti kode yang bertentangan dengan jenis permohonan",
                          "pokok motor ↔ mobil diambil dari distribusi jenis lain",
                          "pokok PKB ×10 atau ÷10, denda tetap",
                          "kd_jenis_kb motor ↔ mobil/barang, nilai tetap",
                          "denda tunggakan 5–18% atau 32–50% dari tunggakan",
                          "pokok & denda tunggakan ×2,5–5 (rasio 25% dipertahankan)",
                          "BBNKB ×5–12 atau ÷5–12"]
    b.table(T3, "Tipe anomali injeksi (100 per tipe per ulangan)", widths=[1.0, 4.4, 2.4, 7.6], size=7.5)

    # ---------------- 3. Hasil dan pembahasan
    b.h1("3. Hasil dan Pembahasan")
    b.h2("3.1 Profil kualitas data dan pergeseran skema kode")
    b.para(f"Gambar 3 menunjukkan bahwa pada 2024 kode status 3/4 digantikan seluruhnya oleh 5/6 (JSD = {dec(d24.loc['kd_status', 'JSD_mentah'])}), dan transaksi daftar ulang yang disertai mutasi berpindah dari kode permohonan 1 ke 2 (pasangan 1|7 → 2|7 dengan volume setara; JSD = {dec(d24.loc['kd_jen_mohon|kd_jen_mutasi', 'JSD_mentah'])}). "
           f"Setelah harmonisasi, JSD turun menjadi {dec(d24.loc['kd_status', 'JSD_harmonisasi'], 4)} dan {dec(d24.loc['kd_jen_mohon|kd_jen_mutasi', 'JSD_harmonisasi'], 4)}. Tanpa langkah ini, ribuan transaksi 2024 akan ditandai keliru sebagai kode tidak konsisten, dan detektor ML mengalami pergeseran distribusi; temuan ini menegaskan pentingnya pemantauan drift pada data operasional [@bayram].")
    b.figure(F("F16_pergeseran_kode"), "Pergeseran skema kode 2024: JSD sebelum/sesudah harmonisasi (kiri) dan perubahan volume pasangan kode (kanan)", width=14.5)
    b.para(f"Secara keseluruhan {num(S['rule_flag'])} transaksi ({dec(S['rule_flag_pct'], 2)}%) melanggar sedikitnya satu aturan (Gambar 4). Pelanggaran terbanyak adalah nilai PKB janggal untuk jenis kendaraan (R09, ±2,0–2,3% per tahun) dan tunggakan tanpa pokok tahun berjalan (R06), yang naik dari {pct(t03.loc[2021, 'R06_pct'])} (2021) menjadi {pct(t03.loc[2024, 'R06_pct'])} (2024). "
           f"Tunggakan tanpa denda (R02) dan tarif denda tunggakan menyimpang (R03) muncul pada 2021–2022 lalu menghilang pada 2024, konsisten dengan dugaan adanya kebijakan relaksasi denda pada masa pandemi yang perlu dikonfirmasi ke Bapenda. "
           f"Sebaliknya, R08 melonjak menjadi {num(t03.loc[2024, 'R08'])} transaksi pada 2024; sebagian besar berasal dari pasangan 2|4 tanpa BBNKB, yaitu akibat penggabungan dua makna kode ke dalam satu kode baru. Hal ini menunjukkan bahwa pengodean ulang tidak hanya mengubah label, tetapi juga menghilangkan informasi semantik. "
           f"Selain itu terdapat {num(S['zero_value'])} transaksi bernilai nol ({dec(100 * S['zero_value'] / S['n_transaksi'], 1)}%) dan {num(S['plate_nonstandard'])} nomor polisi dengan format tidak baku.")
    b.figure(F("F03_prevalensi_aturan"), "Prevalensi pelanggaran aturan domain per tahun (persentase dan jumlah transaksi)", width=14.5)
    T4 = T04.copy(); T4.columns = ["Tahun", "Kelengkapan", "Validitas", "Konsistensi", "Akurasi", "Plausibilitas", "Keunikan", "Indeks DQ"]
    for c in T4.columns[1:]: T4[c] = T4[c].map(lambda v: dec(v, 2))
    T4["Pelanggaran aturan"] = [f"{num(t03.loc[y, 'rule_flag'])} ({dec(t03.loc[y, 'rule_flag_pct'], 2)}%)" for y in T4["Tahun"]]
    b.table(T4, "Kartu skor kualitas data SAMSAT per dimensi (%) dan jumlah pelanggaran aturan", widths=[1.1, 1.8, 1.6, 1.8, 1.5, 1.8, 1.5, 1.6, 2.6])
    b.para(f"Tabel 4 memperlihatkan indeks kualitas data yang tinggi ({dec(t04['Indeks DQ'].min(), 2)}–{dec(t04['Indeks DQ'].max(), 2)}%), tetapi dimensi konsistensi 2024 adalah yang terendah ({dec(t04.loc[2024, 'Konsistensi'], 2)}%). Angka agregat yang tinggi dapat menyamarkan kesalahan bernilai besar; karena itu deteksi tingkat transaksi tetap diperlukan.")
    b.h2("3.2 Kinerja detektor pada anomali injeksi")
    T5 = N["T07"].copy()
    T5["Metode"] = T5["Metode"].replace({"Hibrida": "**Hibrida (usulan)**"})
    T5 = T5[["Metode", "ROC_AUC", "PR_AUC", "P@k", "Recall@5%", "F1", "FPR (normal)"]]
    T5.columns = ["Metode", "ROC-AUC", "PR-AUC", "P@k", "Recall@5%", "F1", "FPR normal"]
    for c in T5.columns[1:]: T5[c] = T5[c].astype(str).str.replace(".", ",", regex=False)
    b.table(T5, "Komparasi kinerja detektor (rerata ± simpangan baku, 10 ulangan, uji 2024)", widths=[2.8, 2.4, 2.4, 2.4, 2.4, 2.4, 1.8], bold_last=True)
    b.para(f"Tabel 5 dan Gambar 5 (kurva ROC dan precision–recall disajikan pada materi suplemen) menunjukkan bahwa detektor hibrida unggul pada hampir semua metrik: PR-AUC {gs('Hibrida', 'PR_AUC')}, ROC-AUC {g('Hibrida', 'ROC_AUC')}, P@k {g('Hibrida', 'P@k')}, dan Recall@5% {g('Hibrida', 'Recall@5%')}. "
           f"Aturan saja memiliki presisi sempurna (1,000) tetapi hanya menangkap {dec(agg.loc['Aturan', ('Recall', 'mean')])} anomali, sedangkan autoencoder menjadi detektor tunggal terbaik (PR-AUC {g('AE', 'PR_AUC')}). "
           f"Isolation Forest (PR-AUC {g('IF', 'PR_AUC')}) dan LOF ({g('LOF', 'PR_AUC')}) lemah karena anomali pada data ini bersifat kontekstual—nilai yang wajar secara marginal tetapi tidak wajar terhadap kode atau jenis kendaraan—sehingga sulit diisolasi oleh partisi acak atau kerapatan lokal pada ruang campuran fitur kontinu–diskret [@xu; @alghushairy]. "
           f"Uji Friedman signifikan (χ² = {dec(w9.Friedman_chi2.iloc[0], 1)}, p = {(f"{S['friedman_p']:.1e}".split('e')[0]).replace('.', ',')} × 10^{int(f"{S['friedman_p']:.1e}".split('e')[1])}^), dan hibrida mengungguli setiap pembanding pada 10 dari 10 ulangan (Wilcoxon-Holm p = {dec(w9.p_holm.max(), 4)}). "
           f"F1 hibrida ({g('Hibrida', 'F1')}) sedikit di bawah aturan ({g('Aturan', 'F1')}) karena ambang P99 menambah positif palsu ({dec(agg.loc['Hibrida', ('Presisi', 'mean')])} presisi); pada praktik audit, metrik berbasis anggaran (P@k, Recall@5%) lebih relevan.")
    b.figure(F("F09_komparasi_metrik"), "Komparasi PR-AUC, P@k, dan F1 antardetektor (rerata ± sd, 10 ulangan)", width=14.5)
    b.h2("3.3 Analisis per tipe anomali, ablasi, dan sensitivitas")
    b.para(f"Gambar 6 menunjukkan sifat saling melengkapi kedua lapis. Aturan menangkap A1, A2, A3, A4, dan A5 hampir sempurna, tetapi gagal pada A8 (BBNKB janggal; recall {dec(T08.loc['A8', 'Aturan'], 2)}) yang tidak memiliki aturan. "
           f"Autoencoder justru menangkap A8 ({dec(T08.loc['A8', 'AE'], 2)}), A7 ({dec(T08.loc['A7', 'AE'], 2)}), dan A6 ({dec(T08.loc['A6', 'AE'], 2)}), tetapi lemah pada kode jenis kendaraan tertukar (A5, {dec(T08.loc['A5', 'AE'], 2)}). "
           f"Hibrida mempertahankan kekuatan keduanya: A8 naik dari {dec(T08.loc['A8', 'Aturan'], 2)} (aturan) menjadi {dec(T08.loc['A8', 'Hibrida'], 2)}, sementara A1–A5 tetap ≥ {dec(T08.loc[['A1', 'A2', 'A3', 'A4', 'A5'], 'Hibrida'].min(), 2)}. "
           "Temuan ini sejalan dengan rekomendasi audit berbasis data bahwa aturan dan deteksi tak terawasi seharusnya dipadukan, bukan dipertentangkan [@westerski; @carcillo].")
    b.figure(F("F08_recall_per_tipe"), "Recall@k per tipe anomali dan per detektor (rerata 10 ulangan)", width=14.0)
    b.para(f"Ablasi (Gambar S1 pada materi suplemen) menunjukkan bahwa fitur domain menaikkan PR-AUC autoencoder dari {ab('AE', 'dasar')} menjadi {ab('AE', 'dasar+domain')}, ensemble dari {ab('Ensemble ML', 'dasar')} menjadi {ab('Ensemble ML', 'dasar+domain')}, dan hibrida dari {ab('Hibrida', 'dasar')} menjadi {ab('Hibrida', 'dasar+domain')}; hanya LOF yang menurun ({ab('LOF', 'dasar')} → {ab('LOF', 'dasar+domain')}) karena penambahan dimensi memperburuk estimasi kerapatan. "
           f"Pada analisis sensitivitas, fusi maksimum persentil memberi PR-AUC ensemble {dec(sens.loc[('Strategi fusi', 'max persentil (dipakai)'), 'PR_AUC'])} dibanding {dec(sens.loc[('Strategi fusi', 'rerata persentil'), 'PR_AUC'])} untuk fusi rerata, karena rerata meredam sinyal autoencoder oleh IF dan LOF yang lemah. "
           f"Arsitektur AE yang lebih lebar (64–16–64) meningkatkan PR-AUC AE menjadi {dec(sens.loc[('AE arsitektur', '64-16-64'), 'PR_AUC'])}, sedangkan PR-AUC hibrida stabil pada {dec(T10.PR_AUC_hibrida.min())}–{dec(T10.PR_AUC_hibrida.max())} untuk seluruh variasi parameter. "
           f"Waktu latih rata-rata pada CPU dua inti adalah {dec(tm['fit_IF'], 2)} detik (IF), {dec(tm['fit_LOF'], 2)} detik (LOF), dan {dec(tm['fit_AE'], 2)} detik (AE) untuk 40.000 transaksi, dan skoring 20.800 transaksi < 1 detik, sehingga layak dijalankan sebagai *batch* harian.")
    b.h2("3.4 Penerapan pada data riil 2021–2024")
    T6 = T12[["tahun", "n", "aturan", "keduanya", "ml_saja", "nihil", "hibrida_total", "hibrida_pct"]].copy()
    T6.columns = ["Tahun", "Transaksi", "Melanggar aturan", "Aturan & ML", "ML saja", "Nilai nol (informatif)", "Daftar audit hibrida", "% daftar audit"]
    for c in T6.columns[1:-1]: T6[c] = T6[c].map(lambda v: num(v))
    T6["% daftar audit"] = T6["% daftar audit"].map(lambda v: dec(v, 2))
    b.table(T6, "Hasil penerapan detektor hibrida pada data riil per tahun", widths=[1.2, 1.9, 2.0, 1.9, 1.6, 2.4, 2.4, 1.9])
    b.para(f"Tabel 6 dan Gambar 7 merangkum penerapan pada {num(S['n_transaksi'])} transaksi. Daftar audit hibrida mencakup {hyb_pct[0]}%, {hyb_pct[1]}%, {hyb_pct[2]}%, dan {hyb_pct[3]}% transaksi pada 2021–2024. "
           f"Sebanyak {num(S['ml_only'])} transaksi ditandai hanya oleh ML; alasan utamanya adalah z-skor nilai per jenis kendaraan yang mendekati batas aturan, besaran denda, dan BBNKB yang tidak lazim (Tabel 7). "
           f"Transaksi nilai nol ({num(int(T12.nihil.sum()))} yang tidak melanggar aturan) dipisahkan sebagai kategori informatif karena kemungkinan merupakan pengesahan tanpa pembayaran. "
           f"Irisan pada Gambar 7 juga menunjukkan bahwa ML ikut menandai {dec(cross.loc['R06', 'persen_juga_ML'], 0)}% transaksi R06 dan {dec(cross.loc['R08', 'persen_juga_ML'], 0)}% transaksi R08, tetapi hanya {dec(cross.loc['R09', 'persen_juga_ML'], 1)}% transaksi R09; artinya batas robust-z ±3,5 lebih ketat daripada pola yang dipelajari model dan dapat dikalibrasi ulang bersama petugas.")
    b.figure(F("F13_temuan_data_riil"), "Komposisi transaksi terflag per tahun (kiri) dan irisan pelanggaran aturan dengan bendera ML (kanan)", width=14.5)
    ex = pd.concat([N["T13"].head(2), N["T13b"].drop_duplicates(["kd_jen_mohon", "kd_jen_mutasi", "kd_jenis_kb", "pkbpok"]).head(4)])
    T7 = pd.DataFrame({"Tahun": ex.tahun, "Mohon|mutasi": ex.kd_jen_mohon.map(lambda v: str(int(v))) + "|" + ex.kd_jen_mutasi.map(lambda v: "NA" if pd.isna(v) else str(int(float(v)))),
                       "Jenis": ex.kd_jenis_kb.astype(int), "PKB pokok (Rp)": ex.pkbpok.map(lambda v: num(v)), "BBNKB (Rp)": ex.bbnpok.map(lambda v: num(v)),
                       "Tunggakan (Rp)": ex.tgk_total.map(lambda v: num(v)), "Aturan": ex.aturan_dilanggar, "Alasan AE (3 teratas)": ex.alasan_AE_top3.str.replace("_", " ")})
    b.table(T7, "Contoh transaksi prioritas audit beserta kode alasan (identitas kendaraan dipseudonimkan)", widths=[1.0, 1.4, 1.0, 1.9, 1.9, 2.0, 1.3, 5.5], size=7)
    b.para("Contoh pada Tabel 7 memperlihatkan bentuk keluaran yang dapat ditindaklanjuti: dua transaksi teratas melanggar R06 dan R08 (tunggakan dibayar tanpa pokok tahun berjalan pada kode yang lazimnya disertai BBNKB), sedangkan transaksi ML-saja mencakup mutasi armada bus dengan nilai identik, registrasi kode kendaraan langka dengan BBNKB tinggi, dan rasio BBNKB/PKB yang tidak lazim. "
           "Tidak semua transaksi tersebut merupakan kesalahan; peran sistem adalah memprioritaskan verifikasi sehingga petugas dapat memeriksa ratusan transaksi alih-alih ratusan ribu.")
    b.h2("3.5 Implikasi dan keterbatasan")
    b.para("Bagi sistem informasi pemerintahan, hasil ini menunjukkan tiga implikasi praktis. Pertama, validasi entri ISQL-WEB SAMSAT dapat diperkuat dengan aturan R01–R11 yang dijalankan harian. Kedua, perubahan skema kode perlu didokumentasikan dalam kamus data berversi agar analitik lintas tahun tidak bias. Ketiga, daftar prioritas hibrida beserta kode alasan dapat menjadi dasar audit internal dan rekonsiliasi penerimaan. "
           "Keterbatasan penelitian ini adalah: anomali dievaluasi melalui injeksi sintetis sehingga kinerja pada kesalahan nyata perlu dikonfirmasi dengan verifikasi petugas; semantik kode diturunkan dari data karena kamus resmi tidak tersedia; dan data hanya mencakup periode Mei–Juni dari wilayah Samsat yang diteliti.")

    # ---------------- 4. Kesimpulan
    b.h1("4. Kesimpulan")
    b.para(f"**Masalah** yang dijawab penelitian ini adalah ketiadaan mekanisme sistematis untuk mendeteksi anomali transaksi dan menilai kualitas data PKB SAMSAT tanpa label kebenaran dan di tengah perubahan skema kode. "
           f"**Solusi** yang diusulkan adalah detektor hibrida tiga lapis yang memadukan sebelas aturan domain, harmonisasi pergeseran kode, dan ensemble Isolation Forest–LOF–autoencoder dengan fusi maksimum persentil. "
           f"**Hasil** pada {num(S['n_transaksi'])} transaksi Banten 2021–2024 menunjukkan PR-AUC {g('Hibrida', 'PR_AUC')} dan ROC-AUC {g('Hibrida', 'ROC_AUC')} pada uji temporal 2024, lebih tinggi daripada aturan saja ({g('Aturan', 'PR_AUC')}) maupun detektor ML tunggal terbaik ({g('AE', 'PR_AUC')}); fitur domain menaikkan PR-AUC ensemble sebesar {dec(T11.loc['Ensemble ML', ('mean', 'dasar+domain')] - T11.loc['Ensemble ML', ('mean', 'dasar')], 2)}; serta terdeteksi pengodean ulang kode status dan permohonan pada 2024 dan {dec(S['rule_flag_pct'], 2)}% transaksi yang melanggar aturan. "
           "Penelitian selanjutnya akan memvalidasi daftar prioritas bersama petugas SAMSAT untuk memperoleh label nyata, memperluas cakupan ke seluruh bulan dan wilayah Banten, serta mengembangkan pembelajaran aktif dari umpan balik petugas.")
    b.blank()
    b.para("**Ucapan terima kasih.** Penulis berterima kasih kepada petugas SAMSAT di Provinsi Banten yang membantu penyediaan data untuk keperluan eksperimen. Kode, notebook Google Colab, dan tabel hasil tersedia dari penulis korespondensi.")
    b.references("Referensi")
