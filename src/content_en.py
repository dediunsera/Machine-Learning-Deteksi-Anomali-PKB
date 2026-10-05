# -*- coding: utf-8 -*-
"""English manuscript content (INOVATIF). All numbers are read from the result tables (N)."""
import os
import pandas as pd
from docx.enum.text import WD_ALIGN_PARAGRAPH

def build(b, N, ROOT):
    F = lambda n: os.path.join(ROOT, "results", "figures_en", n + ".png")
    num = lambda x, d=0: f"{x:,.{d}f}"
    dec = lambda x, d=3: f"{x:.{d}f}"
    pct = lambda x, d=2: dec(x, d) + "%"
    S, agg, T12, T03, T04, DR, T08, T11, T10, cross = N["S"], N["agg"], N["T12"], N["T03"], N["T04"], N["DRIFT"], N["T08"], N["T11"], N["T10"], N["cross"]
    g = lambda meth, met: dec(agg.loc[meth, (met, "mean")])
    gs = lambda meth, met: dec(agg.loc[meth, (met, "mean")]) + " ± " + dec(agg.loc[meth, (met, "std")])
    d24 = DR[DR.tahun == 2024].set_index("field")
    t03 = T03.set_index("tahun"); t12 = T12.set_index("tahun"); t04 = T04.set_index("tahun")
    ab = lambda meth, fs: dec(T11.loc[meth, ("mean", fs)])
    n_tr = S["n_train_bersih"]; tm = N["TM"]; w9 = N["T09"]
    sens = T10.set_index(["parameter", "nilai"])
    hyb_pct = [dec(t12.loc[y, "hibrida_pct"], 2) for y in [2021, 2022, 2023, 2024]]
    MEN = {"Aturan": "Rules", "IF": "IF", "LOF": "LOF", "AE": "AE", "Ensemble ML": "ML ensemble", "Hibrida": "**Hybrid (proposed)**"}

    b.front("Hybrid Rule–Machine Learning Detection of Anomalies in SAMSAT Vehicle Tax Data",
            [("Ahmad Dedi Jubaedi", "1,*"), ("Saleh Dwiyatno", "1"), ("Rahmat", "2")],
            [("1", "Faculty of Computer Science, Universitas Serang Raya, Jl. Raya Cilegon Km. 5, Taktakan, Serang, Banten 42162, Indonesia"),
             ("2", "AMIK Serang, Serang, Banten, Indonesia")],
            "E-mail: dedi@unsera.ac.id (corresponding author), salehdwiyatno@gmail.com, rahmat042@gmail.com")

    b.para("Abstract", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    b.para(f"Motor vehicle tax (PKB) transactions recorded by SAMSAT offices underpin regional revenue reporting, yet their quality is rarely audited systematically. "
           f"**Problem:** inconsistencies such as penalties without arrears, mutation codes that contradict the application type and PKB values implausible for the vehicle type are hidden among hundreds of thousands of records, and no labelled ground truth exists. "
           f"**Solution:** we propose a three-layer hybrid detector that combines eleven domain rules with an ensemble of Isolation Forest, Local Outlier Factor (LOF) and an autoencoder, preceded by coding-scheme drift detection and harmonisation. "
           f"The method was applied to {num(S['n_transaksi'])} ISQL-WEB SAMSAT transactions from Banten (May–June 2021–2024), trained on 2021–2023 and tested on 2024 with eight injected anomaly types over ten runs. "
           f"**Results:** the hybrid detector reached PR-AUC {g('Hibrida', 'PR_AUC')} and ROC-AUC {g('Hibrida', 'ROC_AUC')}, outperforming rules alone ({g('Aturan', 'PR_AUC')}), the autoencoder ({g('AE', 'PR_AUC')}), Isolation Forest ({g('IF', 'PR_AUC')}) and LOF ({g('LOF', 'PR_AUC')}) (Friedman p < 0.001). "
           f"Domain features raised the ensemble PR-AUC from {dec(T11.loc['Ensemble ML', ('mean', 'dasar')], 2)} to {dec(T11.loc['Ensemble ML', ('mean', 'dasar+domain')], 2)}. "
           f"On real data, {dec(S['rule_flag_pct'], 1)}% of transactions violated at least one rule and a 2024 recoding of status and application codes was detected.")
    b.para("**Keywords:** anomaly detection; data quality; Isolation Forest; autoencoder; motor vehicle tax")

    b.h1("1. Introduction")
    b.para("Motor vehicle tax (PKB) and the vehicle transfer fee (BBNKB) are major sources of provincial own-source revenue in Indonesia, and their administration was recently redefined by the law on fiscal relations between the central and regional governments [@uuhkpd]. "
           "Every PKB transaction is recorded by the One-Roof Administration System (SAMSAT); in Banten the records are compiled by the ISQL-WEB SAMSAT application and used for revenue reporting, arrears collection and the evaluation of relief programmes. "
           "Indonesian PKB research has focused on taxpayer compliance and e-services [@ambarwati; @ervina] and on forecasting revenue realisation [@kurniasari], whereas the quality of the transaction data itself is rarely examined. "
           "Data quality is nevertheless a prerequisite for reliable analytics and machine learning [@ehrlinger; @kapoor], and studies of open government data show that missing values and anomalies can be substantial when they are not monitored [@karamanou].")
    b.para("In tax administration, machine learning has been used to detect tax risk, select audit targets and predict arrears [@zheng; @battaglini; @baghdasaryan; @belahouaoui; @yang; @khreis]. "
           "Because labels for fraud or errors are almost always scarce, unsupervised approaches dominate [@savic; @hilal]. "
           "Isolation Forest isolates rare points through random partitions [@hariri; @xu], Local Outlier Factor (LOF) compares local densities [@alghushairy], and autoencoders score the reconstruction error against learned normal patterns [@pang; @ruff]. "
           "Audit systems in government, however, still rely mostly on rules written by officers; rules are easy to explain but only catch anticipated errors, whereas unsupervised models can find new patterns but are hard to explain [@westerski; @lixad].")
    b.para("**Problem statement.** The PKB transaction data of SAMSAT Banten contain inconsistencies that entry validation does not catch, such as penalties without arrears, mutation codes that do not match the application type, transfer fees that contradict the transaction code, and PKB values that are implausible for the vehicle type. "
           "Three obstacles make the problem hard: (i) no ground-truth labels exist, so detector performance cannot be measured directly; (ii) the official code dictionary is not available to analysts and coding schemes may change between years, so static rules become outdated; and (iii) officers need an explainable priority list rather than a black-box score.")
    b.para("**Previous work.** Table 1 summarises the closest studies. Savić et al. [@savic] combined clustering and representation learning to manage income-tax evasion risk, Westerski et al. [@westerski] deployed explainable anomaly detection for public procurement, Karamanou et al. [@karamanou] used Isolation Forest to assess the quality of open government data, while Carcillo et al. [@carcillo] and Bakumenko and Elragal [@bakumenko] combined unsupervised scores with supervised models on labelled financial data. "
           "No study has (a) integrated regional-tax domain rules with an IF–LOF–autoencoder ensemble at SAMSAT transaction level, (b) evaluated it quantitatively without labels through realistic anomaly injection [@steinbuss] under temporal validation, and (c) simultaneously measured data quality and year-to-year coding-scheme drift [@bayram].")
    T1 = pd.DataFrame([
        ["Savić et al. [@savic]", "Personal income tax, Serbia", "Clustering + autoencoder (HUNOD), surrogate model", "Partly", "Internal validation", "No"],
        ["Westerski et al. [@westerski]", "Public procurement", "Explainable anomaly detection", "Yes", "Deployment feedback", "No"],
        ["Karamanou et al. [@karamanou]", "Traffic OGD, Greece", "Isolation Forest + seasonal decomposition", "No", "Descriptive", "Yes"],
        ["Bakumenko & Elragal [@bakumenko]", "Financial journal data", "Supervised & unsupervised ML", "No", "Anomaly labels", "No"],
        ["Carcillo et al. [@carcillo]", "Credit-card transactions", "Outlier scores + supervised classifier", "No", "Fraud labels", "No"],
        ["Kurniasari et al. [@kurniasari]", "BBNKB realisation, Lampung", "Artificial neural network (forecast)", "No", "Forecast error", "No"],
        ["**This study**", "**SAMSAT Banten PKB transactions 2021–2024**", "**11 rules + IF/LOF/AE, max fusion, hybrid score**", "**Yes**", "**8 injected types × 10 runs, temporal test**", "**Yes (6 dimensions + code JSD)**"],
    ], columns=["Study", "Domain/data", "Method", "Domain rules", "Evaluation", "Data quality & drift"])
    b.table(T1, "Comparison with previous studies", widths=[2.9, 2.9, 3.6, 1.5, 2.4, 2.1])
    b.para("**Proposed solution and novelty.** We propose a three-layer hybrid framework for SAMSAT transactions: the first layer is a domain rule engine (R01–R11) whose code profile is derived from reference data; the second is an ensemble of Isolation Forest, LOF and an autoencoder with maximum-percentile fusion; the third merges both into an audit priority list with reason codes. "
           "The novelty lies in: (1) an operational catalogue of PKB data-quality rules mapped to six data-quality dimensions; (2) Jensen–Shannon-divergence-based detection and harmonisation of coding-scheme drift before rules and models are run; (3) domain features (penalty ratios, per-type value z-scores, BBNKB/PKB ratio) that demonstrably strengthen unsupervised detectors; and (4) a label-free evaluation protocol with eight injected anomaly types, ten runs, temporal validation and Friedman–Wilcoxon tests.")

    b.h1("2. Methodology")
    b.para(f"The research stages are shown in Figure 1. All experiments were run in Python (pandas, scikit-learn {S['versions']['sklearn']}) with random seed 42 and cross-checked in Google Colab with a notebook identical to the script.")
    b.figure(F("F01_flowchart_metode"), "Flowchart of the research method (three detection layers and evaluation)", width=11.0)
    b.h2("2.1 Data and pre-processing")
    t01 = N["T01"]
    b.para(f"The data are ISQL-WEB SAMSAT Banten exports for May–June 2021, 2022, 2023 and 2024 (four files, {num(t01.baris_mentah.sum())} rows). After removing {num(t01.baris_header_footer.sum())} title and repeated-header rows, {num(S['n_transaksi'])} transactions with 24 attributes remained: status code, application type (kd_jen_mohon), mutation type (kd_jen_mutasi), location, regency/city, vehicle type (kd_jenis_kb), brand, payment date, BBNKB principal and penalty, current-year PKB principal and penalty, and five years of arrears principal and penalties. "
           "Licence plates were used only for format and duplicate checks and then replaced with SHA-256 pseudonyms; names and addresses were never processed. The data were obtained through SAMSAT officers for experimental purposes and analysed in pseudonymised form.")
    b.h2("2.2 Coding-scheme drift detection")
    b.para("The distribution of each categorical attribute in year *y* was compared with the other reference years using the Jensen–Shannon divergence (JSD) with log~2~ (range 0–1):")
    b.eq("JSD(P‖Q) = ½·KL(P‖M) + ½·KL(Q‖M),   M = ½(P + Q)", 1)
    b.para("Codes that appear or disappear simultaneously with equivalent volume are treated as recoding and harmonised before rules and models are run (Section 3.1).")
    b.h2("2.3 Domain rule engine")
    b.para("Eleven rules (Table 2) were built from PKB tariff logic and code patterns in the 2021–2023 reference years and mapped to the consistency, accuracy, plausibility and uniqueness dimensions [@ehrlinger]. "
           "Because the official code dictionary was unavailable, code consistency (R07, R08) was derived from data: application–mutation pairs with < 0.5% share within their application code (or n < 30) are considered unusual, while pairs that carry BBNKB in ≥ 95% (≤ 5%) of cases define the BBNKB-required (BBNKB-free) reference. "
           "Implausible PKB values for a vehicle type (R09) use a robust z-score:")
    b.eq("z~i~ = (log~10~ PKB~i~ − med~g~) / max(MAD~g~, 0.05),   |z~i~| > 3.5 ⇒ R09", 2)
    b.para("where *g* is the vehicle-type code and med~g~ and MAD~g~ are the median and normal-scaled median absolute deviation in the reference data. The rule score is the weighted sum of violations (weights 1–3, Table 2).")
    T2 = N["T02"][["ID", "Rule", "Dimensi", "Bobot", "Logika"]].copy()
    T2["Dimensi"] = T2["Dimensi"].map({"Konsistensi": "Consistency", "Akurasi": "Accuracy", "Plausibilitas": "Plausibility", "Keunikan": "Uniqueness"})
    T2["Logika"] = ["pkbden_k > 0 and pkbtgk_k = 0 (k = 0..4)", "pkbtgk_k > 0 and pkbden_k = 0",
                    "penalty_k/arrears_k outside [0.245; 0.2525] (k ≥ 1) or > 0.2525 (k = 0)", "pkbden/pkbpok > 0.2525 or pkbden > 0 with pkbpok = 0",
                    "arrears in year k+1 present while year k is empty", "total arrears > 0 and pkbpok = 0",
                    "(kd_jen_mohon, kd_jen_mutasi) pair < 0.5% within application code / n < 30, or rare code (< 50)",
                    "BBNKB-required pair (≥ 95%) without BBNKB, BBNKB-free pair (≤ 5%) with BBNKB, or bbnden > 0 without bbnpok",
                    "|robust z of log10 pkbpok within kd_jenis_kb| > 3.5 (reference median/MAD)", "pkbtgk0/pkbpok > 3 or < 0.3",
                    "exact duplicate (plate, date, codes, amounts)"]
    T2.columns = ["ID", "Rule", "Dimension", "Weight", "Logic"]
    b.table(T2, "Catalogue of domain rules for PKB transaction data quality", widths=[0.9, 3.6, 1.9, 1.0, 8.0], size=7.5)
    b.h2("2.4 Features and unsupervised detectors")
    b.para(f"Detectors were trained only on clean transactions (no rule violation and non-zero value) from 2021–2023 (n = {num(n_tr)}; training sample 40,000) and tested on 2024 to avoid temporal leakage [@kapoor]. "
           "Base features comprise logarithms of PKB, BBNKB and arrears principal/penalty amounts, number of arrears years, code frequencies and status codes. Domain features add penalty/principal, penalty/arrears and arrears/principal ratios, the BBNKB/PKB ratio, the per-type value z-score (equation 2), the BBNKB/PKB z-score per code pair, code-pair frequency and the deviation of BBNKB presence from the pair's norm. "
           "The three detectors are Isolation Forest (300 trees, ψ = 512) with score")
    b.eq("s(x) = 2^−E[h(x)]/c(ψ)^", 3)
    b.para("LOF in novelty mode (k = 35), and a 32–8–32 MLP autoencoder (ReLU, Adam, early stopping) with reconstruction error")
    b.eq("e(x) = (1/d) Σ~j~ (x~j~ − x̂~j~)²", 4)
    b.para("Shallow models and a small network were chosen because tabular data do not necessarily benefit from deep architectures [@shwartz; @borisov].")
    b.h2("2.5 Ensemble fusion and hybrid score")
    b.para("Each detector score *m* is calibrated into a percentile F~m~ against the training-score distribution and fused with the maximum (the mean breaks ties), so that a strong signal from one detector is not diluted by the others [@pang]. The hybrid score puts rule-violating transactions first:")
    b.eq("S~ML~(x) = max~m~ F~m~(s~m~(x));      H(x) = 𝟙[R(x) > 0] + S~ML~(x)", 5)
    b.para("A transaction is flagged by ML when S~ML~ exceeds the 99th percentile of the training data. Reason codes come from the violated rules and the three features with the largest autoencoder reconstruction error [@lixad]. Figure 2 shows how the detector is deployed in the SAMSAT information system.")
    b.figure(F("F02_arsitektur_sistem"), "Deployment architecture of the hybrid detector in the SAMSAT data flow", width=14.5)
    b.h2("2.6 Evaluation design")
    b.para("Because ground truth is unavailable, performance was measured with controlled anomaly injection that mimics real errors [@steinbuss]. In each run, 20,000 clean 2024 transactions were sampled and 800 anomalies (8 types × 100) were inserted (3.85% contamination), for ten runs (Table 3). "
           "Type A8 deliberately has no matching rule, to test the ability to find unanticipated patterns. Metrics are ROC-AUC, PR-AUC (average precision), precision at k = number of anomalies (P@k), recall at a 5% audit budget, and precision, recall and F1 at the operating threshold. "
           "Differences between methods were tested with Friedman and Holm-corrected paired Wilcoxon tests. Sensitivity analysis covered LOF k, number of IF trees, AE architecture, training size and fusion strategy; the ablation compared base features with base + domain features.")
    T3 = N["T06"][["Kode", "Anomaly type", "Aturan terkait"]].copy()
    T3["Aturan terkait"] = T3["Aturan terkait"].str.replace("sebagian", "partial").str.replace("- (tanpa aturan)", "– (no rule)", regex=False)
    T3["Injection"] = ["penalty of 2nd arrears year = 5–25% of principal, without arrears principal", "kd_jen_mutasi replaced by a code contradicting the application type",
                       "motorcycle ↔ car principal drawn from the other type's distribution", "PKB principal ×10 or ÷10, penalty unchanged",
                       "kd_jenis_kb motorcycle ↔ car/goods vehicle, amounts unchanged", "arrears penalty 5–18% or 32–50% of arrears",
                       "arrears principal & penalty ×2.5–5 (25% ratio preserved)", "BBNKB ×5–12 or ÷5–12"]
    T3.columns = ["Code", "Anomaly type", "Related rule", "Injection"]
    b.table(T3, "Injected anomaly types (100 per type per run)", widths=[1.0, 4.4, 2.4, 7.6], size=7.5)

    b.h1("3. Results and Discussion")
    b.h2("3.1 Data-quality profile and coding-scheme drift")
    b.para(f"Figure 3 shows that in 2024 status codes 3/4 were entirely replaced by 5/6 (JSD = {dec(d24.loc['kd_status', 'JSD_mentah'])}), and re-registrations with a mutation moved from application code 1 to 2 (pair 1|7 → 2|7 with equivalent volume; JSD = {dec(d24.loc['kd_jen_mohon|kd_jen_mutasi', 'JSD_mentah'])}). "
           f"After harmonisation the JSD fell to {dec(d24.loc['kd_status', 'JSD_harmonisasi'], 4)} and {dec(d24.loc['kd_jen_mohon|kd_jen_mutasi', 'JSD_harmonisasi'], 4)}. Without this step, thousands of 2024 transactions would be wrongly flagged as inconsistent codes and the ML detectors would suffer distribution shift, underlining the importance of drift monitoring on operational data [@bayram].")
    b.figure(F("F16_pergeseran_kode"), "Coding-scheme drift in 2024: JSD before/after harmonisation (left) and volume of code pairs (right)", width=14.5)
    b.para(f"Overall, {num(S['rule_flag'])} transactions ({dec(S['rule_flag_pct'], 2)}%) violated at least one rule (Figure 4). The most frequent violations were implausible PKB values for the vehicle type (R09, about 2.0–2.3% per year) and arrears without current-year principal (R06), which rose from {pct(t03.loc[2021, 'R06_pct'])} (2021) to {pct(t03.loc[2024, 'R06_pct'])} (2024). "
           f"Arrears without penalty (R02) and off-schedule arrears penalties (R03) appeared in 2021–2022 and disappeared by 2024, consistent with a possible pandemic-era penalty relaxation that should be confirmed with the regional revenue agency. "
           f"Conversely, R08 jumped to {num(t03.loc[2024, 'R08'])} transactions in 2024, mostly from pair 2|4 without BBNKB, i.e. from merging two code meanings into one new code. Recoding therefore does not only relabel values; it also removes semantic information. "
           f"In addition, there were {num(S['zero_value'])} zero-value transactions ({dec(100 * S['zero_value'] / S['n_transaksi'], 1)}%) and {num(S['plate_nonstandard'])} licence plates in a non-standard format.")
    b.figure(F("F03_prevalensi_aturan"), "Prevalence of domain-rule violations per year (percentage and number of transactions)", width=14.5)
    T4 = T04.copy(); T4.columns = ["Year", "Completeness", "Validity", "Consistency", "Accuracy", "Plausibility", "Uniqueness", "DQ index"]
    for c in T4.columns[1:]: T4[c] = T4[c].map(lambda v: dec(v, 2))
    T4["Rule violations"] = [f"{num(t03.loc[y, 'rule_flag'])} ({dec(t03.loc[y, 'rule_flag_pct'], 2)}%)" for y in T4["Year"]]
    b.table(T4, "SAMSAT data-quality scorecard per dimension (%) and number of rule violations", widths=[1.1, 1.9, 1.5, 1.8, 1.5, 1.8, 1.6, 1.5, 2.6])
    b.para(f"Table 4 shows a high data-quality index ({dec(t04['Indeks DQ'].min(), 2)}–{dec(t04['Indeks DQ'].max(), 2)}%), but the 2024 consistency dimension is the lowest ({dec(t04.loc[2024, 'Konsistensi'], 2)}%). High aggregate scores can hide high-value errors, so transaction-level detection remains necessary.")
    b.h2("3.2 Detector performance on injected anomalies")
    T5 = N["T07"].copy(); T5["Metode"] = T5["Metode"].map(MEN)
    T5 = T5[["Metode", "ROC_AUC", "PR_AUC", "P@k", "Recall@5%", "F1", "FPR (normal)"]]
    T5.columns = ["Method", "ROC-AUC", "PR-AUC", "P@k", "Recall@5%", "F1", "FPR normal"]
    b.table(T5, "Detector performance comparison (mean ± standard deviation, 10 runs, 2024 test)", widths=[2.8, 2.4, 2.4, 2.4, 2.4, 2.4, 1.8], bold_last=True)
    b.para(f"Table 5 and Figure 5 (ROC and precision–recall curves are given in the supplementary material) show that the hybrid detector is best on almost every metric: PR-AUC {gs('Hibrida', 'PR_AUC')}, ROC-AUC {g('Hibrida', 'ROC_AUC')}, P@k {g('Hibrida', 'P@k')} and Recall@5% {g('Hibrida', 'Recall@5%')}. "
           f"Rules alone have perfect precision (1.000) but capture only {dec(agg.loc['Aturan', ('Recall', 'mean')])} of the anomalies, while the autoencoder is the best single detector (PR-AUC {g('AE', 'PR_AUC')}). "
           f"Isolation Forest (PR-AUC {g('IF', 'PR_AUC')}) and LOF ({g('LOF', 'PR_AUC')}) are weak because the anomalies here are contextual—values that are plausible marginally but not with respect to the code or vehicle type—and are therefore hard to isolate by random partitions or local density in a mixed continuous–discrete feature space [@xu; @alghushairy]. "
           f"The Friedman test is significant (χ² = {dec(w9.Friedman_chi2.iloc[0], 1)}, p = {(f"{S['friedman_p']:.1e}".split('e')[0]).replace('.', '.')} × 10^{int(f"{S['friedman_p']:.1e}".split('e')[1])}^), and the hybrid beats every competitor in 10 of 10 runs (Wilcoxon–Holm p = {dec(w9.p_holm.max(), 4)}). "
           f"The hybrid F1 ({g('Hibrida', 'F1')}) is slightly below the rules ({g('Aturan', 'F1')}) because the P99 threshold adds false positives (precision {dec(agg.loc['Hibrida', ('Presisi', 'mean')])}); in audit practice, budget-based metrics (P@k, Recall@5%) are more relevant.")
    b.figure(F("F09_komparasi_metrik"), "Comparison of PR-AUC, P@k and F1 across detectors (mean ± sd, 10 runs)", width=14.5)
    b.h2("3.3 Per-type analysis, ablation and sensitivity")
    b.para(f"Figure 6 shows how the two layers complement each other. Rules capture A1–A5 almost perfectly but fail on A8 (implausible BBNKB; recall {dec(T08.loc['A8', 'Aturan'], 2)}), which has no rule. "
           f"The autoencoder captures A8 ({dec(T08.loc['A8', 'AE'], 2)}), A7 ({dec(T08.loc['A7', 'AE'], 2)}) and A6 ({dec(T08.loc['A6', 'AE'], 2)}) but is weak on swapped vehicle-type codes (A5, {dec(T08.loc['A5', 'AE'], 2)}). "
           f"The hybrid keeps both strengths: A8 rises from {dec(T08.loc['A8', 'Aturan'], 2)} (rules) to {dec(T08.loc['A8', 'Hibrida'], 2)}, while A1–A5 stay ≥ {dec(T08.loc[['A1', 'A2', 'A3', 'A4', 'A5'], 'Hibrida'].min(), 2)}. "
           "This agrees with data-driven audit practice in which rules and unsupervised detection should be combined rather than opposed [@westerski; @carcillo].")
    b.figure(F("F08_recall_per_tipe"), "Recall@k per anomaly type and detector (mean of 10 runs)", width=14.0)
    b.para(f"The ablation (Figure S1 in the supplementary material) shows that domain features raise the autoencoder PR-AUC from {ab('AE', 'dasar')} to {ab('AE', 'dasar+domain')}, the ensemble from {ab('Ensemble ML', 'dasar')} to {ab('Ensemble ML', 'dasar+domain')} and the hybrid from {ab('Hibrida', 'dasar')} to {ab('Hibrida', 'dasar+domain')}; only LOF decreases ({ab('LOF', 'dasar')} → {ab('LOF', 'dasar+domain')}) because extra dimensions degrade density estimation. "
           f"In the sensitivity analysis, maximum-percentile fusion gives an ensemble PR-AUC of {dec(sens.loc[('Strategi fusi', 'max persentil (dipakai)'), 'PR_AUC'])} versus {dec(sens.loc[('Strategi fusi', 'rerata persentil'), 'PR_AUC'])} for mean fusion, because averaging dilutes the autoencoder signal with the weak IF and LOF scores. "
           f"A wider AE (64–16–64) increases the AE PR-AUC to {dec(sens.loc[('AE arsitektur', '64-16-64'), 'PR_AUC'])}, while the hybrid PR-AUC remains stable at {dec(T10.PR_AUC_hibrida.min())}–{dec(T10.PR_AUC_hibrida.max())} across all parameter settings. "
           f"Mean training time on a two-core CPU was {dec(tm['fit_IF'], 2)} s (IF), {dec(tm['fit_LOF'], 2)} s (LOF) and {dec(tm['fit_AE'], 2)} s (AE) for 40,000 transactions, and scoring 20,800 transactions took < 1 s, which is feasible as a daily batch.")
    b.h2("3.4 Application to real data 2021–2024")
    T6 = T12[["tahun", "n", "aturan", "keduanya", "ml_saja", "nihil", "hibrida_total", "hibrida_pct"]].copy()
    T6.columns = ["Year", "Transactions", "Rule violations", "Rules & ML", "ML only", "Zero value (informative)", "Hybrid audit list", "% audit list"]
    for c in T6.columns[1:-1]: T6[c] = T6[c].map(lambda v: num(v))
    T6["% audit list"] = T6["% audit list"].map(lambda v: dec(v, 2))
    b.table(T6, "Application of the hybrid detector to real data per year", widths=[1.2, 1.9, 2.0, 1.9, 1.6, 2.4, 2.4, 1.9])
    b.para(f"Table 6 and Figure 7 summarise the application to {num(S['n_transaksi'])} transactions. The hybrid audit list covers {hyb_pct[0]}%, {hyb_pct[1]}%, {hyb_pct[2]}% and {hyb_pct[3]}% of transactions in 2021–2024. "
           f"A total of {num(S['ml_only'])} transactions were flagged only by ML; the main reasons are per-type value z-scores close to the rule limit, penalty amounts and unusual BBNKB (Table 7). "
           f"Zero-value transactions ({num(int(T12.nihil.sum()))} without rule violations) were separated as an informative category because they probably represent endorsements without payment. "
           f"The overlap in Figure 7 also shows that ML flags {dec(cross.loc['R06', 'persen_juga_ML'], 0)}% of R06 and {dec(cross.loc['R08', 'persen_juga_ML'], 0)}% of R08 transactions but only {dec(cross.loc['R09', 'persen_juga_ML'], 1)}% of R09 transactions; the robust-z limit of ±3.5 is therefore stricter than the pattern learned by the model and can be recalibrated with officers.")
    b.figure(F("F13_temuan_data_riil"), "Composition of flagged transactions per year (left) and overlap of rule violations with ML flags (right)", width=14.5)
    ex = pd.concat([N["T13"].head(2), N["T13b"].drop_duplicates(["kd_jen_mohon", "kd_jen_mutasi", "kd_jenis_kb", "pkbpok"]).head(4)])
    T7 = pd.DataFrame({"Year": ex.tahun, "Appl.|mutation": ex.kd_jen_mohon.map(lambda v: str(int(v))) + "|" + ex.kd_jen_mutasi.map(lambda v: "NA" if pd.isna(v) else str(int(float(v)))),
                       "Type": ex.kd_jenis_kb.astype(int), "PKB principal (Rp)": ex.pkbpok.map(lambda v: num(v)), "BBNKB (Rp)": ex.bbnpok.map(lambda v: num(v)),
                       "Arrears (Rp)": ex.tgk_total.map(lambda v: num(v)), "Rules": ex.aturan_dilanggar, "AE reasons (top 3)": ex.alasan_AE_top3.str.replace("_", " ")})
    b.table(T7, "Examples of audit-priority transactions with reason codes (vehicle identities pseudonymised)", widths=[1.0, 1.4, 1.0, 1.9, 1.9, 2.0, 1.3, 5.5], size=7)
    b.para("The examples in Table 7 illustrate actionable output: the two top transactions violate R06 and R08 (arrears paid without current-year principal under a code that normally carries BBNKB), whereas ML-only cases include a bus-fleet transfer with identical amounts, a registration of a rare vehicle code with a high BBNKB, and an unusual BBNKB/PKB ratio. "
           "Not all of them are errors; the system's role is to prioritise verification so that officers examine hundreds rather than hundreds of thousands of transactions.")
    b.h2("3.5 Implications and limitations")
    b.para("For government information systems, three practical implications follow. First, ISQL-WEB SAMSAT entry validation can be strengthened by running rules R01–R11 daily. Second, coding-scheme changes should be documented in a versioned data dictionary so that cross-year analytics are not biased. Third, the hybrid priority list with reason codes can support internal audit and revenue reconciliation. "
           "The limitations are that anomalies were evaluated through synthetic injection, so performance on real errors must be confirmed through officer verification; code semantics were derived from data because the official dictionary was unavailable; and the data cover only May–June in the SAMSAT offices studied.")

    b.h1("4. Conclusion")
    b.para(f"The **problem** addressed is the lack of a systematic mechanism to detect transaction anomalies and assess the quality of SAMSAT PKB data without ground-truth labels and amid coding-scheme changes. "
           f"The proposed **solution** is a three-layer hybrid detector that combines eleven domain rules, code-drift harmonisation and an Isolation Forest–LOF–autoencoder ensemble with maximum-percentile fusion. "
           f"The **results** on {num(S['n_transaksi'])} Banten transactions from 2021–2024 show PR-AUC {g('Hibrida', 'PR_AUC')} and ROC-AUC {g('Hibrida', 'ROC_AUC')} on the 2024 temporal test, higher than rules alone ({g('Aturan', 'PR_AUC')}) and the best single ML detector ({g('AE', 'PR_AUC')}); domain features raised the ensemble PR-AUC by {dec(T11.loc['Ensemble ML', ('mean', 'dasar+domain')] - T11.loc['Ensemble ML', ('mean', 'dasar')], 2)}; a 2024 recoding of status and application codes was detected; and {dec(S['rule_flag_pct'], 2)}% of transactions violated at least one rule. "
           "Future work will validate the priority list with SAMSAT officers to obtain real labels, extend coverage to all months and regions of Banten, and develop active learning from officer feedback.")
    b.blank()
    b.para("**Acknowledgments.** The authors thank SAMSAT officers in Banten Province who helped provide the data for experimental purposes. Code, the Google Colab notebook and result tables are available from the corresponding author.")
    b.references("References")
