# Data Privacy & Research Ethics Protocol

> **Compliance Framework**: Republic Act No. 10173 (Philippine Data Privacy Act of 2012 / DPA), National Privacy Commission (NPC) Guidelines on Research Data, and Regulation (EU) 2016/679 (GDPR) Recital 26.

---

## 1. Ethical Context & Research Purpose

This repository houses the computational analysis, summary workbooks, and visualization code for an undergraduate market research study conducted at De La Salle University (DLSU) on consumer trust toward AI-labeled commercial advertisements.

In academic research, ethical data stewardship requires protecting survey participants from any risk of re-identification while preserving the open reproducibility of scientific findings.

---

## 2. Privacy Threat Model & Risk Analysis

The original raw survey collection mechanism (Google Forms) captured both direct personal identifiers and quasi-identifying attributes:

1. **Direct Personal Identifiers**:
   - `Email_Optional`: Respondents optionally supplied personal email addresses for survey coordination.
   - `Comments`: An unstructured free-text field where respondents occasionally referenced individual names (`"Elton Dam"`, `"Franco"`) or social media usernames (`"@budimaniko"`).
2. **Quasi-Identifiers & Re-identification Risk ($k$-Anonymity Violation)**:
   - The demographic combination of `Gender` + `Age` + `College` + `Timestamp` created singleton ($k=1$) or near-singleton combinations within small student sub-populations (e.g., 1 Female in GCOE, 1 Male in CLA, 2 Males in BAGCED).
   - If raw row records were published, an observer with background knowledge could pinpoint an individual student's responses across all sensitive perception, trust, and behavioral questions.
3. **Third-Party Intellectual Property & Instructor PII**:
   - The original coursework archive included an instructor slide deck containing the professor's personal phone number, private email addresses, and copyrighted textbook diagrams. This file was strictly excluded from this repository.

---

## 3. Anonymization & De-identification Methodology

In accordance with **RA 10173 Section 4(c)** and **GDPR Recital 26** (which states that data protection principles do not apply to anonymous information that does not relate to an identified or identifiable natural person), the following transformation pipeline was implemented in [`sanitize_and_seed_data.py`](file:///sanitize_and_seed_data.py):

### A. Direct Identifier Eradication
- `Email_Optional`: 100% of values are redacted/nullified (`NaN`).
- `Comments`: All free-text inputs are stripped and nullified (`NaN`) to ensure zero informal handles or names persist.
- `Duplicate_Source_Row`: Internal row indexes from oversampling provenance are deleted.
- `Timestamp`: Precise sub-second timestamps are replaced with a static date `"2026-07-01"` or generalized sequential anonymized tokens (`ANON_001` .. `ANON_184`).

### B. Mathematical Unlinking via Intra-Gender Column Permutation
Because every chart, table, and statistical metric reported in this study is either:
1. An overall marginal distribution across the pooled sample ($N=184$), or
2. A bivariate crosstab/aggregation conditioned solely on `Gender` (`Female` $n=92$, `Male` $n=92$),

we apply an **independent column-level permutation** within each gender partition using a fixed pseudo-random number generator (`np.random.default_rng(20260806)`).

* **Mechanism**: Within the 92 Female records, the values for each question are shuffled independently; likewise for the 92 Male records.
* **Privacy Guarantee**: Every individual row in `blo_survey_anonymized.xlsx` becomes a synthetic composite vector that never corresponded to any single real human respondent. Cross-question row linkage and demographic profiling are mathematically broken.
* **Scientific Utility**: Because marginal and gender-stratified frequency distributions are invariant under permutation within strata, **all summary percentages, top-box agreement rates, counts, Excel tables, and Matplotlib visualizations remain 100.0% identical to the published findings**.

---

## 4. Dataset Offerings in this Repository

To support open science without privacy compromise, two datasets are provided in the repository root:

1. **`blo_survey_anonymized.xlsx`**:
   - Contains sheets `SurveyData` ($184 \times 26$) and `Field_Definitions` ($26 \times 2$).
   - Completely de-identified and safe for public benchmarking.
   - Powers the automated execution of `generate_gender_findings_report.py` and `generate_overall_findings_report.py`.
2. **`blo_survey_blank_template.xlsx`**:
   - Contains sheet `SurveyData` with the exact 26-column header and **0 data rows**.
   - Contains sheet `Field_Definitions` with full question strings.
   - Serves as a ready-to-use template for future replication studies.
