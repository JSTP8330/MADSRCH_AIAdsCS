# Consumer Perception and Behavioral Responses to AI-Labeled Commercial Ads
> **MADSRCH Research Project** | De La Salle University (DLSU)  
> *A quantitative study examining the impact of generative AI disclosure on consumer authenticity, brand credibility, engagement, and purchase intent among Filipino undergraduates.*

---

## 📌 Executive Summary

With major digital and social media platforms implementing mandatory AI disclosure badges (e.g., *"Made with AI"*, *"AI-generated"*), this study examines how college digital natives ($N = 184$, balanced 92 Male / 92 Female) perceive and interact with commercial advertisements labeled as AI-generated.

### Core Empirical Takeaways
* **Authenticity Penalty**: **67.9%** of students agree that AI-labeled ads feel *less authentic and genuine*, and **59.8%** consider them more likely to be *misleading*.
* **Immediate Brand Trust Deficit**: **59.8%** of respondents view a brand *more negatively* upon seeing an AI label (only **4.9%** view it positively).
* **Behavioral Avoidance**: **69.0%** actively *ignore* AI-labeled ads. When presented with identical products, **77.8%** prefer *human-created* advertising over AI ads.
* **Commercial Purchase Drop**: **53.3%** report being *less likely to purchase* a product after learning its ad used AI (**23.4%** much less likely).
* **Gender Discrepancy**: Female undergraduates demonstrate significantly higher skepticism and purchase penalties (**61.9%** less likely to buy vs. **44.6%** for males; **67.4%** negative brand view vs. **52.2%** for males).

---

## 📊 Key Findings at a Glance

| Metric / Question | Female (n=92) | Male (n=92) | Overall Pooled (n=184) | Primary Insight |
|:---|:---:|:---:|:---:|:---|
| **View Brand More Negatively** | **67.4%** | 52.2% | **59.8%** | AI disclosure inflicts an immediate brand perception penalty; female consumers react more critically. |
| **Feel Less Authentic (Top-Box 4–5)** | 67.4% | 68.5% | **67.9%** | Strong consensus across genders that AI-generated ads lack human authenticity. |
| **More Likely to be Misleading (Top-Box 4–5)** | 56.5% | **63.0%** | **59.8%** | Male students express greater concern regarding factual accuracy and deception. |
| **Likely to Actively Ignore (Top-Box 4–5)** | 62.0% | **76.1%** | **69.0%** | AI badges trigger ad-blindness and immediate scrolling past content. |
| **Prefer Human-Created Ad** | **80.5%** | 75.0% | **77.8%** | Overwhelming preference for human creative effort over algorithmic generation. |
| **Less Likely to Purchase** | **61.9%** | 44.6% | **53.3%** | Disclosure creates friction against purchase conversion. |

### Top Drivers of Consumer Skepticism
1. **Lack of human emotional connection / empathy** (71.7% overall; 78.3% female)
2. **Concern that the product or visual is fake/unrepresentative** (70.7% overall; 76.1% female)
3. **Perception of brand laziness or cost-cutting** (62.0% overall)
4. **Data privacy and algorithmic targeting concerns** (44.6% overall)

---

## 🔒 Data Privacy & Compliance Statement

This repository adheres to the **Philippine Data Privacy Act of 2012 (RA 10173)** and **GDPR Recital 26**:
* **Direct Identifiers Eradicated**: All optional emails, free-text comments containing names/usernames, and sub-second timestamps have been completely removed.
* **Mathematical Row-Level Unlinking**: An independent column-level permutation was applied within each gender stratum. Every record in `blo_survey_anonymized.xlsx` is a synthetic composite, mathematically preventing cross-question profiling while preserving 100.0% of marginal frequencies, crosstabs, and chart percentages.
* **Blank Reusable Schema**: A 0-row schema template is included in `blo_survey_blank_template.xlsx` for replication studies.
* For full protocol details, see [`docs/PRIVACY_AND_ETHICS.md`](file:///docs/PRIVACY_AND_ETHICS.md).

---

## 📂 Repository Structure

```text
MADSRCH_AIAdsCS/
├── .gitignore                             # Standard Python, Office lock, and OS rules
├── LICENSE                                # MIT License (Code) / CC-BY-4.0 (Documentation)
├── README.md                              # Repository overview, methodology, & findings
├── requirements.txt                       # Minimal dependencies (pandas, openpyxl, matplotlib)
│
├── run_analysis.py                        # Unified CLI runner (--all, --overall, --gender, --verify)
├── sanitize_and_seed_data.py              # Anonymization engine & blank template generator
├── generate_gender_findings_report.py     # Gender-disaggregated (92M / 92F) report & chart generator
├── generate_overall_findings_report.py    # Pooled sample (n=184) report & chart generator
│
├── blo_survey_anonymized.xlsx             # Anonymized research dataset (SurveyData + Field_Definitions)
├── blo_survey_blank_template.xlsx         # Blank 0-row schema template for replication
├── findings_summary.xlsx                  # Male vs Female percentage & headcount tables
├── findings_summary_overall.xlsx          # Pooled overall percentage & headcount tables
│
├── charts/                                # 12 publication-ready PNG charts (Gender split)
├── charts_overall/                        # 10 publication-ready PNG charts (Overall pooled)
│
├── docs/
│   ├── FINDINGS_OUTLINE.md                # Qualitative & quantitative findings (Gender split)
│   ├── FINDINGS_OUTLINE_OVERALL.md        # Qualitative & quantitative findings (Overall pooled)
│   ├── PRIVACY_AND_ETHICS.md              # RA 10173 / GDPR compliance documentation
│   └── QUESTIONNAIRE.PDF                  # Blank 4-page survey questionnaire instrument
│
└── ppt_slides/                            # 13 Full HD (1920x1080) presentation slides (0.png - 12.png)
```

---

## 🚀 Quickstart & Reproduction

### Prerequisites
* Python 3.10 or higher
* Recommended: Virtual environment (`venv`)

### 1. Installation
```bash
git clone https://github.com/<your-username>/MADSRCH_AIAdsCS.git
cd MADSRCH_AIAdsCS

# Optional: Set up virtual environment
python -m venv .venv
# On Windows:
.venv\Scripts\activate
# On macOS/Linux:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Run Analysis & Verification
```bash
# Run both report generators and automated verification
python run_analysis.py

# Or run selectively:
python run_analysis.py --overall   # Generates findings_summary_overall.xlsx and charts_overall/
python run_analysis.py --gender    # Generates findings_summary.xlsx and charts/
python run_analysis.py --verify    # Runs privacy & statistical sanity checks
```

---

## 🔬 Methodology & Measurement Standards

* **Sampling Design**: Purposive sampling of enrolled DLSU undergraduates ($N=184$, balanced 92 Male, 92 Female).
* **Likert Scale Treatment**: The analysis avoids reporting simple arithmetic means for ordinal 1–5 Likert scales (which obscure multimodal polarization). Instead, it evaluates:
  1. **Top-Box Agreement (%)**: Ratings of 4 (*Agree / Likely*) and 5 (*Strongly Agree / Very Likely*).
  2. **100% Stacked Distribution Bars**: Horizontal bars capturing the full response spectrum from 1 (*Strongly Disagree*) to 5 (*Strongly Agree*).
* **Color Palettes**:
  * Female: Coral/Magenta (`#A82470`)
  * Male: Steel Blue (`#136395`)
  * Overall: Saffron Yellow (`#F4C430`)

---

## 📄 References & Questionnaire
* Questionnaire Instrument: [`docs/QUESTIONNAIRE.PDF`](file:///docs/QUESTIONNAIRE.PDF)
* Presentation Slides: [`ppt_slides/`](file:///ppt_slides)
* Course: MADSRCH (Marketing Research), Term 3, DLSU.
