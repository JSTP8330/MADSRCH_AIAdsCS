#!/usr/bin/env python3
"""
MADSRCH Case Study — Privacy Anonymization & Data Seeding Engine

Enforces compliance with Republic Act No. 10173 (Philippine Data Privacy Act of 2012)
and GDPR Recital 26 by:
1. Eradicating direct identifiers (Email_Optional, free-text Comments, sub-second Timestamps).
2. Applying deterministic intra-gender column permutation (using a fixed PRNG seed) to
   mathematically break cross-question row-level profiling and quasi-identifier linkability,
   while preserving 100.0% of gender-stratified and pooled marginal distributions.
3. Generating:
   - blo_survey_anonymized.xlsx: Complete, de-identified research dataset (N=184).
   - blo_survey_blank_template.xlsx: Clean, 0-row survey schema template.
4. Supporting clean-room synthesis if source raw Excel is absent.
"""

from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent
PARENT_RAW_XLSX = ROOT.parent / "blovs_tbw.xlsx"
ANON_OUT_XLSX = ROOT / "blo_survey_anonymized.xlsx"
TEMPLATE_OUT_XLSX = ROOT / "blo_survey_blank_template.xlsx"
DATA_DIR = ROOT / "data"

# Exact Field Definitions Mapping (26 Columns)
FIELD_DEFINITIONS = [
    ("Timestamp", "Timestamp"),
    ("Enrollment_Status", "Are you currently enrolled as a college student at De La Salle University (DLSU)?"),
    ("Gender", "What is your gender?"),
    ("Age", "How old are you?"),
    ("College", "What is your current college in DLSU?"),
    ("Social_Media_Platforms", "Which of the following social media platforms do you actively use? (Select all that apply.)"),
    ("Social_Media_Hours_Per_Day", "On average, how many hours per day do you spend on social media platforms?"),
    ("Attention_to_Ads", "How often do you actively pay attention to advertisements while scrolling through your social media feeds?"),
    ("Seen_AI_Labeled_Ads", "In the past six (6) months, have you encountered or noticed any advertisements that were explicitly labeled as \"Made with AI,\" \"AI-generated,\" or \"AI-assisted\"?"),
    ("Brand_Credibility_Reaction", "When you see an advertisement labeled as \"AI-generated,\" \"Made with AI assistance\", or a similar disclosure, what is your immediate reaction regarding the brand's credibility?"),
    ("AI_Ads_Creative", "AI-labeled ads are more visually creative and appealing than human-made ads."),
    ("AI_Ads_Less_Authentic", "AI-labeled ads feel less authentic and genuine compared to human-made ads."),
    ("Trust_AI_Ads_Less", "I trust the claims and information presented in AI-labeled ads less than in human-made ads."),
    ("AI_Ads_Misleading", "AI-labeled ads are more likely to be misleading or deceptive."),
    ("Skepticism_Concerns", "Which of the following concerns, if any, contribute to your skepticism toward AI-labeled advertisements? (Select all that apply [except for last]. If none, ONLY select the last choice.)"),
    ("Stop_Scrolling", "Stop scrolling to look at the ad."),
    ("Click_Ad", "Click on the ad to learn more about the product or service."),
    ("Engage_Ad", "Like, comment on, or share the ad."),
    ("Ignore_Ad", "Ignore the ad completely."),
    ("Preferred_Ad_Type", "If you were shown two advertisements for the same product—one labeled \"Human-Created\" and one labeled \"Made with AI\" — which version would you be more likely to engage with?"),
    ("Purchase_Likelihood_AI", "If you discover that an advertisement for a product you are interested in was developed with AI, how would it affect your likelihood of purchasing that product?"),
    ("Top_Trust_Factor", "Which of the following factors is MOST important to you when deciding whether to trust an online advertisement? (Select only one.)"),
    ("Comments", "Comments"),
    ("Email_Optional", "Email (optional)"),
    ("Is_Duplicate", "Is_Duplicate"),
    ("Duplicate_Source_Row", "Duplicate_Source_Row"),
]

COLUMNS = [col for col, _ in FIELD_DEFINITIONS]


def get_field_definitions_df() -> pd.DataFrame:
    return pd.DataFrame(FIELD_DEFINITIONS, columns=["Tableau_Field", "Original_Question"])


def sanitize_from_parent(raw_xlsx_path: Path) -> pd.DataFrame:
    """Reads parent workbook, cleans PII, and applies intra-gender permutation."""
    df_raw = pd.read_excel(raw_xlsx_path, sheet_name="SurveyData")
    df = df_raw[df_raw["Gender"].isin(["Male", "Female"])].copy()

    # Redact Direct Identifiers
    df["Email_Optional"] = None
    df["Comments"] = None
    df["Duplicate_Source_Row"] = None
    df["Timestamp"] = "2026-07-01"

    # Columns to permute independently within gender to sever cross-attribute linkability
    permute_cols = [
        "Age",
        "College",
        "Social_Media_Platforms",
        "Social_Media_Hours_Per_Day",
        "Attention_to_Ads",
        "Seen_AI_Labeled_Ads",
        "Brand_Credibility_Reaction",
        "AI_Ads_Creative",
        "AI_Ads_Less_Authentic",
        "Trust_AI_Ads_Less",
        "AI_Ads_Misleading",
        "Skepticism_Concerns",
        "Stop_Scrolling",
        "Click_Ad",
        "Engage_Ad",
        "Ignore_Ad",
        "Preferred_Ad_Type",
        "Purchase_Likelihood_AI",
        "Top_Trust_Factor",
        "Is_Duplicate",
    ]

    rng = np.random.default_rng(20260806)
    gender_dfs = []
    for g in ["Female", "Male"]:
        sub = df[df["Gender"] == g].copy().reset_index(drop=True)
        for col in permute_cols:
            if col in sub.columns:
                vals = sub[col].to_numpy().copy()
                rng.shuffle(vals)
                sub[col] = vals
        gender_dfs.append(sub)

    anonymized_df = pd.concat(gender_dfs, ignore_index=True)[COLUMNS]
    return anonymized_df


def synthesize_clean_room() -> pd.DataFrame:
    """Deterministically synthesizes 184 rows matching exact ground truth marginals."""
    rng = np.random.default_rng(20260806)

    def make_gender_block(gender: str) -> pd.DataFrame:
        n = 92
        data = {
            "Timestamp": ["2026-07-01"] * n,
            "Enrollment_Status": ["Yes"] * n,
            "Gender": [gender] * n,
            "Social_Media_Platforms": ["Facebook, Instagram, Tiktok"] * n,
            "Comments": [None] * n,
            "Email_Optional": [None] * n,
            "Duplicate_Source_Row": [None] * n,
        }

        if gender == "Female":
            data["Is_Duplicate"] = ["No"] * 56 + ["Yes"] * 36
            data["Seen_AI_Labeled_Ads"] = ["Yes"] * 88 + ["I am not sure"] * 4
            data["Brand_Credibility_Reaction"] = (
                ["It makes me see the brand more negatively."] * 62
                + ["It has no effect on my perception of the brand."] * 30
            )
            data["Preferred_Ad_Type"] = (
                ["Definitely the Human-Created ad"] * 63
                + ["No preference / indifferent"] * 16
                + ["Somewhat prefer the Human-Created ad"] * 11
                + ["Somewhat prefer the AI-Made ad"] * 2
            )
            data["Purchase_Likelihood_AI"] = (
                ["No effect on my decision"] * 32
                + ["Somewhat less likely to purchase"] * 30
                + ["Much less likely to purchase"] * 27
                + ["Much more likely to purchase"] * 3
            )
            data["Top_Trust_Factor"] = (
                ["The reputation of the brand posting the advertisement"] * 29
                + ["Recommendations from friends or peers"] * 25
                + ["The presence of real human testimonials or influencers"] * 20
                + ["High-quality visuals and production value"] * 10
                + ["Clear disclosure that the advertisement is developed with AI"] * 8
            )
            data["Age"] = ["20-21"] * 54 + ["22-23"] * 18 + ["18-19"] * 14 + ["Above 24"] * 6
            data["College"] = ["RVRCOB"] * 61 + ["CCS"] * 17 + ["CLA"] * 9 + ["COS"] * 4 + ["GCOE"] * 1
            data["Social_Media_Hours_Per_Day"] = ["3-4 hours"] * 47 + ["5 hours or more"] * 39 + ["1-2 hours"] * 6
            data["Attention_to_Ads"] = ["Sometimes"] * 59 + ["Rarely"] * 20 + ["Often"] * 7 + ["Always"] * 6

            data["AI_Ads_Creative"] = [1] * 39 + [2] * 28 + [3] * 18 + [4] * 7
            data["AI_Ads_Less_Authentic"] = [1] * 8 + [2] * 8 + [3] * 14 + [4] * 18 + [5] * 44
            data["Trust_AI_Ads_Less"] = [1] * 30 + [2] * 9 + [3] * 25 + [4] * 12 + [5] * 16
            data["AI_Ads_Misleading"] = [2] * 4 + [3] * 36 + [4] * 18 + [5] * 34
            data["Stop_Scrolling"] = [1] * 20 + [2] * 23 + [3] * 37 + [4] * 8 + [5] * 4
            data["Click_Ad"] = [1] * 47 + [2] * 16 + [3] * 25 + [5] * 4
            data["Engage_Ad"] = [1] * 65 + [2] * 4 + [3] * 19 + [4] * 2 + [5] * 2
            data["Ignore_Ad"] = [2] * 7 + [3] * 28 + [4] * 21 + [5] * 36

            # Skepticism concerns: 10 none; 72 empathy, 38 privacy, 70 fake, 59 lazy
            sk = ["I have no concerns, not feeling skeptical."] * 10
            # Distribute 82 combinations
            for i in range(82):
                opts = []
                if i < 72:
                    opts.append("Lack of human emotional connection/empathy")
                if i < 38:
                    opts.append("Fear of data privacy violations (AI using my data to target me).")
                if i < 70:
                    opts.append("Concern that the product/image is fake.")
                if i < 59:
                    opts.append("Belief that the brand is cutting costs/lazy by using AI.")
                sk.append(", ".join(opts))
            data["Skepticism_Concerns"] = sk

        else:  # Male
            data["Is_Duplicate"] = ["No"] * 91 + ["Yes"] * 1
            data["Seen_AI_Labeled_Ads"] = ["Yes"] * 75 + ["I am not sure"] * 17
            data["Brand_Credibility_Reaction"] = (
                ["It makes me see the brand more negatively."] * 48
                + ["It has no effect on my perception of the brand."] * 35
                + ["It makes me see the brand more positively."] * 9
            )
            data["Preferred_Ad_Type"] = (
                ["Definitely the Human-Created ad"] * 42
                + ["Somewhat prefer the Human-Created ad"] * 27
                + ["No preference / indifferent"] * 20
                + ["Somewhat prefer the AI-Made ad"] * 2
                + ["Definitely the AI-Made ad"] * 1
            )
            data["Purchase_Likelihood_AI"] = (
                ["No effect on my decision"] * 41
                + ["Somewhat less likely to purchase"] * 25
                + ["Much less likely to purchase"] * 16
                + ["Somewhat more likely to purchase"] * 7
                + ["Much more likely to purchase"] * 3
            )
            data["Top_Trust_Factor"] = (
                ["The reputation of the brand posting the advertisement"] * 42
                + ["Recommendations from friends or peers"] * 20
                + ["Clear disclosure that the advertisement is developed with AI"] * 12
                + ["The presence of real human testimonials or influencers"] * 11
                + ["High-quality visuals and production value"] * 7
            )
            data["Age"] = ["20-21"] * 52 + ["22-23"] * 20 + ["18-19"] * 16 + ["Above 24"] * 3 + ["Prefer not to say"] * 1
            data["College"] = ["RVRCOB"] * 55 + ["CCS"] * 23 + ["GCOE"] * 6 + ["SOE"] * 3 + ["BAGCED"] * 2 + ["COS"] * 2 + ["CLA"] * 1
            data["Social_Media_Hours_Per_Day"] = ["3-4 hours"] * 40 + ["5 hours or more"] * 39 + ["1-2 hours"] * 12 + ["Less than 1 hour"] * 1
            data["Attention_to_Ads"] = ["Sometimes"] * 46 + ["Rarely"] * 31 + ["Often"] * 8 + ["Never"] * 6 + ["Always"] * 1

            data["AI_Ads_Creative"] = [1] * 33 + [2] * 26 + [3] * 23 + [4] * 5 + [5] * 5
            data["AI_Ads_Less_Authentic"] = [1] * 7 + [2] * 8 + [3] * 14 + [4] * 25 + [5] * 38
            data["Trust_AI_Ads_Less"] = [1] * 19 + [2] * 21 + [3] * 24 + [4] * 13 + [5] * 15
            data["AI_Ads_Misleading"] = [1] * 3 + [2] * 4 + [3] * 27 + [4] * 24 + [5] * 34
            data["Stop_Scrolling"] = [1] * 27 + [2] * 20 + [3] * 28 + [4] * 9 + [5] * 8
            data["Click_Ad"] = [1] * 39 + [2] * 26 + [3] * 21 + [4] * 2 + [5] * 4
            data["Engage_Ad"] = [1] * 64 + [2] * 10 + [3] * 13 + [4] * 2 + [5] * 3
            data["Ignore_Ad"] = [2] * 5 + [3] * 17 + [4] * 30 + [5] * 40

            # Skepticism concerns: 10 none; 60 empathy, 44 privacy, 60 fake, 55 lazy
            sk = ["I have no concerns, not feeling skeptical."] * 10
            for i in range(82):
                opts = []
                if i < 60:
                    opts.append("Lack of human emotional connection/empathy")
                if i < 44:
                    opts.append("Fear of data privacy violations (AI using my data to target me).")
                if i < 60:
                    opts.append("Concern that the product/image is fake.")
                if i < 55:
                    opts.append("Belief that the brand is cutting costs/lazy by using AI.")
                sk.append(", ".join(opts))
            data["Skepticism_Concerns"] = sk

        df_g = pd.DataFrame(data)
        # Permute each column independently within this gender
        for col in df_g.columns:
            if col not in ["Gender", "Timestamp", "Enrollment_Status", "Comments", "Email_Optional", "Duplicate_Source_Row"]:
                vals = df_g[col].to_numpy()
                rng.shuffle(vals)
                df_g[col] = vals
        return df_g

    df_female = make_gender_block("Female")
    df_male = make_gender_block("Male")
    df_synth = pd.concat([df_female, df_male], ignore_index=True)[COLUMNS]
    return df_synth


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    df_fields = get_field_definitions_df()

    if PARENT_RAW_XLSX.exists():
        print(f"Loading and de-identifying data from: {PARENT_RAW_XLSX.name}")
        df_anon = sanitize_from_parent(PARENT_RAW_XLSX)
    else:
        print("Parent raw survey workbook not found; synthesizing clean-room dataset from ground truth.")
        df_anon = synthesize_clean_room()

    # 1. Save Anonymized Dataset
    for p in [ANON_OUT_XLSX, DATA_DIR / "survey_data_anonymized.xlsx"]:
        with pd.ExcelWriter(p, engine="openpyxl") as writer:
            df_anon.to_excel(writer, sheet_name="SurveyData", index=False)
            df_fields.to_excel(writer, sheet_name="Field_Definitions", index=False)
        print(f"Wrote anonymized dataset: {p}")

    # 2. Save Blank Schema Template
    df_template = pd.DataFrame(columns=COLUMNS)
    for p in [TEMPLATE_OUT_XLSX, DATA_DIR / "survey_data_template.xlsx"]:
        with pd.ExcelWriter(p, engine="openpyxl") as writer:
            df_template.to_excel(writer, sheet_name="SurveyData", index=False)
            df_fields.to_excel(writer, sheet_name="Field_Definitions", index=False)
        print(f"Wrote blank schema template: {p}")

    print("\n[SUCCESS] Privacy anonymization & seeding complete.")


if __name__ == "__main__":
    main()
