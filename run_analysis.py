#!/usr/bin/env python3
"""
MADSRCH Case Study — Unified Analysis Pipeline Runner

Executes data generation, pooled overall findings, gender-split findings, and verification.

Usage:
    python run_analysis.py          # Runs both report pipelines (default)
    python run_analysis.py --all    # Runs both report pipelines
    python run_analysis.py --overall# Runs pooled overall report only
    python run_analysis.py --gender # Runs gender-split report only
    python run_analysis.py --seed   # Runs data anonymization & template generation
    python run_analysis.py --verify # Performs automated privacy & statistical checks
"""

import argparse
import subprocess
import sys
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent


def run_script(script_name: str) -> bool:
    script_path = ROOT / script_name
    print(f"\n========================================================", flush=True)
    print(f"  Executing: {script_name}", flush=True)
    print(f"========================================================", flush=True)
    try:
        result = subprocess.run([sys.executable, str(script_path)], check=True, cwd=ROOT)
        return result.returncode == 0
    except subprocess.CalledProcessError as e:
        print(f"Error running {script_name}: exited with status {e.returncode}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"Unexpected error running {script_name}: {e}", file=sys.stderr)
        return False


def verify_repository() -> bool:
    print(f"\n========================================================", flush=True)
    print(f"  Verifying Repository Privacy & Statistical Fidelity", flush=True)
    print(f"========================================================", flush=True)
    ok = True

    # 1. Prohibited files check
    prohibited = [
        "references/MADSRCH_FULL.pdf",
        "MADSRCH_FULL.pdf",
        "CASE_STUDY.7z",
        "blovs_excel.xlsx",
        "debug.log",
    ]
    for p in prohibited:
        if (ROOT / p).exists():
            print(f" [FAIL] Prohibited file present: {p}", file=sys.stderr)
            ok = False
        else:
            print(f" [PASS] Prohibited file absent: {p}")

    # 2. Anonymized dataset privacy check
    anon_file = ROOT / "blo_survey_anonymized.xlsx"
    if not anon_file.exists():
        print(f" [FAIL] Missing anonymized workbook: {anon_file}", file=sys.stderr)
        return False

    df = pd.read_excel(anon_file, sheet_name="SurveyData")
    if len(df) != 184:
        print(f" [FAIL] Expected 184 rows, got {len(df)}", file=sys.stderr)
        ok = False
    else:
        print(f" [PASS] Sample size verified: N=184")

    # Direct PII checks
    if df["Email_Optional"].notna().sum() > 0:
        print(f" [FAIL] Email_Optional contains non-null values!", file=sys.stderr)
        ok = False
    else:
        print(f" [PASS] Email_Optional 100% null (PII eradicated)")

    if df["Comments"].notna().sum() > 0:
        print(f" [FAIL] Comments contains non-null values!", file=sys.stderr)
        ok = False
    else:
        print(f" [PASS] Comments 100% null (free-text eradicated)")

    if df["Duplicate_Source_Row"].notna().sum() > 0:
        print(f" [FAIL] Duplicate_Source_Row contains non-null values!", file=sys.stderr)
        ok = False
    else:
        print(f" [PASS] Duplicate_Source_Row 100% null")

    # 3. Blank template check
    tpl_file = ROOT / "blo_survey_blank_template.xlsx"
    if not tpl_file.exists():
        print(f" [FAIL] Missing blank template: {tpl_file}", file=sys.stderr)
        ok = False
    else:
        tpl_df = pd.read_excel(tpl_file, sheet_name="SurveyData")
        if len(tpl_df) != 0 or len(tpl_df.columns) != 26:
            print(f" [FAIL] Blank template shape unexpected: {tpl_df.shape}", file=sys.stderr)
            ok = False
        else:
            print(f" [PASS] Blank template schema verified (0 rows, 26 columns)")

    # 4. Statistical Ground-Truth Check
    female_neg = (df[df["Gender"] == "Female"]["Brand_Credibility_Reaction"] == "It makes me see the brand more negatively.").sum()
    male_neg = (df[df["Gender"] == "Male"]["Brand_Credibility_Reaction"] == "It makes me see the brand more negatively.").sum()
    if female_neg == 62 and male_neg == 48:
        print(f" [PASS] Negative brand reaction counts match ground truth (62 F / 48 M)")
    else:
        print(f" [FAIL] Negative brand reaction counts mismatch: F={female_neg}, M={male_neg}", file=sys.stderr)
        ok = False

    female_human = df[df["Gender"] == "Female"]["Preferred_Ad_Type"].isin(["Definitely the Human-Created ad", "Somewhat prefer the Human-Created ad"]).sum()
    male_human = df[df["Gender"] == "Male"]["Preferred_Ad_Type"].isin(["Definitely the Human-Created ad", "Somewhat prefer the Human-Created ad"]).sum()
    if female_human == 74 and male_human == 69:
        print(f" [PASS] Human ad preference counts match ground truth (74 F / 69 M)")
    else:
        print(f" [FAIL] Human ad preference counts mismatch: F={female_human}, M={male_human}", file=sys.stderr)
        ok = False

    # 5. Charts check
    gender_charts = list((ROOT / "charts").glob("*.png"))
    overall_charts = list((ROOT / "charts_overall").glob("*.png"))
    if len(gender_charts) == 12:
        print(f" [PASS] 12/12 gender-split charts generated")
    else:
        print(f" [WARN] Found {len(gender_charts)}/12 gender-split charts")

    if len(overall_charts) == 10:
        print(f" [PASS] 10/10 pooled overall charts generated")
    else:
        print(f" [WARN] Found {len(overall_charts)}/10 pooled overall charts")

    print("--------------------------------------------------------")
    if ok:
        print(" [ALL VERIFICATION CHECKS PASSED] Repository is clean & compliant.")
    else:
        print(" [VERIFICATION COMPLETED WITH WARNINGS/ERRORS]", file=sys.stderr)
    return ok


def main():
    parser = argparse.ArgumentParser(description="Run MADSRCH analysis and chart generation pipelines.")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--all", action="store_true", help="Run both overall and gender reports")
    group.add_argument("--overall", action="store_true", help="Run overall findings report only")
    group.add_argument("--gender", action="store_true", help="Run gender-split findings report only")
    group.add_argument("--seed", action="store_true", help="Run data anonymization and seeding engine only")
    group.add_argument("--verify", action="store_true", help="Perform verification checks only")

    args = parser.parse_args()

    # Ensure dataset exists before running reports
    if not (ROOT / "blo_survey_anonymized.xlsx").exists() and not (args.seed or args.verify):
        print("Anonymized workbook not found. Seeding first...", flush=True)
        run_script("sanitize_and_seed_data.py")

    overall_ok = True
    gender_ok = True

    if args.seed:
        run_script("sanitize_and_seed_data.py")
        return
    elif args.verify:
        verify_repository()
        return
    elif args.overall:
        overall_ok = run_script("generate_overall_findings_report.py")
    elif args.gender:
        gender_ok = run_script("generate_gender_findings_report.py")
    else:  # --all or default
        overall_ok = run_script("generate_overall_findings_report.py")
        gender_ok = run_script("generate_gender_findings_report.py")

    print("\n--------------------------------------------------------", flush=True)
    if overall_ok and gender_ok:
        print(" [SUCCESS] Analysis pipeline completed successfully.", flush=True)
        print(" Summary workbooks and high-resolution charts generated.", flush=True)
        # Run quick auto-verification
        verify_repository()
    else:
        print(" [WARNING] Pipeline completed with errors.", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
