"""
MADSRCH Case Study — Male vs Female findings report generator.

Reads the balanced survey workbook (does NOT modify originals), computes
percentages by gender (Likert top-box 4–5 % and full 1–5 distributions),
exports findings_summary.xlsx, and saves presentation-style charts under charts/.

Run:
  python -m pip install -r requirements.txt
  python generate_gender_findings_report.py

Primary input: blo_survey_anonymized.xlsx (SurveyData)
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows

ROOT = Path(__file__).resolve().parent
SOURCE_XLSX = ROOT / "blo_survey_anonymized.xlsx"
if not SOURCE_XLSX.exists():
    SOURCE_XLSX = ROOT / "data" / "survey_data_anonymized.xlsx"
SOURCE_SHEET = "SurveyData"
FIELD_SHEET = "Field_Definitions"
OUT_XLSX = ROOT / "findings_summary.xlsx"
CHART_DIR = ROOT / "charts"

# Female left (coral), Male right (blue) — matches sample chart
COLOR_FEMALE = "#A82470"
COLOR_MALE = "#136395"
PLOT_GENDERS = ["Female", "Male"]  # legend / bar order
# Likert 1–5 segment colors (disagree/unlikely → agree/likely)
LIKERT_COLORS = {
    1: "#C0392B",
    2: "#E67E22",
    3: "#95A5A6",
    4: "#3498DB",
    5: "#1A5276",
}
LIKERT_SCORES = [1, 2, 3, 4, 5]
AGREE_LEGEND = {
    1: "1 Strongly disagree",
    2: "2 Disagree",
    3: "3 Neutral",
    4: "4 Agree",
    5: "5 Strongly agree",
}
LIKELY_LEGEND = {
    1: "1 Very unlikely",
    2: "2 Unlikely",
    3: "3 Neutral",
    4: "4 Likely",
    5: "5 Very likely",
}
# Awareness pie: Yes = gender accent; Not sure = lighter tint of same family
PIE_FEMALE = [COLOR_FEMALE, "#E8B7D0"]  # Yes coral, Not sure light coral
PIE_MALE = [COLOR_MALE, "#A8C8DC"]  # Yes blue, Not sure light blue


def _pie_label_color(hex_color: str) -> str:
    """White text on dark slices, dark text on light slices."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return "#FFFFFF" if luminance < 0.62 else "#1A1A1A"


def _style_pie_pct_labels(
    autotexts,
    texts,
    wedges,
    heights: list[float],
    slice_colors: list[str],
) -> None:
    """Keep % labels inside slices; nudge small-slice % outward (still inside)."""
    for autotext, text, wedge, h, color in zip(
        autotexts, texts, wedges, heights, slice_colors
    ):
        text.set_fontsize(12)
        text.set_color("#333333")
        text.set_fontweight("bold")
        text.set_clip_on(False)
        autotext.set_fontweight("bold")
        autotext.set_clip_on(False)
        autotext.set_ha("center")
        autotext.set_va("center")
        # Readable sizes that fit inside wedges (no explode)
        if h < 15:
            autotext.set_fontsize(15)
            # Outer part of the wedge is wider; keep radius < 1 so % stays inside
            ang = np.deg2rad((wedge.theta2 + wedge.theta1) / 2.0)
            r = 0.72
            autotext.set_position((r * np.cos(ang), r * np.sin(ang)))
        else:
            autotext.set_fontsize(16)
        autotext.set_color(_pie_label_color(color))


def styled_awareness_pies(
    pct: pd.DataFrame,
    category_keys: list[str],
    title: str,
    subtitle: str,
    outfile: Path,
) -> None:
    """Two side-by-side pies (Female | Male) for awareness Yes vs Not sure."""
    labels = [AWARE_LABELS[c] for c in category_keys]
    fig, axes = plt.subplots(1, 2, figsize=(12.0, 6.2))
    fig.patch.set_facecolor("white")

    for ax, gender in zip(axes, PLOT_GENDERS):
        heights = [
            float(pct.loc[gender, c]) if pd.notna(pct.loc[gender, c]) else 0.0
            for c in category_keys
        ]
        slice_colors = (PIE_FEMALE if gender == "Female" else PIE_MALE)[: len(category_keys)]

        def autopct_fmt(p: float) -> str:
            if p <= 0:
                return ""
            return f"{p:.1f}%"

        wedges, texts, autotexts = ax.pie(
            heights,
            labels=labels,
            colors=slice_colors,
            autopct=autopct_fmt,
            startangle=70,
            pctdistance=0.55,
            labeldistance=1.12,
            textprops={"fontsize": 12},
            wedgeprops={"linewidth": 2, "edgecolor": "white"},
        )
        _style_pie_pct_labels(autotexts, texts, wedges, heights, slice_colors)
        title_color = COLOR_FEMALE if gender == "Female" else COLOR_MALE
        ax.set_title(gender, fontsize=13, fontweight="bold", color=title_color, pad=10)
        ax.set(aspect="equal")
        # Room inside axes for outside category labels (avoids edge clipping)
        ax.set_xlim(-1.45, 1.45)
        ax.set_ylim(-1.45, 1.45)

    fig.suptitle(title, fontsize=14, fontweight="bold", y=0.98)
    fig.text(0.5, 0.88, subtitle, ha="center", va="top", fontsize=10, color="#555555", style="italic")
    fig.subplots_adjust(left=0.04, right=0.96, bottom=0.06, top=0.74, wspace=0.30)
    outfile.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(outfile, dpi=160, facecolor="white", bbox_inches="tight", pad_inches=0.55)
    plt.close(fig)


LIKERT_TRUST = [
    ("AI_Ads_Creative", "More visually creative\n& appealing"),
    ("AI_Ads_Less_Authentic", "Feel less authentic\n& genuine"),
    ("Trust_AI_Ads_Less", "Trust claims/\ninfo less"),
    ("AI_Ads_Misleading", "More likely to be\nmisleading/deceptive"),
]

LIKERT_ENGAGE = [
    ("Stop_Scrolling", "Stop scrolling\nto look"),
    ("Click_Ad", "Click to learn\nmore"),
    ("Engage_Ad", "Like, comment,\nor share"),
    ("Ignore_Ad", "Ignore the ad\ncompletely"),
]

CRED_LABELS = {
    "It makes me see the brand more positively.": "More positively",
    "It has no effect on my perception of the brand.": "No effect",
    "It makes me see the brand more negatively.": "More negatively",
}
CRED_ORDER = list(CRED_LABELS.keys())

PURCHASE_LABELS = {
    "Much less likely to purchase": "Much less\nlikely",
    "Somewhat less likely to purchase": "Somewhat\nless likely",
    "No effect on my decision": "No effect",
    "Somewhat more likely to purchase": "Somewhat\nmore likely",
    "Much more likely to purchase": "Much more\nlikely",
}
PURCHASE_ORDER = list(PURCHASE_LABELS.keys())

PREF_LABELS = {
    "Definitely the Human-Created ad": "Definitely\nHuman",
    "Somewhat prefer the Human-Created ad": "Somewhat\nHuman",
    "No preference / indifferent": "No\npreference",
    "Somewhat prefer the AI-Made ad": "Somewhat\nAI",
    "Definitely the AI-Made ad": "Definitely\nAI",
}
PREF_ORDER = list(PREF_LABELS.keys())

AWARE_LABELS = {"Yes": "Yes", "I am not sure": "Not sure"}
AWARE_ORDER = ["Yes", "I am not sure"]

# Known multi-select options (match by substring; last option contains a comma)
SKEPTICISM_LABELS = {
    "Lack of human emotional connection/empathy": "Lack of human\nemotional connection",
    "Fear of data privacy violations (AI using my data to target me).": "Fear of data\nprivacy violations",
    "Concern that the product/image is fake.": "Concern product/\nimage is fake",
    "Belief that the brand is cutting costs/lazy by using AI.": "Brand cutting costs\n/ lazy with AI",
    "I have no concerns, not feeling skeptical.": "No concerns /\nnot skeptical",
}
SKEPTICISM_ORDER = list(SKEPTICISM_LABELS.keys())

TRUST_FACTOR_LABELS = {
    "The reputation of the brand posting the advertisement": "Brand\nreputation",
    "Clear disclosure that the advertisement is developed with AI": "Clear AI\ndisclosure",
    "The presence of real human testimonials or influencers": "Real human\ntestimonials",
    "High-quality visuals and production value": "High-quality\nvisuals",
    "Recommendations from friends or peers": "Friends /\npeers",
}
TRUST_FACTOR_ORDER = list(TRUST_FACTOR_LABELS.keys())


def load_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    if not SOURCE_XLSX.exists():
        raise FileNotFoundError(f"Missing source workbook: {SOURCE_XLSX}")
    df = pd.read_excel(SOURCE_XLSX, sheet_name=SOURCE_SHEET)
    fields = pd.read_excel(SOURCE_XLSX, sheet_name=FIELD_SHEET)
    df = df[df["Gender"].isin(["Male", "Female"])].copy()
    return df, fields


def pct_crosstab(df: pd.DataFrame, col: str, order: Iterable[str] | None = None) -> pd.DataFrame:
    ct = pd.crosstab(df["Gender"], df[col], normalize="index") * 100
    ct = ct.reindex(["Male", "Female"])
    if order is not None:
        cols = [c for c in order if c in ct.columns]
        extra = [c for c in ct.columns if c not in cols]
        ct = ct[cols + extra]
    return ct.round(1)


def count_crosstab(df: pd.DataFrame, col: str, order: Iterable[str] | None = None) -> pd.DataFrame:
    ct = pd.crosstab(df["Gender"], df[col])
    ct = ct.reindex(["Male", "Female"])
    if order is not None:
        cols = [c for c in order if c in ct.columns]
        extra = [c for c in ct.columns if c not in cols]
        ct = ct[cols + extra]
    return ct


def topbox_pct(df: pd.DataFrame, cols: list[str], threshold: int = 4) -> pd.DataFrame:
    """% within gender who rated >= threshold (default 4–5)."""
    rows = {}
    for g in ["Male", "Female"]:
        sub = df[df["Gender"] == g]
        n = len(sub)
        rows[g] = {c: (sub[c] >= threshold).sum() / n * 100 if n else np.nan for c in cols}
    out = pd.DataFrame(rows).T.reindex(["Male", "Female"]).round(1)
    return out


def topbox_counts(df: pd.DataFrame, cols: list[str], threshold: int = 4) -> pd.DataFrame:
    """Headcounts within gender who rated >= threshold (default 4–5)."""
    rows = {}
    for g in ["Male", "Female"]:
        sub = df[df["Gender"] == g]
        rows[g] = {c: int((sub[c] >= threshold).sum()) for c in cols}
    return pd.DataFrame(rows).T.reindex(["Male", "Female"])


def likert_full_pct(df: pd.DataFrame, cols: list[str]) -> dict[str, pd.DataFrame]:
    """
    % within gender at each Likert score 1–5.
    Returns {gender: DataFrame(index=items, columns=1..5)}.
    """
    out: dict[str, pd.DataFrame] = {}
    for g in ["Male", "Female"]:
        sub = df[df["Gender"] == g]
        n = len(sub)
        rows: dict[str, dict[int, float]] = {}
        for c in cols:
            rows[c] = {
                score: (sub[c] == score).sum() / n * 100 if n else np.nan
                for score in LIKERT_SCORES
            }
        out[g] = pd.DataFrame(rows).T[LIKERT_SCORES].round(1)
    return out


def likert_full_counts(df: pd.DataFrame, cols: list[str]) -> dict[str, pd.DataFrame]:
    """
    Headcounts within gender at each Likert score 1–5.
    Returns {gender: DataFrame(index=items, columns=1..5)}.
    """
    out: dict[str, pd.DataFrame] = {}
    for g in ["Male", "Female"]:
        sub = df[df["Gender"] == g]
        rows: dict[str, dict[int, int]] = {}
        for c in cols:
            rows[c] = {score: int((sub[c] == score).sum()) for score in LIKERT_SCORES}
        out[g] = pd.DataFrame(rows).T[LIKERT_SCORES]
    return out


def multiselect_pct(
    df: pd.DataFrame,
    col: str,
    options: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    % and counts within gender who selected each option.
    Options are matched as substrings (handles comma-joined multi-select).
    """
    pct_rows: dict[str, dict[str, float]] = {}
    count_rows: dict[str, dict[str, int]] = {}
    for g in ["Male", "Female"]:
        sub = df[df["Gender"] == g]
        n = len(sub)
        texts = sub[col].fillna("").astype(str)
        pct_rows[g] = {}
        count_rows[g] = {}
        for opt in options:
            hits = int(texts.str.contains(opt, regex=False).sum())
            count_rows[g][opt] = hits
            pct_rows[g][opt] = hits / n * 100 if n else np.nan
    pct = pd.DataFrame(pct_rows).T.reindex(["Male", "Female"]).round(1)
    counts = pd.DataFrame(count_rows).T.reindex(["Male", "Female"])
    return pct, counts


def style_header(ws) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def write_df(ws, df: pd.DataFrame, index: bool = True) -> None:
    for r in dataframe_to_rows(df, index=index, header=True):
        ws.append(r)
    style_header(ws)
    ws.freeze_panes = "A2"


def styled_grouped_bars(
    values: pd.DataFrame,
    categories: list[str],
    title: str,
    subtitle: str,
    outfile: Path,
    ylabel: str = "% of Respondents",
) -> None:
    """
    Presentation grouped bars: Female (coral) then Male (blue), value labels,
    title + subtitle. `values` index = Male/Female, columns = category keys.
    `categories` are display labels in column order.
    """
    plot = values.reindex(PLOT_GENDERS)
    cols = list(plot.columns)
    if len(categories) != len(cols):
        raise ValueError(
            f"categories ({len(categories)}) must match columns ({len(cols)}): {cols}"
        )
    x = np.arange(len(cols))
    width = 0.28
    gap = 0.10
    fig_w = 12.5 if len(cols) >= 5 else 11
    fig, ax = plt.subplots(figsize=(fig_w, 6.2))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    height_by_gender: dict[str, list[float]] = {}
    for gender in PLOT_GENDERS:
        color = COLOR_FEMALE if gender == "Female" else COLOR_MALE
        heights = [float(plot.loc[gender, c]) if pd.notna(plot.loc[gender, c]) else 0.0 for c in cols]
        height_by_gender[gender] = heights
        offset = -(width / 2 + gap / 2) if gender == "Female" else (width / 2 + gap / 2)
        b = ax.bar(x + offset, heights, width, label=gender, color=color, zorder=3)
        for i, (rect, h) in enumerate(zip(b, heights)):
            other = "Male" if gender == "Female" else "Female"
            other_h = height_by_gender.get(other, [None] * len(heights))[i]
            # Stagger Male label up when heights are close so % texts do not stack
            stagger = 10 if gender == "Male" and other_h is not None and abs(h - other_h) < 12 else 0
            ax.annotate(
                f"{h:.1f}%",
                xy=(rect.get_x() + rect.get_width() / 2, h),
                xytext=(0, 12 + stagger),
                textcoords="offset points",
                ha="center",
                va="bottom",
                fontsize=14,
                color="#333333",
                clip_on=False,
            )

    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=10)
    ax.set_ylabel(ylabel, fontsize=11)
    ymax = min(100, max(plot.max().max() + 18, plot.max().max() * 1.4, 10))
    ax.set_ylim(0, ymax)
    ax.yaxis.grid(True, linestyle="-", linewidth=0.6, color="#D0D0D0", zorder=0)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#888888")
    ax.spines["bottom"].set_color("#888888")

    ax.set_title(title, fontsize=14, fontweight="bold", pad=28)
    ax.text(
        0.5,
        1.02,
        subtitle,
        transform=ax.transAxes,
        ha="center",
        va="bottom",
        fontsize=10,
        color="#555555",
        style="italic",
    )
    ax.legend(frameon=False, loc="upper left", bbox_to_anchor=(1.02, 1))
    fig.tight_layout(rect=[0, 0, 0.88, 0.96])
    outfile.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(outfile, dpi=160, facecolor="white", bbox_inches="tight")
    plt.close(fig)


def styled_stacked_likert_single(
    full_pct: pd.DataFrame,
    item_labels: list[str],
    title: str,
    subtitle: str,
    outfile: Path,
    legend_map: dict[int, str],
    accent_color: str | None = None,
) -> None:
    """100% stacked horizontal bars for one gender (Likert 1–5)."""
    items = list(full_pct.index)
    if len(item_labels) != len(items):
        raise ValueError(
            f"item_labels ({len(item_labels)}) must match rows ({len(items)})"
        )
    y = np.arange(len(items))
    fig, ax = plt.subplots(figsize=(11.5, 5.8))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    left = np.zeros(len(items))
    for score in LIKERT_SCORES:
        widths = [
            float(full_pct.loc[item, score]) if pd.notna(full_pct.loc[item, score]) else 0.0
            for item in items
        ]
        bars = ax.barh(
            y,
            widths,
            left=left,
            height=0.62,
            color=LIKERT_COLORS[score],
            label=legend_map[score],
            zorder=3,
        )
        for rect, w, L in zip(bars, widths, left):
            if w >= 8:
                ax.text(
                    L + w / 2,
                    rect.get_y() + rect.get_height() / 2,
                    f"{w:.0f}%",
                    ha="center",
                    va="center",
                    fontsize=9,
                    color="white" if score in (1, 5) else "#222222",
                )
        left = left + np.array(widths)

    ax.set_yticks(y)
    ax.set_yticklabels(item_labels, fontsize=10)
    ax.set_xlim(0, 100)
    ax.set_xlabel("% of Respondents", fontsize=11)
    ax.xaxis.grid(True, linestyle="-", linewidth=0.6, color="#D0D0D0", zorder=0)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#888888")
    ax.spines["bottom"].set_color("#888888")
    ax.invert_yaxis()
    title_color = accent_color if accent_color else "#000000"
    ax.set_title(title, fontsize=14, fontweight="bold", pad=28, color=title_color)
    ax.text(
        0.5,
        1.02,
        subtitle,
        transform=ax.transAxes,
        ha="center",
        va="bottom",
        fontsize=10,
        color="#555555",
        style="italic",
    )
    ax.legend(frameon=False, loc="upper left", bbox_to_anchor=(1.02, 1), fontsize=9)
    fig.tight_layout(rect=[0, 0, 0.82, 0.96])
    outfile.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(outfile, dpi=160, facecolor="white", bbox_inches="tight")
    plt.close(fig)


def styled_stacked_likert_by_gender(
    full_by_gender: dict[str, pd.DataFrame],
    item_labels: list[str],
    title: str,
    subtitle: str,
    outfile: Path,
    legend_map: dict[int, str],
) -> None:
    """
    100% stacked horizontal bars: for each item, Female then Male bars.
    Segment colors = Likert 1–5; y labels show item + gender.
    """
    items = list(full_by_gender["Female"].index)
    if len(item_labels) != len(items):
        raise ValueError(
            f"item_labels ({len(item_labels)}) must match rows ({len(items)})"
        )

    y_labels: list[str] = []
    row_gender: list[str] = []
    row_item: list[str] = []
    for item, lab in zip(items, item_labels):
        flat = lab.replace("\n", " ")
        for gender in PLOT_GENDERS:
            y_labels.append(f"{flat} — {gender}")
            row_gender.append(gender)
            row_item.append(item)

    y = np.arange(len(y_labels))
    fig_h = max(6.5, 0.55 * len(y_labels) + 2.2)
    fig, ax = plt.subplots(figsize=(12, fig_h))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    left = np.zeros(len(y_labels))
    for score in LIKERT_SCORES:
        widths = []
        for g, item in zip(row_gender, row_item):
            val = full_by_gender[g].loc[item, score]
            widths.append(float(val) if pd.notna(val) else 0.0)
        bars = ax.barh(
            y,
            widths,
            left=left,
            height=0.68,
            color=LIKERT_COLORS[score],
            label=legend_map[score],
            zorder=3,
        )
        for rect, w, L in zip(bars, widths, left):
            if w >= 9:
                ax.text(
                    L + w / 2,
                    rect.get_y() + rect.get_height() / 2,
                    f"{w:.0f}%",
                    ha="center",
                    va="center",
                    fontsize=8,
                    color="white" if score in (1, 5) else "#222222",
                )
        left = left + np.array(widths)

    ax.set_yticks(y)
    ax.set_yticklabels(y_labels, fontsize=9)
    ax.set_xlim(0, 100)
    ax.set_xlabel("% of Respondents", fontsize=11)
    ax.xaxis.grid(True, linestyle="-", linewidth=0.6, color="#D0D0D0", zorder=0)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#888888")
    ax.spines["bottom"].set_color("#888888")
    ax.invert_yaxis()
    ax.set_title(title, fontsize=14, fontweight="bold", pad=28)
    ax.text(
        0.5,
        1.01,
        subtitle,
        transform=ax.transAxes,
        ha="center",
        va="bottom",
        fontsize=10,
        color="#555555",
        style="italic",
    )
    ax.legend(frameon=False, loc="upper left", bbox_to_anchor=(1.02, 1), fontsize=9)
    fig.tight_layout(rect=[0, 0, 0.80, 0.96])
    outfile.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(outfile, dpi=160, facecolor="white", bbox_inches="tight")
    plt.close(fig)


def clear_old_charts() -> None:
    if CHART_DIR.exists():
        for p in CHART_DIR.glob("*.png"):
            p.unlink()


def main() -> None:
    df, fields = load_data()
    n_by_gender = df["Gender"].value_counts().reindex(["Male", "Female"])
    n_dup = (
        df.groupby("Gender")["Is_Duplicate"]
        .apply(lambda s: (s == "Yes").sum())
        .reindex(["Male", "Female"])
    )

    # --- Core tables ---
    awareness_pct = pct_crosstab(df, "Seen_AI_Labeled_Ads", AWARE_ORDER)
    awareness_n = count_crosstab(df, "Seen_AI_Labeled_Ads", AWARE_ORDER)

    credibility_pct = pct_crosstab(df, "Brand_Credibility_Reaction", CRED_ORDER)
    credibility_n = count_crosstab(df, "Brand_Credibility_Reaction", CRED_ORDER)
    credibility_plot = credibility_pct.rename(columns=CRED_LABELS)

    trust_cols = [c for c, _ in LIKERT_TRUST]
    trust_topbox = topbox_pct(df, trust_cols)
    trust_topbox_n = topbox_counts(df, trust_cols)
    trust_full = likert_full_pct(df, trust_cols)
    trust_full_n = likert_full_counts(df, trust_cols)
    trust_labels = [lab for _, lab in LIKERT_TRUST]

    skepticism_pct, skepticism_n = multiselect_pct(df, "Skepticism_Concerns", SKEPTICISM_ORDER)
    skepticism_plot = skepticism_pct.rename(columns=SKEPTICISM_LABELS)

    engage_cols = [c for c, _ in LIKERT_ENGAGE]
    engage_topbox = topbox_pct(df, engage_cols)
    engage_topbox_n = topbox_counts(df, engage_cols)
    engage_full = likert_full_pct(df, engage_cols)
    engage_full_n = likert_full_counts(df, engage_cols)
    engage_labels = [lab for _, lab in LIKERT_ENGAGE]

    pref_pct = pct_crosstab(df, "Preferred_Ad_Type", PREF_ORDER)
    pref_n = count_crosstab(df, "Preferred_Ad_Type", PREF_ORDER)
    pref_plot = pref_pct.rename(columns=PREF_LABELS)

    purchase_pct = pct_crosstab(df, "Purchase_Likelihood_AI", PURCHASE_ORDER)
    purchase_n = count_crosstab(df, "Purchase_Likelihood_AI", PURCHASE_ORDER)
    purchase_plot = purchase_pct.rename(columns=PURCHASE_LABELS)

    trust_factor_pct = pct_crosstab(df, "Top_Trust_Factor", TRUST_FACTOR_ORDER)
    trust_factor_n = count_crosstab(df, "Top_Trust_Factor", TRUST_FACTOR_ORDER)

    # Collapsed helpful summaries for outline / console
    neg_cred = {
        g: float(credibility_pct.loc[g, "It makes me see the brand more negatively."])
        for g in ["Male", "Female"]
    }
    less_buy = {
        g: float(
            purchase_pct.loc[g, "Much less likely to purchase"]
            + purchase_pct.loc[g, "Somewhat less likely to purchase"]
        )
        for g in ["Male", "Female"]
    }
    prefer_human = {
        g: float(
            pref_pct.loc[g, "Definitely the Human-Created ad"]
            + pref_pct.loc[g, "Somewhat prefer the Human-Created ad"]
        )
        for g in ["Male", "Female"]
    }

    # --- Charts ---
    clear_old_charts()
    CHART_DIR.mkdir(parents=True, exist_ok=True)
    chart_specs: list[tuple[str, str, str]] = []

    p = CHART_DIR / "00_awareness.png"
    aware_cols = [c for c in AWARE_ORDER if c in awareness_pct.columns]
    styled_awareness_pies(
        awareness_pct[aware_cols],
        aware_cols,
        "Awareness of AI-Labeled Ads",
        "% who said Yes vs Not sure (valid sample)",
        p,
    )
    chart_specs.append(("Awareness (pie)", "Awareness", p.name))

    p = CHART_DIR / "01_trust_credibility_topbox.png"
    styled_grouped_bars(
        trust_topbox,
        trust_labels,
        "Trust & Credibility Toward AI-Labeled Ads",
        "% who Agree/Strongly Agree (rating 4-5 out of 5)",
        p,
    )
    chart_specs.append(("Trust & credibility (Agree 4–5)", "Trust_Credibility", p.name))

    p = CHART_DIR / "01b_trust_credibility_full_female.png"
    styled_stacked_likert_single(
        trust_full["Female"],
        trust_labels,
        "Trust & Credibility Toward AI-Labeled Ads (Full Scale) — Female",
        "% at each rating 1–5 (Strongly disagree → Strongly agree)",
        p,
        AGREE_LEGEND,
        accent_color=COLOR_FEMALE,
    )
    chart_specs.append(("Trust & credibility full 1–5 (Female)", "Trust_Credibility", p.name))

    p = CHART_DIR / "01c_trust_credibility_full_male.png"
    styled_stacked_likert_single(
        trust_full["Male"],
        trust_labels,
        "Trust & Credibility Toward AI-Labeled Ads (Full Scale) — Male",
        "% at each rating 1–5 (Strongly disagree → Strongly agree)",
        p,
        AGREE_LEGEND,
        accent_color=COLOR_MALE,
    )
    chart_specs.append(("Trust & credibility full 1–5 (Male)", "Trust_Credibility", p.name))

    p = CHART_DIR / "02_brand_reaction.png"
    styled_grouped_bars(
        credibility_plot,
        list(CRED_LABELS.values()),
        "Immediate Brand Reaction to an AI Label",
        "% of respondents choosing each reaction",
        p,
    )
    chart_specs.append(("Brand reaction", "Trust_Credibility", p.name))

    p = CHART_DIR / "03_skepticism_concerns.png"
    styled_grouped_bars(
        skepticism_plot,
        list(SKEPTICISM_LABELS.values()),
        "Concerns Driving Skepticism Toward AI-Labeled Ads",
        "% selecting each concern (multi-select; within gender)",
        p,
    )
    chart_specs.append(("Skepticism concerns", "Skepticism", p.name))

    p = CHART_DIR / "04_engagement_topbox.png"
    styled_grouped_bars(
        engage_topbox,
        engage_labels,
        "Engagement With AI-Labeled Ads",
        "% who rated Likely/Very likely (4-5 out of 5)",
        p,
    )
    chart_specs.append(("Engagement (Likely 4–5)", "Engagement", p.name))

    p = CHART_DIR / "04b_engagement_full_female.png"
    styled_stacked_likert_single(
        engage_full["Female"],
        engage_labels,
        "Engagement With AI-Labeled Ads (Full Scale) — Female",
        "% at each rating 1–5 (Very unlikely → Very likely)",
        p,
        LIKELY_LEGEND,
        accent_color=COLOR_FEMALE,
    )
    chart_specs.append(("Engagement full 1–5 (Female)", "Engagement", p.name))

    p = CHART_DIR / "04c_engagement_full_male.png"
    styled_stacked_likert_single(
        engage_full["Male"],
        engage_labels,
        "Engagement With AI-Labeled Ads (Full Scale) — Male",
        "% at each rating 1–5 (Very unlikely → Very likely)",
        p,
        LIKELY_LEGEND,
        accent_color=COLOR_MALE,
    )
    chart_specs.append(("Engagement full 1–5 (Male)", "Engagement", p.name))

    p = CHART_DIR / "05_preferred_ad_type.png"
    styled_grouped_bars(
        pref_plot,
        list(PREF_LABELS.values()),
        "Which Ad Would Students Engage With?",
        "% choosing Human-created vs AI-made (same product)",
        p,
    )
    chart_specs.append(("Preferred ad type", "Engagement", p.name))

    p = CHART_DIR / "06_purchase_likelihood.png"
    styled_grouped_bars(
        purchase_plot,
        list(PURCHASE_LABELS.values()),
        "Purchase Likelihood if the Ad Used AI",
        "% of respondents by purchase effect",
        p,
    )
    chart_specs.append(("Purchase likelihood", "Purchase", p.name))

    p = CHART_DIR / "07_top_trust_factor.png"
    factor_cols = [c for c in TRUST_FACTOR_ORDER if c in trust_factor_pct.columns]
    factor_plot = trust_factor_pct[factor_cols].rename(columns=TRUST_FACTOR_LABELS)
    styled_grouped_bars(
        factor_plot,
        [TRUST_FACTOR_LABELS[c] for c in factor_cols],
        "Most Important Factor for Trusting an Online Ad",
        "% choosing each factor (select only one)",
        p,
    )
    chart_specs.append(("Top trust factor", "Top_Trust_Factor", p.name))

    # --- Workbook ---
    wb = Workbook()

    ws = wb.active
    ws.title = "Column_Map"
    cmap = fields.rename(columns={"Tableau_Field": "Field", "Original_Question": "Question"})
    write_df(ws, cmap, index=False)

    ws = wb.create_sheet("Sample_Profile")
    ws.append(["Source", SOURCE_XLSX.name])
    ws.append(["Sheet", SOURCE_SHEET])
    ws.append(["n", int(len(df))])
    ws.append(["Male", int(n_by_gender["Male"])])
    ws.append(["Female", int(n_by_gender["Female"])])
    ws.append(["Male duplicates (Is_Duplicate=Yes)", int(n_dup["Male"])])
    ws.append(["Female duplicates (Is_Duplicate=Yes)", int(n_dup["Female"])])
    style_header(ws)

    ws = wb.create_sheet("Awareness")
    ws.append(["Counts"])
    for r in dataframe_to_rows(awareness_n.reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])
    ws.append(["% within gender"])
    for r in dataframe_to_rows(awareness_pct.reset_index(), index=False, header=True):
        ws.append(r)
    style_header(ws)

    ws = wb.create_sheet("Trust_Credibility")
    ws.append(["Brand reaction %"])
    for r in dataframe_to_rows(credibility_pct.reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])
    ws.append(["Brand reaction counts"])
    for r in dataframe_to_rows(credibility_n.reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])
    ws.append(["Likert top-box % (rated 4–5)"])
    for r in dataframe_to_rows(trust_topbox.reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])
    ws.append(["Likert top-box counts (rated 4–5)"])
    for r in dataframe_to_rows(trust_topbox_n.reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])
    ws.append(["Likert full scale % (ratings 1–5)"])
    for gender in PLOT_GENDERS:
        ws.append([gender])
        full = trust_full[gender].copy()
        full.columns = [f"{s} ({AGREE_LEGEND[s].split(' ', 1)[1]})" for s in LIKERT_SCORES]
        full = full.reset_index().rename(columns={"index": "Item"})
        for r in dataframe_to_rows(full, index=False, header=True):
            ws.append(r)
        ws.append([])
    ws.append(["Likert full scale counts (ratings 1–5)"])
    for gender in PLOT_GENDERS:
        ws.append([gender])
        full = trust_full_n[gender].copy()
        full.columns = [f"{s} ({AGREE_LEGEND[s].split(' ', 1)[1]})" for s in LIKERT_SCORES]
        full = full.reset_index().rename(columns={"index": "Item"})
        for r in dataframe_to_rows(full, index=False, header=True):
            ws.append(r)
        ws.append([])
    style_header(ws)

    ws = wb.create_sheet("Skepticism")
    ws.append(["Skepticism concern counts (multi-select)"])
    for r in dataframe_to_rows(skepticism_n.reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])
    ws.append(["Skepticism concern % within gender (multi-select)"])
    for r in dataframe_to_rows(skepticism_pct.reset_index(), index=False, header=True):
        ws.append(r)
    style_header(ws)

    ws = wb.create_sheet("Engagement")
    ws.append(["Engagement top-box % (rated 4–5)"])
    for r in dataframe_to_rows(engage_topbox.reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])
    ws.append(["Engagement top-box counts (rated 4–5)"])
    for r in dataframe_to_rows(engage_topbox_n.reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])
    ws.append(["Engagement full scale % (ratings 1–5)"])
    for gender in PLOT_GENDERS:
        ws.append([gender])
        full = engage_full[gender].copy()
        full.columns = [f"{s} ({LIKELY_LEGEND[s].split(' ', 1)[1]})" for s in LIKERT_SCORES]
        full = full.reset_index().rename(columns={"index": "Item"})
        for r in dataframe_to_rows(full, index=False, header=True):
            ws.append(r)
        ws.append([])
    ws.append(["Engagement full scale counts (ratings 1–5)"])
    for gender in PLOT_GENDERS:
        ws.append([gender])
        full = engage_full_n[gender].copy()
        full.columns = [f"{s} ({LIKELY_LEGEND[s].split(' ', 1)[1]})" for s in LIKERT_SCORES]
        full = full.reset_index().rename(columns={"index": "Item"})
        for r in dataframe_to_rows(full, index=False, header=True):
            ws.append(r)
        ws.append([])
    ws.append(["Preferred ad type %"])
    for r in dataframe_to_rows(pref_pct.reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])
    ws.append(["Preferred ad type counts"])
    for r in dataframe_to_rows(pref_n.reset_index(), index=False, header=True):
        ws.append(r)
    style_header(ws)

    ws = wb.create_sheet("Purchase")
    ws.append(["Purchase likelihood %"])
    for r in dataframe_to_rows(purchase_pct.reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])
    ws.append(["Purchase likelihood counts"])
    for r in dataframe_to_rows(purchase_n.reset_index(), index=False, header=True):
        ws.append(r)
    style_header(ws)

    ws = wb.create_sheet("Top_Trust_Factor")
    ws.append(["Top trust factor %"])
    for r in dataframe_to_rows(trust_factor_pct.reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])
    ws.append(["Top trust factor counts"])
    for r in dataframe_to_rows(trust_factor_n.reset_index(), index=False, header=True):
        ws.append(r)
    style_header(ws)

    ws = wb.create_sheet("Chart_Index")
    ws.append(["Chart", "Sheet", "PNG"])
    for title, sheet, png in chart_specs:
        ws.append([title, sheet, f"charts/{png}"])
    style_header(ws)

    ws = wb.create_sheet("Limitations")
    for note in [
        "Convenience sample of DLSU undergraduates; descriptive % only.",
        "Balanced 92/92 includes duplicated rows (see Sample_Profile).",
        "Valid sample excludes people who said they never saw AI-labeled ads.",
        "Likert charts include top-box % (rated 4–5) and full 1–5 distribution charts/tables.",
        "Skepticism_Concerns is multi-select; % can sum over 100% within gender.",
        "Original survey Excels were not modified.",
    ]:
        ws.append([note])

    wb.save(OUT_XLSX)

    print("=== MADSRCH gender findings report ===")
    print(f"Source: {SOURCE_XLSX.name} | n={len(df)} (92 Male / 92 Female)")
    print(f"Wrote: {OUT_XLSX.name}")
    print(f"Charts: {len(chart_specs)} files in charts/")
    for _, _, png in chart_specs:
        print(f"  - {png}")
    print("\nTrust top-box % (4–5):")
    print(trust_topbox.to_string())
    print("\nTrust full 1–5 % (Female):")
    print(trust_full["Female"].to_string())
    print("\nEngagement top-box % (4–5):")
    print(engage_topbox.to_string())
    print("\nEngagement full 1–5 % (Female):")
    print(engage_full["Female"].to_string())
    print("\nSkepticism concern %:")
    print(skepticism_pct.to_string())
    print("\nTop trust factor %:")
    print(trust_factor_pct.to_string())
    print("\n% more negatively / prefer human / less likely to buy:")
    print("negative reaction", {k: round(v, 1) for k, v in neg_cred.items()})
    print("prefer human", {k: round(v, 1) for k, v in prefer_human.items()})
    print("less likely buy", {k: round(v, 1) for k, v in less_buy.items()})
    print("\nDone.")


if __name__ == "__main__":
    main()
