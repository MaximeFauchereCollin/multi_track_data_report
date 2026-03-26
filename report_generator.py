"""
What this script does
---------------------
1) Loads an Excel workbook where *each sheet represents a Track*.
2) Cleans & standardizes the merged dataset.
3) Computes summary statistics by Track, Cohort, and Income status.
4) Exports:
   - outputs/cleaned_dataset.csv
   - outputs/summary_report.xlsx
   - outputs/figures/*.png
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


# =============================================================================
# CONFIG
# =============================================================================

# Treat these string values as missing data.
SPECIAL_NULLS: list[str] = ["", " ", "NA", "N/A", "n/a", "NaN", "-", "Waived", "None", "null"]

# Mapping for boolean-like values found in spreadsheets.
BOOL_MAP: dict[str, bool] = {
    "TRUE": True,
    "FALSE": False,
    "Y": True,
    "N": False,
    "YES": True,
    "NO": False,
    "1": True,
    "0": False,
    "1.0": True,
    "0.0": False,
    "T": True,
    "F": False,
}

# Columns used throughout the pipeline.
STRING_COLS: list[str] = ["FirstName", "LastName", "Class", "Cohort", "Track"]
SCORE_COLS: list[str] = ["Math", "English", "Science", "History", "Attendance (%)", "ProjectScore"]
BOOLEAN_COLS: list[str] = ["IncomeStudent", "Passed (Y/N)"]

# Columns required to produce the stats/plots without special-case handling.
REQUIRED_FOR_REPORT: list[str] = [
    "Track",
    "Cohort",
    "Math",
    "English",
    "Science",
    "History",
    "Attendance (%)",
    "ProjectScore",
    "Passed (Y/N)",
    "IncomeStudent",
]


# =============================================================================
# OUTPUT PATHS
# =============================================================================

BASE_DIR = Path(__file__).resolve().parent
OUTDIR = BASE_DIR / "outputs"
FIGDIR = OUTDIR / "figures"
OUTDIR.mkdir(parents=True, exist_ok=True)
FIGDIR.mkdir(parents=True, exist_ok=True)


# =============================================================================
# 1) DATA LOADING
# =============================================================================

def load_all_sheets(xlsx_path: str | Path) -> pd.DataFrame:
    """
    Load all sheets from an Excel workbook and merge them into one DataFrame.

    Each sheet is read into a DataFrame, then we add a "Track" column equal to the sheet name.
    Finally, all sheets are concatenated into a single DataFrame.

    Returns
    -------
    pd.DataFrame
        Concatenated dataset containing all sheets with an added "Track" column.
    """
    xlsx_path = Path(xlsx_path)
    xl = pd.ExcelFile(xlsx_path, engine="openpyxl")

    frames: list[pd.DataFrame] = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(xlsx_path, sheet_name=sheet, engine="openpyxl")
        df["Track"] = sheet  # sheet name becomes the Track label
        frames.append(df)

    return pd.concat(frames, ignore_index=True)


# =============================================================================
# 2) CLEANING
# =============================================================================

def _coerce_bool_series(s: pd.Series) -> pd.Series:
    """
    Coerce a boolean-like column into pandas BooleanDtype.

    Handles:
    - actual bool dtype
    - numeric 0/1 (incl. floats 0.0/1.0)
    - strings ("Y/N", "TRUE/FALSE", etc.)

    Returns a nullable Boolean series (values in {True, False, <NA>}).
    """
    # Already boolean -> standardize to pandas nullable boolean dtype
    if pd.api.types.is_bool_dtype(s):
        return s.astype("boolean")

    # Numeric case: map 0 -> False, 1 -> True, other -> NA
    if pd.api.types.is_numeric_dtype(s):
        out = s.copy()
        out = out.where(~out.isna(), other=pd.NA)
        out = out.map(lambda x: True if x == 1 else (False if x == 0 else pd.NA))
        return out.astype("boolean")

    # String/mixed: normalize and map with BOOL_MAP
    ss = s.astype("string").str.strip().str.upper()
    out = ss.map(BOOL_MAP)
    return out.astype("boolean")


def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean the dataset.

    Steps
    -----
    - Replace SPECIAL_NULLS by pandas NA
    - Normalize string columns (trim, title-case)
    - Validate formats for "Class" and "Cohort"
    - Coerce StudentID and Term, validate allowed formats/ranges
    - Coerce score columns to numeric, validate range [0, 100]
    - Impute missing numeric scores with Track mean (design choice)
    - Coerce boolean-like columns, impute missing with Track mode (design choice)
    - Drop rows missing core columns required by plots/statistics (design choice)

    Notes
    -----
    This function is intentionally opinionated to produce a dataset that is “safe” for the
    downstream stats/plots without adding many special cases.
    """
    df = df.copy()

    # 1) Standardize missing values early so downstream coercions behave consistently
    df = df.replace(SPECIAL_NULLS, pd.NA)

    # 2) Normalize string columns (only those that exist)
    for c in STRING_COLS:
        if c in df.columns:
            df[c] = df[c].astype("string").str.strip().str.title()

    # 3) Validate Class format: e.g. "27A" (2 digits + 1 letter)
    if "Class" in df.columns:
        df.loc[~df["Class"].str.fullmatch(r"\d{2}[A-Za-z]", na=False), "Class"] = pd.NA

    # 4) Validate Cohort format: e.g. "27-28"
    if "Cohort" in df.columns:
        df.loc[~df["Cohort"].str.fullmatch(r"\d{2}-\d{2}", na=False), "Cohort"] = pd.NA

    # 5) StudentID: numeric + enforce 4-digit IDs (keeps IDs consistent)
    if "StudentID" in df.columns:
        df["StudentID"] = pd.to_numeric(df["StudentID"], errors="coerce").astype("Int64")
        df.loc[~df["StudentID"].astype("string").str.fullmatch(r"\d{4}", na=False), "StudentID"] = pd.NA

    # 6) Term: only allow 1 or 2
    if "Term" in df.columns:
        df["Term"] = pd.to_numeric(df["Term"], errors="coerce").astype("Int64")
        df.loc[~df["Term"].isin([1, 2]), "Term"] = pd.NA

    # 7) Scores: coerce, validate range, then impute NA with Track mean
    for c in SCORE_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
            df.loc[(df[c] < 0) | (df[c] > 100), c] = pd.NA

            # Design choice: impute missing numeric scores with mean within the same Track
            # -> preserves per-Track differences and keeps report tables dense
            if "Track" in df.columns:
                grp_mean = df.groupby("Track", dropna=False)[c].transform("mean")
                df[c] = df[c].fillna(grp_mean.round(1))

    # 8) Boolean-like columns: coerce then impute NA with Track mode
    for c in BOOLEAN_COLS:
        if c in df.columns:
            df[c] = _coerce_bool_series(df[c])

            # Design choice: fill missing booleans with the most frequent value per Track
            if "Track" in df.columns:
                def _mode_or_na(x: pd.Series):
                    m = x.mode(dropna=True)
                    return m.iloc[0] if not m.empty else pd.NA

                grp_mode = df.groupby("Track", dropna=False)[c].transform(_mode_or_na)
                df[c] = df[c].fillna(grp_mode)

    # 9) Drop rows missing core columns required by the report
    # Design choice: prefer complete records (stable plots/stats) over maximum row retention.
    subset = [c for c in REQUIRED_FOR_REPORT if c in df.columns]
    df = df.dropna(subset=subset)

    return df


def drop_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove duplicate student-term-track entries (if those columns exist).

    Keeps the first occurrence.
    """
    df = df.copy()
    if {"StudentID", "Term", "Track"}.issubset(df.columns):
        df = df.drop_duplicates(subset=["StudentID", "Term", "Track"], keep="first")
    return df


# =============================================================================
# 3) STATISTICS (OOP)
# =============================================================================

class BaseGroupStatistics:
    """
    Compute statistics and generate plots aggregated by a grouping column.

    Parameters
    ----------
    df : pd.DataFrame
        Clean dataset (assumed to contain required columns).
    group_col : str
        Column name used for grouping (e.g., "Track", "Cohort", "IncomeStudent").
    """

    def __init__(self, df: pd.DataFrame, group_col: str):
        self.df = df
        self.group_col = group_col

    # ---------- TABLES / STATS ----------

    def nb_students(self) -> pd.DataFrame:
        """Count rows per group (interpreted as number of students/records)."""
        return (
            self.df[self.group_col]
            .value_counts(dropna=False)
            .rename_axis(self.group_col)
            .reset_index(name="Nb Students")
        )

    def avg_scores(self) -> pd.DataFrame:
        """Average of subject scores by group."""
        return (
            self.df.groupby(self.group_col, dropna=False)[["Math", "English", "Science", "History"]]
            .mean()
            .rename_axis(self.group_col)
            .reset_index()
        )

    def avg_attendance(self) -> pd.DataFrame:
        """Average attendance by group."""
        return (
            self.df.groupby(self.group_col, dropna=False)["Attendance (%)"]
            .mean()
            .rename_axis(self.group_col)
            .reset_index(name="Avg Attendance (%)")
        )

    def avg_project_scores(self) -> pd.DataFrame:
        """Average project score by group."""
        return (
            self.df.groupby(self.group_col, dropna=False)["ProjectScore"]
            .mean()
            .rename_axis(self.group_col)
            .reset_index(name="Avg Project Score")
        )

    def pass_rate(self) -> pd.DataFrame:
        """
        Pass rate by group.

        Assumes "Passed (Y/N)" is boolean (True/False). Mean of a boolean yields proportion.
        """
        return (
            self.df.groupby(self.group_col, dropna=False)["Passed (Y/N)"]
            .mean()
            .mul(100)
            .rename_axis(self.group_col)
            .reset_index(name="Pass Rate (%)")
        )

    def corr_attendance_project(self) -> pd.DataFrame:
        """
        Correlation between attendance and project score by group (Pearson r in [-1, 1]).
        """
        rows: list[dict] = []
        for key, g in self.df.groupby(self.group_col, dropna=False):
            corr = g["Attendance (%)"].corr(g["ProjectScore"])
            rows.append({self.group_col: key, "Correlation (r)": corr})
        return pd.DataFrame(rows)

    # ---------- PLOTS ----------

    def plot_pass_rate(self, path: Path) -> None:
        """Barplot: pass rate (%) by group."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)

        plt.figure(figsize=(8, 5))
        sns.barplot(
            data=self.pass_rate(),
            x=self.group_col,
            y="Pass Rate (%)",
            hue=self.group_col,      # keep legend consistent; can be removed if you prefer
            palette="pastel",
        )
        plt.title(f"Pass Rate by {self.group_col}")
        plt.xlabel(self.group_col)
        plt.ylabel("Pass Rate (%)")
        plt.ylim(0, 100)
        plt.grid(axis="y", linestyle="--", alpha=0.7)
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()

    def plot_avg_scores(self, path: Path) -> None:
        """Grouped barplot: average subject scores by group."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)

        avg_df = self.avg_scores().melt(
            id_vars=self.group_col,
            value_vars=["Math", "English", "Science", "History"],
            var_name="Subject",
            value_name="Average Score",
        )

        plt.figure(figsize=(10, 5))
        sns.barplot(
            data=avg_df,
            x=self.group_col,
            y="Average Score",
            hue="Subject",
            palette="pastel",
        )
        plt.title(f"Average Scores by {self.group_col} and Subject")
        plt.xlabel(self.group_col)
        plt.ylabel("Average Score")
        plt.ylim(0, 100)
        plt.legend(title="Subject")
        plt.grid(axis="y", linestyle="--", alpha=0.7)
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()

    def plot_history_distribution(self, path: Path) -> None:
        """Histogram: distribution of History grades, split by group."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)

        plt.figure(figsize=(8, 5))
        sns.histplot(
            data=self.df,
            x="History",
            hue=self.group_col,
            kde=True,
            bins=14,
            element="step",
        )
        plt.title(f"Distributions of History grades by {self.group_col}")
        plt.xlabel("History grade")
        plt.ylabel("Number of students")
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()

    def plot_avg_math(self, path: Path) -> None:
        """Barplot: average Math score by group."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)

        plt.figure(figsize=(8, 5))
        sns.barplot(
            data=self.avg_scores()[[self.group_col, "Math"]],
            x=self.group_col,
            y="Math",
            hue=self.group_col,
            palette="pastel",
        )
        plt.title(f"Average Mathematics Scores by {self.group_col}")
        plt.xlabel(self.group_col)
        plt.ylabel("Average Math Score")
        plt.ylim(0, 100)
        plt.grid(axis="y", linestyle="--", alpha=0.7)
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()

    def plot_attendance_vs_project(self, path: Path) -> None:
        """Scatter: Attendance (%) vs ProjectScore, colored by group."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)

        plt.figure(figsize=(8, 5))
        sns.scatterplot(
            data=self.df,
            x="Attendance (%)",
            y="ProjectScore",
            hue=self.group_col,
            palette="pastel",
            alpha=0.7,
            edgecolor="none",
        )
        plt.title(f"Relationship between Attendance and Project Score by {self.group_col}")
        plt.xlabel("Attendance Rate (%)")
        plt.ylabel("Project Score")
        plt.grid(True, linestyle="--", alpha=0.6)
        plt.legend(title=self.group_col, bbox_to_anchor=(1.05, 1), loc="upper left")
        plt.tight_layout()
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()


class TrackStatistics(BaseGroupStatistics):
    """Statistics grouped by Track (sheet name)."""
    def __init__(self, df: pd.DataFrame):
        super().__init__(df, group_col="Track")


class CohortStatistics(BaseGroupStatistics):
    """Statistics grouped by Cohort (e.g., '27-28')."""
    def __init__(self, df: pd.DataFrame):
        super().__init__(df, group_col="Cohort")


class IncomeStatusStatistics(BaseGroupStatistics):
    """Statistics grouped by IncomeStudent (boolean)."""
    def __init__(self, df: pd.DataFrame):
        super().__init__(df, group_col="IncomeStudent")

    # Override: map True/False -> readable labels for the plot legend
    def plot_avg_scores(self, path: Path = FIGDIR / "avg_scores_by_income_status.png") -> None:
        """Average subject scores by income status, with readable labels."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)

        plot_data = self.avg_scores().melt(
            id_vars=self.group_col,
            value_vars=["Math", "English", "Science", "History"],
            var_name="Subject",
            value_name="AverageScore",
        )

        plot_data[self.group_col] = plot_data[self.group_col].map(
            {True: "Income Student", False: "Non-Income Student"}
        )

        plt.figure(figsize=(8, 5))
        sns.barplot(
            data=plot_data,
            x="Subject",
            y="AverageScore",
            hue=self.group_col,
            palette="pastel",
            edgecolor="black",
        )
        plt.title("Average Scores by Subject and Income Status")
        plt.xlabel("Subject")
        plt.ylabel("Average Score")
        plt.ylim(0, 100)
        plt.grid(axis="y", linestyle="--", alpha=0.7)
        plt.legend(title="Income Status")
        plt.tight_layout()
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()


# =============================================================================
# 4) EXPORTS
# =============================================================================

def export_cleaned_data(df: pd.DataFrame, path: Path = OUTDIR / "cleaned_dataset.csv") -> None:
    """Export cleaned dataset to CSV."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False)


def compute_all_stats(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Compute all stats tables and return them as a dict of {sheet_name: dataframe}.

    The returned dict is designed to be directly exported to Excel (one key -> one sheet).
    """

    def _round_numeric(_df: pd.DataFrame, ndigits: int = 1) -> pd.DataFrame:
        """Round numeric columns for nicer report tables (does not affect source df)."""
        out = _df.copy()
        num_cols = out.select_dtypes(include=[np.number, "Float64", "Int64"]).columns
        if len(num_cols):
            out[num_cols] = out[num_cols].astype(float).round(ndigits)
        return out

    ts = TrackStatistics(df)
    cs = CohortStatistics(df)
    ins = IncomeStatusStatistics(df)

    return {
        # Track
        "Track - Counts": ts.nb_students(),
        "Track - Avg Scores": _round_numeric(ts.avg_scores()),
        "Track - Attendance": _round_numeric(ts.avg_attendance()),
        "Track - Project": _round_numeric(ts.avg_project_scores()),
        "Track - Pass Rate": _round_numeric(ts.pass_rate()),
        "Track - Corr (r)": _round_numeric(ts.corr_attendance_project()),
        # Cohort
        "Cohort - Counts": cs.nb_students(),
        "Cohort - Avg Scores": _round_numeric(cs.avg_scores()),
        "Cohort - Attendance": _round_numeric(cs.avg_attendance()),
        "Cohort - Project": _round_numeric(cs.avg_project_scores()),
        "Cohort - Pass Rate": _round_numeric(cs.pass_rate()),
        # Income
        "Income - Counts": ins.nb_students(),
        "Income - Avg Scores": _round_numeric(ins.avg_scores()),
        "Income - Attendance": _round_numeric(ins.avg_attendance()),
        "Income - Project": _round_numeric(ins.avg_project_scores()),
        "Income - Pass Rate": _round_numeric(ins.pass_rate()),
    }


def export_stats_excel(
    stats_dict: dict[str, pd.DataFrame],
    path: Path = OUTDIR / "summary_report.xlsx",
) -> None:
    """Export stats tables to an Excel workbook (one table per sheet)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        for sheet, df_sheet in stats_dict.items():
            # Excel limits sheet names to 31 characters
            df_sheet.to_excel(writer, sheet_name=sheet[:31], index=False)


def export_figures_png(df: pd.DataFrame) -> None:
    """Generate all figures and save them under outputs/figures/."""
    ts = TrackStatistics(df)
    cs = CohortStatistics(df)
    ins = IncomeStatusStatistics(df)

    # Track figures
    ts.plot_pass_rate(FIGDIR / "pass_rate_by_track.png")
    ts.plot_avg_scores(FIGDIR / "avg_scores_by_track.png")
    ts.plot_history_distribution(FIGDIR / "history_grades_by_track.png")
    ts.plot_avg_math(FIGDIR / "avg_math_scores_by_track.png")
    ts.plot_attendance_vs_project(FIGDIR / "attendance_vs_project_by_track.png")

    # Cohort figures
    cs.plot_pass_rate(FIGDIR / "pass_rate_by_cohort.png")
    cs.plot_avg_scores(FIGDIR / "avg_scores_by_cohort.png")

    # Income figures
    ins.plot_avg_scores(FIGDIR / "avg_scores_by_income_status.png")


# =============================================================================
# 5) MAIN
# =============================================================================

def main(xlsx_path: str | Path = "student_grades_2027-2028.xlsx") -> None:
    """
    End-to-end pipeline:
    - load -> clean -> de-duplicate -> compute stats -> export tables & figures
    """
    df = load_all_sheets(xlsx_path)
    df = clean_df(df)
    df = drop_duplicates(df)

    stats = compute_all_stats(df)
    export_cleaned_data(df)
    export_stats_excel(stats)
    export_figures_png(df)


if __name__ == "__main__":
    main()
