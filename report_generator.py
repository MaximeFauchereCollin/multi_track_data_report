import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

FIGDIR = Path("outputs/figures")
FIGDIR.mkdir(parents=True, exist_ok=True)

## 1. Data Loading and Cleaning

def load_all_sheets(xlsx_path):
    # Load all sheets from an Excel workbook and merge them into one DataFrame
    xl = pd.ExcelFile(xlsx_path, engine='openpyxl')
    frames = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(xlsx_path, sheet_name=sheet, engine='openpyxl')
        df["Track"] = sheet  # Keep track of the source sheet
        frames.append(df)
    df_all = pd.concat(frames, ignore_index=True)
    return df_all

SPECIAL_NULLS = ["", " ", "NA", "N/A", "n/a", "NaN", "-", "Waived", "None", "null"]
BOOL_MAP = {
    "TRUE": True, "FALSE": False, "Y": True, "N": False, "YES": True, "NO": False,
    "1": True, "0": False, "1.0": True, "0.0": False
}

def clean_df(df):
    # Replace special null values with pandas NA
    df = df.replace(SPECIAL_NULLS, pd.NA)

    # Normalize string columns (trim, capitalize)
    str_cols = ['FirstName', 'LastName', 'Class', 'Cohort', 'Track']
    for c in str_cols:
        if c in df:
            df[c] = df[c].astype("string").str.strip().str.title()

    # Validate 'Class' and 'Cohort' string formats
    if 'Class' in df:
        df.loc[~df['Class'].str.fullmatch(r'\d{2}[A-Za-z]'), 'Class'] = pd.NA
    if 'Cohort' in df:
        df.loc[~df['Cohort'].str.fullmatch(r'\d{2}-\d{2}'), 'Cohort'] = pd.NA

    # Convert numeric columns safely to integer (nullable type)
    int_cols = ['StudentID', 'Term']
    for c in int_cols:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce").astype("Int64")

    # Validate numeric IDs and terms
    if 'StudentID' in df:
        df.loc[~df['StudentID'].astype(str).str.fullmatch(r'\d{4}'), 'StudentID'] = pd.NA
    if 'Term' in df:
        df.loc[~df['Term'].isin([1, 2]), 'Term'] = pd.NA

    # Clean and impute score columns
    score_cols = ['Math', 'English', 'Science', 'History', 'Attendance (%)', 'ProjectScore']
    for c in score_cols:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
            df.loc[(df[c] < 0) | (df[c] > 100), c] = pd.NA  # Remove invalid scores
            grp_mean = df.groupby('Track')[c].transform('mean')
            df[c] = df[c].fillna(round(grp_mean, 1))  # Impute missing scores with track mean

    # Convert boolean-like columns using predefined map
    boolean_cols = ['IncomeStudent', 'Passed (Y/N)']
    for c in boolean_cols:
        if c in df:
            df[c] = df[c].astype("string").str.upper().str.strip().map(BOOL_MAP).astype("boolean")
            # Fill missing values with most frequent value per track
            grp_mode = df.groupby('Track')[c].transform(lambda x: x.mode(dropna=True).iloc[0] if not x.mode(dropna=True).empty else np.nan)
            df[c] = df[c].fillna(grp_mode)

    # Drop any remaining incomplete rows
    df = df.dropna()
    return df

def drop_duplicates(df):
    # Remove duplicate student-term-track entries (if any)
    if {'StudentID','Term'}.issubset(df.columns):
        df = df.sort_index()
        df = df.drop_duplicates(subset=['StudentID', 'Term', 'Track'], keep='first')
    return df

## 2. Track-Level Summary Statistics

class TrackStatistics:
    def __init__(self, df):
        self.df = df
    def nb_students(self):
        # Count students per track
        return self.df['Track'].value_counts().rename_axis('Track').reset_index(name='Nb Students')
    def avg_scores(self):
        # Compute mean academic scores by track
        return self.df.groupby('Track')[['Math', 'English', 'Science', 'History']].mean().rename_axis('Track').reset_index()
    def avg_attendance(self):
        return self.df.groupby('Track')['Attendance (%)'].mean().rename_axis('Track').reset_index()
    def avg_project_scores(self):
        return self.df.groupby('Track')['ProjectScore'].mean().rename_axis('Track').reset_index()
    def pass_rate(self):
        # Mean boolean -> pass rate percentage
        return (self.df.groupby('Track')['Passed (Y/N)'].mean() * 100).rename_axis('Track').reset_index(name='Pass Rate (%)')
    def corr_attendance_project(self):
        # Compute correlation between attendance and project score per track
        rows = []
        for track, g in self.df.groupby('Track', dropna=False):
            corr = g[['Attendance (%)', 'ProjectScore']].corr().iloc[0, 1]
            rows.append({'Track': track, 'Correlation (%)': 100 * corr})
        return pd.DataFrame(rows)
    def plot_pass_rate(self, path="outputs/figures/pass_rate_by_track.png"):
        # Bar plot of pass rate by track
        plt.figure(figsize=(8, 5))
        sns.barplot(data=self.pass_rate(), x='Track', y='Pass Rate (%)', hue='Track', palette='pastel')
        plt.title("Pass Rate by Track")
        plt.xlabel("Track")
        plt.ylabel("Pass Rate (%)")
        plt.ylim(0, 100)
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()
    def plot_avg_scores(self, path="outputs/figures/avg_scores_by_track.png"):
        # Compare average subject scores per track
        avg_df = self.avg_scores().melt(id_vars='Track',
                                        value_vars=['Math', 'English', 'Science', 'History'],
                                        var_name='Subject',
                                        value_name='Average Score')
        plt.figure(figsize=(10, 5))
        sns.barplot(data=avg_df, x='Track', y='Average Score', hue='Subject', palette='pastel')
        plt.title("Average Scores by Track and Subject")
        plt.xlabel("Track")
        plt.ylabel("Average Score")
        plt.ylim(0, 100)
        plt.legend(title='Subject')
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()
    def plot_history_distribution(self, path="outputs/figures/history_grades_by_track.png"):
        # Show grade distributions for History by track
        plt.figure(figsize=(8, 5))
        sns.histplot(data=self.df, x='History', hue='Track', kde=True, bins=14, element='step')
        plt.title("Distributions of History grades by Track")
        plt.xlabel("History grade")
        plt.ylabel("Number of students")
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()
    def plot_avg_math(self, path="outputs/figures/avg_math_scores_by_track.png"):
        # Plot mean math score by track
        plt.figure(figsize=(8, 5))
        sns.barplot(data=self.avg_scores()[['Track', 'Math']], x='Track', y='Math', hue='Track', palette='pastel')
        plt.title("Average Mathematics Scores by Track")
        plt.xlabel("Track")
        plt.ylabel("Average Math Score")
        plt.ylim(0, 100)
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()
    def plot_attendance_vs_project(self, path="outputs/figures/attendance_vs_project_by_track.png"):
        # Scatterplot showing relationship between attendance and project score
        plt.figure(figsize=(8, 5))
        sns.scatterplot(data=self.df, x="Attendance (%)", y="ProjectScore", hue="Track", palette="pastel", alpha=0.7, edgecolor="none")
        plt.title("Correlation between Attendance and Project Score by Track")
        plt.xlabel("Attendance Rate (%)")
        plt.ylabel("Project Score")
        plt.grid(True, linestyle='--', alpha=0.6)
        plt.legend(title="Track", bbox_to_anchor=(1.05, 1), loc='upper left')
        plt.tight_layout()
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()

## 4. Cohort-Level Analysis

class CohortStatistics:
    def __init__(self, df):
        self.df = df
    # Similar structure to TrackStatistics but grouped by Cohort
    def nb_students(self):
        return self.df['Cohort'].value_counts().rename_axis('Cohort').reset_index(name='Nb Students')
    def avg_scores(self):
        return self.df.groupby('Cohort')[['Math', 'English', 'Science', 'History']].mean().rename_axis('Cohort').reset_index()
    def avg_attendance(self):
        return self.df.groupby('Cohort')['Attendance (%)'].mean().rename_axis('Cohort').reset_index()
    def avg_project_scores(self):
        return self.df.groupby('Cohort')['ProjectScore'].mean().rename_axis('Cohort').reset_index()
    def pass_rate(self):
        return (self.df.groupby('Cohort')['Passed (Y/N)'].mean() * 100).rename_axis('Cohort').reset_index(name='Pass Rate (%)')
    def plot_pass_rate(self, path="outputs/figures/pass_rate_by_cohort.png"):
        # Plot pass rate by cohort
        plt.figure(figsize=(8, 5))
        sns.barplot(data=self.pass_rate(), x='Cohort', y='Pass Rate (%)', hue='Cohort', palette='pastel')
        plt.title("Pass Rate by Cohort")
        plt.xlabel("Cohort")
        plt.ylabel("Pass Rate (%)")
        plt.ylim(0, 100)
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()
    def plot_avg_scores(self, path="outputs/figures/avg_scores_by_cohort.png"):
        # Plot average subject scores per cohort
        avg_df = self.avg_scores().melt(id_vars='Cohort',
                                        value_vars=['Math', 'English', 'Science', 'History'],
                                        var_name='Subject',
                                        value_name='Average Score')
        plt.figure(figsize=(10, 5))
        sns.barplot(data=avg_df, x='Cohort', y='Average Score', hue='Subject', palette='pastel')
        plt.title("Average Scores by Cohort and Subject")
        plt.xlabel("Cohort")
        plt.ylabel("Average Score")
        plt.ylim(0, 100)
        plt.legend(title='Subject')
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()

class IncomeStatusStatistics:
    def __init__(self, df):
        self.df = df
    # Group statistics based on income student status
    def nb_students(self):
        return self.df['IncomeStudent'].value_counts().rename_axis('IncomeStudent').reset_index(name='Nb Students')
    def avg_scores(self):
        return self.df.groupby('IncomeStudent')[['Math', 'English', 'Science', 'History']].mean().rename_axis('IncomeStudent').reset_index()
    def avg_attendance(self):
        return self.df.groupby('IncomeStudent')['Attendance (%)'].mean().rename_axis('IncomeStudent').reset_index()
    def avg_project_scores(self):
        return self.df.groupby('IncomeStudent')['ProjectScore'].mean().rename_axis('IncomeStudent').reset_index()
    def avg_passed_rate(self):
        return (self.df.groupby('IncomeStudent')['Passed (Y/N)'].mean() * 100).rename_axis('IncomeStudent').reset_index(name='Pass rate (%)')
    def plot_avg_scores(self, path="outputs/figures/avg_scores_by_income_status.png"):
        # Compare average subject scores between income and non-income students
        plot_data = self.avg_scores().melt(
            id_vars='IncomeStudent',
            value_vars=['Math', 'English', 'Science', 'History'],
            var_name='Subject',
            value_name='AverageScore'
        )

        plot_data['IncomeStudent'] = plot_data['IncomeStudent'].map({
            True: 'Income Student',
            False: 'Non-Income Student'
        })

        plt.figure(figsize=(8, 5))
        sns.barplot(data=plot_data, x='Subject', y='AverageScore', hue='IncomeStudent', palette='pastel', edgecolor='black')
        plt.title("Average Scores by Subject and Income Status")
        plt.xlabel("Subject")
        plt.ylabel("Average Score")
        plt.ylim(0, 100)
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.legend(title='Income Status')
        plt.tight_layout()
        plt.savefig(path, dpi=200, bbox_inches="tight")
        plt.close()

## 5. Final Report Generation

def export_cleaned_data(df, path="outputs/cleaned_dataset.csv"):
    # Save cleaned dataset to CSV
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False)

def compute_all_stats(df):
    # Helper to round numerical values
    def _round_numeric(_df, ndigits=1):
        out = _df.copy()
        num_cols = out.select_dtypes(include=[np.number, "Float64", "Int64"]).columns
        if len(num_cols):
            out[num_cols] = out[num_cols].astype(float).round(ndigits)
        return out

    # Instantiate statistic classes
    ts = TrackStatistics(df)
    cs = CohortStatistics(df)
    ins = IncomeStatusStatistics(df)

    # Compute and round all relevant statistics
    track_counts = ts.nb_students()
    track_avgs = _round_numeric(ts.avg_scores())
    track_att = _round_numeric(ts.avg_attendance())
    track_proj = _round_numeric(ts.avg_project_scores())
    track_pass = _round_numeric(ts.pass_rate())
    track_corr = _round_numeric(ts.corr_attendance_project())

    cohort_counts = cs.nb_students()
    cohort_avgs = _round_numeric(cs.avg_scores())
    cohort_att = _round_numeric(cs.avg_attendance())
    cohort_proj = _round_numeric(cs.avg_project_scores())
    cohort_pass = _round_numeric(cs.pass_rate())

    income_counts = ins.nb_students()
    income_avgs = _round_numeric(ins.avg_scores())
    income_att = _round_numeric(ins.avg_attendance())
    income_proj = _round_numeric(ins.avg_project_scores())
    income_pass = _round_numeric(ins.avg_passed_rate())

    # Store results in a structured dictionary for Excel export
    stats_dict = {
        "Track - Counts": track_counts,
        "Track - Avg Scores": track_avgs,
        "Track - Attendance": track_att,
        "Track - Project": track_proj,
        "Track - Pass Rate": track_pass,
        "Track - Corr (%)": track_corr,

        "Cohort - Counts": cohort_counts,
        "Cohort - Avg Scores": cohort_avgs,
        "Cohort - Attendance": cohort_att,
        "Cohort - Project": cohort_proj,
        "Cohort - Pass Rate": cohort_pass,

        "Income - Counts": income_counts,
        "Income - Avg Scores": income_avgs,
        "Income - Attendance": income_att,
        "Income - Project": income_proj,
        "Income - Pass Rate": income_pass,
    }

    return stats_dict

def export_stats_excel(stats_dict, path="outputs/summary_report.xlsx"):
    # Export all computed statistics into a multi-sheet Excel report
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        for sheet, df in stats_dict.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

def export_figures_png(df):
    # Generate all predefined plots and export as PNG
    ts = TrackStatistics(df)
    cs = CohortStatistics(df)
    ins = IncomeStatusStatistics(df)

    ts.plot_history_distribution()
    ts.plot_avg_math()
    ts.corr_attendance_project()
    ts.plot_avg_scores()
    ts.plot_pass_rate()
    ts.plot_attendance_vs_project()

    cs.plot_avg_scores()
    cs.plot_pass_rate()

    ins.plot_avg_scores()

def main(xlsx_path="student_grades_2027-2028.xlsx"):
    # Main workflow for data cleaning, analysis, and export
    df = load_all_sheets(xlsx_path)
    df = clean_df(df)
    df = drop_duplicates(df)
    stats = compute_all_stats(df)
    export_cleaned_data(df)
    export_stats_excel(stats)
    export_figures_png(df)

if __name__ == "__main__":
    main()
