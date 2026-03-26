"""
Purpose
-------
This app provides a simple UI around the `report_generator.py` pipeline:
- Upload an Excel workbook (.xlsx) where each sheet represents a Track.
- Run the cleaning + stats + export pipeline on demand (button).
- Display KPIs, summary tables, and saved figures.
- Offer downloads for cleaned CSV + Excel report.
"""

from __future__ import annotations

from pathlib import Path
import hashlib

import streamlit as st
import pandas as pd

# Import the core pipeline functions and output directories.
# NOTE: This assumes report_generator.py is in the same project and on PYTHONPATH.
from report_generator import (
    load_all_sheets,
    clean_df,
    drop_duplicates,
    compute_all_stats,
    export_cleaned_data,
    export_stats_excel,
    export_figures_png,
    FIGDIR,
    OUTDIR,
)

# Configure Streamlit page-level settings (title/icon/layout).
st.set_page_config(page_title="Student Analytics", page_icon="📊", layout="wide")


# =============================================================================
# HELPERS
# =============================================================================

def _bytes_sha256(b: bytes) -> str:
    """
    Return a short SHA256 digest (12 hex chars) of a bytes object.

    Used to generate stable filenames for uploaded content.
    """
    return hashlib.sha256(b).hexdigest()[:12]


def _persist_upload_to_disk(uploaded_file) -> Path:
    """
    Persist the uploaded file to disk using a deterministic content-hash filename.

    Why:
    - Streamlit reruns the script often; keeping a stable file path helps avoid surprises.
    - Users may upload files with identical names; content hashing prevents collisions.
    - This also avoids relying on Streamlit's temp objects if you want reproducible exports.

    Returns
    -------
    Path
        Path to the persisted .xlsx file in outputs/uploads/
    """
    data = uploaded_file.getvalue()
    digest = _bytes_sha256(data)

    upload_dir = OUTDIR / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)

    # Minimal filename sanitization to avoid path traversal / invalid path separators.
    safe_name = uploaded_file.name.replace("/", "_").replace("\\", "_")
    path = upload_dir / f"{digest}__{safe_name}"

    # Only write once (idempotent) so reruns don't constantly rewrite the same file.
    if not path.exists():
        path.write_bytes(data)

    return path


def _offer_downloads(clean_csv_path: Path, report_xlsx_path: Path) -> None:
    """
    Render download buttons for the generated artifacts.

    We read bytes from disk rather than holding large blobs in memory/state.
    """
    c1, c2 = st.columns(2)

    with open(clean_csv_path, "rb") as f:
        c1.download_button(
            "⬇️ Download cleaned dataset (CSV)",
            data=f.read(),
            file_name=clean_csv_path.name,
            mime="text/csv",
            use_container_width=True,
        )

    with open(report_xlsx_path, "rb") as f:
        c2.download_button(
            "⬇️ Download summary report (XLSX)",
            data=f.read(),
            file_name=report_xlsx_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def _show_fig_if_exists(path: Path, caption: str) -> None:
    """
    Show an image if it exists on disk.

    The try/except is here for Streamlit version compatibility:
    - newer Streamlit uses use_container_width
    - older Streamlit used use_column_width
    """
    if not path.exists():
        return

    try:
        st.image(str(path), caption=caption, use_container_width=True)
    except TypeError:
        # Backward compatibility for older Streamlit versions
        st.image(str(path), caption=caption, use_column_width=True)


def _run_pipeline(xlsx_path: Path) -> dict:
    """
    Run the full pipeline once and return all artifacts for display.

    Notes
    -----
    - This function does not cache on purpose. The "Run analysis" button is the control.
    - Exports are written to OUTDIR/ and FIGDIR/ (as defined by report_generator.py).

    Returns
    -------
    dict
        {
          "df": cleaned dataframe,
          "stats": dict[str, pd.DataFrame],
          "cleaned_csv": Path to outputs/cleaned_dataset.csv,
          "report_xlsx": Path to outputs/summary_report.xlsx,
        }
    """
    # Load raw data from all sheets (Track = sheet name)
    raw_df = load_all_sheets(str(xlsx_path))

    # Clean and de-duplicate (see report_generator for the exact assumptions)
    df = clean_df(raw_df)
    df = drop_duplicates(df)

    # Build report tables
    stats = compute_all_stats(df)

    # Export artifacts to disk
    cleaned_csv = OUTDIR / "cleaned_dataset.csv"
    report_xlsx = OUTDIR / "summary_report.xlsx"

    # NOTE: these export functions accept Path in your generator; we pass Path directly.
    export_cleaned_data(df, path=cleaned_csv)
    export_stats_excel(stats, path=report_xlsx)
    export_figures_png(df)

    return {
        "df": df,
        "stats": stats,
        "cleaned_csv": cleaned_csv,
        "report_xlsx": report_xlsx,
    }


# =============================================================================
# SIDEBAR (INPUTS)
# =============================================================================

st.sidebar.title("📁 Upload Excel")
uploaded = st.sidebar.file_uploader(
    "Excel file (.xlsx) with multiple sheets (one per Track)",
    type=["xlsx"],
)
st.sidebar.caption(
    "Tip: each sheet should represent a Track; keep the columns expected by your script."
)

st.title("🎓 Student Analytics – Dashboard")
st.write(
    "Upload your Excel file, then click **Run analysis**. The pipeline cleans the data, "
    "computes statistics, generates figures, and provides downloads (CSV + XLSX report)."
)


# =============================================================================
# MAIN (CONTROL FLOW)
# =============================================================================

# Stop early until the user provides a file.
if uploaded is None:
    st.info("Upload a file in the sidebar to get started.")
    st.stop()

# Persist file to disk (stable path) so we can re-run deterministically.
xlsx_path = _persist_upload_to_disk(uploaded)

# Button: avoids running heavy work on every Streamlit rerun.
run_clicked = st.button("▶️ Run analysis", type="primary", use_container_width=True)

# Session state: keep results across reruns until the user changes file or re-runs.
if "artifacts" not in st.session_state:
    st.session_state["artifacts"] = None
if "last_xlsx_path" not in st.session_state:
    st.session_state["last_xlsx_path"] = None

# If the uploaded file changes (new hash/name), invalidate existing results.
if st.session_state["last_xlsx_path"] != str(xlsx_path):
    st.session_state["artifacts"] = None
    st.session_state["last_xlsx_path"] = str(xlsx_path)

# Run pipeline on demand.
if run_clicked:
    with st.spinner("Running pipeline..."):
        try:
            st.session_state["artifacts"] = _run_pipeline(xlsx_path)
            st.success("Processing complete ✅")
        except Exception as e:
            # Show an error + the full stack trace for debugging in the app.
            st.error(f"An error occurred: {e}")
            st.exception(e)
            st.stop()

artifacts = st.session_state["artifacts"]
if artifacts is None:
    st.warning("Click **Run analysis** to generate results.")
    st.stop()

# Unpack artifacts for display.
df: pd.DataFrame = artifacts["df"]
stats: dict[str, pd.DataFrame] = artifacts["stats"]
cleaned_csv: Path = artifacts["cleaned_csv"]
report_xlsx: Path = artifacts["report_xlsx"]


# =============================================================================
# KPIs
# =============================================================================

st.subheader("🔢 Key Metrics")

c1, c2, c3, c4 = st.columns(4)

# These KPIs assume the cleaned dataset still contains these columns.
total_students = len(df)
n_tracks = df["Track"].nunique() if "Track" in df.columns else 0
n_cohorts = df["Cohort"].nunique() if "Cohort" in df.columns else 0

# Mean of boolean -> pass rate (proportion). Multiply by 100 for percentage.
pass_rate_overall = (
    (df["Passed (Y/N)"].mean() * 100) if "Passed (Y/N)" in df.columns else float("nan")
)

c1.metric("Total students", f"{total_students}")
c2.metric("Tracks", f"{n_tracks}")
c3.metric("Cohorts", f"{n_cohorts}")
c4.metric(
    "Overall pass rate",
    f"{pass_rate_overall:.1f} %" if pd.notna(pass_rate_overall) else "—",
)

st.divider()


# =============================================================================
# SUMMARY TABLES
# =============================================================================

st.subheader("📋 Summary Tables")
tabs = st.tabs(["Track", "Cohort", "Income Status"])

with tabs[0]:
    colA, colB = st.columns(2)
    colA.dataframe(stats.get("Track - Counts", pd.DataFrame()), use_container_width=True)
    colB.dataframe(stats.get("Track - Pass Rate", pd.DataFrame()), use_container_width=True)

    st.dataframe(stats.get("Track - Avg Scores", pd.DataFrame()), use_container_width=True)

    cA, cB = st.columns(2)
    cA.dataframe(stats.get("Track - Attendance", pd.DataFrame()), use_container_width=True)
    cB.dataframe(stats.get("Track - Project", pd.DataFrame()), use_container_width=True)

    # Backward/alternative key handling:
    # If you renamed the correlation sheet, this keeps the app resilient.
    corr_df = stats.get("Track - Corr (%)")
    if corr_df is None:
        corr_df = stats.get("Track - Corr (r)", pd.DataFrame())

    st.markdown("#### Attendance Rate vs Project Score Correlation")
    st.dataframe(corr_df, use_container_width=True)

with tabs[1]:
    colA, colB = st.columns(2)
    colA.dataframe(stats.get("Cohort - Counts", pd.DataFrame()), use_container_width=True)
    colB.dataframe(stats.get("Cohort - Pass Rate", pd.DataFrame()), use_container_width=True)

    st.dataframe(stats.get("Cohort - Avg Scores", pd.DataFrame()), use_container_width=True)

    cA, cB = st.columns(2)
    cA.dataframe(stats.get("Cohort - Attendance", pd.DataFrame()), use_container_width=True)
    cB.dataframe(stats.get("Cohort - Project", pd.DataFrame()), use_container_width=True)

with tabs[2]:
    colA, colB = st.columns(2)
    colA.dataframe(stats.get("Income - Counts", pd.DataFrame()), use_container_width=True)
    colB.dataframe(stats.get("Income - Pass Rate", pd.DataFrame()), use_container_width=True)

    st.dataframe(stats.get("Income - Avg Scores", pd.DataFrame()), use_container_width=True)

    cA, cB = st.columns(2)
    cA.dataframe(stats.get("Income - Attendance", pd.DataFrame()), use_container_width=True)
    cB.dataframe(stats.get("Income - Project", pd.DataFrame()), use_container_width=True)

st.divider()


# =============================================================================
# VISUALIZATIONS
# =============================================================================

st.subheader("📈 Visualizations")

# Filenames are expected to be produced by export_figures_png(df) in report_generator.py.
fig_paths = [
    ("pass_rate_by_track.png", "Pass Rate by Track"),
    ("avg_scores_by_track.png", "Average Scores by Track and Subject"),
    ("history_grades_by_track.png", "Distributions of History Grades by Track"),
    ("avg_math_scores_by_track.png", "Average Mathematics Scores by Track"),
    ("attendance_vs_project_by_track.png", "Attendance vs Project Score by Track"),
    ("avg_scores_by_cohort.png", "Average Scores by Cohort and Subject"),
    ("pass_rate_by_cohort.png", "Pass Rate by Cohort"),
    ("avg_scores_by_income_status.png", "Average Scores by Subject and Income Status"),
]

# 2-column layout for images.
cols = st.columns(2)
for i, (fname, caption) in enumerate(fig_paths):
    p = FIGDIR / fname
    with cols[i % 2]:
        _show_fig_if_exists(p, caption)

st.divider()


# =============================================================================
# DOWNLOADS
# =============================================================================

st.subheader("📦 Downloads")
_offer_downloads(cleaned_csv, report_xlsx)

st.divider()


# =============================================================================
# CLEANED DATA PREVIEW
# =============================================================================

st.subheader("🧹 Cleaned Data (preview)")
st.dataframe(df.head(50), use_container_width=True)
