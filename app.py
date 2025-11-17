from pathlib import Path
import streamlit as st
import pandas as pd

# Import the module where the original code lives
from report_generator import (
    load_all_sheets, clean_df, drop_duplicates,
    compute_all_stats, export_cleaned_data, export_stats_excel, export_figures_png, FIGDIR
)

st.set_page_config(
    page_title="Student Analytics",
    page_icon="📊",
    layout="wide"
)

# ---------- Sidebar ----------
st.sidebar.title("📁 Upload Excel")
uploaded = st.sidebar.file_uploader(
    "Excel file (.xlsx) with multiple sheets (one per Track)",
    type=["xlsx"]
)
st.sidebar.caption("Tip: each sheet should represent a Track; keep the columns expected by your script.")

st.title("🎓 Student Analytics – Dashboard")
st.write(
    "Upload your Excel file. The pipeline will clean the data, compute statistics, "
    "generate figures, and provide downloads (CSV + XLSX report)."
)

# ---------- Helpers ----------
@st.cache_data(show_spinner=False)
def _persist_upload_to_disk(uploaded_file) -> Path:
    """Save the uploaded file locally so pandas/openpyxl can read from a path."""
    tmp_path = Path(st.session_state.get("_last_upload_path", "uploaded.xlsx"))
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    st.session_state["_last_upload_path"] = str(tmp_path)
    return tmp_path

def _offer_downloads(clean_csv_path: Path, report_xlsx_path: Path):
    c1, c2 = st.columns(2)
    with open(clean_csv_path, "rb") as f:
        c1.download_button(
            "⬇️ Download cleaned dataset (CSV)",
            data=f.read(),
            file_name=clean_csv_path.name,
            mime="text/csv",
            use_container_width=True
        )
    with open(report_xlsx_path, "rb") as f:
        c2.download_button(
            "⬇️ Download summary report (XLSX)",
            data=f.read(),
            file_name=report_xlsx_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

def _show_fig_if_exists(path: Path, caption: str):
    if path.exists():
        st.image(str(path), caption=caption, use_column_width=True)

# ---------- Main flow ----------
if uploaded is None:
    st.info("Upload a file in the sidebar to get started.")
else:
    try:
        # 1) Save locally + raw load
        xlsx_path = _persist_upload_to_disk(uploaded)
        raw_df = load_all_sheets(str(xlsx_path))

        # 2) Clean + drop duplicates
        df = clean_df(raw_df)
        df = drop_duplicates(df)

        # 3) Stats + Exports
        stats = compute_all_stats(df)
        cleaned_csv = Path("outputs/cleaned_dataset.csv")
        report_xlsx = Path("outputs/summary_report.xlsx")
        export_cleaned_data(df, path=str(cleaned_csv))
        export_stats_excel(stats, path=str(report_xlsx))

        # 4) Figures
        export_figures_png(df)
        st.success("Processing complete ✅")

        # ---------- KPIs ----------
        st.subheader("🔢 Key Metrics")
        c1, c2, c3, c4 = st.columns(4)
        try:
            total_students = len(df)
            n_tracks = df["Track"].nunique() if "Track" in df else 0
            n_cohorts = df["Cohort"].nunique() if "Cohort" in df else 0
            pass_rate_overall = (df["Passed (Y/N)"].mean() * 100) if "Passed (Y/N)" in df else float("nan")
        except Exception:
            total_students, n_tracks, n_cohorts, pass_rate_overall = 0, 0, 0, float("nan")

        c1.metric("Total students", f"{total_students}")
        c2.metric("Tracks", f"{n_tracks}")
        c3.metric("Cohorts", f"{n_cohorts}")
        c4.metric("Overall pass rate", f"{pass_rate_overall:.1f} %" if pd.notna(pass_rate_overall) else "—")

        st.divider()

        # ---------- Summary tables ----------
        st.subheader("📋 Summary Tables")
        tabs = st.tabs(["Track", "Cohort", "Income Status"])

        with tabs[0]:
            colA, colB = st.columns(2)
            colA.dataframe(stats["Track - Counts"], use_container_width=True)
            colB.dataframe(stats["Track - Pass Rate"], use_container_width=True)

            st.dataframe(stats["Track - Avg Scores"], use_container_width=True)

            cA, cB = st.columns(2)
            cA.dataframe(stats["Track - Attendance"], use_container_width=True)
            cB.dataframe(stats["Track - Project"], use_container_width=True)

            st.markdown("#### Attendance Rate vs Project Score Correlation")
            st.dataframe(stats["Track - Corr (%)"], use_container_width=True)

        with tabs[1]:
            colA, colB = st.columns(2)
            colA.dataframe(stats["Cohort - Counts"], use_container_width=True)
            colB.dataframe(stats["Cohort - Pass Rate"], use_container_width=True)
            st.dataframe(stats["Cohort - Avg Scores"], use_container_width=True)
            cA, cB = st.columns(2)
            cA.dataframe(stats["Cohort - Attendance"], use_container_width=True)
            cB.dataframe(stats["Cohort - Project"], use_container_width=True)

        with tabs[2]:
            colA, colB = st.columns(2)
            colA.dataframe(stats["Income - Counts"], use_container_width=True)
            colB.dataframe(stats["Income - Pass Rate"], use_container_width=True)
            st.dataframe(stats["Income - Avg Scores"], use_container_width=True)
            cA, cB = st.columns(2)
            cA.dataframe(stats["Income - Attendance"], use_container_width=True)
            cB.dataframe(stats["Income - Project"], use_container_width=True)

        st.divider()

        # ---------- Visualizations (generated PNGs) ----------
        st.subheader("📈 Visualizations")
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
        grid_cols = st.columns(2)
        for i, (fname, caption) in enumerate(fig_paths):
            p = FIGDIR / fname
            with grid_cols[i % 2]:
                _show_fig_if_exists(p, caption)

        st.divider()

        # ---------- Downloads ----------
        st.subheader("📦 Downloads")
        _offer_downloads(cleaned_csv, report_xlsx)

        # ---------- Cleaned data preview ----------
        st.subheader("🧹 Cleaned Data (preview)")
        st.dataframe(df.head(50), use_container_width=True)

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.exception(e)
