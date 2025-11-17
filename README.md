# 🎓 Student Analytics Dashboard

## 📘 Description
**Student Analytics** is a Python-based data processing and visualization tool that analyzes student performance across multiple academic tracks and cohorts.  
It combines data cleaning, statistical computation, and visualization — and provides an **interactive Streamlit dashboard** for exploration.

The application was built in two main components:
- **`report_generator.py`** – Core data processing and report generation pipeline.  
- **`app.py`** – Streamlit web application for interactive analytics.

---

## ⚙️ Instructions for Running the Script

### 🧩 Requirements
Install dependencies (preferably in a virtual environment):
```bash
pip install -r requirements.txt
```

### ▶️ Option 1: Run the Python script directly
This option executes the full analysis pipeline on your Excel file and generates:
- A cleaned dataset (`outputs/cleaned_dataset.csv`)
- A summary report (`outputs/summary_report.xlsx`)
- PNG charts in `outputs/figures/`

Run:
```bash
python report_generator.py
```

Make sure your Excel file (e.g. `student_grades_2027-2028.xlsx`) is in the same directory, with **one sheet per Track** (each sheet must have the same column structure).

### ▶️ Option 2: Launch the Streamlit Dashboard
To run the interactive dashboard:
```bash
streamlit run app.py
```
Then open the provided local URL (usually `http://localhost:8501`) in your browser.

Upload your `.xlsx` file in the sidebar — the app will:
- Clean and merge all sheets  
- Compute statistics per Track, Cohort, and Income status  
- Generate figures  
- Display interactive KPIs, tables, and visualizations  
- Allow you to download the processed CSV and Excel reports

---

## 🔍 Implemented Functionalities

### Data Cleaning
- Merge all Excel sheets into a unified DataFrame.
- Handle missing or invalid entries (`NA`, `n/a`, `-`, etc.).
- Validate and normalize key fields:
  - `Class`, `Cohort`, `StudentID`, `Term`
- Impute missing numerical values with **track-level means**.
- Impute missing categorical values with **track-level mode**.
- Normalize boolean fields (`Passed (Y/N)`, `IncomeStudent`) with consistent mappings.

### Statistical Analysis
- Compute per-Track, per-Cohort, and per-Income summaries:
  - Student counts
  - Average subject scores
  - Attendance rates
  - Project scores
  - Pass rates
  - Correlation between Attendance and Project Score

### Visualization
Automatically generates publication-ready figures:
- Pass rate by track / cohort
- Average subject scores by track, cohort, and income status
- Math and History grade distributions
- Attendance vs. project score by track scatterplots

### Reporting
- Cleaned CSV export (`outputs/cleaned_dataset.csv`)
- Multi-sheet Excel summary (`outputs/summary_report.xlsx`)
- High-quality `.png` figures stored in `outputs/figures/`

---

## ⚠️ Assumptions & Limitations
- Each **sheet** in the Excel file represents a unique **Track** and must contain the same columns:
  ```
  StudentID | FirstName | LastName | Class | Cohort | Term | Math | English | Science | History | Attendance (%) | ProjectScore | IncomeStudent | Passed (Y/N)
  ```
- **Scores** are expected on a scale of 0–100.
- Boolean fields accept `Y/N`, `Yes/No`, `True/False`, `1/0`, etc.
- Invalid or missing rows are dropped if still incomplete after cleaning.
- Figures and reports are generated in `outputs/` (folders created automatically).
- The Streamlit app must be run in an environment where local file I/O is permitted.

- The following format conventions must be respected:
  - **StudentID**: four-digit code (e.g., 1234)
  - **Cohort**: e.g., 25-26 or 26-27.
  - **Class**: two-digit code followed by a letter (e.g., 11A)
  - **Term**: 1 or 2
- Scores are expected on a scale of 0–100.
- Boolean fields accept Y/N, Yes/No, True/False, 1/0, etc.
- Missing numeric values are imputed using the mean of the corresponding column.
- Missing categorical values are imputed using the mode (most frequent value).
- Invalid or missing rows are dropped if still incomplete after cleaning.
- Figures and reports are generated in the outputs/ directory (folders created automatically).
- The Streamlit app must be run in an environment where local file I/O is permitted.

---

## 🌟 Additional Features
- Streamlit interface for real-time exploration and download.
- Automatic figure generation and export.
- Cached file upload to optimize dashboard performance.
- Modular design: `report_generator.py` can be reused independently in other projects.
