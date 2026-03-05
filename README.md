# 🎓 Student Analytics Dashboard

## 📘 Description
**Student Analytics** is a Python-based data processing and visualization tool that analyzes student performance across multiple academic tracks and cohorts.  
It combines data cleaning, statistical computation, and visualization — and provides an **interactive Streamlit dashboard** for exploration.

The application was built in two main components:
- **`report_generator.py`** – Core data processing and report generation pipeline  
- **`app.py`** – Streamlit web application for interactive analytics  

---

## ⚙️ Instructions for Running the Script

### 🧩 Requirements
Python 3.10+ recommended.

Install dependencies (preferably in a virtual environment):

```bash
pip install -r requirements.txt
```

---

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

---

### ▶️ Option 2: Launch the Streamlit Dashboard

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
- Impute missing numeric values with **Track-level means**.
- Impute missing boolean/categorical values with **Track-level mode**.
- Normalize boolean fields (`Passed (Y/N)`, `IncomeStudent`) with consistent mappings.
- Remove duplicate `(StudentID, Term, Track)` entries.

---

### Statistical Analysis
Computed per Track / Cohort / Income status:
- Student counts
- Average subject scores
- Attendance rates
- Project scores
- Pass rates
- Correlation between Attendance and Project Score (Pearson r)

---

### Visualization
Generates `.png` figures in `outputs/figures/`, including:
- Pass rate by Track and Cohort
- Average subject scores by Track, Cohort, and Income status
- History grade distributions by Track
- Average Math scores by Track
- Attendance vs Project Score scatterplots by Track

---

## ⚠️ Assumptions & Limitations

Each **sheet** in the Excel file represents one **Track** and must contain the same columns:

```
StudentID | FirstName | LastName | Class | Cohort | Term | 
Math | English | Science | History | 
Attendance (%) | ProjectScore | IncomeStudent | Passed (Y/N)
```

### Format conventions
- **StudentID**: four-digit code (e.g., `1234`)
- **Cohort**: format `YY-YY` (e.g., `25-26`)
- **Class**: two-digit code followed by a letter (e.g., `11A`)
- **Term**: 1 or 2
- **Scores**: expected on a 0–100 scale (values outside are treated as missing)
- Boolean fields accept `Y/N`, `Yes/No`, `True/False`, `1/0`, etc.

### Processing assumptions
- Missing numeric values are imputed using the **mean within the same Track**.
- Missing boolean values are imputed using the **mode within the same Track**.
- Rows still incomplete after cleaning (for required report columns) are dropped.
- Reports and figures are generated in the `outputs/` directory (created automatically).
- The Streamlit app must be run in an environment where local file I/O is permitted.

---

## 🧱 Design Principles
- Modular pipeline (`report_generator.py` reusable independently)
- Clear separation between data processing and UI
- Deterministic file handling for dashboard uploads
- Exportable and reproducible outputs