"""
HR Data ETL Pipeline — Staffing & Termination Data Cleaning
============================================================
Description:
    Reads raw Oracle HCM exports (Excel format), cleans and restructures
    the data for downstream analysis in Power BI.

    Part 1: Staffing register — combines monthly files, selects key columns
    Part 2: Termination data — parses nested org hierarchy, extracts clean fields

Input files:
    - "Staffing_Register_<month>.xlsx" — monthly headcount files (row 4 = header)
    - "Termination_Q1.xlsx"           — termination register (row 4 = header)

Output:
    - Output/staffing_clean.xlsx   — cleaned staffing register
    - Output/termination_clean.xlsx — cleaned termination list

Author: Korlan Shaketova
"""

import pandas as pd
import glob
import os

# ─────────────────────────────────────────────
# PART 1: STAFFING REGISTER
# ─────────────────────────────────────────────

# Read all monthly staffing files matching the pattern
all_files = glob.glob("Staffing_Register_*.xlsx")

records = []
for file in all_files:
    df = pd.read_excel(file, header=3)
    df.columns = df.columns.str.strip()

    # Extract month name from filename
    month = os.path.basename(file).replace("Staffing_Register_", "").replace(".xlsx", "")
    df["Month"] = month
    records.append(df)

staffing = pd.concat(records, ignore_index=True)

# Select only relevant columns for analysis
columns_to_keep = [
    "Division_1",
    "Division_2",
    "Division_3",
    "Division_4",
    "Division_5",
    "Register_Type",
    "Position",
    "Position_EN",
    "Personnel_Category",
    "Grade",
    "Planned_Headcount",
    "Working_Conditions",
    "Salary_Per_Employee",
    "Total_Payroll_KZT",
    "Shortage_Allowance",
    "Salary_With_Allowance",
    "Total_Payroll_With_Allowance_KZT",
    "Month"
]

# Drop rows with no position (empty rows from merged Excel cells)
staffing_clean = staffing.dropna(subset=["Position"])

# Save result
os.makedirs("Output", exist_ok=True)
staffing_clean.to_excel("Output/staffing_clean.xlsx", index=False)
print(f"Staffing register saved: {staffing_clean.shape[0]} rows, {staffing_clean.shape[1]} columns")


# ─────────────────────────────────────────────
# PART 2: TERMINATION DATA
# ─────────────────────────────────────────────

term = pd.read_excel("Termination_Q1.xlsx", header=3)
term.columns = term.columns.str.strip()

# Rename unnamed columns from Oracle export
term = term.rename(columns={
    "Position_From": "Position_Raw",
    "Unnamed: 11": "Category",
    "Unnamed: 15": "Org_Path",
    "Unnamed: 16": "Division_Type"
})

# Drop header artifacts and empty rows
term = term.dropna(subset=["Employee_ID"])

# ── Parse org hierarchy ──────────────────────
# Oracle exports org path as "Company->Site->Department->Team"
# We reverse it so the most specific division comes first (Подр_1)

def parse_org_path(row):
    """
    Splits Oracle org path string by '->' separator,
    reverses order (most specific first), pads to 5 levels.
    Example:
        "KAZ Minerals->VCM->Mine->Drilling" 
        → ["Drilling", "Mine", "VCM", "KAZ Minerals", None]
    """
    parts = [x for x in row if x is not None and str(x) != "nan"]
    parts = parts[::-1]  # reverse: most specific first
    return pd.Series(parts + [None] * (5 - len(parts)))

org_split = term["Org_Path"].astype(str).str.split("->", expand=True)
org_parsed = org_split.apply(parse_org_path, axis=1)
org_parsed.columns = ["Div_1", "Div_2", "Div_3", "Div_4", "Div_5"]

term = pd.concat([term, org_parsed], axis=1)

# ── Parse position and employee ID ──────────
# Oracle stores them as "Job Title..123456" — split on ".."
term[["Position_Clean", "Employee_ID_Check"]] = (
    term["Position_Raw"]
    .str.split(r"\.\.", expand=True, n=1)
)

# Keep only valid rows (numeric employee ID)
term = term[term["Employee_ID_Check"].str.isnumeric() == True]

# ── Build final clean table ──────────────────
output_columns = [
    "Employee_ID",
    "Last_Name",
    "First_Name",
    "Middle_Name",
    "Birth_Date",
    "Gender",
    "Position_Clean",
    "Div_1",
    "Div_2",
    "Div_3",
    "Div_4",
    "Division_Type",
    "Hire_Date",
    "Termination_Date",
    "Termination_Basis",
    "Termination_Reason"
]

term_clean = term[output_columns].copy()

# Combine name fields into single full name column
term_clean["Full_Name"] = (
    term_clean["Last_Name"].fillna("") + " " +
    term_clean["First_Name"].fillna("") + " " +
    term_clean["Middle_Name"].fillna("")
).str.strip()

term_clean = term_clean.drop(columns=["Last_Name", "First_Name", "Middle_Name"])

# Save result
term_clean.to_excel("Output/termination_clean.xlsx", index=False)
print(f"Termination data saved: {term_clean.shape[0]} rows, {term_clean.shape[1]} columns")
