"""
Financial Resilience Data Cleaning Script
Reads the raw Excel file, fills missing values, cleans column names,
and exports a cleaned Excel file ready for analysis.

Requirements:
    pip install pandas openpyxl
"""

import pandas as pd

# 1. Load Data 
INPUT_FILE  = "C:\\Users\\dell\\Downloads\\Fiancial Resilience (1) (1).xlsx"
OUTPUT_FILE = "Financial_Resilience_Cleaned.xlsx"
SHEET       = "Clean"

df = pd.read_excel(INPUT_FILE, sheet_name=SHEET)

print("=" * 60)
print(f"Loaded sheet '{SHEET}' → {df.shape[0]} rows × {df.shape[1]} columns")
print("=" * 60)

# 2. Fix Column Names 
# Strip leading/trailing whitespace and remove stray newline characters
df.columns = (
    df.columns
      .str.strip()
      .str.replace(r"^\n+", "", regex=True)  # remove leading \n
      .str.replace(r"\s+", " ", regex=True)  # collapse internal whitespace
)

print("\n[1] Column names cleaned.")

# 3. Missing Value Summary (before cleaning) 
missing_before = df.isnull().sum()
missing_cols   = missing_before[missing_before > 0]

print("\n[2] Missing values BEFORE cleaning:")
print(missing_cols.to_string())

# 4. Fill Missing Values 

# -- 4a. Categorical / free-text columns → "Not Specified"
# These columns are blank because the question was not applicable to the
# respondent (e.g. someone who does not save has no "Saving Method").
categorical_na_cols = [
    "Largest Expense",         # 75 missing  — not all respondents answered
    "Shared Largest Expense",  # 149 missing — only filled if sharing finances
    "Saving Method",           # 71 missing  — only for those who save
    "No Saving Reason",        # 97 missing  — only for non-savers
]

for col in categorical_na_cols:
    if col in df.columns:
        df[col] = df[col].fillna("Not Specified")

# -- 4b. Monthly Savings → "Not Specified"
# Missing because the respondent does not save or chose not to answer.
if "Monthly Savings" in df.columns:
    df["Monthly Savings"] = df["Monthly Savings"].fillna("Not Specified")

# -- 4c. Finance Support → most-common (mode) imputation
# Only 1 value missing — safe to use the mode.
if "Finance Support" in df.columns:
    mode_val = df["Finance Support"].mode()[0]
    df["Finance Support"] = df["Finance Support"].fillna(mode_val)
    print(f"\n[3] 'Finance Support' 1 missing value filled with mode: '{mode_val}'")

# 5. Standardise Text Columns
# Strip extra whitespace from all object (string) columns.
obj_cols = df.select_dtypes(include="object").columns
df[obj_cols] = df[obj_cols].apply(lambda s: s.str.strip())

print("\n[4] Whitespace stripped from all text columns.")

# 6. Fix Known Value Inconsistencies 
# Normalise free-text "Other:" entries → "Other" (remove trailing colon)
for col in obj_cols:
    df[col] = df[col].replace({"Other:": "Other"}, regex=False)

# Normalise near-duplicate responses in Finance Support
if "Finance Support" in df.columns:
    df["Finance Support"] = df["Finance Support"].replace({
        "No response": "No Response",
        "none":        "None",
    })

print("[5] Value inconsistencies normalised.")

# 7. Drop Entirely Empty Rows (safety check) 
before = len(df)
df.dropna(how="all", inplace=True)
after  = len(df)
if before != after:
    print(f"\n[6] Dropped {before - after} completely empty rows.")
else:
    print("[6] No completely empty rows found.")

# 8. Reset Index
df.reset_index(drop=True, inplace=True)

# 9. Missing Value Summary (after cleaning)
missing_after = df.isnull().sum()
remaining     = missing_after[missing_after > 0]

print("\n[7] Missing values AFTER cleaning:")
if remaining.empty:
    print("    ✓ No missing values remain.")
else:
    print(remaining.to_string())

# 10. Export Cleaned Data
df.to_excel(OUTPUT_FILE, index=False, sheet_name="Cleaned_Data")

print(f"\n{'=' * 60}")
print(f"✓ Cleaned file saved → {OUTPUT_FILE}")
print(f"  Rows: {df.shape[0]}  |  Columns: {df.shape[1]}")
print("=" * 60)