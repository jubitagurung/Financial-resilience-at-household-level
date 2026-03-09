"""
Financial Resilience - Full Data Cleaning Script
=================================================
Fixes all remaining issues after the first cleaning pass:
  - Typo/case duplicates in categorical columns
  - Corrupted text & stray values
  - Trailing commas in multi-select columns
  - Truncated & free-text entries
  - Duplicate education level categories
  - Inconsistent casing

Requirements:
    pip install pandas openpyxl
"""

import pandas as pd

# ── 0. Load Data ──────────────────────────────────────────────────────────────
INPUT_FILE  = "C:\\Users\\dell\\OneDrive\\Desktop\\data cleaning\\Financial_Resilience_Cleaned.xlsx"
OUTPUT_FILE = "Financial_Resilience_Final.xlsx"

df = pd.read_excel(INPUT_FILE)
print("=" * 65)
print(f"Loaded  →  {df.shape[0]} rows  ×  {df.shape[1]} columns")
print("=" * 65)

# ── 1. Urban Area — fix typos & case ─────────────────────────────────────────
df["urban area"] = df["urban area"].replace({
    "Prefer not say"  : "Prefer not to say",
    "Prefer notto say": "Prefer not to say",
    "yes"             : "Yes",
})
print("\n[1] 'urban area' typos fixed.")

# ── 2. Finance Support — corrupted text, stray strings ───────────────────────
df["Finance Support"] = df["Finance Support"].replace({
    "Financial literac+AD127y programs or workshops": "Financial literacy programs or workshops",
    "nan"                                           : "Not Specified",
    "saving habits, we Bhutanese seems to lack"     : "Other",
    "Just the income source"                        : "Other",
})
# Fill any real NaN that slipped through
df["Finance Support"] = df["Finance Support"].fillna("Not Specified")
print("[2] 'Finance Support' corrupted/stray values fixed.")

# ── 3. Budget Difficulty — strip trailing commas ─────────────────────────────
df["Budget Difficulty"] = (
    df["Budget Difficulty"]
      .str.strip()
      .str.rstrip(",")
      .str.strip()
)
print("[3] 'Budget Difficulty' trailing commas removed.")

# ── 4. Finance Challenges — strip trailing commas ────────────────────────────
df["Finance Challenges"] = (
    df["Finance Challenges"]
      .str.strip()
      .str.rstrip(",")
      .str.strip()
)
print("[4] 'Finance Challenges' trailing commas removed.")

# ── 5. Finance Motivation — fix truncation, casing, free-text ────────────────
# Long free-text responses to bucket as 'Other'
free_text_motivations = [
    "Helping families and these are some of thing that helps us keep calm and motivated",
    "Managing household finance will motivate in personal and proficianal level to manage "
    "and budget the financial deficiency. It can also help and habituates us in saving and "
    "to operate personal business in proficianal way.",
    "Day to day household problems",
]

df["Finance Motivation"] = df["Finance Motivation"].replace({
    "Family Responsibility / Suppor": "Family Responsibility / Support",
    "financial responsibility"       : "Family Responsibility / Support",
    "No response"                    : "No Response",
    **{v: "Other" for v in free_text_motivations},
})
print("[5] 'Finance Motivation' truncations, casing & free-text fixed.")

# ── 6. Education Level — merge duplicate categories ──────────────────────────
df["Education Level"] = df["Education Level"].replace({
    "Undergraduate": "College Undergraduate",
})
print("[6] 'Education Level' duplicates merged  (Undergraduate → College Undergraduate).")

# ── 7. Monthly Savings — consistent title case ───────────────────────────────
df["Monthly Savings"] = df["Monthly Savings"].replace({
    "prefer not to save": "Prefer Not to Save",
    "prefer not to say" : "Prefer Not to Say",
})
print("[7] 'Monthly Savings' casing standardised.")

# ── 8. No Saving Reason — consolidate near-duplicate student/unemployed ───────
df["No Saving Reason"] = df["No Saving Reason"].replace({
    "No way of coming income as a student": "Student / Not Employed",
    "Still a student"                     : "Student / Not Employed",
    "Am not employed"                     : "Student / Not Employed",
    "No way of coming income"             : "Student / Not Employed",
})
print("[8] 'No Saving Reason' near-duplicate student/unemployed entries consolidated.")

# ── 9. Saving Method — fix leftover 'Other:' in multi-select strings ──────────
df["Saving Method"] = df["Saving Method"].str.replace(
    r"\bOther:\b", "Other", regex=True
)
print("[9] 'Saving Method' trailing 'Other:' corrected to 'Other'.")

# ── 10. Global whitespace & NaN tidy-up ──────────────────────────────────────
str_cols = df.select_dtypes(include="object").columns
df[str_cols] = df[str_cols].apply(lambda s: s.str.strip())
df[str_cols] = df[str_cols].fillna("Not Specified")
print("[10] All text columns stripped & remaining NaNs filled.")

# ── 11. Final Validation ──────────────────────────────────────────────────────
print("\n--- Missing values after cleaning ---")
missing = df.isnull().sum()
remaining = missing[missing > 0]
if remaining.empty:
    print("    ✓  No missing values.")
else:
    print(remaining.to_string())

print("\n--- Unique value counts (key columns) ---")
check_cols = [
    "urban area", "Finance Support", "Monthly Savings",
    "Education Level", "Finance Motivation", "No Saving Reason",
]
for col in check_cols:
    print(f"  {col}  →  {df[col].nunique()} unique values")

# ── 12. Export ────────────────────────────────────────────────────────────────
df.to_excel(OUTPUT_FILE, index=False, sheet_name="Final_Cleaned")

print(f"\n{'=' * 65}")
print(f"✓  Final cleaned file saved  →  {OUTPUT_FILE}")
print(f"   Rows: {df.shape[0]}   |   Columns: {df.shape[1]}")
print("=" * 65)
