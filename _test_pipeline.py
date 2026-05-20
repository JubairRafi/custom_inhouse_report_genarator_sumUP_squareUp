"""
_test_pipeline.py — end-to-end headless test (no Streamlit)
Run: python _test_pipeline.py
"""
import datetime, sys

from data_loader import load_outlet_file
from report_engine import (
    top10_combined,
    top10_single,
    category_totals_combined,
)
from excel_exporter import build_excel

START = datetime.date(2025, 1, 1)
END   = datetime.date(2025, 12, 31)

print("Loading sample files...")
kings_df    = load_outlet_file("sample_kings_road.xlsx")
stmartin_df = load_outlet_file("sample_st_martin.xlsx")
print(f"  Kings Road : {len(kings_df)} rows")
print(f"  St Martin  : {len(stmartin_df)} rows")

print("\nGenerating reports...")
r1, t1 = top10_combined(kings_df, stmartin_df, START, END)
r2, t2 = top10_single(kings_df, "Kings", START, END)
r3, t3 = top10_single(stmartin_df, "St Martin", START, END)

print("\n  Generating Category Totals report...")
rc_cat, tc_cat = category_totals_combined(kings_df, stmartin_df, START, END)
print(f"  [Cat Totals] '{tc_cat}' generated with columns={rc_cat.columns.tolist()}")

# ── Assertions ────────────────────────────────────────────────────────────
errors = []

for label, df, qty_cols in [
    ("Combined", r1, ["Kings Road", "St Martin", "Total"]),
    ("Kings",    r2, ["Qty"]),
    ("St Martin",r3, ["Qty"]),
]:
    n_data = len(df) - 1          # exclude total row
    if n_data < 1 or n_data > 10:
        errors.append(f"{label}: expected 1–10 data rows, got {n_data}")

    # Serial numbers should run 1..n_data
    sls = list(df["#SL"].iloc[:n_data])
    if sls != list(range(1, n_data + 1)):
        errors.append(f"{label}: #SL mismatch — {sls}")

    # Total row should match sum of data rows
    for col in qty_cols:
        data_sum = int(df[col].iloc[:n_data].sum())
        total_val = int(df[col].iloc[-1])
        if data_sum != total_val:
            errors.append(f"{label}/{col}: total row {total_val} != sum {data_sum}")

    print(f"  [{label}] '{t1 if label=='Combined' else (t2 if label=='Kings' else t3)}'")
    print(f"    data rows={n_data}, columns={list(df.columns)}")

    if label == "Combined":
        # Verify Total = Kings Road + St Martin for each data row
        for i, row in df.iloc[:n_data].iterrows():
            if int(row["Kings Road"]) + int(row["St Martin"]) != int(row["Total"]):
                errors.append(f"Combined row {i}: Kings+StMartin != Total")

# Assertions for Category Totals
if tc_cat not in rc_cat.columns:
    errors.append(f"Category Totals: missing '{tc_cat}' column")
if "Kings" not in rc_cat.columns:
    errors.append("Category Totals: missing 'Kings' column")
if "St Martin" not in rc_cat.columns:
    errors.append("Category Totals: missing 'St Martin' column")
if "Total" not in rc_cat.columns:
    errors.append("Category Totals: missing 'Total' column")
if rc_cat.columns[0] != tc_cat:
    errors.append(f"Category Totals: first column expected '{tc_cat}', got '{rc_cat.columns[0]}'")

n_data_cat = len(rc_cat) - 1 # exclude total row
if n_data_cat < 1:
    errors.append(f"Category Totals: expected at least 1 data row, got {n_data_cat}")

# Total row should match sum of data rows for Category Totals
for col in ["Kings", "St Martin", "Total"]:
    data_sum = int(rc_cat[col].iloc[:n_data_cat].sum())
    total_val = int(rc_cat[col].iloc[-1])
    if data_sum != total_val:
        errors.append(f"Category Totals/{col}: total row {total_val} != sum {data_sum}")

# Verify Total = Kings + St Martin for each data row in Category Totals
for i, row in rc_cat.iloc[:n_data_cat].iterrows():
    if int(row["Kings"]) + int(row["St Martin"]) != int(row["Total"]):
        errors.append(f"Category Totals row {i}: Kings+StMartin != Total")

print(f"  [Category Totals] '{tc_cat}'")
print(f"    data rows={n_data_cat}, columns={list(rc_cat.columns)}")


# ── Excel export ──────────────────────────────────────────────────────────
print("\nBuilding Excel export...")
excel_bytes = build_excel(r1, t1, r2, t2, r3, t3)
out_path = "test_output.xlsx"
with open(out_path, "wb") as f:
    f.write(excel_bytes)
print(f"  Written {len(excel_bytes):,} bytes -> {out_path}")

# ── Final result ──────────────────────────────────────────────────────────
if errors:
    print("\nFAILED — errors found:")
    for e in errors:
        print(f"  - {e}")
    sys.exit(1)
else:
    print("\nAll assertions passed. Pipeline is working correctly.")
