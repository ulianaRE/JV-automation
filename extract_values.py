import sys
import pandas as pd
import json

# 🔧 Constants
EXCEL_FILE = "spreadsheet_input.xlsx"
SHEET_NAME = "sheetname_input"
OUTPUT_JSON = "extracted_values.json"

# 🔍 Helper functions
def find_row_index(df, row_label):
    """Find index of the row where the first column matches the label"""
    for i, val in enumerate(df.iloc[:, 0]):
        if str(val).strip() == row_label:
            return i
    return None

def find_col_index(df, col_label):
    """Find index of the column where any row contains the label"""
    for col in df.columns:
        if df[col].astype(str).str.strip().eq(col_label).any():
            return df.columns.get_loc(col)
    return None

def extract_cross_value(df, row_label, col_label, key_name):
    """Extracts value from intersection of row and column"""
    row_idx = find_row_index(df, row_label)
    col_idx = find_col_index(df, col_label)
    value = ""
    if row_idx is not None and col_idx is not None:
        cell = df.iloc[row_idx, col_idx]
        value = "" if pd.isna(cell) else str(cell).strip()
    values[key_name] = value
    print(f"{key_name} = {value or 'NOT FOUND'}")

def extract_adjacent_value(df, label, key_name):
    """Extracts value adjacent (right) to the given row label"""
    row_idx = find_row_index(df, label)
    value = ""
    if row_idx is not None:
        if len(df.columns) > 1:
            cell = df.iloc[row_idx, 1]
            value = "" if pd.isna(cell) else str(cell).strip()
    values[key_name] = value
    print(f"{key_name} = {value or 'NOT FOUND'}")

# Load sheet name
if len(sys.argv) < 2:
    raise ValueError("No sheet name provided to extract_values.py")

selected_sheet = sys.argv[1]

# 🧾 Load Excel with only 1 sheet
df = pd.read_excel(EXCEL_FILE, sheet_name=selected_sheet)

# ✨ Print separator for clarity
print("\n")

# 📦 Data extraction
values = {}

# Extracts the address of the property
extract_adjacent_value(df, "Property:", "property_address")
# Funding partner (Party B) details
extract_cross_value(df, "Entity or Name", "Party B", "funding_partner1_entity")
extract_cross_value(df, "Address for JV", "Party B", "funding_partner1_address")
extract_cross_value(df, "Phone #", "Party B", "funding_partner1_phone")
extract_cross_value(df, "Email", "Party B", "funding_partner1_email")
# COE and Title Company details
extract_adjacent_value(df, "COE", "coe_date")
extract_cross_value(df, "Entity or Name", "Title Company", "title_company_entity")
extract_cross_value(df, "Phone #", "Title Company", "title_company_phone")
extract_cross_value(df, "Name", "Title Company", "title_company_name")
extract_cross_value(df, "Email", "Title Company", "title_company_email")
# Funding & ROI
extract_cross_value(df, "Funding Amount", "Party A", "owner_partner_funding")
extract_cross_value(df, "Funding Amount", "Party B", "funding_partner1_funding")
extract_cross_value(df, "ROI", "Party B", "funding_partner1_ROI")
# Maturity and grace dates
extract_adjacent_value(df, "Maturity Date", "maturity_date")
extract_adjacent_value(df, "Grace period date", "grace_period_date")
# Late fees
extract_cross_value(df, "Extension Fee (per month)", "Party B", "funding_partner1_late_fee")
extract_cross_value(df, "Extension Fee (per month)", "Party A", "owner_partner_late_fee")
# Entity state
extract_cross_value(df, "Entity Formation State", "Party B", "funding_partner1_state")

# 💾 Write to JSON
with open(OUTPUT_JSON, "w") as f:
    json.dump(values, f, indent=4)

# ✅ Summary block
print("\nExtracted values summary:")
print(json.dumps(values, indent=4))
