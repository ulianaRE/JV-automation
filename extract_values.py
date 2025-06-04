import pandas as pd
import json

# === File Path & Sheet Name ===
XLSX_PATH = "spreadsheet_input.xlsx"
SHEET_NAME = "1 PML"

values = {}

def find_row_index(df, row_label):
    for i in range(df.shape[0]):
        if str(df.iloc[i, 0]).strip() == row_label:
            return i
    return None

def find_col_index(df, col_label):
    for j in range(df.shape[1]):
        for i in range(df.shape[0]):
            if str(df.iloc[i, j]).strip() == col_label:
                return j
    return None

def extract_cross_value(df, row_label, col_label, key_name):
    row_idx = find_row_index(df, row_label)
    col_idx = find_col_index(df, col_label)
    if row_idx is not None and col_idx is not None:
        value = str(df.iat[row_idx, col_idx]).strip()
        print(f"{key_name} = {value}")
        values[key_name] = value
    else:
        print(f"{key_name} = NOT FOUND")
        values[key_name] = None

def extract_adjacent_value(df, label, key_name):
    for row in range(df.shape[0]):
        for col in range(df.shape[1] - 1):
            if str(df.iat[row, col]).strip() == label:
                value = str(df.iat[row, col + 1]).strip()
                print(f"{key_name} = {value}")
                values[key_name] = value
                return
    print(f"{key_name} = NOT FOUND")
    values[key_name] = None

def main():
    df = pd.read_excel(XLSX_PATH, sheet_name=SHEET_NAME, header=None)

    extract_adjacent_value(df, "Property:", "property_address")
    extract_cross_value(df, "Entity or Name", "Funding Partner #1", "funding_partner1_entity")
    extract_cross_value(df, "Address", "Funding Partner #1", "funding_partner1_address")
    extract_cross_value(df, "Phone #", "Funding Partner #1", "funding_partner1_phone")
    extract_cross_value(df, "Email", "Funding Partner #1", "funding_partner1_email")
    extract_adjacent_value(df, "COE", "coe_date")
    extract_cross_value(df, "Entity or Name", "Title Company", "title_company_entity")
    extract_cross_value(df, "Phone #", "Title Company", "title_company_phone")
    extract_cross_value(df, "Name", "Title Company", "title_company_name")
    extract_cross_value(df, "Email", "Title Company", "title_company_email")
    extract_cross_value(df, "Funding Amount", "Owner Partner", "owner_partner_funding")
    extract_cross_value(df, "Funding Amount", "Funding Partner #1", "funding_partner1_funding")
    extract_cross_value(df, "ROI", "Funding Partner #1", "funding_partner1_ROI")
    extract_adjacent_value(df, "Maturity Date", "maturity_date")
    extract_adjacent_value(df, "Grace Period", "grace_period")
    extract_cross_value(df, "Extension Fee (per month)", "Funding Partner #1", "funding_partner1_late_fee")
    extract_cross_value(df, "Extension Fee (per month)", "Owner Partner", "owner_partner_late_fee")
    extract_adjacent_value(df, "Funds released at COE", "funds_released_at_COE")

    # Save and print results
    with open("extracted_values.json", "w") as f:
        json.dump(values, f, indent=2)

    print("\n📦 Extracted Values Dictionary:")
    for k, v in values.items():
        print(f"  {k}: {v}")

if __name__ == "__main__":
    main()

