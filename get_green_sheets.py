from openpyxl import load_workbook

def get_green_sheets(file_path):
    wb = load_workbook(file_path)
    green_sheets = []
    for sheet in wb.worksheets:
        tab_color = sheet.sheet_properties.tabColor
        if tab_color and tab_color.rgb and tab_color.rgb.upper().startswith("FF00FF00"):  # bright green
            green_sheets.append(sheet.title)
    return green_sheets

# Manual test mode
#if __name__ == "__main__":
#    test_file = "spreadsheet_input.xlsx"  # Replace with actual filename
#    sheets = get_green_sheets(test_file)
#    print("Green sheets:", sheets)