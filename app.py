import streamlit as st
import subprocess
import sys
import os
import shutil

# Constants
TEMP_DIR = "temp"
EXCEL_PATH = os.path.join(TEMP_DIR, "spreadsheet_input.xlsx")
TEMPLATE_PATH = os.path.join(TEMP_DIR, "template.docx")
OUTPUT_DOC = "filled_agreement.docx"

# UI
st.set_page_config(page_title="JV Agreement Automation Tool", page_icon="🧾")
st.title("🧾 JV Agreement Automation Tool")
st.write("Please upload your Excel and Word template files.")

# Ensure temp dir exists
os.makedirs(TEMP_DIR, exist_ok=True)

# File uploads
xlsx = st.file_uploader("Upload Excel (.xlsx)", type="xlsx")
docx = st.file_uploader("Upload Word Template (.docx)", type="docx")

# Save uploaded files if provided
if xlsx:
    with open(EXCEL_PATH, "wb") as f:
        f.write(xlsx.getbuffer())
    st.success("✅ Excel uploaded and saved.")

if docx:
    with open(TEMPLATE_PATH, "wb") as f:
        f.write(docx.getbuffer())
    st.success("✅ Word template uploaded and saved.")

# Trigger processing
if xlsx and docx and st.button("Generate JV Agreement"):
    subprocess.run([sys.executable, "run_all.py"], check=False)
    if os.path.exists(OUTPUT_DOC):
        with open(OUTPUT_DOC, "rb") as f:
            st.download_button("📥 Download Agreement", f, file_name="JV_Agreement_Final.docx")
    else:
        st.error("❌ Agreement generation failed. Please check your files and try again.")
