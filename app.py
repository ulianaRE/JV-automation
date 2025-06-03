import streamlit as st
import subprocess
import os
import json
import tempfile

st.title("JV Agreement Filler")
st.write("Upload your input files and click below to generate the filled Word document.")

uploaded_docx = st.file_uploader("Upload DOCX file", type="docx")
uploaded_xlsx = st.file_uploader("Upload XLSX file", type="xlsx")

if uploaded_docx and uploaded_xlsx:
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "JV Agreement - 1 PML.docx")
        xlsx_path = os.path.join(tmpdir, "input.xlsx")

        with open(docx_path, "wb") as f:
            f.write(uploaded_docx.read())

        with open(xlsx_path, "wb") as f:
            f.write(uploaded_xlsx.read())

        # Extract values to JSON
        extract_script = "extract_values_from_spreadsheet_to_json.py"
        subprocess.run(["python3", extract_script, xlsx_path], cwd=".", check=True)

        # Run all fill scripts
        for script in sorted(f for f in os.listdir(".") if f.startswith("fill_docx_v") and f.endswith(".py")):
            subprocess.run(["python3", script], cwd=".", check=True)

        filled_docx_path = "JV Agreement - 1 PML - filled v99.docx"
        if os.path.exists(filled_docx_path):
            with open(filled_docx_path, "rb") as f:
                st.download_button("Download Filled Document", f, file_name="JV Agreement - 1 PML - filled.docx")
        else:
            st.error("Something went wrong — output file not found.")
