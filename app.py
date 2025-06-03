import streamlit as st
import subprocess
import os
import json
from pathlib import Path

st.title("📄 JV Agreement Automation")

# Upload inputs
docx_file = st.file_uploader("Upload the Word document (.docx)", type=["docx"])
xlsx_file = st.file_uploader("Upload the Excel spreadsheet (.xlsx)", type=["xlsx"])

if docx_file and xlsx_file:
    # Save uploaded files locally
    docx_path = Path("JV Agreement - 1 PML.docx")
    xlsx_path = Path("Project Info Sheet for AI.xlsx")
    json_path = Path("extracted_values.json")

    with open(docx_path, "wb") as f:
        f.write(docx_file.read())
    with open(xlsx_path, "wb") as f:
        f.write(xlsx_file.read())

    st.write("✅ Files uploaded successfully.")

    # Run the extraction script
    extract_script = "extract_values_from_1pml_spreadsheet_to_json.py"
    st.write("🔄 Extracting data from spreadsheet...")

    try:
        result = subprocess.run(
            ["python3", extract_script, str(xlsx_path)],
            cwd=".",
            check=True,
            capture_output=True,
            text=True
        )
        st.success("✅ Data extracted from Excel and written to JSON.")
        st.code(result.stdout, language="bash")
    except subprocess.CalledProcessError as e:
        st.error("❌ Error running spreadsheet extraction script.")
        st.code(e.stderr or "No stderr output", language="bash")
        st.stop()

    # Run all filling scripts
    st.write("🛠️ Running document filler scripts...")

    fill_scripts = sorted(Path(".").glob("fill_*.py"))
    for script in fill_scripts:
        if script.name == "app.py":
            continue
        st.write(f"▶️ Running `{script.name}`...")
        try:
            result = subprocess.run(
                ["python3", script.name],
                cwd=".",
                check=True,
                capture_output=True,
                text=True
            )
            st.code(result.stdout, language="bash")
        except subprocess.CalledProcessError as e:
            st.error(f"❌ Error in `{script.name}`")
            st.code(e.stderr or "No stderr output", language="bash")

    # Offer final docx for download
    filled_versions = sorted(Path(".").glob("*filled v*.docx"), key=os.path.getmtime, reverse=True)
    if filled_versions:
        latest_file = filled_versions[0]
        with open(latest_file, "rb") as f:
            st.download_button(
                label=f"📥 Download: {latest_file.name}",
                data=f,
                file_name=latest_file.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("⚠️ No filled document found to download.")
