import streamlit as st
import subprocess
import sys
import os
import shutil

# Setup
st.set_page_config(page_title="JV Agreement Automation Tool", page_icon="🧾")
st.title("🧾 JV Agreement Automation Tool")
st.write("Upload your Excel and Word template files to generate a custom JV Agreement.")

# File uploader
uploaded_excel = st.file_uploader("Upload Excel (.xlsx)", type="xlsx")
uploaded_docx = st.file_uploader("Upload JV Agreement Template (.docx)", type="docx")

TEMP_DIR = "temp"
OUTPUT_DOC = os.path.join(TEMP_DIR, "filled_agreement.docx")
os.makedirs(TEMP_DIR, exist_ok=True)

if uploaded_excel and uploaded_docx:
    excel_path = os.path.join(TEMP_DIR, uploaded_excel.name)
    docx_path = os.path.join(TEMP_DIR, uploaded_docx.name)

    # Save uploaded files
    with open(excel_path, "wb") as f:
        f.write(uploaded_excel.getbuffer())
    with open(docx_path, "wb") as f:
        f.write(uploaded_docx.getbuffer())

    st.success("✅ Files uploaded!")

    if st.button("Generate JV Agreement"):
        try:
            # Set expected filenames for run_all.py
            os.rename(excel_path, "input.xlsx")
            os.rename(docx_path, "template.docx")

            # Run the automation pipeline
            result = subprocess.run([sys.executable, "run_all.py"], check=True, capture_output=True, text=True)
            st.success("🎉 Agreement generated successfully!")

            # Move output to temp for download
            shutil.move("filled_agreement.docx", OUTPUT_DOC)
            with open(OUTPUT_DOC, "rb") as file:
                st.download_button("📥 Download Agreement", file, file_name="JV_Agreement_Final.docx")
        except subprocess.CalledProcessError as e:
            st.error("❌ Error during generation.")
            st.text(e.stdout)
            st.text(e.stderr)
else:
    st.info("📤 Please upload both the Excel and Word files to begin.")
