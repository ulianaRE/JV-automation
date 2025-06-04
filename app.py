import streamlit as st
import subprocess
import sys
import os
import shutil

# === CONSTANTS ===
INPUT_EXCEL = "spreadsheet_input.xlsx"
TEMPLATE_DOCX = "template.docx"
OUTPUT_FILENAME = "filled_agreement.docx"
TEMP_DIR = "temp"
OUTPUT_DOC = os.path.join(TEMP_DIR, OUTPUT_FILENAME)

# === UI CONFIG ===
st.set_page_config(page_title="JV Agreement Automation Tool", page_icon="🧾")
st.title("🧾 JV Agreement Automation Tool")
st.write("Upload your Excel and Word template files to generate a custom JV Agreement.")

# === FILE UPLOADS ===
uploaded_excel = st.file_uploader("Upload Excel (.xlsx)", type="xlsx")
uploaded_docx = st.file_uploader("Upload JV Agreement Template (.docx)", type="docx")

os.makedirs(TEMP_DIR, exist_ok=True)

if uploaded_excel and uploaded_docx:
    excel_path = os.path.join(TEMP_DIR, uploaded_excel.name)
    docx_path = os.path.join(TEMP_DIR, uploaded_docx.name)

    with open(excel_path, "wb") as f:
        f.write(uploaded_excel.getbuffer())
    with open(docx_path, "wb") as f:
        f.write(uploaded_docx.getbuffer())

    st.success("✅ Files uploaded!")

    if st.button("Generate JV Agreement"):
        try:
            # Move uploaded files to expected names
            shutil.move(excel_path, INPUT_EXCEL)
            shutil.move(docx_path, TEMPLATE_DOCX)

            # Run the backend pipeline
            result = subprocess.run([sys.executable, "run_all.py"], check=True, capture_output=True, text=True)

            if os.path.exists(OUTPUT_FILENAME):
                shutil.move(OUTPUT_FILENAME, OUTPUT_DOC)

                # 📥 DOWNLOAD
                with open(OUTPUT_DOC, "rb") as file:
                    st.download_button("📥 Download Agreement", file, file_name="JV_Agreement_Final.docx")
            else:
                st.error("❌ The agreement was not generated.")
                st.text(result.stdout)
                st.text(result.stderr)

        except subprocess.CalledProcessError as e:
            st.error("❌ An error occurred during document generation.")
            st.text(e.stdout)
            st.text(e.stderr)
else:
    st.info("📤 Please upload both the Excel and Word files to proceed.")
