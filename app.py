import streamlit as st
import subprocess
import sys
import os

# === CONSTANTS ===
TEMP_DIR = "temp"
INPUT_EXCEL = os.path.join(TEMP_DIR, "spreadsheet_input.xlsx")
TEMPLATE_DOCX = os.path.join(TEMP_DIR, "template.docx")
OUTPUT_FILENAME = "filled_agreement.docx"
OUTPUT_DOC = os.path.join(TEMP_DIR, OUTPUT_FILENAME)

# === UI CONFIG ===
st.set_page_config(page_title="JV Agreement Automation Tool", page_icon="🧾")
st.title("🧾 JV Agreement Automation Tool")
st.write("Hi Marcia! Let's run it! Please upload your files.")

# === FILE UPLOADS ===
uploaded_excel = st.file_uploader("Upload Excel (.xlsx)", type="xlsx")
uploaded_docx = st.file_uploader("Upload JV Agreement Template (.docx)", type="docx")

os.makedirs(TEMP_DIR, exist_ok=True)

def save_uploaded_file(uploaded_file, path):
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

if uploaded_excel and uploaded_docx:
    save_uploaded_file(uploaded_excel, INPUT_EXCEL)
    save_uploaded_file(uploaded_docx, TEMPLATE_DOCX)
    st.success("✅ Files uploaded!")

    if st.button("Generate JV Agreement"):
        try:
            result = subprocess.run([sys.executable, "run_all.py"], check=True, capture_output=True, text=True)

            if os.path.exists(OUTPUT_FILENAME):
                os.replace(OUTPUT_FILENAME, OUTPUT_DOC)
                with open(OUTPUT_DOC, "rb") as file:
                    st.download_button("📥 Download Agreement", file, file_name="JV_Agreement_Final.docx")
            else:
                st.error("❌ The agreement was not generated.")
                st.code(result.stdout)
                st.code(result.stderr)

        except subprocess.CalledProcessError as e:
            st.error("❌ An error occurred during document generation.")
            st.code(e.stdout)
            st.code(e.stderr)
else:
    st.info("📤 Please upload both the Excel and Word files to proceed.")
