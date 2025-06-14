import streamlit as st
import subprocess
import sys
import os
import shutil

# === CONSTANTS ===
INPUT_EXCEL = "spreadsheet_input.xlsx"
TEMPLATE_DOCX = "template.docx"
OUTPUT_FILENAME = "filled_agreement.docx"
LOG_FILE = "run_all.log"
TEMP_DIR = "temp"
OUTPUT_DOC = os.path.join(TEMP_DIR, OUTPUT_FILENAME)

# === UI CONFIG ===
st.set_page_config(page_title="JV Agreement Automation Tool", page_icon="🧾")
st.title("🧾 JV Agreement Automation Tool")
st.write("Hi Marcia! Let's run it! Please upload your files.")

os.makedirs(TEMP_DIR, exist_ok=True)

# === Reset state if new files uploaded ===
def reset_on_upload(file_key):
    if uploaded := st.session_state.get(file_key):
        st.session_state.generated = False

# === FILE UPLOADS ===
st.file_uploader("Upload Excel (.xlsx)", type="xlsx", key="excel", on_change=lambda: reset_on_upload("excel"))
st.file_uploader("Upload JV Agreement Template (.docx)", type="docx", key="docx", on_change=lambda: reset_on_upload("docx"))

uploaded_excel = st.session_state.get("excel")
uploaded_docx = st.session_state.get("docx")

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

            # Run backend pipeline
            result = subprocess.run([sys.executable, "run_all.py"], check=True, capture_output=True, text=True)

            if os.path.exists(OUTPUT_FILENAME):
                shutil.move(OUTPUT_FILENAME, OUTPUT_DOC)
                st.session_state.generated = True
            else:
                st.error("❌ The agreement was not generated.")
                st.text(result.stdout)
                st.text(result.stderr)

        except subprocess.CalledProcessError as e:
            st.error("❌ An error occurred during document generation.")
            st.text(e.stdout)
            st.text(e.stderr)

# === Show Download Buttons if files were generated ===
if st.session_state.get("generated"):
    if os.path.exists(OUTPUT_DOC):
        with open(OUTPUT_DOC, "rb") as f:
            st.download_button("📥 Download Agreement", f, file_name="JV_Agreement_Final.docx")

    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, "rb") as f:
            st.download_button("📝 Download Log File", f, file_name="run_all.log")
else:
    if not uploaded_excel or not uploaded_docx:
        st.info("📤 Please upload both the Excel and Word files to proceed.")
