import streamlit as st
import subprocess
import sys
import os
import shutil
from get_green_sheets import get_green_sheets

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

# === Reset State on Upload ===
def reset_on_upload(file_key):
    if uploaded := st.session_state.get(file_key):
        st.session_state.generated = False
        st.session_state.green_sheets = None
        st.session_state.selected_sheet = None
        st.session_state.ready_to_generate = False

# === FILE UPLOADS ===
st.file_uploader("Upload JV Agreement Template (.docx)", type="docx", key="docx", on_change=lambda: reset_on_upload("docx"))
st.file_uploader("Upload Excel (.xlsx)", type="xlsx", key="excel", on_change=lambda: reset_on_upload("excel"))

uploaded_docx = st.session_state.get("docx")
uploaded_excel = st.session_state.get("excel")

if uploaded_excel and uploaded_docx:
    # Save files
    excel_path = os.path.join(TEMP_DIR, uploaded_excel.name)
    docx_path = os.path.join(TEMP_DIR, uploaded_docx.name)

    with open(docx_path, "wb") as f:
        f.write(uploaded_docx.getbuffer())
    with open(excel_path, "wb") as f:
        f.write(uploaded_excel.getbuffer())

    st.success("✅ Files uploaded!")

    # Extract green sheets
    if "green_sheets" not in st.session_state or st.session_state.green_sheets is None:
        with st.spinner("🔍 Extracting green-labeled sheets..."):
            st.session_state.green_sheets = get_green_sheets(excel_path)

    if st.session_state.green_sheets:
        selected = st.selectbox(
            "📗 Choose a green-labeled sheet to process:",
            options=["-- Select a sheet --"] + st.session_state.green_sheets,
            index=0
        )
        if selected != "-- Select a sheet --":
            st.session_state.selected_sheet = selected
            st.session_state.ready_to_generate = True
        else:
            st.session_state.ready_to_generate = False
    else:
        st.warning("⚠️ No green-labeled sheets found in the uploaded Excel file.")
        st.stop()

# === Generate Button ===
if st.session_state.get("ready_to_generate") and not st.session_state.get("generated"):
    if st.button("🚀 Generate JV Agreement"):
        with st.spinner("🛠️ Generating your JV Agreement... please wait..."):
            try:
                shutil.move(docx_path, TEMPLATE_DOCX)
                shutil.move(excel_path, INPUT_EXCEL)

                selected_sheet = st.session_state.selected_sheet
                result = subprocess.run(
                    [sys.executable, "run_all.py", selected_sheet],
                    check=True,
                    capture_output=True,
                    text=True
                )

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

# === DOWNLOAD ZONE ===
if st.session_state.get("generated"):
    col1, col2 = st.columns([2, 1])
    with col1:
        if os.path.exists(OUTPUT_DOC):
            with open(OUTPUT_DOC, "rb") as f:
                st.download_button("📥 Download Agreement", f, file_name="JV_Agreement_Final.docx")
    with col2:
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, "rb") as f:
                st.download_button("📝 Log", f, file_name="run_all.log")
