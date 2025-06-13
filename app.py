import streamlit as st
import subprocess
import sys
import os
import tempfile

# === UI CONFIG ===
st.set_page_config(page_title="JV Agreement Automation Tool", page_icon="🧾")
st.title("🧾 JV Agreement Automation Tool")
st.write("Hi Marcia! Let's run it! Please upload your files.")

# === FILE UPLOADS ===
uploaded_excel = st.file_uploader("Upload Excel (.xlsx)", type="xlsx")
uploaded_docx = st.file_uploader("Upload JV Agreement Template (.docx)", type="docx")

if uploaded_excel and uploaded_docx:
    st.success("✅ Files uploaded!")

    if st.button("Generate JV Agreement"):
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save uploaded files to temp directory
            excel_path = os.path.join(temp_dir, "input.xlsx")
            docx_path = os.path.join(temp_dir, "template.docx")
            output_path = os.path.join(temp_dir, "filled_agreement.docx")

            with open(excel_path, "wb") as f:
                f.write(uploaded_excel.getbuffer())
            with open(docx_path, "wb") as f:
                f.write(uploaded_docx.getbuffer())

            # Run the backend script
            try:
                result = subprocess.run(
                    [sys.executable, "run_all.py", excel_path, docx_path, output_path],
                    capture_output=True,
                    text=True,
                    check=True
                )

                if os.path.exists(output_path):
                    with open(output_path, "rb") as file:
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
