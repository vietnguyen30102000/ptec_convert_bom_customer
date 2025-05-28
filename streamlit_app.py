import streamlit as st
import pandas as pd
import tempfile
import os
from pathlib import Path
from convertexcel import main_process  # Replace with your actual script path

st.set_page_config(page_title="BOM Converter", layout="centered")
st.title("📄 BOM Converter Tool")

st.markdown("""
Upload your **BOM Excel file** with the required sheets: **'BOM'** and **'MFG'**.

✅ The tool will:
- Validate and merge BOM and MFG data
- Apply your company Excel template
- Highlight and format the sheet
- Export a new downloadable file
""")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.warning("⚠️ Uploaded files are temporary and auto-deleted after processing.")
    
    if st.button("🚀 Run Conversion"):
        with st.spinner("Processing BOM..."):
            try:
                # Save uploaded file to a temp path
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(uploaded_file.read())
                    input_path = tmp.name

                # Run your main conversion logic
                output_path = main_process(input_path)

                # Load for download
                with open(output_path, "rb") as f:
                    st.success("✅ Conversion complete!")
                    st.download_button(
                        label="📥 Download Processed File",
                        data=f,
                        file_name="Completed_Template.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # ✅ Auto-cleanup
                os.remove(input_path)
                os.remove(output_path)

            except Exception as e:
                st.error(f"❌ An error occurred:\n\n{e}")
else:
    st.info("⬆️ Please upload your file to get started.")
