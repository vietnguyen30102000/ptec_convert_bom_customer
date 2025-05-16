import streamlit as st
import pandas as pd
import tempfile
from pathlib import Path
from convertexcel import main_process  # Change to your script's filename without .py

st.set_page_config(page_title="BOM Converter", layout="centered")

st.title("📄 BOM Converter Tool")
st.markdown("""
Upload your BOM Excel file. It must include **'BOM'** and **'MFG'** sheets.

The tool will:
- Validate and merge BOM and MFG
- Apply your company template
- Highlight & style the result
- Return a completed Excel file
""")

uploaded_file = st.file_uploader("📁 Upload BOM Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 Run BOM Conversion"):
        with st.spinner("Processing... Please wait."):
            try:
                # Save uploaded file to temp path
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(uploaded_file.read())
                    temp_input_path = tmp.name

                # Call your main process function
                output_path = main_process(temp_input_path)

                # Let user download the result
                with open(output_path, "rb") as out_file:
                    st.success("✅ Conversion complete! Download your result below:")
                    st.download_button(
                        label="📥 Download Completed Template",
                        data=out_file,
                        file_name="Completed_Template.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"❌ An error occurred:\n\n{str(e)}")
else:
    st.info("⬆️ Upload a file to begin.")
