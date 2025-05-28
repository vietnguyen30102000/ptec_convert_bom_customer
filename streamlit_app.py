import streamlit as st
import tempfile
import os
from convertexcel import main_process  # Make sure this is the correct import path

st.set_page_config(page_title="BOM Converter", layout="centered")

st.title("📄 BOM Converter Tool")

st.markdown("""
Upload your **BOM Excel file** containing the required sheets: **'BOM'** and **'MFG'**.

The tool will:
- Merge and validate your BOM and MFG data
- Format with your company template
- Generate a downloadable result
""")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 Run BOM Conversion"):
        with st.spinner("Processing... Please wait."):
            try:
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
                    tmp_input.write(uploaded_file.read())
                    input_path = tmp_input.name

                # Call processing logic (returns output path)
                output_path = main_process(input_path)

                # Read and offer download
                with open(output_path, "rb") as f:
                    st.success("✅ Conversion complete!")
                    st.download_button(
                        label="📥 Download Converted File",
                        data=f,
                        file_name="Completed_Template.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # Auto-delete temp files after download button rendered
                os.remove(input_path)
                os.remove(output_path)

            except Exception as e:
                st.error(f"❌ An error occurred:\n\n{e}")
else:
    st.info("⬆️ Upload a file to get started.")
