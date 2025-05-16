import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook

import os
from tkinter import Tk, filedialog
import datetime

# =====================
# STEP 1: Load customer BOM Excel
# =====================
def load_customer_bom():
    """Opens a file dialog to let user select BOM Excel file, loads all sheets as dict of DataFrames."""
    Tk().withdraw()  # hide tkinter window
    customer_bom_path = filedialog.askopenfilename(title="Select Customer BoM file")
    all_sheets = pd.read_excel(customer_bom_path, sheet_name=None)
    return all_sheets

# =====================
# STEP 2: Check required sheets
# =====================
def validate_required_sheets(all_sheets, required_sheets):
    """Checks that required sheets (BOM, MFG) exist in the Excel file."""
    missing = [s for s in required_sheets if s not in all_sheets]
    if missing:
        raise ValueError(f"❌ Missing required sheet(s): {', '.join(missing)}")
    return all_sheets['BOM'], all_sheets['MFG']

# =====================
# STEP 3: Load company Excel template
# =====================
def load_template(path):
    """Loads company template Excel as DataFrame."""
    try:
        template_df = pd.read_excel(path)
        print("✅ Template loaded successfully.")
        return template_df
    except Exception as e:
        raise FileNotFoundError(f"❌ Failed to load template: {e}")

# =====================
# STEP 4: Validate required columns in DataFrames
# =====================
def validate_required_columns(df, name, required_cols):
    """
    Validates that required columns exist.
    Cleans the columns: strips whitespace + uppercases (for consistency).
    Checks for empty values in required columns and fails if found.
    """
    # Check that required columns exist
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"❌ {name} is missing required columns: {', '.join(missing_cols)}")

    # Clean the required columns (strip spaces + uppercase for consistency)
    for col in required_cols:
        df[col] = df[col].astype(str).str.strip().str.upper()

    # Check for empty values in required columns
    empty_rows = df[df[required_cols].isin(['']).any(axis=1)]
    if not empty_rows.empty:
        error_details = empty_rows[required_cols].reset_index()
        print(f"❌ Found rows in {name} with missing data:\n", error_details)
        raise ValueError(
            f"❌ {name} contains {len(empty_rows)} row(s) with missing data in columns: {', '.join(required_cols)}.\n"
            f"Please correct the data before proceeding."
        )
    else:
        print(f"✅ All rows in {name} have required columns filled.")


# =====================
# STEP 5: Merge BOM and MFG DataFrames
# =====================
def merge_bom_mfg(bom_df, mfg_df, bom_keys, mfg_keys):
    """Performs left join between BOM and MFG DataFrames."""
    merged = bom_df.merge(
        mfg_df,
        left_on=bom_keys,
        right_on=mfg_keys,
        how='left',
        suffixes=('', '_MFG') # suffixes=('', '_MFG'): if any columns (like DESCRIPTION) exist in both, the MFG version gets _MFG added so nothing is overwritten.

    )
    # Drop duplicate merge columns except 'ORIGINAL'
    merged.drop(columns=[c for c in merged.columns if c.upper().startswith('ORIGINAL') and c != 'ORIGINAL'], inplace=True)
    return merged
################
# If multiple MFG rows match same BOM row → merge duplicates BOM row once per match.
# If BOM or MFG has duplicate keys → may duplicate merged rows.
################



# =====================
# STEP 6: Map customer columns to company template
# =====================
def map_columns_to_template(combined_df, column_mapping):
    """Renames columns from customer BOM to company internal template columns; fills missing values with 'N/A'."""
    template_df = combined_df.rename(columns=column_mapping) # rename columns to match template
    template_df = template_df.reindex(columns=column_mapping.values()) # reindex to match template order
    template_df.fillna('', inplace=True) # fill missing values with empty string
    return template_df
# =====================
# STEP 7: Write filled Excel template with formatting
# =====================

def write_filled_template(template_df, bom_df, company_template_path, output_path, columns_to_extract):
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Alignment, Font
    from openpyxl.utils import get_column_letter

    def write_dataframe_to_sheet(sheet, df, bold=True, center_headers=True):
        header_font = Font(name='Arial', size=10, bold=bold)
        regular_font = Font(name='Arial', size=10)

        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = sheet.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center' if center_headers else 'left', vertical='center')

        for row_idx, row in enumerate(df.itertuples(index=False), start=2):
            for col_idx, value in enumerate(row, start=1):
                cell = sheet.cell(row=row_idx, column=col_idx, value=value)
                cell.font = regular_font
                cell.alignment = Alignment(horizontal='left', vertical='center')

        for i, col in enumerate(df.columns, start=1):
            max_len = max(df[col].astype(str).map(len).max(), len(col))
            sheet.column_dimensions[get_column_letter(i)].width = max_len + 2

        sheet.freeze_panes = sheet['A2']

    workbook = load_workbook(company_template_path)

    # -- Formatted-BOM --
    if 'Formatted-BOM' in workbook.sheetnames:
        workbook.remove(workbook['Formatted-BOM'])
    sheet = workbook.create_sheet('Formatted-BOM')
    write_dataframe_to_sheet(sheet, template_df)

    # Apply special styles to Formatted-BOM
    bold_font = Font(name='Arial', size=10, bold=True)
    regular_font = Font(name='Arial', size=10)
    highlight_fill = PatternFill(start_color="FF92D050", end_color="FF92D050", fill_type="solid")
    center_align = Alignment(horizontal='center', vertical='center')

    level_idx = template_df.columns.get_loc('Level') + 1 if 'Level' in template_df.columns else None
    desc_idx = template_df.columns.get_loc('Description') + 1 if 'Description' in template_df.columns else None
    dwg_idx = template_df.columns.get_loc('Dwg_Item') + 1 if 'Dwg_Item' in template_df.columns else None
    align_cols = ['Level', 'Dwg_Item', 'Customer_Part', 'REV', 'UM', 'Unit_Qty']
    align_indices = [template_df.columns.get_loc(c) + 1 for c in align_cols if c in template_df.columns]

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        level = sheet.cell(row=row[0].row, column=level_idx).value if level_idx else None
        is_lvl_0 = (level == 0)

        for cell in row:
            if is_lvl_0:
                cell.fill = highlight_fill
                cell.font = bold_font
            if cell.col_idx in align_indices or (is_lvl_0 and desc_idx and cell.col_idx == desc_idx):
                cell.alignment = center_align
            if dwg_idx and cell.col_idx == dwg_idx:
                cell.number_format = 'General' if is_lvl_0 else '000'
            if level_idx and cell.col_idx == level_idx and level == 0:
                cell.alignment = Alignment(horizontal='left', vertical='center')

        sheet.row_dimensions[row[0].row].height = 13

    # -- BOM-Extract Sheet --
    if columns_to_extract:
        if 'BOM-Extract' in workbook.sheetnames:
            workbook.remove(workbook['BOM-Extract'])
        extract_df = bom_df[columns_to_extract]
        extract_sheet = workbook.create_sheet('BOM-Extract')
        write_dataframe_to_sheet(extract_sheet, extract_df)

        # Highlight ITEM_NUMBER == 0
        part_idx = extract_df.columns.get_loc('PART_NUMBER') + 1 if 'PART_NUMBER' in extract_df.columns else None
        item_idx = extract_df.columns.get_loc('ITEM_NUMBER') + 1 if 'ITEM_NUMBER' in extract_df.columns else None
        blue_fill = PatternFill(start_color="FF00B0F0", end_color="FF00B0F0", fill_type="solid")

        for row in extract_sheet.iter_rows(min_row=2, max_row=extract_sheet.max_row):
            item_value = extract_sheet.cell(row=row[0].row, column=item_idx).value if item_idx else None
            if item_value == 0 and part_idx:
                extract_sheet.cell(row=row[0].row, column=part_idx).fill = blue_fill

        print(f"✅ Extracted columns {columns_to_extract} into 'BOM-Extract' with highlight.")

    # -- Raw-BOM Sheet --
    if 'Raw-BOM' in workbook.sheetnames:
        workbook.remove(workbook['Raw-BOM'])
    raw_sheet = workbook.create_sheet('Raw-BOM')
    write_dataframe_to_sheet(raw_sheet, bom_df)

    print("✅ Raw BOM sheet added.")

        # === Set desired sheet order ===
    desired_order = ["QUOTE", "Formatted-BOM", "Raw-BOM", "Master File BOM", "BOM-Extract"]

    # Ensure only existing sheets are included and preserve their order
    ordered_sheets = [sheet for name in desired_order for sheet in workbook.worksheets if sheet.title == name]
    # Add any remaining sheets not listed
    unordered_sheets = [s for s in workbook.worksheets if s not in ordered_sheets]

    # Apply final order
    workbook._sheets = ordered_sheets + unordered_sheets


    workbook.save(output_path)


        

# =====================
# STEP 8: Open output file (Windows only)
# =====================
def open_output_file(output_path):
    """Opens the completed Excel file using system default app (works on Windows)."""
    if os.path.exists(output_path):
        os.startfile(output_path)
    else:
        print("❌ Completed file not saved.")

# =====================
# STEP 9: Generate console summary report
# =====================
def generate_summary_report(bom_df, combined_df):
    """Prints summary report: counts, missing MPN/Mfr, etc."""
    ###############
    # Only keep rows that actually have a part number filled in (ignoring blanks, NaN, or just spaces).

    mask = combined_df['PART_NUMBER'].notna() & (combined_df['PART_NUMBER'].astype(str).str.strip() != '')
    completed_df = combined_df[mask]

    total_rows_bom = len(bom_df)
    total_rows_output = len(completed_df)

    missing_mpn = completed_df[completed_df['MFG_PART_NUM'].isna() | (completed_df['MFG_PART_NUM'].astype(str).str.strip() == '')]
    missing_mfr = completed_df[completed_df['MANUFACTURER_NAME'].isna() | (completed_df['MANUFACTURER_NAME'].astype(str).str.strip() == '')]
    missing_both = completed_df[
        (completed_df['MFG_PART_NUM'].isna() | (completed_df['MFG_PART_NUM'].astype(str).str.strip() == '')) &
        (completed_df['MANUFACTURER_NAME'].isna() | (completed_df['MANUFACTURER_NAME'].astype(str).str.strip() == ''))
    ]

    print("\n🔎 Summary Report:")
    print(f"📄 Total rows in BOM: {total_rows_bom}")
    print(f"📦 Total rows in output: {total_rows_output}")
    print(f"✔️ Total valid parts: {len(completed_df)}")
    print(f"⚠️ Missing MPN: {len(missing_mpn)}")
    print(missing_mpn['PART_NUMBER'].tolist())
    print(f"⚠️ Missing Mfr: {len(missing_mfr)}")
    print(missing_mfr['PART_NUMBER'].tolist())
    print(f"🚨 Missing BOTH MPN and Mfr: {len(missing_both)}")
    print(missing_both['PART_NUMBER'].tolist())

# =====================
# MAIN WORKFLOW
# =====================
def main():
    """Main execution flow: loads data, validates, merges, writes, reports."""
    required_sheets = ['BOM', 'MFG']
    company_template_path = "Renew_Template.xlsx"  # <<< customize path if needed
    output_path = "Completed_Template.xlsx"        # <<< customize output path if needed

    # Mapping customer column names → internal template column names
    column_mapping = {
        'LEVEL': 'Level',
        'ITEM_NUMBER': 'Dwg_Item',
        'PART_NUMBER': 'Customer_Part',
        'REVISION': 'REV',
        'DESCRIPTION': 'Description',
        'MANUFACTURER_NAME': 'Mfr',
        'MFG_PART_NUM': 'MPN',
        'UOM': 'UM',
        'QUANTITY': 'Unit_Qty',
    }

    # --- Load customer BOM Excel ---
    all_sheets = load_customer_bom()

    # --- Validate required sheets exist ---
    bom_df, mfg_df = validate_required_sheets(all_sheets, required_sheets)

    # --- Load company template Excel ---
    template_df = load_template(company_template_path)

    # --- Validate input DataFrames are not empty ---
    if bom_df.empty:
        raise ValueError("❌ BOM sheet is empty.")
    # if template_df.empty:
    #     raise ValueError("❌ Company template is empty.")

    # --- Optional: add extra blank rows if BOM longer than template ---
    if len(bom_df) > len(template_df):
        extra_rows = len(bom_df) - len(template_df)
        empty_rows = pd.DataFrame('', index=range(extra_rows), columns=template_df.columns)
        template_df = pd.concat([template_df, empty_rows], ignore_index=True)
        print(f"➕ Added {extra_rows} empty row(s) to template.")
    else:
        print("✅ Template has enough rows.")

    # --- Determine merge keys dynamically ---
    if bom_df['ORIGINAL'].nunique() == 1:
        print("Only one unique 'ORIGINAL' in BOM → merging on 'PART_NUMBER' only.")
        bom_key_cols = ['PART_NUMBER']
        mfg_key_cols = ['PART_NUMBER']
    else:
        print(f"Found {bom_df['ORIGINAL'].nunique()} unique 'ORIGINAL' values in BOM → merging on ['ORIGINAL', 'PART_NUMBER'].")
        bom_key_cols = ['ORIGINAL', 'PART_NUMBER']
        mfg_key_cols = ['ORIGINAL', 'PART_NUMBER']

    # --- Validate required columns exist in input ---
    validate_required_columns(bom_df, 'BOM', bom_key_cols)
    validate_required_columns(mfg_df, 'MFG', mfg_key_cols)

    # --- Merge BOM + MFG data ---
    combined_df = merge_bom_mfg(bom_df, mfg_df, bom_key_cols, mfg_key_cols)

    # --- Map customer columns to company template format ---
    template_filled_df = map_columns_to_template(combined_df, column_mapping)

    # --- Write filled Excel template + format ---
    columns_to_extract = ['QUANTITY', 'CRITICAL_PART', 'PART_NUMBER', 'ITEM_NUMBER']
    write_filled_template(template_filled_df,bom_df, company_template_path, output_path, columns_to_extract)

    # --- Open completed file ---
    open_output_file(output_path)

    print("✅ Done: Template filled successfully!")

    # --- Generate console summary report ---
    generate_summary_report(bom_df, combined_df)



#### ==================== Run in streamlit app ====================
def main_process(input_file_path):
    """Used for Streamlit — takes uploaded Excel file path, processes, and returns output path."""
    required_sheets = ['BOM', 'MFG']
    company_template_path = "Renew_Template.xlsx"
    output_path = "Completed_Template.xlsx"

    # Mapping customer column names → internal template column names
    column_mapping = {
        'LEVEL': 'Level',
        'ITEM_NUMBER': 'Dwg_Item',
        'PART_NUMBER': 'Customer_Part',
        'REVISION': 'REV',
        'DESCRIPTION': 'Description',
        'MANUFACTURER_NAME': 'Mfr',
        'MFG_PART_NUM': 'MPN',
        'UOM': 'UM',
        'QUANTITY': 'Unit_Qty',
    }

    # --- Load Excel from Streamlit-uploaded file ---
    all_sheets = pd.read_excel(input_file_path, sheet_name=None)

    bom_df, mfg_df = validate_required_sheets(all_sheets, required_sheets)
    template_df = load_template(company_template_path)

    if bom_df.empty:
        raise ValueError("❌ BOM sheet is empty.")

    # Ensure enough rows in template
    if len(bom_df) > len(template_df):
        extra_rows = len(bom_df) - len(template_df)
        empty_rows = pd.DataFrame('', index=range(extra_rows), columns=template_df.columns)
        template_df = pd.concat([template_df, empty_rows], ignore_index=True)

    # Determine merge keys
    if bom_df['ORIGINAL'].nunique() == 1:
        bom_key_cols = ['PART_NUMBER']
        mfg_key_cols = ['PART_NUMBER']
    else:
        bom_key_cols = ['ORIGINAL', 'PART_NUMBER']
        mfg_key_cols = ['ORIGINAL', 'PART_NUMBER']

    validate_required_columns(bom_df, 'BOM', bom_key_cols)
    validate_required_columns(mfg_df, 'MFG', mfg_key_cols)

    combined_df = merge_bom_mfg(bom_df, mfg_df, bom_key_cols, mfg_key_cols)
    template_filled_df = map_columns_to_template(combined_df, column_mapping)

    columns_to_extract = ['QUANTITY', 'CRITICAL_PART', 'PART_NUMBER', 'ITEM_NUMBER']
    write_filled_template(template_filled_df, bom_df, company_template_path, output_path, columns_to_extract)

    return output_path



if __name__ == "__main__":
    main()
