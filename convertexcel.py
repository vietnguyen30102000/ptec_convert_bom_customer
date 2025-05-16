import pandas as pd
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
import os

# =====================
# Load Excel sheets
# =====================
def load_customer_bom_from_path(file_path):
    """Loads all sheets from the Excel file at a given path."""
    all_sheets = pd.read_excel(file_path, sheet_name=None)
    return all_sheets

def validate_required_sheets(all_sheets, required_sheets):
    missing = [s for s in required_sheets if s not in all_sheets]
    if missing:
        raise ValueError(f"❌ Missing required sheet(s): {', '.join(missing)}")
    return all_sheets['BOM'], all_sheets['MFG']

def load_template(path):
    try:
        return pd.read_excel(path)
    except Exception as e:
        raise FileNotFoundError(f"❌ Failed to load template: {e}")

# =====================
# Data validation & merging
# =====================
def validate_required_columns(df, name, required_cols):
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"❌ {name} is missing required columns: {', '.join(missing_cols)}")

    for col in required_cols:
        df[col] = df[col].astype(str).str.strip().str.upper()

    empty_rows = df[df[required_cols].isin(['']).any(axis=1)]
    if not empty_rows.empty:
        raise ValueError(f"❌ {name} contains empty values in: {', '.join(required_cols)}")

def merge_bom_mfg(bom_df, mfg_df, bom_keys, mfg_keys):
    merged = bom_df.merge(
        mfg_df,
        left_on=bom_keys,
        right_on=mfg_keys,
        how='left',
        suffixes=('', '_MFG')
    )
    merged.drop(columns=[c for c in merged.columns if c.upper().startswith('ORIGINAL') and c != 'ORIGINAL'], inplace=True)
    return merged

def map_columns_to_template(combined_df, column_mapping):
    template_df = combined_df.rename(columns=column_mapping)
    template_df = template_df.reindex(columns=column_mapping.values())
    template_df.fillna('', inplace=True)
    return template_df

# =====================
# Excel Writing
# =====================
def write_filled_template(template_df, bom_df, template_path, output_path, columns_to_extract):
    def write_dataframe_to_sheet(sheet, df, bold=True):
        header_font = Font(name='Arial', size=10, bold=bold)
        regular_font = Font(name='Arial', size=10)

        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = sheet.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        for row_idx, row in enumerate(df.itertuples(index=False), start=2):
            for col_idx, value in enumerate(row, start=1):
                cell = sheet.cell(row=row_idx, column=col_idx, value=value)
                cell.font = regular_font
                cell.alignment = Alignment(horizontal='left', vertical='center')

        for i, col in enumerate(df.columns, start=1):
            width = max(df[col].astype(str).map(len).max(), len(col))
            sheet.column_dimensions[get_column_letter(i)].width = width + 2

        sheet.freeze_panes = sheet['A2']

    workbook = load_workbook(template_path)

    # Format 'Formatted-BOM'
    if 'Formatted-BOM' in workbook.sheetnames:
        workbook.remove(workbook['Formatted-BOM'])
    sheet = workbook.create_sheet('Formatted-BOM')
    write_dataframe_to_sheet(sheet, template_df)

    # Style rules
    bold_font = Font(name='Arial', size=10, bold=True)
    highlight_fill = PatternFill(start_color="FF92D050", end_color="FF92D050", fill_type="solid")
    center_align = Alignment(horizontal='center', vertical='center')

    level_idx = template_df.columns.get_loc('Level') + 1 if 'Level' in template_df.columns else None
    desc_idx = template_df.columns.get_loc('Description') + 1 if 'Description' in template_df.columns else None
    dwg_idx = template_df.columns.get_loc('Dwg_Item') + 1 if 'Dwg_Item' in template_df.columns else None
    align_cols = ['Level', 'Dwg_Item', 'Customer_Part', 'REV', 'UM', 'Unit_Qty']
    align_indices = [template_df.columns.get_loc(c) + 1 for c in align_cols if c in template_df.columns]

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        is_lvl_0 = sheet.cell(row=row[0].row, column=level_idx).value == 0 if level_idx else False
        for cell in row:
            if is_lvl_0:
                cell.fill = highlight_fill
                cell.font = bold_font
            if cell.col_idx in align_indices or (is_lvl_0 and desc_idx and cell.col_idx == desc_idx):
                cell.alignment = center_align
            if dwg_idx and cell.col_idx == dwg_idx:
                cell.number_format = 'General' if is_lvl_0 else '000'

        sheet.row_dimensions[row[0].row].height = 13

    # BOM-Extract sheet
    if columns_to_extract:
        if 'BOM-Extract' in workbook.sheetnames:
            workbook.remove(workbook['BOM-Extract'])
        extract_df = bom_df[columns_to_extract]
        extract_sheet = workbook.create_sheet('BOM-Extract')
        write_dataframe_to_sheet(extract_sheet, extract_df)

    # Raw-BOM
    if 'Raw-BOM' in workbook.sheetnames:
        workbook.remove(workbook['Raw-BOM'])
    raw_sheet = workbook.create_sheet('Raw-BOM')
    write_dataframe_to_sheet(raw_sheet, bom_df)

    # Set final sheet order
    desired_order = ["QUOTE", "Formatted-BOM", "Raw-BOM", "Master File BOM", "BOM-Extract"]
    workbook._sheets = [s for name in desired_order for s in workbook.worksheets if s.title == name] + \
                       [s for s in workbook.worksheets if s.title not in desired_order]

    workbook.save(output_path)

# =====================
# Entry Point for Streamlit
# =====================
def main_process(input_file_path):
    required_sheets = ['BOM', 'MFG']
    company_template_path = "Renew_Template.xlsx"
    output_path = "Completed_Template.xlsx"

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

    all_sheets = load_customer_bom_from_path(input_file_path)
    bom_df, mfg_df = validate_required_sheets(all_sheets, required_sheets)
    template_df = load_template(company_template_path)

    if bom_df.empty:
        raise ValueError("❌ BOM sheet is empty.")

    if len(bom_df) > len(template_df):
        extra_rows = len(bom_df) - len(template_df)
        empty_rows = pd.DataFrame('', index=range(extra_rows), columns=template_df.columns)
        template_df = pd.concat([template_df, empty_rows], ignore_index=True)

    if bom_df['ORIGINAL'].nunique() == 1:
        bom_keys = ['PART_NUMBER']
        mfg_keys = ['PART_NUMBER']
    else:
        bom_keys = ['ORIGINAL', 'PART_NUMBER']
        mfg_keys = ['ORIGINAL', 'PART_NUMBER']

    validate_required_columns(bom_df, 'BOM', bom_keys)
    validate_required_columns(mfg_df, 'MFG', mfg_keys)

    combined_df = merge_bom_mfg(bom_df, mfg_df, bom_keys, mfg_keys)
    filled_df = map_columns_to_template(combined_df, column_mapping)

    columns_to_extract = ['QUANTITY', 'CRITICAL_PART', 'PART_NUMBER', 'ITEM_NUMBER']
    write_filled_template(filled_df, bom_df, company_template_path, output_path, columns_to_extract)

    return output_path
