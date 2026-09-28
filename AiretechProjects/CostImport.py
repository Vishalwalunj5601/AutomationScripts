"""
===============================================================================
Airetech PPM Conversion - Project Cost Import FBDI Generation Script
===============================================================================
Target FBDI Template: ImportCosts DEV1.xlsx / ImportCosts Mapping.xlsm
Target Output Directory:
  C:\\Users\\vishal walunj\\Vishal\\1D_AIR_CONTROL_CONCEPTS\\AiretechProjects\\DEV1\\DataFiles\\Cost Import

Key Rules:
  - Source File: Airetech TAB jobs as of 9-11-26.xlsx
  - Cost Column: "To Date Costs"
  - Document: "GL Costed and Accounted Costs"
  - Document Entry: "GL Costed and Accounted Costs"
  - Transaction Type: "MISCELLANEOUS"
  - Business Unit: "Airetech"
  - Source Application: "Project Miscellaneous Costs"
  - Cutover Accounting Date: "2026/07/31"
  - Filter: Only jobs with To Date Costs > 0
  - Standalone Execution: Runs seamlessly with or without external mapping templates.
===============================================================================
"""

import os
import openpyxl
import pandas as pd
import datetime

# -----------------------------------------------------------------------------
# DIRECTORIES & FILE PATHS
# -----------------------------------------------------------------------------
datafiles_dir = r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\FINAL LOAD\DATAFILE\Cost Import"
os.makedirs(datafiles_dir, exist_ok=True)

out_populated = os.path.join(datafiles_dir, "ImportCosts_FINAL_LOAD.xlsx")
out_xlsm = os.path.join(datafiles_dir, "ImportCosts_FINAL_LOAD.xlsm")

# Candidate Source Files
candidates_src = [
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\FINAL LOAD\SOURCE\TAB Pull from Jonas 9-18.xlsx",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "TAB Pull from Jonas 9-18.xlsx"),
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\DEV1\SourceFile\Airetech TAB jobs as of 9-11-26.xlsx",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "Airetech TAB jobs as of 9-11-26.xlsx"),
]
p_src = next((c for c in candidates_src if os.path.exists(c)), candidates_src[0])

# Candidate Reference Templates
candidates_ref = [
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\DEV1\DataFiles\Cost Import\ImportCosts DEV1 v1.xlsx",
    r"C:\Users\vishal walunj\Downloads\Reference_Mapping_Files\ImportCosts Mapping.xlsm",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "ImportCosts Mapping.xlsm")
]
p_ref = next((c for c in candidates_ref if os.path.exists(c)), None)

# -----------------------------------------------------------------------------
# DEFAULT VALUES & CONFIGURATION
# -----------------------------------------------------------------------------
TRANSACTION_TYPE = "MISCELLANEOUS"
BUSINESS_UNIT_NAME = "Airetech"
TPA_SOURCE = "Project Miscellaneous Costs"
DOCUMENT = "Conversion Costs"
DOCUMENT_ENTRY = "Conversion Costs"
TASK_NAME = "Labor"
EXPENDITURE_TYPE = "Labor"
UOM_DEFAULT = "Currency"
BILLABLE_DEFAULT = "Y"
CURRENCY_CODE = "USD"
CUTOVER_DATE = "2026/07/31"
RAW_COST_ACCOUNT = "1200.1221.200.000.512001.0000.000.1065.0000"

# -----------------------------------------------------------------------------
# EMBEDDED COMPLETE ORACLE FBDI HEADERS (104 Columns)
# -----------------------------------------------------------------------------
COST_IMPORT_HEADERS = [
    '* Transaction Type', 'Business Unit Name', 'Business Unit ID',
    'Third-Party Application Transaction Source', 'Third-Party Application Transaction Source ID',
    'Document', 'Document ID', 'Document Entry', 'Document Entry ID',
    '* Expenditure Batch', 'Batch Ending Date', 'Batch Description',
    '* Expenditure Item Date', 'Person Number', 'Person Name', 'Person ID',
    'Human Resources Assignment', 'Human Resources Assignment ID',
    'Project Number', 'Project Name', 'Project ID', 'Task Number', 'Task Name',
    'Task ID', 'Expenditure Type', 'Expenditure Type ID',
    'Expenditure Organization', 'Expenditure Organization ID',
    'Contract Number', 'Contract Name', 'Contract ID',
    'Funding Source Number', 'Funding Source Name', '* Quantity',
    'Unit of Measure', 'Unit of Measure Code', 'Work Type', 'Work Type ID',
    'Billable', 'Capitalizable', 'Accrual Item',
    '* Original Transaction Reference', 'Unmatched Negative Transaction',
    'Reversed Original Transaction', 'Expenditure Item Comment', 'Accounting Date',
    'Transaction Currency Code', 'Transaction Currency',
    'Raw Cost in Transaction Currency', 'Burdened Cost in Transaction Currency',
    'Raw Cost Credit CCID', 'Raw Cost Credit Account',
    'Raw Cost Debit CCID', 'Raw Cost Debit Account',
    'Burdened Cost Credit CCID', 'Burdened Cost Credit Account',
    'Burdened Cost Debit CCID', 'Burdened Cost Debit Account',
    'Burden Cost Credit CCID', 'Burden Cost Credit Account',
    'Burden Cost Debit CCID', 'Burden Cost Debit Account',
    'Provider Ledger Currency Code', 'Provider Ledger Currency',
    'Raw Cost in Provider Ledger Currency', 'Burdened Cost in Provider Ledger Currency',
    'Provider Ledger Conversion Rate Type', 'Provider Ledger Conversion Rate Date',
    'Provider Ledger Conversion Date Type', 'Provider Ledger Currency Conversion Rate',
    'Provider Ledger Currency Conversion Rounding Limit', 'Converted',
    'Context Category', 'User Defined Attribute 1', 'User Defined Attribute 2',
    'User Defined Attribute 3', 'User Defined Attribute 4', 'User Defined Attribute 5',
    'User Defined Attribute 6', 'User Defined Attribute 7', 'User Defined Attribute 8',
    'User Defined Attribute 9', 'User Defined Attribute 10', 'Funding Source ID',
    'Reserved Attribute 2', 'Reserved Attribute 3', 'Reserved Attribute 4',
    'Reserved Attribute 5', 'Reserved Attribute 6', 'Reserved Attribute 7',
    'Reserved Attribute 8', 'Reserved Attribute 9', 'Reserved Attribute 10',
    'Attribute Category', 'Attribute 1', 'Attribute 2', 'Attribute 3',
    'Attribute 4', 'Attribute 5', 'Attribute 6', 'Attribute 7',
    'Attribute 8', 'Attribute 9', 'Attribute 10'
]


# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def clean_job_no(val):
    """Preserves original job number string, padding numeric values to 10 digits."""
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    if val_str.isdigit():
        return val_str.zfill(10)
    return val_str


def parse_cost_value(val):
    """Parses cost string/number cleanly into float."""
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace(',', '').replace('$', '')
    try:
        return float(val_str)
    except Exception:
        return 0.0


def parse_date_fixed(val):
    """Parses various date string formats and returns YYYY/MM/DD."""
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    for fmt in ('%b %d,%Y', '%b %d, %Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%m/%d/%Y'):
        try:
            dt = datetime.datetime.strptime(val_str, fmt)
            return dt.strftime('%Y/%m/%d')
        except ValueError:
            pass
    try:
        dt = pd.to_datetime(val_str)
        return dt.strftime('%Y/%m/%d')
    except Exception:
        return str(val_str)


def get_org(off):
    """Derives Owning Business Unit Organization from Office abbreviation."""
    off_str = str(off).strip().upper()
    if 'TUL' in off_str:
        return 'Tulsa-Airetech-Services'
    elif 'SPG' in off_str:
        return 'Springdale-Airetech-Services'
    elif 'LR' in off_str:
        return 'Little Rock-Airetech-Services'
    return 'Springdale-Airetech-Services'


def safe_save(wb, path):
    """Saves workbook safely, creating a _clean alternative if file is locked in Excel."""
    try:
        wb.save(path)
        print(f"  [OK] Successfully saved: {path}")
        return path
    except Exception as e:
        base, ext = os.path.splitext(path)
        alt = f"{base}_clean{ext}"
        wb.save(alt)
        print(f"  [WARN] File locked in Excel ({e}): Saved to alternative -> {alt}")
        return alt


def create_cost_fbdi_workbook():
    """Creates a complete, clean Oracle Cost Import FBDI workbook from embedded definitions."""
    wb_new = openpyxl.Workbook()
    # Sheet 1: Instructions and CSV Generation
    ws_inst = wb_new.active
    ws_inst.title = "Instructions and CSV Generation"
    ws_inst.cell(row=1, column=1, value="Oracle Fusion Project Cost Transactions Interface")

    # Sheet 2: PJC_TXN_XFACE_STAGE_ALL
    ws_cost = wb_new.create_sheet(title="PJC_TXN_XFACE_STAGE_ALL")
    ws_cost.views.sheetView[0].showGridLines = True
    ws_cost.cell(row=2, column=1, value="Project Miscellaneous Cost Transactions Interface")
    ws_cost.cell(row=3, column=1, value="* Required")
    for col_idx, header in enumerate(COST_IMPORT_HEADERS, start=1):
        ws_cost.cell(row=4, column=col_idx, value=header)

    return wb_new


# -----------------------------------------------------------------------------
# MAIN GENERATION PROCESS
# -----------------------------------------------------------------------------
def main():
    print("=" * 80)
    print("  Airetech PPM Conversion - Project Cost Import FBDI Script Starting")
    print("=" * 80)

    if not os.path.exists(p_src):
        raise FileNotFoundError(f"Source file not found at: {p_src}")

    # Load Source File
    df_src = pd.read_excel(p_src)
    df_src = df_src[df_src['Job No.'].notna() & (df_src['Job No.'].astype(str).str.strip() != 'Job No.')].reset_index(drop=True)
    print(f"  [OK] Loaded source dataset: {p_src} ({len(df_src)} records)")

    # Parse Cost Column: "To Date Costs"
    cost_col_name = "To Date Costs" if "To Date Costs" in df_src.columns else "Actual Costs"
    print(f"  [OK] Referring Cost from column: '{cost_col_name}'")
    df_src['Parsed_Cost'] = df_src[cost_col_name].apply(parse_cost_value)

    # Filter records with positive costs
    df_costs = df_src[df_src['Parsed_Cost'] > 0].reset_index(drop=True)
    print(f"  [OK] Found {len(df_costs)} jobs with positive '{cost_col_name}' > 0 to import.")

    # Initialize FBDI Workbook
    if p_ref and os.path.exists(p_ref):
        try:
            print(f"  [INFO] Loading base FBDI template: {p_ref}")
            wb = openpyxl.load_workbook(p_ref)
            ws_cost = wb["PJC_TXN_XFACE_STAGE_ALL"]
            # Clear existing data rows starting from row 5 down to max_row
            for r in range(5, max(ws_cost.max_row + 1, 50)):
                for c in range(1, 105):
                    ws_cost.cell(row=r, column=c, value='')
        except Exception as e:
            print(f"  [WARN] Could not load base template ({e}). Creating fresh FBDI workbook...")
            wb = create_cost_fbdi_workbook()
            ws_cost = wb["PJC_TXN_XFACE_STAGE_ALL"]
    else:
        print("  [INFO] Standalone Mode: Generating complete FBDI workbook from embedded definitions...")
        wb = create_cost_fbdi_workbook()
        ws_cost = wb["PJC_TXN_XFACE_STAGE_ALL"]

    # Ensure headers at Row 4
    ws_cost.cell(row=2, column=1, value="Project Miscellaneous Cost Transactions Interface")
    ws_cost.cell(row=3, column=1, value="* Required")
    for col_idx, header in enumerate(COST_IMPORT_HEADERS, start=1):
        ws_cost.cell(row=4, column=col_idx, value=header)

    start_row = 5
    print("\n  Populating Cost Import FBDI Records...")
    for i in range(len(df_costs)):
        row = df_costs.iloc[i]
        r_idx = start_row + i

        j_num = clean_job_no(row['Job No.'])
        cost_val = float(row['Parsed_Cost'])
        org = get_org(row.get('Office (Proj Organization)', ''))

        # Batch name formatted as conversion_costs_YYYYMMDD per feedback
        batch_name = "conversion_costs_20260918"

        # Hours comments
        hrs_val = str(row['To Date Hrs']).strip() if ('To Date Hrs' in row and pd.notna(row['To Date Hrs'])) else ""
        comment = f"Legacy Hours: {hrs_val}" if hrs_val else ""

        # Row Population
        ws_cost.cell(row=r_idx, column=1, value=TRANSACTION_TYPE)           # * Transaction Type
        ws_cost.cell(row=r_idx, column=2, value=BUSINESS_UNIT_NAME)          # Business Unit Name
        ws_cost.cell(row=r_idx, column=4, value=TPA_SOURCE)                  # Third-Party App Txn Source
        ws_cost.cell(row=r_idx, column=6, value=DOCUMENT)                    # Document = "Conversion Costs"
        ws_cost.cell(row=r_idx, column=8, value=DOCUMENT_ENTRY)              # Document Entry = "Conversion Costs"
        ws_cost.cell(row=r_idx, column=10, value=batch_name)                 # * Expenditure Batch
        ws_cost.cell(row=r_idx, column=13, value=CUTOVER_DATE)               # * Expenditure Item Date (2026/07/31)
        ws_cost.cell(row=r_idx, column=19, value=j_num)                      # Project Number
        ws_cost.cell(row=r_idx, column=23, value=TASK_NAME)                  # Task Name = "Labor"
        ws_cost.cell(row=r_idx, column=25, value=EXPENDITURE_TYPE)           # Expenditure Type = "Labor"
        ws_cost.cell(row=r_idx, column=27, value=org)                        # Expenditure Organization
        ws_cost.cell(row=r_idx, column=34, value=cost_val)                   # * Quantity
        ws_cost.cell(row=r_idx, column=35, value=UOM_DEFAULT)                # Unit of Measure = "Currency"
        ws_cost.cell(row=r_idx, column=39, value=BILLABLE_DEFAULT)           # Billable = "Y"
        ws_cost.cell(row=r_idx, column=42, value=f"TXN_COST_{j_num}_01")     # * Original Transaction Reference
        ws_cost.cell(row=r_idx, column=45, value=comment)                    # Expenditure Item Comment
        ws_cost.cell(row=r_idx, column=46, value=CUTOVER_DATE)               # Accounting Date (2026/07/31)
        ws_cost.cell(row=r_idx, column=47, value=CURRENCY_CODE)              # Transaction Currency Code = "USD"
        ws_cost.cell(row=r_idx, column=49, value=cost_val)                   # Raw Cost in Transaction Currency
        ws_cost.cell(row=r_idx, column=52, value=RAW_COST_ACCOUNT)           # Raw Cost Credit Account
        ws_cost.cell(row=r_idx, column=54, value=RAW_COST_ACCOUNT)           # Raw Cost Debit Account
        ws_cost.cell(row=r_idx, column=63, value=CURRENCY_CODE)              # Provider Ledger Currency Code = "USD"
        ws_cost.cell(row=r_idx, column=65, value=cost_val)                   # Raw Cost in Provider Ledger Currency
        ws_cost.cell(row=r_idx, column=68, value=CUTOVER_DATE)               # Provider Ledger Conversion Rate Date

    # Save to FINAL LOAD DataFiles directory
    safe_save(wb, out_populated)
    if p_ref and p_ref.endswith('.xlsm'):
        safe_save(wb, out_xlsm)

    print("\n" + "=" * 80)
    print("  [OK] COST IMPORT FBDI GENERATION COMPLETED SUCCESSFULLY!")
    print(f"  Target File:     {out_populated}")
    print(f"  Records Created: {len(df_costs)} cost transactions")
    print(f"  Cost Source:     Column '{cost_col_name}'")
    print(f"  Document:        '{DOCUMENT}'")
    print(f"  Document Entry:  '{DOCUMENT_ENTRY}'")
    print("=" * 80)


if __name__ == "__main__":
    main()
