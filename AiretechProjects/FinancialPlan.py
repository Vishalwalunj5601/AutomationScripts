"""
===============================================================================
Airetech PPM Conversion - Financial Project Plan FBDI Generation Script
Target FBDI: FinancialProjectPlansImportTemplate.xlsm
Source File: Airetech TAB jobs as of 9-11-26.xlsx
Sheet Populated: PJO_FIN_PROJ_PLAN_XFACE
===============================================================================
"""

import os
import openpyxl
import pandas as pd
import datetime

# -----------------------------------------------------------------------------
# DIRECTORIES & PATHS
# -----------------------------------------------------------------------------
datafiles_dir = r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\FINAL LOAD\DATAFILE\Financial Plan Import"
os.makedirs(datafiles_dir, exist_ok=True)

out_populated = os.path.join(datafiles_dir, "FinancialProjectPlan_Populated.xlsx")
out_xlsm = os.path.join(datafiles_dir, "FinancialProjectPlansImportTemplate_FINAL_LOAD.xlsm")

# Source File candidate locations (checks configured path, current script directory, etc.)
candidates_src = [
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\FINAL LOAD\SOURCE\TAB Pull from Jonas 9-18.xlsx",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "TAB Pull from Jonas 9-18.xlsx"),
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\DEV1\SourceFile\Airetech TAB jobs as of 9-11-26.xlsx",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "Airetech TAB jobs as of 9-11-26.xlsx")
]
p_src = next((c for c in candidates_src if os.path.exists(c)), candidates_src[0])

# Reference / Mapping File candidate locations (optional - script runs standalone if not present)
candidates_ref = [
    r"C:\Users\vishal walunj\Downloads\Reference_Mapping_Files\FinancialProjectPlansImportTemplate Mapping.xlsm",
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\DEV1\DataFiles\Financial Plan Import\FinancialProjectPlansImportTemplate  DEV1 - V1.xlsm",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "FinancialProjectPlansImportTemplate Mapping.xlsm"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "FinancialProjectPlansImportTemplate.xlsm"),
]
p_ref = next((c for c in candidates_ref if os.path.exists(c)), None)

# -----------------------------------------------------------------------------
# EMBEDDED COMPLETE ORACLE FBDI HEADERS & STRUCTURE
# -----------------------------------------------------------------------------
FIN_PLAN_HEADERS = [
    '* Processing Mode', 'Project Number', 'Project Name', 'Task Number', 'Task Name',
    '* Resource Name', 'Resource Assignment Level', 'Period Name', 'Start Date', 'Finish Date',
    '*Planning Currency', 'Quantity', 'Raw Cost', 'Burdened Cost', '*Source Line Reference',
    'Attribute Category', 'Attribute1', 'Attribute2', 'Attribute3', 'Attribute4',
    'Attribute5', 'Attribute6', 'Attribute7', 'Attribute8', 'Attribute9', 'Attribute10',
    'Attribute11', 'Attribute12', 'Attribute13', 'Attribute14', 'Attribute15', 'Attribute16',
    'Attribute17', 'Attribute18', 'Attribute19', 'Attribute20', 'Attribute21', 'Attribute22',
    'Attribute23', 'Attribute24', 'Attribute25', 'Attribute26', 'Attribute27', 'Attribute28',
    'Attribute29', 'Attribute30'
]


# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def clean_job_no(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    if val_str.isdigit():
        return val_str.zfill(10)
    return val_str


def parse_date_fixed(val):
    if pd.isna(val): return ""
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
    except:
        return str(val_str)


def add_10_years(s_date_str):
    """Calculates end date as start date + 10 years in YYYY/MM/DD format."""
    if not s_date_str:
        return ""
    try:
        dt = datetime.datetime.strptime(str(s_date_str).strip(), '%Y/%m/%d')
        try:
            dt_plus_10 = dt.replace(year=dt.year + 10)
        except ValueError:
            # Leap year handling (e.g. Feb 29 -> Feb 28)
            dt_plus_10 = dt.replace(year=dt.year + 10, day=28)
        return dt_plus_10.strftime('%Y/%m/%d')
    except Exception:
        try:
            dt = pd.to_datetime(s_date_str)
            dt_plus_10 = dt.replace(year=dt.year + 10)
            return dt_plus_10.strftime('%Y/%m/%d')
        except:
            return str(s_date_str)


def safe_save(wb, path):
    try:
        wb.save(path)
        print("SUCCESS SAVED:", path)
        return path
    except Exception as e:
        base, ext = os.path.splitext(path)
        alt = f"{base}_clean{ext}"
        wb.save(alt)
        print("SAVED TO CLEAN ALT PATH:", alt)
        return alt


def create_fin_plan_workbook():
    """Creates a complete Oracle Financial Plan FBDI workbook from embedded definitions."""
    wb_new = openpyxl.Workbook()
    ws = wb_new.active
    ws.title = "PJO_FIN_PROJ_PLAN_XFACE"
    ws.cell(row=2, column=1, value="Financial Project Plan Resource Assignment")
    ws.cell(row=3, column=1, value="* Required ")
    for col_idx, header in enumerate(FIN_PLAN_HEADERS, start=1):
        ws.cell(row=5, column=col_idx, value=header)
    return wb_new


# -----------------------------------------------------------------------------
# MAIN PROCESSING
# -----------------------------------------------------------------------------
print("=== GENERATING FINANCIAL PROJECT PLAN FBDI ===")

# Read source file and filter out repeated headers and empty rows
df_src = pd.read_excel(p_src)
df_src = df_src[df_src['Job No.'].notna() & (df_src['Job No.'].astype(str).str.strip() != 'Job No.')].reset_index(drop=True)
print(f"Loaded {len(df_src)} clean records from source file: {p_src}")

# Initialize FBDI Workbook (runs standalone if no mapping/reference file is attached)
if p_ref and os.path.exists(p_ref):
    try:
        print(f"Found base template file: {p_ref}")
        wb_d = openpyxl.load_workbook(p_ref)
        ws_d = wb_d["PJO_FIN_PROJ_PLAN_XFACE"]
        # Clear existing rows 11 to 100 across all columns
        for r in range(11, 101):
            for c in range(1, 157):
                ws_d.cell(row=r, column=c, value='')
    except Exception as e:
        print(f"Notice: Could not load template file ({e}). Generating fresh FBDI workbook from embedded code...")
        wb_d = create_fin_plan_workbook()
        ws_d = wb_d["PJO_FIN_PROJ_PLAN_XFACE"]
else:
    print("Standalone Mode: No external mapping/template file attached. Generating full FBDI workbook from embedded code...")
    wb_d = create_fin_plan_workbook()
    ws_d = wb_d["PJO_FIN_PROJ_PLAN_XFACE"]

start_row = 11
for i in range(len(df_src)):
    row = df_src.iloc[i]
    r_idx = start_row + i

    j_num = clean_job_no(row['Job No.'])
    j_name = str(row['Job Name']).strip()
    s_date = parse_date_fixed(row['Start Date'])
    e_date = add_10_years(s_date)  # Finish Date = Start Date + 10 years

    # Raw Cost / Estimated Total Cost = 75% of Contract Amount ("Contract Amount" / "Curr Est Rev")
    rev_val = row['Contract Amount'] if ('Contract Amount' in row and pd.notna(row['Contract Amount']) and str(row['Contract Amount']).strip() != '') else (
              row['Curr Est Rev'] if ('Curr Est Rev' in row and pd.notna(row['Curr Est Rev']) and str(row['Curr Est Rev']).strip() != '') else 0.0)
    try:
        curr_rev = float(rev_val)
    except:
        curr_rev = 0.0
    raw_cost = round(curr_rev * 0.75, 2)

    # Populate PJO_FIN_PROJ_PLAN_XFACE columns
    ws_d.cell(row=r_idx, column=1, value="Create")          # * Processing Mode
    ws_d.cell(row=r_idx, column=2, value=j_num)             # Project Number
    ws_d.cell(row=r_idx, column=3, value=j_name)            # Project Name
    ws_d.cell(row=r_idx, column=4, value="")                # Task Number (Blank)
    ws_d.cell(row=r_idx, column=5, value="Labor")           # Task Name
    ws_d.cell(row=r_idx, column=6, value="Labor Resource")  # * Resource Name
    ws_d.cell(row=r_idx, column=7, value="LINE")            # Resource Assignment Level
    ws_d.cell(row=r_idx, column=8, value="")                # Period Name (Blank)
    ws_d.cell(row=r_idx, column=9, value=s_date)            # Start Date (YYYY/MM/DD)
    ws_d.cell(row=r_idx, column=10, value=e_date)           # Finish Date (Start Date + 10 Years)
    ws_d.cell(row=r_idx, column=11, value="USD")            # *Planning Currency
    ws_d.cell(row=r_idx, column=12, value="")               # Quantity (Blank)
    ws_d.cell(row=r_idx, column=13, value=raw_cost)         # Raw Cost (75% of Contract Amount)
    ws_d.cell(row=r_idx, column=14, value="")               # Burdened Cost (Blank)
    ws_d.cell(row=r_idx, column=15, value=j_num)            # *Source Line Reference = Job No.

print(f"Populated {len(df_src)} financial project plan records.")
safe_save(wb_d, out_populated)
if p_ref and p_ref.endswith('.xlsm'):
    safe_save(wb_d, out_xlsm)
print(f"=== FINANCIAL PLAN FBDI GENERATION COMPLETE: {out_populated} ===")
