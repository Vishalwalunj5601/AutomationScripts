import os
import openpyxl
import pandas as pd
import datetime

# -----------------------------------------------------------------------------
# DIRECTORIES & PATHS
# -----------------------------------------------------------------------------
datafiles_dir = r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\FINAL LOAD\DATAFILE\Project Import"
os.makedirs(datafiles_dir, exist_ok=True)

out_populated = os.path.join(datafiles_dir, "ProjectImportTemplate_Populated.xlsx")
out_xlsm = os.path.join(datafiles_dir, "ProjectImportTemplate_FINAL_LOAD.xlsm")

# Source File candidate locations (checks configured path, current script directory, etc.)
candidates_src = [
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\FINAL LOAD\SOURCE\TAB Pull from Jonas 9-18.xlsx",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "TAB Pull from Jonas 9-18.xlsx"),
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\DEV1\SourceFile\Airetech TAB jobs as of 9-11-26.xlsx",
]
p_src = next((c for c in candidates_src if os.path.exists(c)), candidates_src[0])

# Reference / Mapping File candidate locations (optional - script runs standalone if not present)
candidates_ref = [
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\FINAL LOAD\DATAFILE\Project Import\ProjectImportTemplate  DEV1 - V1.xlsm",
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\DEV1\DataFiles\ProjectImport\ProjectImportTemplate  DEV1 - V1.xlsm",
    r"C:\Users\vishal walunj\Downloads\Reference_Mapping_Files\ProjectImportTemplate Mapping.xlsm",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "ProjectImportTemplate  DEV1 - V1.xlsm"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "ProjectImportTemplate Mapping.xlsm"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "ProjectImportTemplate.xlsm"),
]
p_ref = next((c for c in candidates_ref if os.path.exists(c)), None)


# -----------------------------------------------------------------------------
# EMBEDDED COMPLETE ORACLE FBDI HEADERS & STRUCTURE
# (Allows running completely standalone without requiring any mapping/template file)
# -----------------------------------------------------------------------------
PROJECTS_HEADERS = [
    '*Project Name', 'Project Number', 'Source Template Number', 'Source Template Name',
    'Source Application Code', 'Source Project Reference', 'Project Calendar', 'EPS Element',
    'View-Only Project Plan Access Code', 'Schedule Type Code', 'Organization', 'Legal Entity',
    'Project Description', 'Project Manager Number', 'Project Manager Name', 'Project Manager Email',
    'Cascade Option', 'Project Start Date', 'Project Finish Date', 'Closed Date',
    'Project Plan Baseline Name', 'Project Plan Baseline Description', 'Project Plan Baseline Date',
    'Project Status', 'Priority Code', 'Outline Display Level', 'Planning Project',
    'Service Type Code', 'Work Type', 'Limit to Transaction Controls', 'Enable Budgetary Controls',
    'Project Currency', 'Project Currency Conversion Rate Type', 'Project Currency Conversion Date Type Code',
    'Project Currency Conversion Date', 'Allow Capitalized Interest', 'Capitalized Interest Rate Schedule',
    'Capitalized Interest Stop Date', 'Asset Cost Allocation Method Code', 'Capital Event Processing Method Code',
    'Allow Charges from All Provider Business Units', 'Process Cross-Charge Transactions for Labor',
    'Labor Transfer Price Schedule', 'Labor Transfer Price Fixed Date',
    'Process Cross-Charge Transactions for Nonlabor', 'Nonlabor Transfer Price Schedule',
    'Nonlabor Transfer Price Fixed Date', 'Burden Schedule', 'Burden Schedule Fixed Date',
    'KPI Notifications Enabled', 'KPI Notifications Enabled for Project Manager',
    'Include Notes in KPI Notifications', 'Copy Team Members from Template',
    'Copy Classifications from Template', 'Copy Attachments from Template',
    'Copy Descriptive Flexfields from Template', 'Copy Tasks from Template',
    'Copy Task Attachments from Template', 'Copy Task Descriptive Flexfields from Template',
    'Copy Task Assignments from Template', 'Copy Transaction Controls from Template',
    'Copy Assets from Template', 'Copy Asset Assignments from Template',
    'Copy Costing Overrides from Template', 'Opportunity ID', 'Opportunity Number',
    'Opportunity Customer Number', 'Opportunity Customer ID', 'Opportunity Amount',
    'Opportunity Currency Code', 'Opportunity Win Confidence Percent', 'Opportunity Name',
    'Opportunity Description', 'Opportunity Customer Name', 'Opportunity Status',
    'Attribute Category', 'Attribute1', 'Attribute2', 'Attribute3', 'Attribute4',
    'Attribute5', 'Attribute6', 'Attribute7', 'Attribute8', 'Attribute9',
    'Attribute10', 'Attribute11', 'Attribute12', 'Attribute13', 'Attribute14',
    'Attribute15', 'Attribute16', 'Attribute17', 'Attribute18', 'Attribute19',
    'Attribute20', 'Attribute21', 'Attribute22', 'Attribute23', 'Attribute24',
    'Attribute25', 'Attribute26', 'Attribute27', 'Attribute28', 'Attribute29',
    'Attribute30', 'Attribute31', 'Attribute32', 'Attribute33', 'Attribute34',
    'Attribute35', 'Attribute36', 'Attribute37', 'Attribute38', 'Attribute39',
    'Attribute40', 'Attribute41', 'Attribute42', 'Attribute43', 'Attribute44',
    'Attribute45', 'Attribute46', 'Attribute47', 'Attribute48', 'Attribute49',
    'Attribute50', 'Attribute1_Number', 'Attribute2_Number', 'Attribute3_Number',
    'Attribute4_Number', 'Attribute5_Number', 'Attribute6_Number', 'Attribute7_Number',
    'Attribute8_Number', 'Attribute9_Number', 'Attribute10_Number', 'Attribute11_Number',
    'Attribute12_Number', 'Attribute13_Number', 'Attribute14_Number', 'Attribute15_Number',
    'Attribute1_Date', 'Attribute2_Date', 'Attribute3_Date', 'Attribute4_Date',
    'Attribute5_Date', 'Attribute6_Date', 'Attribute7_Date', 'Attribute8_Date',
    'Attribute9_Date', 'Attribute10_Date', 'Attribute11_Date', 'Attribute12_Date',
    'Attribute13_Date', 'Attribute14_Date', 'Attribute15_Date'
]

TASKS_HEADERS = [
    '*Project Name', 'Project Number', '*Task Name', '*Task Number', 'Source Task Reference',
    'Financial Task', 'Task Description', 'Parent Task Number', 'Planned Start Date',
    'Planned End Date', 'Planned Effort in Hours', 'Duration in Days', 'Milestone', 'Critical',
    'Chargeable', 'Billable', 'Capitalizable', 'Limit to Transaction Controls',
    'Service Type Code', 'Work Type', 'Task Manager', 'Allow Cross Charge Flag',
    'Cross Charge Process Labor Flag', 'Cross Charge Process Non Labor Flag',
    'Receive Project Invoice Flag', 'Organization', 'Requirement Code', 'Sprint', 'Priority',
    'Schedule Mode', 'Project Plan Baseline Start Date', 'Project Plan Baseline Finish Date',
    'Project Plan Baseline Effort in Hours', 'Project Plan Baseline Duration',
    'Project Plan Baseline Allocation', 'Project Plan Baseline Labor Cost Amount',
    'Project Plan Baseline Labor Bill Amount', 'Project Plan Baseline Expense Cost Amount',
    'Constraint Type Code', 'Constraint Date', 'Attribute Category', 'Attribute1',
    'Attribute2', 'Attribute3', 'Attribute4', 'Attribute5', 'Attribute6', 'Attribute7',
    'Attribute8', 'Attribute9', 'Attribute10', 'Attribute11', 'Attribute12', 'Attribute13',
    'Attribute14', 'Attribute15', 'Attribute16', 'Attribute17', 'Attribute18', 'Attribute19',
    'Attribute20', 'Attribute21', 'Attribute22', 'Attribute23', 'Attribute24', 'Attribute25',
    'Attribute26', 'Attribute27', 'Attribute28', 'Attribute29', 'Attribute30', 'Attribute31',
    'Attribute32', 'Attribute33', 'Attribute34', 'Attribute35', 'Attribute36', 'Attribute37',
    'Attribute38', 'Attribute39', 'Attribute40', 'Attribute41', 'Attribute42', 'Attribute43',
    'Attribute44', 'Attribute45', 'Attribute46', 'Attribute47', 'Attribute48', 'Attribute49',
    'Attribute50', 'Attribute1_Number', 'Attribute2_Number', 'Attribute3_Number',
    'Attribute4_Number', 'Attribute5_Number', 'Attribute6_Number', 'Attribute7_Number',
    'Attribute8_Number', 'Attribute9_Number', 'Attribute10_Number', 'Attribute11_Number',
    'Attribute12_Number', 'Attribute13_Number', 'Attribute14_Number', 'Attribute15_Number',
    'Attribute1_Date', 'Attribute2_Date', 'Attribute3_Date', 'Attribute4_Date',
    'Attribute5_Date', 'Attribute6_Date', 'Attribute7_Date', 'Attribute8_Date',
    'Attribute9_Date', 'Attribute10_Date', 'Attribute11_Date', 'Attribute12_Date',
    'Attribute13_Date', 'Attribute14_Date', 'Source Application Code', 'Attribute15_Date'
]

PROJECT_TEAM_MEMBERS_HEADERS = [
    '*Project Name', 'Team Member Number', 'Team Member Name', 'Team Member Email',
    '*Project Role', 'Start Date', 'End Date', 'Percent Allocation', 'Effort in Hours',
    'Cost Rate', 'Bill Rate', 'Track Time', 'Assignment Type', 'Billable Percent',
    'Billable Percent Reason'
]

# -----------------------------------------------------------------------------
# EMBEDDED COMPLETE MAPPING SPECIFICATION & RULES
# (Embedded directly into the code rather than relying on external mapping files)
# -----------------------------------------------------------------------------
TASKS = [
    "Material", "Labor", "Subcontracting", "Misc Costs", "P&R",
    "Warranty", "Admin and OH", "Submittal", "As Built", "Software Engineering"
]

MAPPING_SPECIFICATION = {
    "Projects": {
        1:  {"Header": "*Project Name", "Mapping": "Source: Job Name", "Rule": "Direct mapping"},
        2:  {"Header": "Project Number", "Mapping": "Source: Job No.", "Rule": "Original job number (e.g. 0000066580)"},
        4:  {"Header": "Source Template Name", "Mapping": "Default: Airetech Billable Project Template", "Rule": "Fixed template name"},
        5:  {"Header": "Source Application Code", "Mapping": "Default: ORA_PROJECT_SERVICE", "Rule": "Fixed source app code"},
        11: {"Header": "Organization", "Mapping": "Derived from Office (Proj Organization)", "Rule": "TUL -> Tulsa-Airetech-Services, SPG -> Springdale-Airetech-Services, LR -> Little Rock-Airetech-Services"},
        13: {"Header": "Project Description", "Mapping": "Left Blank", "Rule": "Optional"},
        15: {"Header": "Project Manager Name", "Mapping": "Left Blank on Project Sheet", "Rule": "Loaded via Project Team Members sheet"},
        16: {"Header": "Project Manager Email", "Mapping": "Left Blank on Project Sheet", "Rule": "Loaded via Project Team Members sheet"},
        18: {"Header": "Project Start Date", "Mapping": "Source: Start Date", "Rule": "Formatted YYYY/MM/DD"},
        19: {"Header": "Project Finish Date", "Mapping": "Source: Est Compl'n", "Rule": "Formatted YYYY/MM/DD"},
        24: {"Header": "Project Status", "Mapping": "Left Blank", "Rule": "Removed per user instruction"},
        57: {"Header": "Copy Tasks from Template", "Mapping": "Default: Y", "Rule": "ONLY Copy Tasks = Y; all other copy flags Blank"},
        76: {"Header": "Attribute Category", "Mapping": "Left Blank", "Rule": "Removed per user instruction"},
        77: {"Header": "Attribute1 (quoteNumber)", "Mapping": "Source: Job No.", "Rule": "Attribute1 = Original Job No."},
        78: {"Header": "Attribute2 (contractType)", "Mapping": "Default: New Construction", "Rule": "Default 'New Construction' for all projects"},
        79: {"Header": "Attribute3 (productEstimated)", "Mapping": "Left Blank", "Rule": "Blank"},
        80: {"Header": "Attribute4", "Mapping": "Left Blank", "Rule": "Blank"},
        81: {"Header": "Attribute5 (distributionTo)", "Mapping": "Default: OWNER", "Rule": "Default 'OWNER' for all projects"},
        82: {"Header": "Attribute6 (costPerTripDay)", "Mapping": "Left Blank", "Rule": "Blank"}
    },
    "Tasks": {
        "TaskList": TASKS,
        "Columns": {
            1:  {"Header": "*Project Name", "Mapping": "Source: Job Name"},
            2:  {"Header": "Project Number", "Mapping": "Source: Job No."},
            3:  {"Header": "*Task Name", "Mapping": "Standard 10 Tasks"},
            4:  {"Header": "*Task Number", "Mapping": "Standard 10 Tasks"},
            6:  {"Header": "Financial Task", "Mapping": "Default: Y"},
            9:  {"Header": "Planned Start Date", "Mapping": "Source: Start Date (YYYY/MM/DD)"},
            10: {"Header": "Planned End Date", "Mapping": "Default: 2026/12/31"},
            15: {"Header": "Chargeable", "Mapping": "Default: Y"},
            16: {"Header": "Billable", "Mapping": "Default: Y"},
            42: {"Header": "Progress Status", "Mapping": "Default: NOT_STARTED"}
        }
    },
    "Project Team Members": {
        "Project Manager": {
            "Office_Mapping": {
                "TUL": {"Name": "Alex Borio", "Email": "aborio@airetechcorp.com", "HireDate": "2021/03/03"},
                "SPG": {"Name": "Darren Holliman", "Email": "dholliman@airetechcorp.com", "HireDate": "2019/04/15"},
                "LR":  {"Name": "Jason Draper", "Email": "jdraper@airetechcorp.com", "HireDate": "1996/09/23"},
                "Default": {"Name": "Darren Holliman", "Email": "dholliman@airetechcorp.com", "HireDate": "2019/04/15"}
            },
            "Role": "Project Manager"
        },
        "Sales Lead": {
            "Default_Name": "Air Balance",
            "Default_Email": "air.balance@airetech.com",
            "Role": "Sales Lead",
            "Allocation": 100,
            "HireDate": "2026/07/28"
        }
    }
}


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
    # calculates end date as start date + 10 years
    if not s_date_str:
        return ""
    try:
        dt = datetime.datetime.strptime(str(s_date_str).strip(), '%Y/%m/%d')
        return dt.replace(year=dt.year + 10).strftime('%Y/%m/%d')
    except:
        return str(s_date_str)


def get_org(off):
    off_str = str(off).strip().upper()
    if 'TUL' in off_str:
        return 'Tulsa-Airetech-Services'
    elif 'SPG' in off_str:
        return 'Springdale-Airetech-Services'
    elif 'LR' in off_str:
        return 'Little Rock-Airetech-Services'
    return 'Springdale-Airetech-Services'


def get_pm_info(off):
    off_str = str(off).strip().upper()
    pm_spec = MAPPING_SPECIFICATION["Project Team Members"]["Project Manager"]["Office_Mapping"]
    if 'TUL' in off_str:
        target = pm_spec['TUL']
    elif 'SPG' in off_str:
        target = pm_spec['SPG']
    elif 'LR' in off_str:
        target = pm_spec['LR']
    else:
        target = pm_spec['Default']
    return target['Name'], target['Email'], target.get('HireDate', '')


def get_later_date(date_str1, date_str2):
    """Returns whichever date is later (formatted YYYY/MM/DD)."""
    if not date_str1:
        return date_str2
    if not date_str2:
        return date_str1
    d1 = parse_date_fixed(date_str1)
    d2 = parse_date_fixed(date_str2)
    return d1 if d1 >= d2 else d2


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


def create_fbdi_workbook():
    """Creates a complete, clean Oracle Project Import FBDI workbook from embedded definitions."""
    wb_new = openpyxl.Workbook()
    # Sheet 1: Projects
    ws_p = wb_new.active
    ws_p.title = "Projects"
    ws_p.cell(row=2, column=1, value="Projects")
    ws_p.cell(row=3, column=1, value="* Required")
    for col_idx, header in enumerate(PROJECTS_HEADERS, start=1):
        ws_p.cell(row=4, column=col_idx, value=header)

    # Sheet 2: Tasks
    ws_t = wb_new.create_sheet(title="Tasks")
    ws_t.cell(row=2, column=1, value="Project Tasks")
    ws_t.cell(row=3, column=1, value="* Required")
    for col_idx, header in enumerate(TASKS_HEADERS, start=1):
        ws_t.cell(row=4, column=col_idx, value=header)

    # Sheet 3: Project Team Members
    ws_m = wb_new.create_sheet(title="Project Team Members")
    ws_m.cell(row=2, column=1, value="Project Team Members")
    ws_m.cell(row=3, column=1, value="* Required")
    for col_idx, header in enumerate(PROJECT_TEAM_MEMBERS_HEADERS, start=1):
        ws_m.cell(row=4, column=col_idx, value=header)

    return wb_new


print("=== GENERATING EXACT PROJECT IMPORT FBDI MATCHING USER REVISED SPECIFICATION ===")

# Read source file and filter out repeated headers and empty rows
df_src = pd.read_excel(p_src)
df_src = df_src[df_src['Job No.'].notna() & (df_src['Job No.'].astype(str).str.strip() != 'Job No.')].reset_index(drop=True)
print(f"Loaded {len(df_src)} clean records from source file: {p_src}")

# Initialize FBDI Workbook (irrespective of external mapping file being attached)
if p_ref and os.path.exists(p_ref):
    try:
        print(f"Found base template file: {p_ref}")
        wb = openpyxl.load_workbook(p_ref, keep_vba=True)
        ws_p = wb["Projects"]
        ws_t = wb["Tasks"]
        ws_m = wb["Project Team Members"]

        # Determine start row on Projects sheet based on whether mapping info rows exist
        if ws_p.cell(row=5, column=1).value == "Mapping":
            start_row_p = 8
            clear_start_p = 8
        else:
            start_row_p = 5
            clear_start_p = 5

        # Clear existing data rows starting from data start row onwards
        for r in range(clear_start_p, max(ws_p.max_row + 1, 500)):
            for c in range(1, 157):
                ws_p.cell(row=r, column=c, value='')
        for r in range(5, max(ws_t.max_row + 1, 500)):
            for c in range(1, 115):
                ws_t.cell(row=r, column=c, value='')
        for r in range(5, max(ws_m.max_row + 1, 800)):
            for c in range(1, 20):
                ws_m.cell(row=r, column=c, value='')
    except Exception as e:
        print(f"Notice: Could not load template file ({e}). Generating fresh FBDI workbook from embedded code...")
        wb = create_fbdi_workbook()
        ws_p = wb["Projects"]
        ws_t = wb["Tasks"]
        ws_m = wb["Project Team Members"]
        start_row_p = 5
else:
    print("Standalone Mode: No external mapping/template file attached. Generating full FBDI workbook from embedded code...")
    wb = create_fbdi_workbook()
    ws_p = wb["Projects"]
    ws_t = wb["Tasks"]
    ws_m = wb["Project Team Members"]
    start_row_p = 5


for i in range(len(df_src)):
    row = df_src.iloc[i]
    r_idx = start_row_p + i

    j_num = clean_job_no(row['Job No.'])
    j_name = str(row['Job Name']).strip()
    s_date = parse_date_fixed(row['Start Date'])
    e_date = add_10_years(s_date)  # Project End Date = Start Date + 10 years (YYYY/MM/DD)
    off_raw = row['Office (Proj Organization)']
    org = get_org(off_raw)

    # Defaults per user instructions:
    ctype = "New Construction"      # Default for all projects
    dist_to = "CONTRACTOR"          # Attribute5: default = CONTRACTOR (updated per business feedback)

    # 1-2. Identifiers
    ws_p.cell(row=r_idx, column=1, value=j_name)  # *Project Name
    ws_p.cell(row=r_idx, column=2, value=j_num)   # Project Number

    # 4-5. Template & Application Code
    ws_p.cell(row=r_idx, column=4, value="Airetech Billable Project Template")  # Source Template Name
    ws_p.cell(row=r_idx, column=5, value="AIR Conversion")                      # Source Application Code: exact lookup

    # 6. Source Project Reference
    ws_p.cell(row=r_idx, column=6, value=j_num)

    # 11. Organization
    ws_p.cell(row=r_idx, column=11, value=org)  # Organization

    # 13. Project Description: Populated with verified Job Name
    ws_p.cell(row=r_idx, column=13, value=j_name)  # Project Description

    # 15, 16. PM Name, PM Email -> BLANK on Project Sheet (loaded via Project Team Members sheet)
    ws_p.cell(row=r_idx, column=15, value='')  # Project Manager Name (Blank)
    ws_p.cell(row=r_idx, column=16, value='')  # Project Manager Email (Blank)

    # 18-19. Dates: End Date = Start Date + 10 years
    ws_p.cell(row=r_idx, column=18, value=s_date)  # Project Start Date
    ws_p.cell(row=r_idx, column=19, value=e_date)  # Project Finish Date (Start Date + 10 years)

    # 24. Project Status -> REMOVED / BLANK per user instruction!
    ws_p.cell(row=r_idx, column=24, value='')  # Project Status (Blank)

    # 53-64. Copy Flags: ONLY Col 57 (Copy Tasks from Template) = 'Y', ALL OTHERS BLANK
    for col_c in range(53, 65):
        ws_p.cell(row=r_idx, column=col_c, value="Y" if col_c == 57 else "")

    # 76-82. DFF Attributes
    ws_p.cell(row=r_idx, column=76, value='')       # Attribute Category (Blank)
    ws_p.cell(row=r_idx, column=77, value=j_num)    # Attribute1: quoteNumber = job no
    ws_p.cell(row=r_idx, column=78, value=ctype)    # Attribute2: contractType = New Construction
    ws_p.cell(row=r_idx, column=79, value='')       # Attribute3: productEstimated (Blank)
    ws_p.cell(row=r_idx, column=80, value='')       # Attribute4 (Blank)
    ws_p.cell(row=r_idx, column=81, value=dist_to)  # Attribute5: default = CONTRACTOR
    ws_p.cell(row=r_idx, column=82, value='')       # Attribute6: costPerTripDay (Blank)

# -------------------------------------------------------------
# 2. TASKS SHEET
# Since "Copy Tasks from Template = Y" is enabled, tasks are cloned
# automatically by Oracle PPM. Clearing task rows to avoid duplicate/error.
# -------------------------------------------------------------
for r in range(5, max(ws_t.max_row + 1, 51)):
    for c in range(1, 50):
        ws_t.cell(row=r, column=c, value='')

# -------------------------------------------------------------
# 3. POPULATE PROJECT TEAM MEMBERS SHEET
# -------------------------------------------------------------
for r in range(5, max(ws_m.max_row + 1, 51)):
    for c in range(1, 15):
        ws_m.cell(row=r, column=c, value='')

start_row_m = 5
m_count = 0
for i in range(len(df_src)):
    row = df_src.iloc[i]
    j_name = str(row['Job Name']).strip()
    off_raw = row['Office (Proj Organization)']
    s_date = parse_date_fixed(row['Start Date'])

    # Project Manager derived from Office:
    # TUL -> Alex Borio, SPG -> Darren Holliman, LR -> Jason Draper
    pm_name, pm_email, pm_hire_date = get_pm_info(off_raw)
    pm_start_date = get_later_date(pm_hire_date, s_date)

    # 1. Project Manager Team Member
    r_idx = start_row_m + m_count
    ws_m.cell(row=r_idx, column=1, value=j_name)
    ws_m.cell(row=r_idx, column=3, value=pm_name)
    ws_m.cell(row=r_idx, column=4, value=pm_email)
    ws_m.cell(row=r_idx, column=5, value="Project Manager")
    ws_m.cell(row=r_idx, column=6, value=pm_start_date)
    m_count += 1

    # 2. Sales Lead Team Member (Default: Air Balance)
    # Air Balance Enterprise Resource From Date is 07-28-2026 (2026/07/28).
    # Start date must be whichever is later between hire/from date (2026/07/28) and project start date.
    sales_hire_date = MAPPING_SPECIFICATION["Project Team Members"]["Sales Lead"].get("HireDate", "2026/07/28")
    sales_start_date = get_later_date(sales_hire_date, s_date)

    r_idx = start_row_m + m_count
    ws_m.cell(row=r_idx, column=1, value=j_name)
    ws_m.cell(row=r_idx, column=3, value="Air Balance")
    ws_m.cell(row=r_idx, column=4, value="air.balance@airetech.com")
    ws_m.cell(row=r_idx, column=5, value="Sales Lead")
    ws_m.cell(row=r_idx, column=6, value=sales_start_date)
    ws_m.cell(row=r_idx, column=8, value=100)
    m_count += 1

safe_save(wb, out_populated)
if p_ref and p_ref.endswith('.xlsm'):
    safe_save(wb, out_xlsm)
print(f"=== PROJECT IMPORT FBDI GENERATION COMPLETE: {out_populated} ===")


