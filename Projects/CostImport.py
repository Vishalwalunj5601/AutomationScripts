import os
import pandas as pd
import warnings
from datetime import datetime

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ================= CONFIG =================
INPUT_DIR = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\03 Project Financial Workbooks for active projects"
OUTPUT_FILE = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\DataFiles\Cost\PROD_COST_V1.xlsx"

SHEET_NAME = "Project Schedule"
TURNOVER_SHEET_KEYWORD = "turnover information"

NATURAL_ACC_FILE = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\Pre\NATRUAL_ACC.xlsx"

RAW_COST_CREDIT_TEMPLATE = {
    "Richmond-C&J-Building Automation": "1230.1215.250.625.{NaturalAccount}.0000.000.1025.0000",
    "Richmond-C&J-Services": "1230.1215.200.550.{NaturalAccount}.0000.000.1065.0000",
}

TRANSACTION_TYPE = "MISCELLANEOUS"
BUSINESS_UNIT_NAME = "CJBS"
TPA_SOURCE = "Project Miscellaneous Costs"
DOCUMENT = "Project Miscellaneous Costs"
DOCUMENT_ENTRY = "Project Miscellaneous Cost"

EXPENDITURE_ORG_SERVICE = "Richmond-C&J-Services"
EXPENDITURE_ORG_NON_SERVICE = "Richmond-C&J-Building Automation"

UOM_DEFAULT = "Currency"
BILLABLE_DEFAULT = "Y"


# ================= HELPERS =================
def normalize(val):
    if pd.isna(val):
        return ""
    s = str(val)
    s = s.replace("\xa0", " ").replace("\u2007", " ").replace("\u202f", " ")
    s = " ".join(s.split())
    return s.strip()


def normalize_key(val):
    return normalize(val).lower()


def find_turnover_sheet(xls):
    for sheet in xls.sheet_names:
        if TURNOVER_SHEET_KEYWORD in normalize_key(sheet):
            return sheet
    return None


def extract_project_type_from_turnover(xls):
    sheet = find_turnover_sheet(xls)
    if not sheet:
        return ""
    df = pd.read_excel(xls, sheet_name=sheet, header=None)
    for _, row in df.iterrows():
        if normalize_key(row.iloc[0]) in {"project type", "project type:"}:
            vals = row.iloc[1:].dropna().tolist()
            if vals:
                return " ".join(map(str, vals)).strip()
    return ""


def extract_project_name_number_from_schedule(df):
    project_name = ""
    project_number = ""
    for i in range(min(40, len(df))):
        row = df.iloc[i].astype(str).tolist()
        if "Project Name:" in row:
            project_name = normalize(row[row.index("Project Name:") + 1])
        if "Project Number:" in row:
            project_number = normalize(row[row.index("Project Number:") + 1])
    return project_name, project_number


# 🔥 NEW: Mileage inputs extractor
def extract_mileage_inputs(xls):
    cost_per_trip = 0.0
    labor_days = 0.0

    # Project Schedule B9
    try:
        proj_df = pd.read_excel(xls, sheet_name="Project Schedule", header=None)
        val = proj_df.iloc[8, 1]  # B9
        cost_per_trip = float(val) if pd.notna(val) else 0.0
    except:
        pass

    # Labor Schedule A19
    try:
        labor_sheet = [s for s in xls.sheet_names if "labor schedule" in normalize_key(s)]
        if labor_sheet:
            labor_df = pd.read_excel(xls, sheet_name=labor_sheet[0], header=None)
            val = labor_df.iloc[18, 0]  # A19
            labor_days = float(val) if pd.notna(val) else 0.0
    except:
        pass

    return cost_per_trip, labor_days


def find_schedule_table_header(df):
    for i in range(len(df)):
        if normalize_key(df.iloc[i, 0]) == "task description":
            return i
    return None


def add_line(lines, task_name, amount, exp_type):
    amount = round(float(amount), 2)
    if amount != 0:
        lines.append({"Task Name": task_name, "Amount": amount, "Expenditure Type": exp_type})


def load_natural_account_map(path):
    df = pd.read_excel(path)
    cols = {c.strip().lower(): c for c in df.columns}
    exp_col = cols.get("expenditure type")
    nat_col = cols.get("natural account")

    mapping = {}
    for _, r in df.iterrows():
        k = normalize_key(r[exp_col])
        v = normalize(r[nat_col])
        if k and v:
            mapping[k] = v
    return mapping


def build_raw_cost_credit_account(exp_org, natural_account):
    template = RAW_COST_CREDIT_TEMPLATE.get(exp_org, "")
    if not template or not natural_account:
        return ""
    return template.replace("{}", str(natural_account).strip())


def detect_task_from_row(row, scan_cols=10):
    cells = [normalize(row.iloc[i]) for i in range(min(scan_cols, len(row)))]
    joined = " ".join([c for c in cells if c]).lower()

    if "material shipping" in joined:
        return "Material Shipping"
    if "material tax to be paid" in joined:
        return "Material Tax to be Paid"
    if "material tax" in joined:
        return "Material Tax"
    if "material" in joined:
        return "Material"

    if "subcontractor" in joined:
        return "Subcontractor"
    if "graphics" in joined:
        return "Graphics"
    if "subcontracting" in joined:
        return "Subcontracting"

    if "labor" in joined:
        return "Labor"
    if "misc" in joined:
        return "Misc Costs"
    if "mileage" in joined:
        return "Mileage"
    if "p&r" in joined or joined.strip() == "pr":
        return "P&R"
    if "warranty" in joined:
        return "Warranty"
    if "admin" in joined:
        return "Admin / OH"

    for c in cells:
        if c:
            return c
    return ""


def aggregate_schedule_costs_final(df):
    header_row = find_schedule_table_header(df)
    if header_row is None:
        return []

    header = [normalize(x) for x in df.iloc[header_row]]
    if "Cost to Date (Spent)" not in header:
        return []

    cost_col = header.index("Cost to Date (Spent)")
    data = df.iloc[header_row + 1:].copy()

    data["TaskDetected"] = data.apply(lambda r: detect_task_from_row(r), axis=1)
    data["TaskKey"] = data["TaskDetected"].apply(normalize_key)
    data["Cost"] = pd.to_numeric(data.iloc[:, cost_col], errors="coerce").fillna(0)

    totals = {}

    for _, r in data.iterrows():
        tk = r["TaskKey"]
        cost = float(r["Cost"])

        if cost == 0:
            continue

        totals[tk] = totals.get(tk, 0) + cost

    lines = []
    add_line(lines, "Material", totals.get("material", 0), "Material")
    add_line(lines, "Material", totals.get("material tax", 0), "Accrued Use Tax")
    add_line(lines, "Material", totals.get("material shipping", 0), "Freight")
    add_line(lines, "Labor", totals.get("labor", 0), "Labor")
    add_line(lines, "Subcontracting", totals.get("subcontractor", 0), "Subcontract Quotes")
    add_line(lines, "Subcontracting", totals.get("graphics", 0), "Miscellaneous Graphics Cost")
    add_line(lines, "Misc Costs", totals.get("misc costs", 0), "Miscellaneous Expenses")
    add_line(lines, "Mileage", totals.get("mileage", 0), "Mileage")
    add_line(lines, "P&R", totals.get("p&r", 0), "Risk")
    add_line(lines, "Warranty", totals.get("warranty", 0), "Project Warranty")
    add_line(lines, "Admin / OH", totals.get("admin / oh", 0), "Admin / OH")

    return lines


# ================= MAIN =================
natural_map = load_natural_account_map(NATURAL_ACC_FILE)
today_str = datetime.today().strftime("%Y/%m/%d")
rows_out = []

for file in os.listdir(INPUT_DIR):
    if not file.lower().endswith((".xlsx", ".xlsm")):
        continue

    path = os.path.join(INPUT_DIR, file)
    xls = pd.ExcelFile(path)

    sched_df = pd.read_excel(path, sheet_name=SHEET_NAME, header=None)
    _, project_number = extract_project_name_number_from_schedule(sched_df)

    project_type = extract_project_type_from_turnover(xls)
    is_service = "service" in str(project_type).lower()
    exp_org = EXPENDITURE_ORG_SERVICE if is_service else EXPENDITURE_ORG_NON_SERVICE

    # 🔥 NEW Mileage Calculation
    cost_per_trip, labor_days = extract_mileage_inputs(xls)
    calculated_mileage = round(cost_per_trip * labor_days, 2)

    lines = aggregate_schedule_costs_final(sched_df)

    # 🔥 OVERRIDE Mileage
    # 🔥 FORCE Mileage (independent of schedule)

    mileage_found = False

    for ln in lines:
        if normalize_key(ln["Expenditure Type"]) == "mileage":
            ln["Amount"] = calculated_mileage
            mileage_found = True

    # ✅ If not found → ADD it
    if not mileage_found and calculated_mileage != 0:
        lines.append({
            "Task Name": "Mileage",
            "Amount": calculated_mileage,
            "Expenditure Type": "Mileage"
        })

    exp_batch = f"DATA_CONVERSION_01_{project_number}"

    for ln in lines:
        amt = ln["Amount"]
        if round(amt, 2) == 0:
            continue

        exp_type = ln["Expenditure Type"]
        natural_acc = natural_map.get(normalize_key(exp_type), "")
        raw_cost_credit_acct = build_raw_cost_credit_account(exp_org, natural_acc)

        rows_out.append({
            "* Transaction Type": TRANSACTION_TYPE,
            "Business Unit Name": BUSINESS_UNIT_NAME,
            "Third-Party Application Transaction Source": TPA_SOURCE,
            "Document": DOCUMENT,
            "Document Entry": DOCUMENT_ENTRY,
            "* Expenditure Batch": exp_batch,
            "* Expenditure Item Date": today_str,
            "Project Number": project_number,
            "Task Name": ln["Task Name"],
            "Expenditure Type": exp_type,
            "Expenditure Organization": exp_org,
            "* Quantity": amt,
            "Unit of Measure": UOM_DEFAULT,
            "Billable": BILLABLE_DEFAULT,
            "Natural Account": natural_acc,
            "Raw Cost Credit Account": raw_cost_credit_acct
        })

out_df = pd.DataFrame(rows_out)
out_df.to_excel(OUTPUT_FILE, index=False)

print("✅ Done with NEW Mileage Logic (B9 * A19)")
print(f"📊 Rows: {len(out_df)}")