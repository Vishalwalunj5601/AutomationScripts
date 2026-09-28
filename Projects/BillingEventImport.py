import os
import re
import pandas as pd
import warnings
from datetime import datetime

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ===================== CONFIG =====================
DATA_F_DIR = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\03 Project Financial Workbooks for active projects"
BILLING_EVENTS_DIR = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\Pre\INVOICES"

OUTPUT_UNPAID_FILE  = os.path.join(BILLING_EVENTS_DIR, "UNPAID_CONVERSION.xlsx")
OUTPUT_INVOICE_FILE = os.path.join(BILLING_EVENTS_DIR, "PAID_INVOICE.xlsx")

TURNOVER_SHEET_KEYWORD = "turnover information"

PROJECT_NUM_REGEX = re.compile(r"\b\d{4}-\d{4}\b")

AR_FULLNAME_REQUIRED = "Accounts Receivable (A/R)"
AR_PAID_STATUS_UNPAID = "Unpaid"
AR_PAID_STATUS_PAID = "Paid"


# ===================== HELPERS =====================
def normalize_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def normalize_key(x):
    return normalize_str(x).lower()


def to_number(x):
    if pd.isna(x):
        return None
    try:
        s = str(x).replace(",", "").replace("$", "").strip()
        if s == "":
            return None
        return float(s)
    except:
        return None


def format_date_yyyymmdd(x):
    if pd.isna(x):
        return None
    if isinstance(x, (pd.Timestamp, datetime)):
        return x.strftime("%m/%d/%Y")
    dt = pd.to_datetime(x, errors="coerce")
    if pd.isna(dt):
        return None
    return dt.strftime("%m/%d/%Y")


def extract_project_number_from_customer(customer_text: str):
    m = PROJECT_NUM_REGEX.search(customer_text or "")
    return m.group(0) if m else None


def find_turnover_sheet(xls: pd.ExcelFile):
    for sheet in xls.sheet_names:
        if TURNOVER_SHEET_KEYWORD in normalize_key(sheet):
            return sheet
    return None


def get_project_number_from_dataf_file(file_path: str):
    try:
        xls = pd.ExcelFile(file_path)
    except:
        return None

    sheet = find_turnover_sheet(xls)
    if not sheet:
        return None

    try:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
    except:
        return None

    for _, row in df.iterrows():
        if normalize_key(row.iloc[0]) == "project number:":
            if len(row) > 1 and pd.notna(row.iloc[1]):
                return str(row.iloc[1]).strip()
    return None


# ===================== STEP 1: VALID PROJECTS =====================
valid_projects = set()

print("🔎 Building project list from DATA-F...\n")

for f in os.listdir(DATA_F_DIR):
    if not f.lower().endswith((".xlsx", ".xls", ".xlsm")):
        continue

    full_path = os.path.join(DATA_F_DIR, f)
    pn = get_project_number_from_dataf_file(full_path)

    if pn:
        valid_projects.add(pn)

print(f"✅ Projects found in DATA-F: {len(valid_projects)}\n")


# ===================== STEP 2: READ BILLING EVENTS =====================
unpaid_rows = []
paid_raw_rows = []

idx_unpaid = 1

for file in os.listdir(BILLING_EVENTS_DIR):

    if not file.lower().endswith((".xlsx", ".xls", ".xlsm", ".csv")):
        continue

    full_path = os.path.join(BILLING_EVENTS_DIR, file)

    if os.path.abspath(full_path) in {
        os.path.abspath(OUTPUT_UNPAID_FILE),
        os.path.abspath(OUTPUT_INVOICE_FILE),
    }:
        continue

    print(f"📂 Processing: {file}")

    try:
        if file.lower().endswith(".csv"):
            df = pd.read_csv(full_path, dtype=str).fillna("")
        else:
            df = pd.read_excel(full_path, dtype=str).fillna("")
    except Exception as e:
        print(f"❌ Failed to read {file}: {e}")
        continue

    df.columns = [str(c).strip() for c in df.columns]

    required_cols = [
        "Customer",
        "Transaction type",
        "Open balance amount",
        "Amount",
        "Transaction date",
        "Num",
        "Name",
        "Memo/Description",
        "Full name",
        "A/R paid",
    ]

    if not all(col in df.columns for col in required_cols):
        continue

    for _, r in df.iterrows():

        customer_val = normalize_str(r.get("Customer"))
        txn_type = normalize_str(r.get("Transaction type")).upper()

        open_bal = to_number(r.get("Open balance amount"))
        amount_val = to_number(r.get("Amount"))

        trx_date = format_date_yyyymmdd(r.get("Transaction date"))

        full_name = normalize_str(r.get("Full name"))
        ar_paid = normalize_str(r.get("A/R paid"))

        project_number = extract_project_number_from_customer(customer_val)

        if not project_number:
            continue

        if project_number not in valid_projects:
            continue

        if txn_type != "INVOICE":
            continue

        if full_name != AR_FULLNAME_REQUIRED:
            continue

        # ================= UNPAID =================
        if ar_paid.lower() == AR_PAID_STATUS_UNPAID.lower():

            if open_bal is None or open_bal <= 0:
                continue

            desc = f"{r['Num']}-{r['Name']}-{r['Memo/Description']}".strip("-")

            unpaid_rows.append({
                "Source Name": "Data Conversion",
                "Source Reference": f"{idx_unpaid}_{project_number}",
                "Organization": "CJBS",
                "*Contract Type": "AIR_Customer Contract for Project",
                "*Contract Number": project_number,
                "*Contract Line Number": 1,
                "*Event Type": "Conversion",
                "Description": desc,
                "Completion Date": trx_date,
                "*Bill Transaction Currency": "USD",
                "Event Amount in Bill Transaction Currency": open_bal,
                "Project Number": project_number
            })

            idx_unpaid += 1

        # ================= PAID =================
        elif ar_paid.lower() == AR_PAID_STATUS_PAID.lower():

            if amount_val is None or amount_val <= 0:
                continue

            paid_raw_rows.append({
                "Project Number": project_number,
                "Amount": amount_val,
                "Completion Date": trx_date
            })


# ===================== STEP 3: WRITE UNPAID =====================
unpaid_df = pd.DataFrame(unpaid_rows)

if not unpaid_df.empty:
    unpaid_df = unpaid_df.sort_values(["Project Number"])

unpaid_df.to_excel(OUTPUT_UNPAID_FILE, index=False)

print(f"\n✅ UNPAID Rows: {len(unpaid_df)}")


# ===================== STEP 4: SUM PAID =====================
paid_df = pd.DataFrame(paid_raw_rows)

invoice_rows = []

if not paid_df.empty:

    grouped = paid_df.groupby("Project Number")["Amount"].sum().reset_index()

    idx_invoice = 1

    for _, r in grouped.iterrows():

        invoice_rows.append({
            "Source Name": "Data Conversion",
            "Source Reference": f"INV_{idx_invoice}_{r['Project Number']}",
            "Organization": "CJBS",
            "*Contract Type": "AIR_Customer Contract for Project",
            "*Contract Number": r["Project Number"],
            "*Contract Line Number": 1,
            "*Event Type": "Invoice",
            "Description": "Paid",
            "Completion Date": None,
            "*Bill Transaction Currency": "USD",
            "Event Amount in Bill Transaction Currency": r["Amount"],
            "Project Number": r["Project Number"]
        })

        idx_invoice += 1

invoice_df = pd.DataFrame(invoice_rows)

invoice_df.to_excel(OUTPUT_INVOICE_FILE, index=False)

print(f"✅ PAID rows after SUM: {len(invoice_df)}")
print("🎉 DONE")