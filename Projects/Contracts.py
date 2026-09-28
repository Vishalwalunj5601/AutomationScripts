import os
import pandas as pd
import warnings
from datetime import datetime
from dateutil.relativedelta import relativedelta

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ================== CONFIG ==================
DATA_DIR = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\03 Project Financial Workbooks for active projects"
OUTPUT_FILE = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\DataFiles\Contracts\ContractsF_Load.xlsx"

# DFF source for AttributeNumber2/3
DFF_FILE = r"C:\Users\Vishal Walunj\Desktop\customerdata\AIR-MOCK\Projects\Projects\DATA-FILE_V2\CONTRACT_DFF.xlsx"

# Active project list (for Sales Reps + Sell Price)
ACTIVE_PROJECT_FILE = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\CJBS Active Project List_06_25_26.xlsx"

# Fusion customer/site/address reference file
FUSION_FILE = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\Pre\fusion.xlsx"

# Turnover fields we care about
FIELDS = {
    "Project Name": "project name",
    "Project Number": "project number",
    "Project Type": "project type",
    "Project Start Date": "project start date",
    "Customer": "customer",
    "Project Manager": "project manager"
}

# ================== GENERIC HELPERS ==================
def normalize(text):
    return str(text).lower().replace(":", "").strip()

def to_number(val):
    if pd.isna(val):
        return None
    try:
        s = str(val).replace(",", "").replace("$", "").strip()
        if s == "":
            return None
        return float(s)
    except Exception:
        return None

def format_date_mmddyyyy(value):
    if pd.isna(value):
        return None

    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%m/%d/%Y")

    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%m/%d/%Y")

def parse_date(value):
    """Return datetime or None."""
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime()

def find_turnover_sheet(file_path):
    xls = pd.ExcelFile(file_path)
    for sheet in xls.sheet_names:
        low = sheet.lower()
        if "turnover" in low and "step" in low:
            return sheet
    return None

def find_project_schedule_sheet(file_path):
    xls = pd.ExcelFile(file_path)
    for sheet in xls.sheet_names:
        low = sheet.lower()
        if "project" in low and "schedule" in low:
            return sheet
    return None

def find_billing_invoice_sheet(file_path):
    xls = pd.ExcelFile(file_path)
    for sheet in xls.sheet_names:
        low = sheet.lower()
        if "billing" in low and "invoice" in low:
            return sheet
    return None

def find_sov_step3_sheet(file_path):
    xls = pd.ExcelFile(file_path)
    for sheet in xls.sheet_names:
        low = sheet.lower()
        if "sov" in low and "step" in low:
            return sheet
    return None

def find_billing_sov_sheet(file_path):
    """
    Finds billing SOV sheet.
    Looks for names like:
    - Billing SOV
    - Billing Schedule of Values
    """
    xls = pd.ExcelFile(file_path)
    for sheet in xls.sheet_names:
        low = sheet.lower()
        if "billing" in low and "sov" in low:
            return sheet
        if "billing" in low and "schedule of values" in low:
            return sheet
    return None

# ---------------- SALES PRICE VALIDATION ----------------
def is_valid_sales_price(val):
    """
    Sell Price must be a valid positive number.
    If invalid / blank / non-numeric / 0 -> return False.
    """
    try:
        if pd.isna(val):
            return False
        s = str(val).replace(",", "").strip()
        if s == "":
            return False
        num = float(s)
        return num > 0
    except Exception:
        return False

# ================== LOAD CONTRACT_DFF FOR HEADER DFF ==================
dff_df = pd.read_excel(DFF_FILE)
dff_df.columns = [c.strip() for c in dff_df.columns]

def get_header_dff_amount(project_number: str):
    """
    From CONTRACT_DFF.xlsx:
      - PROJECT column contains ':{PROJECT_NUMBER}' (anywhere in the string)
      - Transaction Type = INVOICE  (case-insensitive)
      - A/R Paid = PAID            (case-insensitive)
      - Sum Amount
    """
    if dff_df.empty:
        return None

    df = dff_df.copy()

    df["PROJECT_str"] = df["PROJECT"].astype(str)
    df["TransactionType_upper"] = df["Transaction Type"].astype(str).str.upper().str.strip()
    df["ARPaid_upper"] = df["A/R Paid"].astype(str).str.upper().str.strip()

    pattern = f":{str(project_number).strip()}"
    project_mask = df["PROJECT_str"].str.contains(pattern, case=False, na=False)

    mask = (
        project_mask &
        (df["TransactionType_upper"] == "INVOICE") &
        (df["ARPaid_upper"] == "PAID")
    )

    sub = df.loc[mask]
    if sub.empty:
        return None

    total = 0.0
    for val in sub["Amount"]:
        num = to_number(val)
        if num is not None:
            total += num

    return total if total != 0 else None

# ================== LOAD ACTIVE PROJECT LIST (SALES REPS + SELL PRICE) ==================
sales_info_map = {}  # {project_number: {"reps": [..], "sell_price": ..}}

def load_active_project_sales():
    global sales_info_map
    if not os.path.exists(ACTIVE_PROJECT_FILE):
        print(f"⚠️ Active project file not found: {ACTIVE_PROJECT_FILE}")
        sales_info_map = {}
        return

    df = pd.read_excel(ACTIVE_PROJECT_FILE)
    df.columns = [c.strip() for c in df.columns]

    pn_col = "Project Number"
    rep1_col = "Sales Rep 1"
    rep2_col = "Sales Rep 2"

    price_col = None
    for cand in ["Sell Price", "Sell price", "Sale Price", "Sale price"]:
        if cand in df.columns:
            price_col = cand
            break

    if pn_col not in df.columns:
        print("⚠️ 'Project Number' column not found in Active Project List.")
        sales_info_map = {}
        return

    for _, row in df.iterrows():
        proj_num_raw = row.get(pn_col)
        if pd.isna(proj_num_raw):
            continue
        proj_num = str(proj_num_raw).strip()
        if not proj_num or proj_num.lower() == "nan":
            continue

        reps = []
        for col in [rep1_col, rep2_col]:
            if col not in df.columns:
                continue
            val = row.get(col)
            if pd.isna(val):
                continue
            s = str(val).strip()
            if not s:
                continue
            if s in {"0", "0.0"}:
                continue
            if s.lower() in {"none", "null"}:
                continue
            reps.append(s)

        sell_price = row.get(price_col) if price_col else None

        sales_info_map[proj_num] = {"reps": reps, "sell_price": sell_price}

    print(f"✅ Loaded sales info for {len(sales_info_map)} projects from Active Project List.")

load_active_project_sales()

# ================== LOAD FUSION CUSTOMER / SITE MAP ==================
fusion_map = {}  # {project_number: {"bill_to": {...} or None, "ship_to": {...} or None}}

def load_fusion_map():
    global fusion_map
    fusion_map = {}

    if not os.path.exists(FUSION_FILE):
        print(f"⚠️ Fusion file not found: {FUSION_FILE}")
        return

    df = pd.read_excel(FUSION_FILE)
    df.columns = [c.strip() for c in df.columns]

    required = ["Project Number", "PARTY_NAME", "ACCOUNT_NUMBER", "SITE_USE_CODE", "LOCATION", "ADDRESS1"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"⚠️ Fusion file missing columns: {missing}")
        return

    df["ProjectNumber_str"] = df["Project Number"].astype(str).str.strip()
    df["SITE_USE_CODE_upper"] = df["SITE_USE_CODE"].astype(str).str.upper().str.strip()

    for proj_num, g in df.groupby("ProjectNumber_str"):
        bill = g[g["SITE_USE_CODE_upper"] == "BILL_TO"].head(1)
        ship = g[g["SITE_USE_CODE_upper"] == "SHIP_TO"].head(1)

        def row_to_dict(rdf):
            if rdf.empty:
                return None
            r = rdf.iloc[0]
            return {
                "party_name": "" if pd.isna(r["PARTY_NAME"]) else str(r["PARTY_NAME"]).strip(),
                "account_number": "" if pd.isna(r["ACCOUNT_NUMBER"]) else str(r["ACCOUNT_NUMBER"]).strip(),
                "location": "" if pd.isna(r["LOCATION"]) else str(r["LOCATION"]).strip(),
                "address1": "" if pd.isna(r["ADDRESS1"]) else str(r["ADDRESS1"]).strip(),
            }

        fusion_map[proj_num] = {
            "bill_to": row_to_dict(bill),
            "ship_to": row_to_dict(ship)
        }

    print(f"✅ Loaded fusion mapping for {len(fusion_map)} projects.")

load_fusion_map()

def get_fusion_bill_ship(project_number: str):
    """
    Returns (bill_to_dict, ship_to_dict).
    If not found -> (None, None)
    """
    key = str(project_number).strip()
    rec = fusion_map.get(key)
    if not rec:
        return None, None
    return rec.get("bill_to"), rec.get("ship_to")

# ================== FIELD EXTRACTORS FROM PROJECT FILES ==================
def extract_basic_fields_retention_duration(file_path):
    """
    From Turnover sheet:
      - Project Name, Number, Type, Start Date, Customer
      - Retention % (label contains 'retention' and '%')
      - Project Duration months (label contains 'duration' and 'month')
    """
    sheet = find_turnover_sheet(file_path)
    if not sheet:
        return None, None, None

    df = pd.read_excel(file_path, sheet_name=sheet, header=None)

    result = {key: None for key in FIELDS}
    retention = None
    duration_months = None

    for i in range(len(df)):
        label = normalize(df.iloc[i, 0])

        for field, expected in FIELDS.items():
            if expected and label == expected:
                values = df.iloc[i, 1:].dropna().tolist()
                if values:
                    if field == "Project Start Date":
                        result[field] = format_date_mmddyyyy(values[0])
                    else:
                        result[field] = " ".join(map(str, values))

        if "retention" in label and "%" in label:
            values = df.iloc[i, 1:].dropna().tolist()
            if values:
                retention = str(values[0]).strip()

        if ("duration" in label) and ("month" in label):
            values = df.iloc[i, 1:].dropna().tolist()
            if values:
                dm = to_number(values[0])
                if dm is not None:
                    duration_months = int(dm)

    return result, retention, duration_months

def extract_total_completed_stored_to(file_path):
    sheet = find_billing_invoice_sheet(file_path)
    if not sheet:
        return None

    df = pd.read_excel(file_path, sheet_name=sheet, header=None)

    for i in range(len(df)):
        label = normalize(df.iloc[i, 0])
        if "total completed" in label and "stored to" in label:
            row = df.iloc[i]
            numeric_vals = []
            for val in row[1:]:
                num = to_number(val)
                if num is not None:
                    numeric_vals.append(num)
            if numeric_vals:
                return numeric_vals[-1]
            break

    return None

def extract_original_sell_and_change_orders(file_path):
    sheet = find_project_schedule_sheet(file_path)
    if not sheet:
        return None, None

    df = pd.read_excel(file_path, sheet_name=sheet, header=None)

    original_sell = None
    change_orders = None

    for i in range(len(df)):
        label = normalize(df.iloc[i, 0])

        if "original sell price" in label:
            original_sell = to_number(df.iloc[i, 1])

        if "change orders to" in label:
            change_orders = to_number(df.iloc[i, 1])

    return original_sell, change_orders

def extract_sov_lines(file_path):
    sheet = find_sov_step3_sheet(file_path)
    if not sheet:
        return []

    df = pd.read_excel(file_path, sheet_name=sheet, header=None)

    header_row_idx = None
    price_col = None

    for i in range(len(df)):
        row = df.iloc[i]
        for j, val in enumerate(row):
            if isinstance(val, str) and "schedule of values on customer bill" in val.lower():
                header_row_idx = i
                break
        if header_row_idx is not None:
            break

    if header_row_idx is None:
        return []

    for i in range(header_row_idx, min(header_row_idx + 5, len(df))):
        row = df.iloc[i]
        for j, val in enumerate(row):
            if isinstance(val, str) and "sell price" in val.lower():
                price_col = j
                header_row_idx = i
                break
        if price_col is not None:
            break

    if price_col is None:
        return []

    sov_lines = []
    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i]
        amount = to_number(row[price_col])
        if amount is None:
            continue

        text_before = " ".join(
            str(v).lower()
            for v in row[:price_col]
            if isinstance(v, str)
        )
        if "total" in text_before:
            continue

        desc = row[price_col - 1] if price_col - 1 >= 0 else ""
        if pd.isna(desc):
            text_desc = ""
            for val in row:
                if isinstance(val, str) and val.strip():
                    text_desc = val.strip()
                    break
            desc = text_desc

        sov_lines.append((str(desc).strip(), float(amount)))

    return sov_lines

def extract_po_number_from_billing_sov(file_path):
    """
    Extract PO Number from Billing SOV sheet.
    Example expected text:
        SUBCONTRACT / PO:    AS25-0620
    Returns:
        AS25-0620
    """
    sheet = find_billing_sov_sheet(file_path)
    if not sheet:
        return None

    df = pd.read_excel(file_path, sheet_name=sheet, header=None)

    for i in range(len(df)):
        for j in range(df.shape[1]):
            val = df.iloc[i, j]
            if pd.isna(val):
                continue

            text = str(val).strip()
            low = text.lower()

            if "subcontract / po" in low:
                # Case 1: value is in same cell -> "SUBCONTRACT / PO: AS25-0620"
                if ":" in text:
                    right = text.split(":", 1)[1].strip()
                    if right:
                        return right

                # Case 2: next column has PO number
                if j + 1 < df.shape[1]:
                    next_val = df.iloc[i, j + 1]
                    if not pd.isna(next_val):
                        next_text = str(next_val).strip()
                        if next_text:
                            return next_text

    return None

# ================== PROJECT END DATE LOGIC ==================
def calc_project_end_date_mmddyyyy(start_date_mmddyyyy: str, duration_months: int):
    """
    End Date = Project Start Date + Project Duration (months)

    If calculated end date <= 2026/06/30
    then use 2026/09/30
    """
    if not start_date_mmddyyyy or duration_months is None:
        return None

    start_dt = parse_date(start_date_mmddyyyy)
    if not start_dt:
        return None

    calc_end = start_dt + relativedelta(months=+int(duration_months))

    cutoff_date = datetime(2026, 6, 30)
    forced_end_date = datetime(2026, 9, 30)

    if calc_end <= cutoff_date:
        final_end = forced_end_date
    else:
        final_end = calc_end

    return final_end.strftime("%m/%d/%Y")

# ================== MAIN PER-FILE PROCESSING ==================
def process_file(file_path):
    base_name = os.path.basename(file_path)

    basic, retention, duration_months = extract_basic_fields_retention_duration(file_path)
    if not basic:
        print(f"⚠️  Skipping {base_name}: Turnover sheet not found.")
        return None

    project_name = basic["Project Name"]
    project_number = basic["Project Number"]
    project_type = basic["Project Type"] or ""
    start_date = basic["Project Start Date"]  # MM/DD/YYYY
    source_customer = basic["Customer"]
    project_manager = basic.get("Project Manager")
    contract_puid = f"PUID_{project_number}"
    party_role_puid = f"SPUID_{project_number}"
    line_puid = f"LINE_PUID_{project_number}"

    dff2_total_ar = extract_total_completed_stored_to(file_path)
    sov_lines = extract_sov_lines(file_path)
    original_sell, change_orders = extract_original_sell_and_change_orders(file_path)
    po_number = extract_po_number_from_billing_sov(file_path)

    total_sov_amount = sum(amount for _, amount in sov_lines)

    if original_sell is not None and change_orders is not None and total_sov_amount:
        expected_total = original_sell + change_orders
        diff = abs(expected_total - total_sov_amount)
        if diff > 0.01:
            print(
                f"❌ WARNING in {base_name} (Project {project_number}): "
                f"SOV total {total_sov_amount:,.2f} "
                f"!= Original Sell + Change Orders {expected_total:,.2f}"
            )
        else:
            print(
                f"✅ SOV total matches for {base_name} (Project {project_number}). "
                f"Total = {total_sov_amount:,.2f}"
            )
    else:
        if sov_lines:
            print(
                f"⚠️  Could not validate totals (missing Original Sell / Change Orders) "
                f"in {base_name} (Project {project_number})."
            )

    attr_amount = get_header_dff_amount(project_number)

    # -------- FUSION LOOKUP --------
    bill_to, ship_to = get_fusion_bill_ship(project_number)

    bill_party = bill_to["party_name"] if bill_to else ""
    bill_acct = bill_to["account_number"] if bill_to else ""
    bill_loc = bill_to["location"] if bill_to else ""
    bill_addr = bill_to["address1"] if bill_to else ""

    ship_party = ship_to["party_name"] if ship_to else ""
    ship_acct = ship_to["account_number"] if ship_to else ""
    ship_loc = ship_to["location"] if ship_to else ""
    ship_addr = ship_to["address1"] if ship_to else ""

    customer_for_contract = bill_party or ship_party or (source_customer or "")

    end_date = calc_project_end_date_mmddyyyy(start_date, duration_months)

    is_service = "service" in project_type.lower()

    # -------- SALES CREDIT PREP --------
    header_sales_credit_rows = []
    line_sales_credit_rows = []

    sales_info = sales_info_map.get(project_number, {"reps": [], "sell_price": None})
    reps = sales_info.get("reps", [])
    sell_price = sales_info.get("sell_price", None)

    if is_valid_sales_price(sell_price) and len(reps) > 0:
        n = len(reps)
        percent_each = round(100.0 / n, 2)

        for idx, rep_name in enumerate(reps, start=1):
            sc_puid = f"SC{idx}_PUID_{project_number}"

            header_sales_credit_rows.append({
                "ContractPuid": contract_puid,
                "SalesCreditPuid": sc_puid,
                "ExternalSource": f"SBP_{project_number}",
                "ExternalKey": f"KBP_{project_number}",
                "Percent": percent_each,
                "SalesrepName": rep_name,
                "SalesCreditType": "Quota Sales Credit",
                "StartDate": start_date
            })

            line_sales_credit_rows.append({
                "LinePuid": line_puid,
                "SalesCreditPuid": sc_puid,
                "ExternalSource": f"SBP_{project_number}",
                "ExternalKey": f"KBP_{project_number}",
                "SalesrepName": rep_name,
                "SalesCreditType": "Non-quota Sales Credit",
                "StartDate": start_date,
                "Percent": percent_each
            })

        contract_sales_rep = ", ".join(reps)

    else:
        default_rep = "Christopher Desoto"
        sc_puid = f"SC1_PUID_{project_number}"

        header_sales_credit_rows.append({
            "ContractPuid": contract_puid,
            "SalesCreditPuid": sc_puid,
            "ExternalSource": f"SBP_{project_number}",
            "ExternalKey": f"KBP_{project_number}",
            "Percent": 100.0,
            "SalesrepName": default_rep,
            "SalesCreditType": "Quota Sales Credit",
            "StartDate": start_date
        })

        line_sales_credit_rows.append({
            "LinePuid": line_puid,
            "SalesCreditPuid": sc_puid,
            "ExternalSource": f"SBP_{project_number}",
            "ExternalKey": f"KBP_{project_number}",
            "SalesrepName": default_rep,
            "SalesCreditType": "Non-quota Sales Credit",
            "StartDate": start_date,
            "Percent": 100.0
        })

        contract_sales_rep = default_rep

        print(
            f"ℹ️ Using default Sales Rep '{default_rep}' for project {project_number} "
            f"(valid_price={is_valid_sales_price(sell_price)}, reps={reps})"
        )

    # -------- CONTRACT HEADER --------
    contract_row = {
        "ContractTypeName": "AIR_Customer Contract for Project",
        "StsCode": "ACTIVE",
        "CurrencyCode": "USD",
        "StartDate": start_date,
        "EndDate": end_date,
        "LegalEntityName": "C&J Building Solutions, LLC",
        "ContractPuid": contract_puid,
        "OrgName": "CJBS",
        "InvConvRateType": "Corporate",
        "InvTrxTypeName": "PA Invoice" if is_service else "AIA-PA-Invoice",

        "BillToAccount": bill_party,
        "BillToAccountNumber": bill_acct,
        "GeneratedInvoiceStatus": "R" if is_service else "D",
        "ContractNumber": project_number,
        "Cognomen": project_name,

        "PartyName": bill_party,
        "BillToSiteUseLocation": bill_loc,
        "BillToAddress": bill_addr,

        "ShipToAccount": ship_party,
        "ShipToAccountNumber": ship_acct,
        "ShipToSiteUseLocation": ship_loc,
        "ShipToAddress": ship_addr,

        "OwningOrgName": "Richmond-C&J-Services" if is_service
                         else "Richmond-C&J-Building Automation",
        "AuthoringPartyMeaning": "Internal",
        "TaxExemptionControlMeaning": "Standard",
        "LedgerCurrency": "USD",
        "AuthoringPartyCode": "INTERNAL",
        "AccessLevel": "CREATE",
        "ContributionPercent": 100,

        "Description": f"Retention %:{retention}" if retention else None,

        "DFF1 Retainange": retention,
        "DFF2  : Total AR Inv Amt": dff2_total_ar,
        "DFF3: Recent Receipt": None,

        "AttributeNumber2": attr_amount,
        "AttributeNumber3": attr_amount,

        # NEW
        "SalesRepName": contract_sales_rep,
        "PONumber": po_number
    }

    # -------- CONTRACT PARTY --------
    contract_party_rows = [
        {
            "ContractPuid": contract_puid,
            "PartyRolePuid": f"CPUID_{project_number}",
            "Role": "All eligible customers",
            "PartyName": customer_for_contract
        },
        {
            "ContractPuid": contract_puid,
            "PartyRolePuid": party_role_puid,
            "Role": "Supplier",
            "PartyName": "CJBS"
        }
    ]

    # -------- CONTRACT PARTY CONTACT --------
    party_contact_row = {
        "PartyRolePuid": party_role_puid,
        "ContactPuid": contract_puid,
        "Role": "Contract administrator",
        "PartyContactName": project_manager if pd.notna(project_manager) and str(project_manager).strip() else "Christopher Desoto",
        "OwnerYn": "Y",
        "StartDate": start_date,
        "AccessLevelName": "Full"
    }

    # -------- PROJECT BILL PLAN --------
    billplan_rows = []

    # Row 1 – AIR Amount Based Invoice
    billplan_rows.append({
        "ContractPuid": contract_puid,
        "BillingCycleName": "Immediate",
        "LaborInvoiceFormatName": "Labor",
        "NlInvoiceFormatName": "Nonlabor",
        "EventsInvoiceFormatName": "Event",
        "ExtSource": f"SBP_{project_number}",
        "ExtKey": f"KBP_{project_number}",
        "BillPlanLang": "US",
        "SourceLangFlag": "Y",
        "BillMethodName": "Amount Based Invoice",
        "BillMethodFlag": "I",
        "BillPlanName": "Amount Based Invoice",

        "BillSetNum": 1,
        "BillToCustAcctNumber": bill_acct,
        "BillToSiteUseLocation": bill_loc,

        "InvCurrOptMeaning": "Contract",
        "PaymentTermName": "IMMEDIATE"
    })

    # Row 2 – Second Bill Plan
    bill_method_name_second = "Amount Based Revenue" if is_service else "Amount Based Revenue"
    billplan_rows.append({
        "ContractPuid": contract_puid,
        "BillingCycleName": "",
        "LaborInvoiceFormatName": "",
        "NlInvoiceFormatName": "",
        "EventsInvoiceFormatName": "",
        "ExtSource": f"SBP01_{project_number}",
        "ExtKey": f"KBP01_{project_number}",
        "BillPlanLang": "US",
        "SourceLangFlag": "Y",
        "BillMethodName": bill_method_name_second,
        "BillMethodFlag": "",
        "BillPlanName": bill_method_name_second,

        "BillSetNum": "",
        "BillToCustAcctNumber": "",
        "BillToSiteUseLocation": "",

        "InvCurrOptMeaning": "",
        "PaymentTermName": ""
    })

    # -------- CONTRACT LINE --------
    contract_line_row = {
        "LineNumber": 1,
        "ContractPuid": contract_puid,
        "LinePuid": line_puid,
        "StsCode": "DRAFT",
        "LineTypeName": "Free-form, project",
        "StartDate": start_date,
        "EndDate": end_date,
        "ItemName": "Contract",

        "ShipToAccount": ship_party,
        "ShipToAccountNumber": ship_acct,
        "ShipToSite": ship_loc,

        "BillPlan": "Bill Plan_1",
        "RevenuePlan": "Revenue Plan_1",
        "BillPlanExternalSource": f"SBP_{project_number}",
        "BillPlanExternalKey": f"KBP_{project_number}",
        "RevenuePlanExternalSource": f"SBP01_{project_number}",
        "RevenuePlanExternalKey": f"KBP01_{project_number}",
        "LineAmount": total_sov_amount
    }

    # -------- CONTRACT-PROJECT LINK --------
    project_link_row = {
        "LinePuid": line_puid,
        "ExtSource": f"SBP_{project_number}",
        "ExtKey": f"KBP_{project_number}",
        "ProjectNumber": project_number,
        "ActiveFlag": "Y",
        "FundingAmount": total_sov_amount
    }

    # -------- HEADER DFF SHEET ROW --------
    dff_row = {
        "ContractPuid": contract_puid,
        "AttributeNumber2": attr_amount,
        "AttributeNumber3": attr_amount
    }

    return (
        contract_row,
        contract_party_rows,
        party_contact_row,
        billplan_rows,
        contract_line_row,
        project_link_row,
        dff_row,
        header_sales_credit_rows,
        line_sales_credit_rows
    )

# ================== RUN FOR ALL FILES ==================
contract_rows = []
contract_party_rows_all = []
party_contact_rows_all = []
billplan_rows_all = []
contract_line_rows_all = []
project_link_rows_all = []
header_dff_rows_all = []
header_sales_credit_rows_all = []
line_sales_credit_rows_all = []

for file in os.listdir(DATA_DIR):
    if not file.lower().endswith((".xlsx", ".xls", ".xlsm")):
        continue

    full_path = os.path.join(DATA_DIR, file)
    if os.path.abspath(full_path) == os.path.abspath(OUTPUT_FILE):
        continue

    result = process_file(full_path)
    if not result:
        continue

    (
        contract_row,
        contract_parties,
        party_contact_row,
        billplans,
        contract_line_row,
        project_link_row,
        dff_row,
        header_sales_credit_rows,
        line_sales_credit_rows
    ) = result

    contract_rows.append(contract_row)
    contract_party_rows_all.extend(contract_parties)
    party_contact_rows_all.append(party_contact_row)
    billplan_rows_all.extend(billplans)
    contract_line_rows_all.append(contract_line_row)
    project_link_rows_all.append(project_link_row)
    header_dff_rows_all.append(dff_row)
    header_sales_credit_rows_all.extend(header_sales_credit_rows)
    line_sales_credit_rows_all.extend(line_sales_credit_rows)

# Build DataFrames
contract_df = pd.DataFrame(contract_rows)
contract_party_df = pd.DataFrame(contract_party_rows_all)
party_contact_df = pd.DataFrame(party_contact_rows_all)
billplan_df = pd.DataFrame(billplan_rows_all)
contract_line_df = pd.DataFrame(contract_line_rows_all)
project_link_df = pd.DataFrame(project_link_rows_all)
header_dff_df = pd.DataFrame(header_dff_rows_all)
header_sales_credit_df = pd.DataFrame(header_sales_credit_rows_all)
line_sales_credit_df = pd.DataFrame(line_sales_credit_rows_all)

# Write to Excel
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    contract_df.to_excel(writer, sheet_name="Contract", index=False)
    contract_party_df.to_excel(writer, sheet_name="ContractParty", index=False)
    party_contact_df.to_excel(writer, sheet_name="ContractPartyContact", index=False)
    billplan_df.to_excel(writer, sheet_name="ProjectBillPlan", index=False)
    contract_line_df.to_excel(writer, sheet_name="ContractLine", index=False)
    project_link_df.to_excel(writer, sheet_name="ContractProject", index=False)
    header_dff_df.to_excel(writer, sheet_name="ContractHeaderDFF", index=False)
    header_sales_credit_df.to_excel(writer, sheet_name="ContractHeaderSalesCredit", index=False)
    line_sales_credit_df.to_excel(writer, sheet_name="ContractLineSalesCredit", index=False)

print("✅ Contract creation completed.")
print(f"   Contracts: {len(contract_df)}")
print(f"   ContractParty rows: {len(contract_party_df)}")
print(f"   ContractPartyContact rows: {len(party_contact_df)}")
print(f"   ProjectBillPlan rows: {len(billplan_df)}")
print(f"   ContractLine rows: {len(contract_line_df)}")
print(f"   ContractProject rows: {len(project_link_df)}")
print(f"   ContractHeaderDFF rows: {len(header_dff_df)}")
print(f"   ContractHeaderSalesCredit rows: {len(header_sales_credit_df)}")
print(f"   ContractLineSalesCredit rows: {len(line_sales_credit_df)}")