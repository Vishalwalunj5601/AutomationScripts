"""
===============================================================================
Airetech PPM Conversion — Contract Import FBDI Generation Script
===============================================================================
Target FBDI Template: ContractMapping.xlsx
Target Output Directory:
  C:\\Users\\vishal walunj\\Vishal\\1D_AIR_CONTROL_CONCEPTS\\AiretechProjects\\DEV1\\DataFiles\\Contracts

Outputs Generated:
  1. Consolidated Data File: ContractImport_Populated.xlsx (9 Sheets)
  2. Individual Excel Files per Sheet (9 Separate .xlsx files):
     - ContractHeader.xlsx
     - ContractParty.xlsx
     - ContractPartyContacts.xlsx
     - ProjectBillPlan.xlsx
     - ContractLine.xlsx
     - ContractProject.xlsx
     - ContractHeaderDFF.xlsx
     - ContractHeaderSalesCredit.xlsx
     - ContractLineSalesCredit.xlsx
  3. Mapping Workbook: Contract_Mapping.xlsx (9 Sheets with Row 3 Mapping Notes)

Key Rules:
  - Dates: Formatted strictly as MM/DD/YYYY.
  - End Date = Start Date + 10 years (MM/DD/YYYY).
  - Standalone execution without requiring external mapping/scripts folders.
===============================================================================
"""

import os
import re
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =============================================================================
# 1. FILE PATH CONFIGURATION (Standalone & Flexible Path Resolution)
# =============================================================================
CANDIDATE_SOURCES = [
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\FINAL LOAD\SOURCE\TAB Pull from Jonas 9-18.xlsx",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "TAB Pull from Jonas 9-18.xlsx"),
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\DEV1\SourceFile\Airetech TAB jobs as of 9-11-26.xlsx",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "Airetech TAB jobs as of 9-11-26.xlsx"),
]
SOURCE_FILE = next((c for c in CANDIDATE_SOURCES if os.path.exists(c)), CANDIDATE_SOURCES[0])

CANDIDATE_CUST_FILES = [
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\DEV1\Fusion\AiretechCustomer - DEV1.xlsx",
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\AiretechCustomer.xlsx",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "AiretechCustomer - DEV1.xlsx"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "AiretechCustomer.xlsx"),
]
CUST_REF_FILE = next((c for c in CANDIDATE_CUST_FILES if os.path.exists(c)), CANDIDATE_CUST_FILES[0])

# Dedicated Target Output Directories for Contracts (Separate EXCEL and CSV subfolders)
OUTPUT_DATA_DIR = r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\FINAL LOAD\DATAFILE\Contracts"
OUTPUT_CONTRACTS_DIR = r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\AiretechProjects\FINAL LOAD\CONTRACTS"

OUTPUT_DATA_EXCEL_DIR = os.path.join(OUTPUT_DATA_DIR, "EXCEL")
OUTPUT_DATA_CSV_DIR = os.path.join(OUTPUT_DATA_DIR, "CSV")
OUTPUT_CONTRACTS_EXCEL_DIR = os.path.join(OUTPUT_CONTRACTS_DIR, "EXCEL")
OUTPUT_CONTRACTS_CSV_DIR = os.path.join(OUTPUT_CONTRACTS_DIR, "CSV")

for d in [OUTPUT_DATA_DIR, OUTPUT_CONTRACTS_DIR, OUTPUT_DATA_EXCEL_DIR, OUTPUT_DATA_CSV_DIR, OUTPUT_CONTRACTS_EXCEL_DIR, OUTPUT_CONTRACTS_CSV_DIR]:
    os.makedirs(d, exist_ok=True)

# Output Files
OUT_POPULATED_FILE = os.path.join(OUTPUT_DATA_EXCEL_DIR, "ContractImport_Populated.xlsx")
OUT_MAPPING_FILE = os.path.join(OUTPUT_DATA_EXCEL_DIR, "Contract_Mapping.xlsx")


# =============================================================================
# 2. HELPER FUNCTIONS
# =============================================================================
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


def parse_date_mmddyyyy(val):
    """Parses various date string formats and returns MM/DD/YYYY."""
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if not val_str:
        return ""
    for fmt in ('%b %d,%Y', '%b %d, %Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%m/%d/%Y', '%Y/%m/%d'):
        try:
            dt = datetime.datetime.strptime(val_str, fmt)
            return dt.strftime('%m/%d/%Y')
        except ValueError:
            pass
    try:
        dt = pd.to_datetime(val_str)
        return dt.strftime('%m/%d/%Y')
    except Exception:
        return str(val_str)


def add_10_years_mmddyyyy(s_date_str):
    """Calculates end date as start date + 10 years in MM/DD/YYYY format."""
    if not s_date_str:
        return ""
    try:
        dt = datetime.datetime.strptime(str(s_date_str).strip(), '%m/%d/%Y')
        try:
            dt_plus_10 = dt.replace(year=dt.year + 10)
        except ValueError:
            # Leap year handling (e.g. Feb 29 -> Feb 28)
            dt_plus_10 = dt.replace(year=dt.year + 10, day=28)
        return dt_plus_10.strftime('%m/%d/%Y')
    except Exception:
        try:
            dt = pd.to_datetime(s_date_str)
            dt_plus_10 = dt.replace(year=dt.year + 10)
            return dt_plus_10.strftime('%m/%d/%Y')
        except Exception:
            return str(s_date_str)


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


def get_pm_info(off):
    """Derives Project Manager name from Office abbreviation."""
    off_str = str(off).strip().upper()
    if 'TUL' in off_str:
        return 'Alex Borio'
    elif 'SPG' in off_str:
        return 'Darren Holliman'
    elif 'LR' in off_str:
        return 'Jason Draper'
    return 'Darren Holliman'


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


def load_customer_reference_map():
    """Loads Customer Master Reference map from CUST_REF_FILE."""
    cust_lookup = {}
    if os.path.exists(CUST_REF_FILE):
        print(f"  [INFO] Loading Customer Master Reference from: {CUST_REF_FILE}")
        try:
            df_c = pd.read_excel(CUST_REF_FILE, sheet_name="DFMiner Export")
        except Exception:
            df_c = pd.read_excel(CUST_REF_FILE)

        df_c.columns = [c.strip() for c in df_c.columns]

        party_id_col = 'PARTY_ID' if 'PARTY_ID' in df_c.columns else ('Party ID' if 'Party ID' in df_c.columns else 'PartyId')
        party_col = 'PARTY_NAME' if 'PARTY_NAME' in df_c.columns else 'Party Name'
        party_num_col = 'PARTY_NUMBER' if 'PARTY_NUMBER' in df_c.columns else ('Party Number' if 'Party Number' in df_c.columns else 'PartyNumber')
        site_use_col = 'SITE_USE_CODE' if 'SITE_USE_CODE' in df_c.columns else 'Site Use Code'
        acct_col = 'ACCOUNT_NUMBER' if 'ACCOUNT_NUMBER' in df_c.columns else 'Account Number'
        loc_col = 'LOCATION' if 'LOCATION' in df_c.columns else 'Location'
        addr_col = 'ADDRESS1' if 'ADDRESS1' in df_c.columns else 'Address1'

        df_c["PARTY_NAME_norm"] = df_c[party_col].astype(str).str.strip().str.upper()

        for party_norm, group in df_c.groupby("PARTY_NAME_norm"):
            bill_rows = group[group[site_use_col].astype(str).str.upper().str.strip() == "BILL_TO"]
            bill_row = bill_rows.head(1)

            # Match SHIP_TO under the SAME customer account number as BILL_TO whenever available
            if not bill_row.empty and acct_col in bill_row.iloc[0]:
                bill_acct_val = str(bill_row.iloc[0][acct_col]).strip()
                same_acct_ship = group[(group[site_use_col].astype(str).str.upper().str.strip() == "SHIP_TO") &
                                       (group[acct_col].astype(str).str.strip() == bill_acct_val)]
                if not same_acct_ship.empty:
                    ship_row = same_acct_ship.head(1)
                else:
                    ship_row = group[group[site_use_col].astype(str).str.upper().str.strip() == "SHIP_TO"].head(1)
            else:
                ship_row = group[group[site_use_col].astype(str).str.upper().str.strip() == "SHIP_TO"].head(1)

            def r_dict(rdf):
                if rdf.empty:
                    return None
                r = rdf.iloc[0]
                pnum = ""
                if party_num_col in r and pd.notna(r[party_num_col]):
                    val = r[party_num_col]
                    if isinstance(val, (int, float)):
                        pnum = str(int(val))
                    else:
                        pnum = str(val).strip()
                        if pnum.endswith(".0"):
                            pnum = pnum[:-2]
                pid = ""
                if party_id_col in r and pd.notna(r[party_id_col]):
                    val = r[party_id_col]
                    if isinstance(val, (int, float)):
                        pid = str(int(val))
                    else:
                        pid = str(val).strip()
                        if pid.endswith(".0"):
                            pid = pid[:-2]
                ca_id = ""
                if "CUST_ACCOUNT_ID" in r and pd.notna(r["CUST_ACCOUNT_ID"]):
                    val = r["CUST_ACCOUNT_ID"]
                    if isinstance(val, (int, float)):
                        ca_id = str(int(val))
                    else:
                        ca_id = str(val).strip()
                        if ca_id.endswith(".0"):
                            ca_id = ca_id[:-2]
                su_id = ""
                if "SITE_USE_ID" in r and pd.notna(r["SITE_USE_ID"]):
                    val = r["SITE_USE_ID"]
                    if isinstance(val, (int, float)):
                        su_id = str(int(val))
                    else:
                        su_id = str(val).strip()
                        if su_id.endswith(".0"):
                            su_id = su_id[:-2]
                return {
                    "party_id": pid,
                    "party_name": "" if pd.isna(r[party_col]) else str(r[party_col]).strip(),
                    "party_number": pnum,
                    "account_number": "" if acct_col not in r or pd.isna(r[acct_col]) else str(r[acct_col]).strip(),
                    "cust_account_id": ca_id,
                    "site_use_id": su_id,
                    "location": "" if loc_col not in r or pd.isna(r[loc_col]) else str(r[loc_col]).strip(),
                    "address1": "" if addr_col not in r or pd.isna(r[addr_col]) else str(r[addr_col]).strip(),
                }
            cust_lookup[party_norm] = {"bill_to": r_dict(bill_row), "ship_to": r_dict(ship_row)}

        print(f"  [OK] Loaded Customer Reference Map with {len(cust_lookup)} unique customer accounts.")
    else:
        print(f"  [WARN] Customer reference file not found at {CUST_REF_FILE}. Using fallback site keys.")

    return cust_lookup


CUSTOMER_ALIASES = {
    'FAYETTEVILLE MECH CONT': 'FAYETTEVILLE MECHANICAL CONTRACTOR',
    'ENFRA MCC, LLC': 'ENFRA MCC LLC',
    'CHEROKEE NATION ENT DIRECT BILL': 'CHEROKEE NATION ENT   DIRECT BILL',
    'J.E.N ENTERPRISES LLC DBA TRIAD SER': 'JEN ENTERPRISES LLC',
}


def clean_alphanumeric(s):
    """Strips all non-alphanumeric characters for clean name matching."""
    return re.sub(r'[^A-Z0-9]', '', str(s).upper())


def get_customer_sites(cust_lookup, j_num, src_cust):
    """Performs accurate hierarchical Customer -> SITE_USE_CODE (BILL_TO / SHIP_TO) lookup."""
    src_norm = str(src_cust).strip().upper()
    matched_entry = None

    # Step 1: Direct Exact Match
    if src_norm in cust_lookup:
        matched_entry = cust_lookup[src_norm]
    # Step 2: Known Alias Match
    elif src_norm in CUSTOMER_ALIASES and CUSTOMER_ALIASES[src_norm] in cust_lookup:
        matched_entry = cust_lookup[CUSTOMER_ALIASES[src_norm]]
    else:
        # Step 3: Normalized Alphanumeric Match (handles punctuation/spacing differences like LLC, commas, etc.)
        src_clean = clean_alphanumeric(src_norm)
        for p_norm, data in cust_lookup.items():
            if clean_alphanumeric(p_norm) == src_clean:
                matched_entry = data
                break

        # Step 4: Substring / Containment Match (longest matching name first, minimum 8 characters to prevent false positives)
        if not matched_entry:
            candidates = []
            for p_norm, data in cust_lookup.items():
                if len(src_norm) >= 8 and (src_norm in p_norm or p_norm in src_norm):
                    candidates.append((len(p_norm), data))
            if candidates:
                candidates.sort(key=lambda x: x[0], reverse=True)
                matched_entry = candidates[0][1]

    b_info = (matched_entry.get("bill_to") or {}) if matched_entry else {}
    s_info = (matched_entry.get("ship_to") or {}) if matched_entry else {}

    # If ship_to is missing in master but bill_to exists, use bill_to site for shipping
    if not s_info and b_info:
        s_info = b_info
    # Conversely, if bill_to is missing in master but ship_to exists, use ship_to
    if not b_info and s_info:
        b_info = s_info

    bill_party = b_info.get("party_name") or src_cust
    bill_party_num = b_info.get("party_number") or ""
    bill_party_id = b_info.get("party_id") or ""
    bill_cust_acct_id = b_info.get("cust_account_id") or ""
    bill_site_use_id = b_info.get("site_use_id") or ""
    bill_acct = b_info.get("account_number") or f"ACCT_{j_num}"
    bill_loc = b_info.get("location") or f"LOC_BILL_{j_num}"
    bill_addr = b_info.get("address1") or f"Site Address {j_num}"

    ship_party = s_info.get("party_name") or src_cust
    ship_party_num = s_info.get("party_number") or ""
    ship_party_id = s_info.get("party_id") or ""
    ship_acct = s_info.get("account_number") or f"ACCT_{j_num}"
    ship_loc = s_info.get("location") or f"LOC_SHIP_{j_num}"
    ship_addr = s_info.get("address1") or f"Site Address {j_num}"

    return bill_party, bill_acct, bill_loc, bill_addr, ship_party, ship_acct, ship_loc, ship_addr, bill_party_num, bill_party_id, bill_cust_acct_id, bill_site_use_id


# =============================================================================
# 3. MAIN CONVERSION PROCESS
# =============================================================================
def main():
    print("=" * 80)
    print("  Airetech PPM Conversion - Contract Import FBDI Script Starting")
    print("=" * 80)

    # Ensure Target Contracts Output Directory exists
    os.makedirs(OUTPUT_DATA_DIR, exist_ok=True)

    if not os.path.exists(SOURCE_FILE):
        raise FileNotFoundError(f"Source data file not found at: {SOURCE_FILE}")

    # Read Source File and filter out repeated headers or blank rows
    df_src = pd.read_excel(SOURCE_FILE)
    df_src = df_src[df_src['Job No.'].notna() & (df_src['Job No.'].astype(str).str.strip() != 'Job No.')].reset_index(drop=True)
    print(f"  [OK] Loaded source dataset: {SOURCE_FILE} ({len(df_src)} records)")

    # Load Customer Reference Map
    cust_lookup = load_customer_reference_map()

    # Data lists for all 9 Sheets
    contract_rows = []
    contract_party_rows = []
    party_contact_rows = []
    billplan_rows = []
    contract_line_rows = []
    project_link_rows = []
    header_dff_rows = []
    header_sales_credit_rows = []
    line_sales_credit_rows = []

    print("\n  [1/3] Building 9 Data Sheets with MM/DD/YYYY Dates (End Date = Start Date + 10 Years)...")
    for i in range(len(df_src)):
        row = df_src.iloc[i]
        j_num = clean_job_no(row['Job No.'])

        # User Keys
        puid = f"PUID_{j_num}"
        s_puid = f"S_PUID_{j_num}"
        c_puid = f"CPUID_{j_num}"
        line_puid = f"LINE_PUID_{j_num}"
        sbp_key = f"SBP_{j_num}"
        kbp_key = f"KBP_{j_num}"
        sbp01_key = f"SBP01_{j_num}"
        kbp01_key = f"KBP01_{j_num}"

        j_name = str(row['Job Name']).strip() if ('Job Name' in row and pd.notna(row['Job Name'])) else f"Job_{j_num}"
        s_date = parse_date_mmddyyyy(row['Start Date']) if 'Start Date' in row else ""
        e_date = add_10_years_mmddyyyy(s_date)

        src_cust = str(row['Customer Name']).strip() if ('Customer Name' in row and pd.notna(row['Customer Name'])) else j_name

        # Project Manager derivation
        if 'Project Manager' in row and pd.notna(row['Project Manager']) and str(row['Project Manager']).strip():
            pm_name = str(row['Project Manager']).strip()
        else:
            pm_name = get_pm_info(row.get('Office (Proj Organization)', ''))

        # Salesperson: standardized persona for Airetech conversion contracts
        sales = "Air Balance"

        # Contract Amount derivation
        if 'Contract Amount' in row and pd.notna(row['Contract Amount']):
            contract_amt = float(row['Contract Amount'])
        elif 'Curr Est Rev' in row and pd.notna(row['Curr Est Rev']):
            contract_amt = float(row['Curr Est Rev'])
        else:
            contract_amt = 0.0

        org = get_org(row.get('Office (Proj Organization)', ''))

        # ATTRIBUTE2 (Paid Invoice Amount = Rev Billed - A/R Owing)
        rev_billed = float(row['Rev Billed']) if ('Rev Billed' in row and pd.notna(row['Rev Billed'])) else 0.0
        ar_owing = float(row['A/R Owing']) if ('A/R Owing' in row and pd.notna(row['A/R Owing'])) else 0.0
        paid_inv_amt = rev_billed - ar_owing

        aia_raw = str(row['AIA']).strip().upper() if ('AIA' in row and pd.notna(row['AIA'])) else 'N'
        inv_type = "AIA-PA-Invoice" if aia_raw == 'Y' else "PA Invoice"
        inv_status = "D" if inv_type == "AIA-PA-Invoice" else "R"

        # Customer Site Lookup
        bill_party, bill_acct, bill_loc, bill_addr, ship_party, ship_acct, ship_loc, ship_addr, bill_party_num, bill_party_id, bill_cust_acct_id, bill_site_use_id = get_customer_sites(cust_lookup, j_num, src_cust)

        # 1. ContractHeader
        contract_rows.append({
            "ContractTypeName": "Airetech_Customer Contract for Project",
            "StsCode": "DRAFT",
            "CurrencyCode": "USD",
            "StartDate": s_date,
            "EndDate": e_date,
            "LegalEntityName": "Airetech Corporation",
            "ContractPuid": puid,
            "OrgName": "Airetech",
            "InvConvRateType": "Corporate",
            "InvTrxTypeName": inv_type,
            "BillToAccount": bill_party,
            "BillToAccountNumber": bill_acct,
            "GeneratedInvoiceStatus": inv_status,
            "ContractNumber": j_num,
            "Cognomen": j_name,
            "PartyName": bill_party,
            "BillToSiteUseLocation": bill_loc,
            "BillToAddress": bill_addr,
            "ShipToAccount": ship_party,
            "ShipToAccountNumber": ship_acct,
            "ShipToSiteUseLocation": ship_loc,
            "ShipToAddress": ship_addr,
            "OwningOrgName": org,
            "AuthoringPartyMeaning": "Internal",
            "TaxExemptionControlMeaning": "Standard",
            "LedgerCurrency": "USD",
            "AuthoringPartyCode": "INTERNAL",
            "AccessLevel": "CREATE",
            "ContributionPercent": 100,
            "Description": None,
            "BillingSalesrepName": sales,
            "CustomerPONumber": None
        })

        # 2. ContractParty
        supp_party_num = ""
        if "AIRETECH CORPORATION" in cust_lookup:
            b_supp = cust_lookup["AIRETECH CORPORATION"].get("bill_to") or {}
            supp_party_num = str(b_supp.get("party_number") or "").strip()

        contract_party_rows.extend([
            {
                "ContractPuid": puid,
                "PartyRolePuid": c_puid,
                "Role": "All eligible customers",
                "PartyName": "" if bill_party_id else bill_party,
                "PartyNumber": str(bill_party_num).strip() if bill_party_num else "",
                "PartyId": str(bill_party_id).strip() if bill_party_id else ""
            },
            {
                "ContractPuid": puid,
                "PartyRolePuid": s_puid,
                "Role": "Supplier",
                "PartyName": "Airetech",
                "PartyNumber": supp_party_num,
                "PartyId": ""
            }
        ])

        # 3. ContractPartyContacts
        party_contact_rows.append({
            "PartyRolePuid": s_puid,
            "ContactPuid": puid,
            "Role": "Contract administrator",
            "PartyContactName": pm_name,
            "OwnerYn": "Y",
            "StartDate": s_date,
            "AccessLevelName": "Full"
        })

        # 4. ProjectBillPlan
        billplan_rows.extend([
            {
                "ContractPuid": puid,
                "BillingCycleName": "Immediate",
                "LaborInvoiceFormatName": "Labor",
                "NlInvoiceFormatName": "Nonlabor",
                "EventsInvoiceFormatName": "Event",
                "ExtSource": sbp_key,
                "ExtKey": kbp_key,
                "BillPlanLang": "US",
                "SourceLangFlag": "Y",
                "BillMethodName": "Amount Based Invoice",
                "BillMethodFlag": "I",
                "BillPlanName": "Amount Based Invoice",
                "BillSetNum": 1,
                "BillToCustAcctNumber": bill_acct,
                "BillToSiteUseLocation": bill_loc,
                "BillToCustAcctId": bill_cust_acct_id,
                "BillToSiteUseId": bill_site_use_id,
                "InvCurrOptMeaning": "Contract",
                "PaymentTermName": "IMMEDIATE"
            },
            {
                "ContractPuid": puid,
                "BillingCycleName": "",
                "LaborInvoiceFormatName": "",
                "NlInvoiceFormatName": "",
                "EventsInvoiceFormatName": "",
                "ExtSource": sbp01_key,
                "ExtKey": kbp01_key,
                "BillPlanLang": "US",
                "SourceLangFlag": "Y",
                "BillMethodName": "Amount Based Revenue",
                "BillMethodFlag": "R",
                "BillPlanName": "Amount Based Revenue",
                "BillSetNum": "",
                "BillToCustAcctNumber": "",
                "BillToSiteUseLocation": "",
                "BillToCustAcctId": "",
                "BillToSiteUseId": "",
                "InvCurrOptMeaning": "",
                "PaymentTermName": ""
            }
        ])

        # 5. ContractLine
        contract_line_rows.append({
            "LineNumber": "Conversion Contract Line",
            "ContractPuid": puid,
            "LinePuid": line_puid,
            "StsCode": "DRAFT",
            "LineTypeName": "Free-form, project",
            "StartDate": s_date,
            "EndDate": e_date,
            "ItemName": "Contract",
            "ShipToAccount": ship_party,
            "ShipToAccountNumber": ship_acct,
            "ShipToSite": ship_loc,
            "BillPlan": "Bill Plan",
            "RevenuePlan": "Revenue Plan",
            "BillPlanExternalSource": sbp_key,
            "BillPlanExternalKey": kbp_key,
            "RevenuePlanExternalSource": sbp01_key,
            "RevenuePlanExternalKey": kbp01_key,
            "LineAmount": contract_amt
        })

        # 6. ContractProject
        project_link_rows.append({
            "LinePuid": line_puid,
            "ExtSource": sbp_key,
            "ExtKey": kbp_key,
            "ProjectNumber": j_num,
            "ActiveFlag": "Y",
            "FundingAmount": contract_amt
        })

        # 7. ContractHeaderDFF
        header_dff_rows.append({
            "ContractPuid": puid,
            "AttributeNumber1": paid_inv_amt
        })

        # 8. ContractHeaderSalesCredit
        header_sales_credit_rows.append({
            "ContractPuid": puid,
            "SalesCreditPuid": f"SC_{j_num}",
            "ExternalSource": sbp_key,
            "ExternalKey": kbp_key,
            "Percent": 100,
            "SalesrepName": sales,
            "SalesCreditType": "Quota Sales Credit",
            "StartDate": s_date
        })

        # 9. ContractLineSalesCredit
        line_sales_credit_rows.append({
            "ContractPuid": puid,
            "SalesCreditPuid": f"SC_LINE_{j_num}",
            "ExternalSource": sbp_key,
            "ExternalKey": kbp_key,
            "Percent": 100,
            "SalesrepName": sales,
            "SalesCreditType": "Non-quota Sales Credit",
            "StartDate": s_date
        })

    data_dict = {
        "ContractHeader": contract_rows,
        "ContractParty": contract_party_rows,
        "ContractPartyContacts": party_contact_rows,
        "ProjectBillPlan": billplan_rows,
        "ContractLine": contract_line_rows,
        "ContractProject": project_link_rows,
        "ContractHeaderDFF": header_dff_rows,
        "ContractHeaderSalesCredit": header_sales_credit_rows,
        "ContractLineSalesCredit": line_sales_credit_rows
    }

    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_hdr = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_reg = Font(name="Segoe UI", size=9.5, color="000000")
    thin_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # -------------------------------------------------------------------------
    # Save Consolidated 9-Sheet Workbook: ContractImport_Populated.xlsx
    # -------------------------------------------------------------------------
    wb_p = openpyxl.Workbook()
    wb_p.remove(wb_p.active)

    for sheetname, row_list in data_dict.items():
        ws = wb_p.create_sheet(title=sheetname)
        ws.views.sheetView[0].showGridLines = True
        if row_list:
            headers = list(row_list[0].keys())
            ws.row_dimensions[1].height = 24
            for c_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=1, column=c_idx, value=h)
                cell.font = font_hdr
                cell.fill = fill_hdr
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            for r_idx, r_data in enumerate(row_list, start=2):
                ws.row_dimensions[r_idx].height = 20
                for c_idx, val in enumerate(r_data.values(), start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = font_reg
                    cell.border = thin_border
                    if isinstance(val, (int, float)) and not any(k in headers[c_idx-1] for k in ["Percent", "BillSetNum", "PartyNumber", "LineNumber"]):
                        cell.number_format = "$#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

    print("\n  Saving Consolidated Populated Workbook...")
    safe_save(wb_p, OUT_POPULATED_FILE)
    safe_save(wb_p, os.path.join(OUTPUT_CONTRACTS_EXCEL_DIR, "ContractImport_Populated.xlsx"))

    # -------------------------------------------------------------------------
    # Save Each Sheet as a Separate Excel File in OUTPUT_DATA_EXCEL_DIR & OUTPUT_CONTRACTS_EXCEL_DIR
    # -------------------------------------------------------------------------
    print("\n  [2/3] Saving Each Sheet as Separate Files in EXCEL and CSV subfolders...")
    for sheetname, row_list in data_dict.items():
        wb_single = openpyxl.Workbook()
        ws_single = wb_single.active
        ws_single.title = sheetname
        ws_single.views.sheetView[0].showGridLines = True
        if row_list:
            headers = list(row_list[0].keys())
            ws_single.row_dimensions[1].height = 24
            for c_idx, h in enumerate(headers, start=1):
                cell = ws_single.cell(row=1, column=c_idx, value=h)
                cell.font = font_hdr
                cell.fill = fill_hdr
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            for r_idx, r_data in enumerate(row_list, start=2):
                ws_single.row_dimensions[r_idx].height = 20
                for c_idx, val in enumerate(r_data.values(), start=1):
                    cell = ws_single.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = font_reg
                    cell.border = thin_border
                    if isinstance(val, (int, float)) and not any(k in headers[c_idx-1] for k in ["Percent", "BillSetNum", "PartyNumber", "LineNumber"]):
                        cell.number_format = "$#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

        # Save single .xlsx files strictly to EXCEL folders
        safe_save(wb_single, os.path.join(OUTPUT_DATA_EXCEL_DIR, f"{sheetname}.xlsx"))
        safe_save(wb_single, os.path.join(OUTPUT_CONTRACTS_EXCEL_DIR, f"{sheetname}.xlsx"))

        # Save single .csv files strictly to CSV folders
        df_sheet = pd.DataFrame(row_list)
        df_sheet.to_csv(os.path.join(OUTPUT_DATA_CSV_DIR, f"{sheetname}.csv"), index=False, encoding='utf-8')
        df_sheet.to_csv(os.path.join(OUTPUT_CONTRACTS_CSV_DIR, f"{sheetname}.csv"), index=False, encoding='utf-8')

    # Create CONTRACTS.zip containing CSV files strictly in CSV folders
    import zipfile
    for csv_dir in [OUTPUT_DATA_CSV_DIR, OUTPUT_CONTRACTS_CSV_DIR]:
        zip_path = os.path.join(csv_dir, "CONTRACTS.zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for sheetname in data_dict.keys():
                csv_file = os.path.join(csv_dir, f"{sheetname}.csv")
                if os.path.exists(csv_file):
                    zf.write(csv_file, arcname=f"{sheetname}.csv")
        print(f"  [OK] Successfully created zip package: {zip_path}")

    # -------------------------------------------------------------------------
    # Save 9-Sheet Mapping Workbook with Row 3 Notes in Contracts directory
    # -------------------------------------------------------------------------
    print("\n  [3/3] Building 9 Mapping Sheets for Contract_Mapping.xlsx...")
    wb_m = openpyxl.Workbook()
    wb_m.remove(wb_m.active)

    sheet_mappings = {
        "ContractHeader": (
            ["ContractTypeName", "StsCode", "CurrencyCode", "StartDate", "EndDate", "LegalEntityName", "ContractPuid", "OrgName", "InvConvRateType", "InvTrxTypeName", "BillToAccount", "BillToAccountNumber", "GeneratedInvoiceStatus", "ContractNumber", "Cognomen", "PartyName", "BillToSiteUseLocation", "BillToAddress", "ShipToAccount", "ShipToAccountNumber", "ShipToSiteUseLocation", "ShipToAddress", "OwningOrgName", "AuthoringPartyMeaning", "TaxExemptionControlMeaning", "LedgerCurrency", "AuthoringPartyCode", "AccessLevel", "ContributionPercent", "Description", "BillingSalesrepName", "CustomerPONumber"],
            {
                1: "Default: Airetech_Customer Contract for Project",
                2: "Default: DRAFT",
                3: "Default: USD",
                4: "Source: Start Date (Formatted MM/DD/YYYY)",
                5: "End Date = Start Date + 10 Years (Formatted MM/DD/YYYY)",
                6: "Default Legal Entity: Airetech Corporation",
                7: "USER KEY = PUID_{PROJECT_NUMBER}",
                8: "Default BU: Airetech",
                9: "Default: Corporate",
                10: "Derived from AIA: N -> PA Invoice, Y -> AIA-PA-Invoice",
                11: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx PARTY_NAME (BILL_TO)",
                12: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx ACCOUNT_NUMBER (BILL_TO)",
                13: "Derived: PA Invoice -> R, AIA-PA-Invoice -> D",
                14: "Source: Job No. (Numeric clean)",
                15: "Source: Job Name",
                16: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx PARTY_NAME (Bill-to Customer)",
                17: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx LOCATION (BILL_TO)",
                18: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx ADDRESS1 (BILL_TO)",
                19: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx PARTY_NAME (SHIP_TO)",
                20: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx ACCOUNT_NUMBER (SHIP_TO)",
                21: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx LOCATION (SHIP_TO)",
                22: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx ADDRESS1 (SHIP_TO)",
                23: "Derived from Office: TUL -> Tulsa-Airetech-Services, SPG -> Springdale-Airetech-Services, LR -> Little Rock-Airetech-Services",
                24: "Default = 'Internal'",
                25: "Default = 'Standard'",
                26: "Default = 'USD'",
                27: "Default = 'INTERNAL'",
                28: "Default = 'CREATE'",
                29: "Default = '100'",
                30: "Retention % (Not Provided in source)",
                31: "Source: Sales Person / 1ST SALESM (BillingSalesrepName)",
                32: "Customer PO Number (CustomerPONumber)"
            }
        ),
        "ContractParty": (
            ["ContractPuid", "PartyRolePuid", "Role", "PartyName", "PartyNumber"],
            {
                1: "USER KEY = PUID_{PROJECT_NUMBER}",
                2: "USER KEY = CPUID_{PROJECT_NUMBER} (Customer) / S_PUID_{PROJECT_NUMBER} (Supplier)",
                3: "Default: 'All eligible customers' (Row 1) / 'Supplier' (Row 2)",
                4: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx PARTY_NAME / 'Airetech Corporation'",
                5: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx PARTY_NUMBER"
            }
        ),
        "ContractPartyContacts": (
            ["PartyRolePuid", "ContactPuid", "Role", "PartyContactName", "OwnerYn", "StartDate", "AccessLevelName"],
            {
                1: "USER KEY = S_PUID_{PROJECT_NUMBER}",
                2: "USER KEY = PUID_{PROJECT_NUMBER}",
                3: "Default = 'Contract administrator'",
                4: "Source: Project Manager / Derived from Office PM",
                5: "Default = 'Y'",
                6: "Source: Start Date (Formatted MM/DD/YYYY)",
                7: "Default = 'Full'"
            }
        ),
        "ProjectBillPlan": (
            ["ContractPuid", "BillingCycleName", "LaborInvoiceFormatName", "NlInvoiceFormatName", "EventsInvoiceFormatName", "ExtSource", "ExtKey", "BillPlanLang", "SourceLangFlag", "BillMethodName", "BillMethodFlag", "BillPlanName", "BillSetNum", "BillToCustAcctNumber", "BillToSiteUseLocation", "InvCurrOptMeaning", "PaymentTermName"],
            {
                1: "USER KEY = PUID_{PROJECT_NUMBER}",
                2: "Default: Immediate (Row 1) / Blank (Row 2)",
                3: "Default: Labor",
                4: "Default: Nonlabor",
                5: "Default: Event",
                6: "USER KEY = SBP_{PROJECT_NUMBER} (Row 1) / SBP01_{PROJECT_NUMBER} (Row 2)",
                7: "USER KEY = KBP_{PROJECT_NUMBER} (Row 1) / KBP01_{PROJECT_NUMBER} (Row 2)",
                8: "Default: US",
                9: "Default: Y",
                10: "Default: Amount Based Invoice (Row 1) / Amount Based Revenue (Row 2)",
                11: "BillMethodFlag: I (Row 1: Amount Based Invoice) / R (Row 2: Amount Based Revenue)",
                12: "Default: Amount Based Invoice (Row 1) / Amount Based Revenue (Row 2)",
                13: "Default: 1",
                14: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx ACCOUNT_NUMBER (BILL_TO)",
                15: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx LOCATION (BILL_TO)",
                16: "Default: Contract",
                17: "Default: IMMEDIATE"
            }
        ),
        "ContractLine": (
            ["LineNumber", "ContractPuid", "LinePuid", "StsCode", "LineTypeName", "StartDate", "EndDate", "ItemName", "ShipToAccount", "ShipToAccountNumber", "ShipToSite", "BillPlan", "RevenuePlan", "BillPlanExternalSource", "BillPlanExternalKey", "RevenuePlanExternalSource", "RevenuePlanExternalKey", "LineAmount"],
            {
                1: "Default: Conversion Contract Line",
                2: "USER KEY = PUID_{PROJECT_NUMBER}",
                3: "USER KEY = LINE_PUID_{PROJECT_NUMBER}",
                4: "Default: DRAFT",
                5: "Default: Free-form, project",
                6: "Source: Start Date (Formatted MM/DD/YYYY)",
                7: "End Date = Start Date + 10 Years (Formatted MM/DD/YYYY)",
                8: "Default: Contract",
                9: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx PARTY_NAME (SHIP_TO)",
                10: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx ACCOUNT_NUMBER (SHIP_TO)",
                11: "Hierarchical Match: Customer Name -> AiretechCustomer.xlsx LOCATION (SHIP_TO)",
                12: "Default: Bill Plan",
                13: "Default: Revenue Plan",
                14: "USER KEY = SBP_{PROJECT_NUMBER}",
                15: "USER KEY = KBP_{PROJECT_NUMBER}",
                16: "USER KEY = SBP01_{PROJECT_NUMBER}",
                17: "USER KEY = KBP01_{PROJECT_NUMBER}",
                18: "Source: Contract Amount / Curr Est Rev"
            }
        ),
        "ContractProject": (
            ["LinePuid", "ExtSource", "ExtKey", "ProjectNumber", "ActiveFlag", "FundingAmount"],
            {
                1: "USER KEY = LINE_PUID_{PROJECT_NUMBER}",
                2: "USER KEY = SBP_{PROJECT_NUMBER}",
                3: "USER KEY = KBP_{PROJECT_NUMBER}",
                4: "Source: Job No. (Numeric clean)",
                5: "Default: Y",
                6: "Source: Contract Amount / Curr Est Rev"
            }
        ),
        "ContractHeaderDFF": (
            ["ContractPuid", "AttributeNumber1"],
            {
                1: "USER KEY = PUID_{PROJECT_NUMBER}",
                2: "Paid Invoice Amount = Rev Billed - A/R Owing"
            }
        ),
        "ContractHeaderSalesCredit": (
            ["ContractPuid", "SalesCreditPuid", "ExternalSource", "ExternalKey", "Percent", "SalesrepName", "SalesCreditType", "StartDate"],
            {
                1: "USER KEY = PUID_{PROJECT_NUMBER}",
                2: "USER KEY = SC_{PROJECT_NUMBER}",
                3: "USER KEY = SBP_{PROJECT_NUMBER}",
                4: "USER KEY = KBP_{PROJECT_NUMBER}",
                5: "Default: 100 (% Allocation)",
                6: "Source: Sales Person / 1ST SALESM (BillingSalesrepName)",
                7: "Default: Quota Sales Credit",
                8: "Source: Start Date (Formatted MM/DD/YYYY)"
            }
        ),
        "ContractLineSalesCredit": (
            ["ContractPuid", "SalesCreditPuid", "ExternalSource", "ExternalKey", "Percent", "SalesrepName", "SalesCreditType", "StartDate"],
            {
                1: "USER KEY = PUID_{PROJECT_NUMBER}",
                2: "USER KEY = SC_LINE_{PROJECT_NUMBER}",
                3: "USER KEY = SBP_{PROJECT_NUMBER}",
                4: "USER KEY = KBP_{PROJECT_NUMBER}",
                5: "Default: 100 (% Allocation)",
                6: "Source: Sales Person / 1ST SALESM (BillingSalesrepName)",
                7: "Default: Non-quota Sales Credit",
                8: "Source: Start Date (Formatted MM/DD/YYYY)"
            }
        )
    }

    fill_note = PatternFill(start_color="F2F4F8", end_color="F2F4F8", fill_type="solid")
    font_note = Font(name="Segoe UI", size=9, italic=True, color="1F4E78")

    for sheetname, (headers, notes_dict) in sheet_mappings.items():
        ws = wb_m.create_sheet(title=sheetname)
        ws.views.sheetView[0].showGridLines = True

        # Row 1: Column Headers
        ws.row_dimensions[1].height = 24
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c_idx, value=h)
            cell.font = font_hdr
            cell.fill = fill_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        # Row 3: Explicit Mapping Notes
        ws.row_dimensions[3].height = 28
        for c_idx, h in enumerate(headers, start=1):
            note = notes_dict.get(c_idx, "")
            cell = ws.cell(row=3, column=c_idx, value=note)
            cell.font = font_note
            cell.fill = fill_note
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = thin_border

    safe_save(wb_m, OUT_MAPPING_FILE)
    safe_save(wb_m, os.path.join(OUTPUT_CONTRACTS_EXCEL_DIR, "Contract_Mapping.xlsx"))

    print("\n" + "=" * 80)
    print("  [OK] CONTRACT CONVERSION COMPLETED SUCCESSFULLY!")
    print(f"  Target Excel Directory: {OUTPUT_CONTRACTS_EXCEL_DIR}")
    print(f"  Target CSV Directory:   {OUTPUT_CONTRACTS_CSV_DIR}")
    print(f"  Consolidated File: {OUT_POPULATED_FILE}")
    print(f"  Separate Files Generated: 9 individual .xlsx files in EXCEL and 9 individual .csv files in CSV")
    print(f"  Mapping Workbook: {OUT_MAPPING_FILE}")
    print("=" * 80)


if __name__ == "__main__":
    main()