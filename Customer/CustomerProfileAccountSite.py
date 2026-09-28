import os
import re
import html
import time
import warnings
import traceback
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from xml.etree import ElementTree as ET

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =========================================================
# CONFIG
# =========================================================

INPUT_FILE = r"C:\AIR_CONTROL_CONCEPTS\Customer_1C\Airetech\CustomerProfile_AiretechDelta.xlsx"
OUTPUT_FILE = r"C:\AIR_CONTROL_CONCEPTS\Customer_1C\Airetech\CustomerProfile_AiretechReportDelta.xlsx"

# If sheet name is None, first sheet will be used
INPUT_SHEET = None

SERVICE_URL = "https://ibyqjb-dev3.fa.ocs.oraclecloud.com/fscmService/ReceivablesCustomerProfileService"
USERNAME = "Conversion"
PASSWORD = "Welcome@12345"

PROFILE_CLASS_NAME = "DEFAULT"

# If your source "CreditLimit" column contains Y/N instead of numeric amount,
# script will treat it as CreditChecking flag automatically.
DEFAULT_CREDIT_CHECKING = "N"

REQUEST_TIMEOUT = 120
SLEEP_BETWEEN_CALLS = 0.2

# Account-level update before site-level create
DO_ACCOUNT_LEVEL_UPDATE = True

# Site create behavior
# First try create with CustomerAccountId
TRY_SITE_CREATE_WITH_CUST_ACCOUNT_ID = True

# If site create fails with duplicate CustAccountId,
# retry create without CustomerAccountId
RETRY_SITE_CREATE_WITHOUT_CUST_ACCOUNT_ID = True

# If create still fails, try site-level update as fallback
TRY_SITE_UPDATE_ON_CREATE_FAILURE = True


# =========================================================
# HELPERS
# =========================================================

def clean_str(val):
    if pd.isna(val):
        return ""
    return str(val).strip()


def is_blank(val):
    return clean_str(val) == ""


def is_yes_no(val):
    v = clean_str(val).upper()
    return v in {"Y", "N"}


def is_number_like(val):
    s = clean_str(val)
    if not s:
        return False
    s = s.replace(",", "").replace("$", "")
    try:
        float(s)
        return True
    except Exception:
        return False


def xml_escape(val):
    return html.escape(clean_str(val), quote=True)


def infer_credit_values(source_value):
    """
    Source column is named CreditLimit, but in your data it may contain:
    - Y/N  -> treat as CreditChecking
    - numeric -> treat as CreditLimit
    """
    raw = clean_str(source_value)

    if is_yes_no(raw):
        return {
            "credit_checking": raw.upper(),
            "credit_limit": "",
            "credit_currency_code": ""
        }

    if is_number_like(raw):
        return {
            "credit_checking": DEFAULT_CREDIT_CHECKING,
            "credit_limit": raw.replace(",", "").replace("$", ""),
            "credit_currency_code": ""
        }

    return {
        "credit_checking": DEFAULT_CREDIT_CHECKING,
        "credit_limit": "",
        "credit_currency_code": ""
    }


def extract_fault_text(xml_text):
    if not xml_text:
        return "Empty response"

    try:
        root = ET.fromstring(xml_text)
        faultstring = root.find(".//faultstring")
        if faultstring is not None and faultstring.text:
            return faultstring.text.strip()
    except Exception:
        pass

    m = re.search(r"<faultstring.*?>(.*?)</faultstring>", xml_text, flags=re.I | re.S)
    if m:
        return html.unescape(re.sub(r"<.*?>", "", m.group(1))).strip()

    return xml_text[:1000]


def extract_success_value(xml_text, tag_name):
    try:
        root = ET.fromstring(xml_text)
        for elem in root.iter():
            if elem.tag.endswith(tag_name):
                return elem.text.strip() if elem.text else ""
    except Exception:
        pass
    return ""


def post_soap(payload):
    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": ""
    }

    resp = requests.post(
        SERVICE_URL,
        data=payload.encode("utf-8"),
        headers=headers,
        auth=HTTPBasicAuth(USERNAME, PASSWORD),
        timeout=REQUEST_TIMEOUT
    )

    return resp.status_code, resp.text


def build_account_update_payload(account_number, payment_terms, credit_checking, credit_limit=""):
    extra_credit_limit = ""
    if clean_str(credit_limit):
        extra_credit_limit += f"\n            <cus:CreditLimit>{xml_escape(credit_limit)}</cus:CreditLimit>"

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
                  xmlns:typ="http://xmlns.oracle.com/apps/financials/receivables/customers/customerProfileService/types/"
                  xmlns:cus="http://xmlns.oracle.com/apps/financials/receivables/customers/customerProfileService/">
   <soapenv:Header/>
   <soapenv:Body>
      <typ:updateCustomerProfile>
         <typ:customerProfile>
            <cus:AccountNumber>{xml_escape(account_number)}</cus:AccountNumber>
            <cus:PaymentTerms>{xml_escape(payment_terms)}</cus:PaymentTerms>
            <cus:CreditChecking>{xml_escape(credit_checking)}</cus:CreditChecking>{extra_credit_limit}
         </typ:customerProfile>
      </typ:updateCustomerProfile>
   </soapenv:Body>
</soapenv:Envelope>"""


def build_site_create_payload(account_number, site_number, party_id, customer_account_id,
                              payment_terms, credit_checking, credit_limit="",
                              include_customer_account_id=True):
    cust_acc_block = ""
    if include_customer_account_id and clean_str(customer_account_id):
        cust_acc_block = f"\n            <cus:CustomerAccountId>{xml_escape(customer_account_id)}</cus:CustomerAccountId>"

    credit_limit_block = ""
    if clean_str(credit_limit):
        credit_limit_block += f"\n            <cus:CreditLimit>{xml_escape(credit_limit)}</cus:CreditLimit>"

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
                  xmlns:typ="http://xmlns.oracle.com/apps/financials/receivables/customers/customerProfileService/types/"
                  xmlns:cus="http://xmlns.oracle.com/apps/financials/receivables/customers/customerProfileService/">
   <soapenv:Header/>
   <soapenv:Body>
      <typ:createCustomerProfile>
         <typ:customerProfile>
            <cus:AccountNumber>{xml_escape(account_number)}</cus:AccountNumber>
            <cus:SiteNumber>{xml_escape(site_number)}</cus:SiteNumber>
            <cus:PartyId>{xml_escape(party_id)}</cus:PartyId>{cust_acc_block}
            <cus:ProfileClassName>{xml_escape(PROFILE_CLASS_NAME)}</cus:ProfileClassName>
            <cus:PaymentTerms>{xml_escape(payment_terms)}</cus:PaymentTerms>
            <cus:CreditChecking>{xml_escape(credit_checking)}</cus:CreditChecking>{credit_limit_block}
         </typ:customerProfile>
      </typ:createCustomerProfile>
   </soapenv:Body>
</soapenv:Envelope>"""


def build_site_update_payload(account_number, site_number, payment_terms, credit_checking, credit_limit=""):
    credit_limit_block = ""
    if clean_str(credit_limit):
        credit_limit_block += f"\n            <cus:CreditLimit>{xml_escape(credit_limit)}</cus:CreditLimit>"

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
                  xmlns:typ="http://xmlns.oracle.com/apps/financials/receivables/customers/customerProfileService/types/"
                  xmlns:cus="http://xmlns.oracle.com/apps/financials/receivables/customers/customerProfileService/">
   <soapenv:Header/>
   <soapenv:Body>
      <typ:updateCustomerProfile>
         <typ:customerProfile>
            <cus:AccountNumber>{xml_escape(account_number)}</cus:AccountNumber>
            <cus:SiteNumber>{xml_escape(site_number)}</cus:SiteNumber>
            <cus:PaymentTerms>{xml_escape(payment_terms)}</cus:PaymentTerms>
            <cus:CreditChecking>{xml_escape(credit_checking)}</cus:CreditChecking>{credit_limit_block}
         </typ:customerProfile>
      </typ:updateCustomerProfile>
   </soapenv:Body>
</soapenv:Envelope>"""


def classify_response(status_code, response_text):
    fault = extract_fault_text(response_text)

    if 200 <= status_code < 300 and "<Fault>" not in response_text and "<env:Fault>" not in response_text:
        return "SUCCESS", ""

    fault_upper = fault.upper()

    if "JBO-25020" in fault_upper or "VIEW ROW WITH KEY NULL IS NOT FOUND" in fault_upper:
        return "NOT_FOUND", fault

    if "AR-855641" in fault_upper or "UNIQUE VALUE FOR THE CUSTACCOUNTID COLUMN" in fault_upper:
        return "DUPLICATE_CUST_ACCOUNT_ID", fault

    if "ALREADY EXISTS" in fault_upper or "DUPLICATE" in fault_upper:
        return "DUPLICATE", fault

    return "ERROR", fault


# =========================================================
# MAIN LOGIC
# =========================================================

def load_input():
    if INPUT_SHEET is None:
        df = pd.read_excel(INPUT_FILE, dtype=str)
    else:
        df = pd.read_excel(INPUT_FILE, sheet_name=INPUT_SHEET, dtype=str)

    required = [
        "PARTY_NAME",
        "PaymentTerms",
        "CreditLimit",
        "PARTY_ID",
        "PARTY_SITE_NUMBER",
        "ACCOUNT_NUMBER",
        "CUST_ACCOUNT_ID",
    ]

    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    for col in required:
        df[col] = df[col].fillna("").astype(str).map(lambda x: x.strip())

    # remove blank key rows
    df = df[df["ACCOUNT_NUMBER"].str.strip() != ""].copy()

    return df.reset_index(drop=True)


def process_account_updates(df):
    account_logs = []

    unique_accounts = (
        df.sort_values(["ACCOUNT_NUMBER", "PARTY_SITE_NUMBER"], na_position="last")
          .drop_duplicates(subset=["ACCOUNT_NUMBER"], keep="first")
          .copy()
    )

    total = len(unique_accounts)

    for idx, (_, row) in enumerate(unique_accounts.iterrows(), start=1):
        account_number = clean_str(row["ACCOUNT_NUMBER"])
        payment_terms = clean_str(row["PaymentTerms"])

        credit_info = infer_credit_values(row["CreditLimit"])
        credit_checking = credit_info["credit_checking"]
        credit_limit = credit_info["credit_limit"]

        payload = build_account_update_payload(
            account_number=account_number,
            payment_terms=payment_terms,
            credit_checking=credit_checking,
            credit_limit=credit_limit
        )

        try:
            status_code, response_text = post_soap(payload)
            status, message = classify_response(status_code, response_text)
        except Exception as e:
            status_code = ""
            response_text = ""
            status = "ERROR"
            message = f"{type(e).__name__}: {e}"

        print(
            f"[ACCOUNT {idx}/{total}] "
            f"PARTY={clean_str(row['PARTY_NAME'])} | "
            f"ACCOUNT={account_number} | "
            f"STATUS={status}"
        )

        account_logs.append({
            "IDX": idx,
            "TOTAL": total,
            "LEVEL": "ACCOUNT",
            "PARTY_NAME": clean_str(row["PARTY_NAME"]),
            "ACCOUNT_NUMBER": account_number,
            "PARTY_SITE_NUMBER": "",
            "PARTY_ID": clean_str(row["PARTY_ID"]),
            "CUST_ACCOUNT_ID": clean_str(row["CUST_ACCOUNT_ID"]),
            "PaymentTerms": payment_terms,
            "DerivedCreditChecking": credit_checking,
            "DerivedCreditLimit": credit_limit,
            "ACTION": "UPDATE_ACCOUNT_PROFILE",
            "STATUS_CODE": status_code,
            "RESULT": status,
            "MESSAGE": message,
            "RAW_RESPONSE": response_text[:5000]
        })

        time.sleep(SLEEP_BETWEEN_CALLS)

    return pd.DataFrame(account_logs)


def process_site_profiles(df):
    site_logs = []

    unique_sites = (
        df.sort_values(["ACCOUNT_NUMBER", "PARTY_SITE_NUMBER"], na_position="last")
          .drop_duplicates(subset=["ACCOUNT_NUMBER", "PARTY_SITE_NUMBER"], keep="first")
          .copy()
    )

    total = len(unique_sites)

    for idx, (_, row) in enumerate(unique_sites.iterrows(), start=1):
        party_name = clean_str(row["PARTY_NAME"])
        account_number = clean_str(row["ACCOUNT_NUMBER"])
        site_number = clean_str(row["PARTY_SITE_NUMBER"])
        party_id = clean_str(row["PARTY_ID"])
        cust_account_id = clean_str(row["CUST_ACCOUNT_ID"])
        payment_terms = clean_str(row["PaymentTerms"])

        credit_info = infer_credit_values(row["CreditLimit"])
        credit_checking = credit_info["credit_checking"]
        credit_limit = credit_info["credit_limit"]

        final_result = "ERROR"
        final_message = ""
        final_action = ""
        final_status_code = ""
        final_response_text = ""

        try:
            # -------------------------------------------------
            # STEP 1: CREATE SITE PROFILE
            # -------------------------------------------------
            include_cust_account_id = TRY_SITE_CREATE_WITH_CUST_ACCOUNT_ID

            payload = build_site_create_payload(
                account_number=account_number,
                site_number=site_number,
                party_id=party_id,
                customer_account_id=cust_account_id,
                payment_terms=payment_terms,
                credit_checking=credit_checking,
                credit_limit=credit_limit,
                include_customer_account_id=include_cust_account_id
            )

            status_code, response_text = post_soap(payload)
            status, message = classify_response(status_code, response_text)

            final_action = "CREATE_SITE_PROFILE_WITH_CUST_ACCOUNT_ID" if include_cust_account_id else "CREATE_SITE_PROFILE"
            final_status_code = status_code
            final_response_text = response_text

            # -------------------------------------------------
            # STEP 2: RETRY CREATE WITHOUT CUSTOMERACCOUNTID
            # -------------------------------------------------
            if status == "DUPLICATE_CUST_ACCOUNT_ID" and include_cust_account_id and RETRY_SITE_CREATE_WITHOUT_CUST_ACCOUNT_ID:
                payload_retry = build_site_create_payload(
                    account_number=account_number,
                    site_number=site_number,
                    party_id=party_id,
                    customer_account_id="",
                    payment_terms=payment_terms,
                    credit_checking=credit_checking,
                    credit_limit=credit_limit,
                    include_customer_account_id=False
                )

                status_code_retry, response_text_retry = post_soap(payload_retry)
                status_retry, message_retry = classify_response(status_code_retry, response_text_retry)

                final_action = "CREATE_SITE_PROFILE_WITHOUT_CUST_ACCOUNT_ID"
                final_status_code = status_code_retry
                final_response_text = response_text_retry

                status = status_retry
                message = message_retry

            # -------------------------------------------------
            # STEP 3: FALLBACK TO UPDATE SITE PROFILE
            # -------------------------------------------------
            if status != "SUCCESS" and TRY_SITE_UPDATE_ON_CREATE_FAILURE:
                payload_update = build_site_update_payload(
                    account_number=account_number,
                    site_number=site_number,
                    payment_terms=payment_terms,
                    credit_checking=credit_checking,
                    credit_limit=credit_limit
                )

                status_code_upd, response_text_upd = post_soap(payload_update)
                status_upd, message_upd = classify_response(status_code_upd, response_text_upd)

                if status_upd == "SUCCESS":
                    final_action = "UPDATE_SITE_PROFILE_FALLBACK"
                    final_status_code = status_code_upd
                    final_response_text = response_text_upd
                    status = status_upd
                    message = message_upd

            final_result = status
            final_message = message

        except Exception as e:
            final_result = "ERROR"
            final_message = f"{type(e).__name__}: {e}\n{traceback.format_exc(limit=1)}"

        print(
            f"[SITE {idx}/{total}] "
            f"PARTY={party_name} | "
            f"ACCOUNT={account_number} | "
            f"SITE={site_number} | "
            f"STATUS={final_result}"
        )

        site_logs.append({
            "IDX": idx,
            "TOTAL": total,
            "LEVEL": "SITE",
            "PARTY_NAME": party_name,
            "ACCOUNT_NUMBER": account_number,
            "PARTY_SITE_NUMBER": site_number,
            "PARTY_ID": party_id,
            "CUST_ACCOUNT_ID": cust_account_id,
            "PaymentTerms": payment_terms,
            "DerivedCreditChecking": credit_checking,
            "DerivedCreditLimit": credit_limit,
            "ACTION": final_action,
            "STATUS_CODE": final_status_code,
            "RESULT": final_result,
            "MESSAGE": final_message,
            "RAW_RESPONSE": final_response_text[:5000]
        })

        time.sleep(SLEEP_BETWEEN_CALLS)

    return pd.DataFrame(site_logs)


def main():
    print("Reading input file...")
    df = load_input()

    # Add helper key
    df["SITE_KEY"] = df["ACCOUNT_NUMBER"].astype(str).str.strip() + "|" + df["PARTY_SITE_NUMBER"].astype(str).str.strip()

    print(f"Total input rows: {len(df)}")
    print(f"Unique accounts  : {df['ACCOUNT_NUMBER'].nunique()}")
    print(f"Unique sites     : {df['SITE_KEY'].nunique()}")

    account_log_df = pd.DataFrame()
    if DO_ACCOUNT_LEVEL_UPDATE:
        print("\nUpdating account-level profiles first...\n")
        account_log_df = process_account_updates(df)
    else:
        print("\nSkipping account-level updates...\n")

    print("\nCreating/updating site-level profiles next...\n")
    site_log_df = process_site_profiles(df)

    all_logs = pd.concat([account_log_df, site_log_df], ignore_index=True)

    success_df = all_logs[all_logs["RESULT"] == "SUCCESS"].copy()
    error_df = all_logs[all_logs["RESULT"] != "SUCCESS"].copy()

    summary = pd.DataFrame([
        {"Metric": "Total Input Rows", "Value": len(df)},
        {"Metric": "Unique Accounts", "Value": df["ACCOUNT_NUMBER"].nunique()},
        {"Metric": "Unique Sites", "Value": df["SITE_KEY"].nunique()},
        {"Metric": "Account Calls", "Value": len(account_log_df)},
        {"Metric": "Site Calls", "Value": len(site_log_df)},
        {"Metric": "Successful Calls", "Value": len(success_df)},
        {"Metric": "Errored Calls", "Value": len(error_df)},
    ])

    print("\nWriting recon file...")
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        df.to_excel(writer, sheet_name="InputData", index=False)
        all_logs.to_excel(writer, sheet_name="AllResults", index=False)
        success_df.to_excel(writer, sheet_name="Success", index=False)
        error_df.to_excel(writer, sheet_name="Errors", index=False)

    print("\nDone.")
    print(f"Recon file saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()