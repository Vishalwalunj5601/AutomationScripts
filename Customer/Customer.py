import pandas as pd

INPUT_FILE = r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\Customer\Airetech\SourceFile\Airetech_Customer -so Equipment.xlsx"
OUTPUT_FILE = r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\Customer\Airetech\SourceFile\Airetech_Customer -So Equipment-retr.xlsx"

BASE_IDX = 500000

df = pd.read_excel(INPUT_FILE)

customer_rows = []
contact_rows = []

recon_customer_rows = []
recon_site_rows = []

cust_idx_map = {}
customer_site_counter = {}
customer_person_counter = {}
customer_primary_assigned = {}
customer_sites = {}

# ✅ NEW: Purpose merge map
site_purpose_map = {}

current_idx = BASE_IDX


def split_name(full_name):
    # split full name components
    parts = str(full_name).strip().split(" ")
    if len(parts) == 1:
        return parts[0], "", "."
    elif len(parts) == 2:
        return parts[0], "", parts[1]
    else:
        return parts[0], " ".join(parts[1:-1]), parts[-1]


def has_valid_email(email):
    return pd.notna(email) and str(email).strip() != ""


# ==========================================
# PASS 1 → VALIDATION + PURPOSE MERGE
# ==========================================
for _, r in df.iterrows():

    customer = str(r.get("entity_name_ora")).strip()
    ship_flag = str(r.get("Ship to's")).strip()
    bill_flag = str(r.get("Bill-to")).strip()

    site1 = r.get("SITE1_address1_ora")
    site2 = r.get("SITE2_address1_ora")

    reason = None
    if not customer or customer.upper() == "NAN":
        reason = "MISSING_CUSTOMER_NAME"
    elif pd.isna(site1) and pd.isna(site2):
        reason = "NO_SITE_ADDRESS"
    elif ship_flag != "Y" and bill_flag != "Y":
        reason = "NO_PURPOSE_FLAG"
    elif pd.isna(r.get("country_ora")):
        reason = "MISSING_COUNTRY"

    if reason:
        recon_customer_rows.append({
            "Customer Name": customer,
            "Reason": reason,
            "Status": "EXCLUDED"
        })
        continue

    if customer not in cust_idx_map:
        cust_idx_map[customer] = current_idx
        customer_site_counter[customer] = 0
        customer_person_counter[customer] = 1
        customer_primary_assigned[customer] = False
        customer_sites[customer] = []
        current_idx += 1

    sites = []
    if pd.notna(site1) and str(site1) != "":
        sites.append(str(site1))
    if pd.notna(site2) and str(site2) != "":
        sites.append(str(site2))

    for address in sites:
        key = (customer, address)

        if key not in site_purpose_map:
            site_purpose_map[key] = {
                "SHIP": False,
                "BILL": False,
                "row": r
            }

        if ship_flag == "Y":
            site_purpose_map[key]["SHIP"] = True

        if bill_flag == "Y":
            site_purpose_map[key]["BILL"] = True


# ==========================================
# PASS 2 → CREATE ADFDI
# ==========================================
for (customer, address_str), val in site_purpose_map.items():

    r = val["row"]
    ship_flag = val["SHIP"]
    bill_flag = val["BILL"]

    idx = cust_idx_map[customer]

    site_index = customer_site_counter[customer]
    customer_site_counter[customer] += 1

    site_ref = f"AT-S{site_index}-{idx}"
    loc_ref = f"AT-LOC{site_index}-{idx}"
    ship_site_use_ref = f"AT-SH{site_index}-{idx}"
    bill_site_use_ref = f"AT-BT{site_index}-{idx}"

    customer_sites[customer].append(site_ref)

    city_val = str(r.get("city_ora")).strip()
    site_name = (
        f"AIRETECH-SITE_12{city_val}{str(site_index + 1).zfill(2)}"
        if city_val and city_val.upper() != "NAN"
        else f"CITY-{str(site_index + 1).zfill(2)}"
    )

    base_customer = {
        "*Source System": "CSV",
        "*Customer Number": "*",
        "Customer Source Reference": f"AT-{idx}",
        "*Customer Name": customer,
        "*Account Number": "*",
        "Account Source Reference": f"AT-ACC-{idx}",
        "Account Description": customer,
        "Account Type": "R",
        "Customer Class": "CUSTOMER",
        "Customer Profile Class": "DEFAULT",
        "*Site Number": "*",
        "Site Source Reference": site_ref,
        "Site Name": site_name,
        "*Account Address Set": "ENTERPRISE",
        "*Location Source Reference": loc_ref,
        "*Address Line 1": address_str,
        "Address Line 2": r.get("address2_ora"),
        "Address Line 3": r.get("address3_ora"),
        "City": r.get("city_ora"),
        "State": r.get("state_ora"),
        "Postal Code": r.get("zip_ora"),
        "*Country": r.get("country_ora"),
        "*Identifying Address": "Y",
        "Site Language": "US",
        "Tax Identifier": r.get("TaxIdentifier_ora"),
        "Payment Term": r.get("PaymentTerm_ora"),
        "Bill-to Site Use Reference": ""
    }

    # ✅ SHIP-TO
    if ship_flag:
        customer_rows.append({
            **base_customer,
            "*Purpose": "SHIP_TO",
            "*Site Purpose Source Reference": ship_site_use_ref,
            "*Site Purpose Primary Indicator": "N",
            "Bill-to Site Use Reference": bill_site_use_ref if bill_flag else ""
        })

    # ✅ BILL-TO
    if bill_flag:
        customer_rows.append({
            **base_customer,
            "*Purpose": "BILL_TO",
            "*Site Purpose Source Reference": bill_site_use_ref,
            "*Site Purpose Primary Indicator": "Y",
            "Bill-to Site Use Reference": ""
        })


# ==========================================
# CONTACTS (UNCHANGED)
# ==========================================
for _, r in df.iterrows():

    customer = str(r.get("entity_name_ora")).strip()
    if customer not in cust_idx_map:
        continue

    idx = cust_idx_map[customer]

    contacts = [
        ("ContactPerson1_ora", "PhoneNumber1_ora", None, "EmailAddress1_ora"),
        ("ContactPerson2_ora", None, "PhoneNumber2_ora", "EmailAddress2_ora")
    ]

    for name_col, phone_col, mobile_col, email_col in contacts:

        name = r.get(name_col)
        if not name or str(name).strip().upper() == "NAN":
            continue

        phone = r.get(phone_col) if phone_col else None
        mobile = r.get(mobile_col) if mobile_col else None
        email = r.get(email_col)
        email_ok = has_valid_email(email)

        first, middle, last = split_name(name)

        person_idx = customer_person_counter[customer]
        person_ref = f"AT-PER{person_idx}-{idx}"
        customer_person_counter[customer] += 1

        primary_contact = "Y" if not customer_primary_assigned[customer] else "N"
        customer_primary_assigned[customer] = True

        for site_ref in customer_sites.get(customer, []):

            contact_rows.append({
                "*Source System": "CSV",
                "*Account Number": "*",
                "Site Number": site_ref,
                "Person Source Reference": person_ref,
                "*First Name": first,
                "Middle Name": middle,
                "*Last Name": last,
                "Primary Contact Indicator": primary_contact,
                "Unformatted Phone Number": phone,
                "Mobile Phone Number": mobile,
                "E-Mail Address": email if email_ok else ""
            })


# ==========================================
# DATAFRAMES + SUMMARY (UNCHANGED)
# ==========================================
adfdi_customer_df = pd.DataFrame(customer_rows)
adfdi_contact_df = pd.DataFrame(contact_rows)

recon_customer_df = pd.DataFrame(recon_customer_rows)
recon_site_df = pd.DataFrame(recon_site_rows)

summary_df = pd.DataFrame({
    "Metric": ["Generated Customers", "Generated Sites"],
    "Count": [len(adfdi_customer_df), len(site_purpose_map)]
})

with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
    adfdi_customer_df.to_excel(writer, sheet_name="ADFDi_Customers", index=False)
    adfdi_contact_df.to_excel(writer, sheet_name="ADFDi_Contacts", index=False)
    recon_customer_df.to_excel(writer, sheet_name="RECON_EXCLUDED_CUSTOMERS", index=False)
    recon_site_df.to_excel(writer, sheet_name="RECON_EXCLUDED_SITES", index=False)
    summary_df.to_excel(writer, sheet_name="SUMMARY_STATS", index=False)

print("✅ ADFDi generated with correct BILL-TO + SHIP-TO merging")
print("✅ ADFDi generated with correct BILL-TO + SHIP-TO merging")
