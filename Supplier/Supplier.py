import pandas as pd
import re

# ==========================================================
# FILE PATHS
# ==========================================================
Bussiness_Unit= input("Enter Bussiness Unit: ")
INPUT_FILE = r"C:\Users\Shruti Pawar\Workspace\DataConversion\SupplierInputFile.xlsx"
OUTPUT_FILE = r"C:\Users\Shruti Pawar\Workspace\1D_Load\Supplier\Etairos\Output\Etairos_SupplierOutput.xlsx"

# ==========================================================
# READ SOURCE FILE
# ==========================================================

df = pd.read_excel(INPUT_FILE)

df = df.fillna("")

# ==========================================================
# HELPER FUNCTIONS
# ==========================================================

def clean_digits(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""

    value = str(value)

    # Remove extension
    value = value.split("#")[0]

    digits = re.sub(r"\D", "", value)

    return digits


def split_phone(phone):

    digits = clean_digits(phone)

    if digits == "":
        return "", "", ""

    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]

    if len(digits) < 10:
        return "", "", ""

    country_code = "1"
    area_code = digits[:3]
    phone_number = digits[-7:]

    return country_code, area_code, phone_number


def split_fax(fax):

    digits = clean_digits(fax)

    if digits == "":
        return "", "", ""

    if len(digits) < 10:
        return "", "", ""

    area_code = digits[:3]
    fax_number = digits[-7:]

    return digits, area_code, fax_number


# ==========================================================
# ADDRESS NAME GENERATION
# ==========================================================

city_counter = {}

address_names = []

for _, row in df.iterrows():

    city = str(row["City"]).strip().upper()

    if city not in city_counter:
        city_counter[city] = 0
        address_name = city
    else:
        city_counter[city] += 1
        address_name = city + str(city_counter[city]).zfill(2)

    address_names.append(address_name)

df["ADDRESS_NAME"] = address_names

# ==========================================================
# PRIMARY PAY LOGIC
# ==========================================================

supplier_counter = {}

primary_pay_list = []

for _, row in df.iterrows():

    supplier = str(row["DataConversion Name"]).strip()

    supplier_counter.setdefault(supplier, 0)
    supplier_counter[supplier] += 1

supplier_occurrence = (
    df.groupby("DataConversion Name")
      .size()
      .to_dict()
)

supplier_seen = {}

for _, row in df.iterrows():

    supplier = str(row["DataConversion Name"]).strip()

    supplier_seen.setdefault(supplier, 0)
    supplier_seen[supplier] += 1

    if supplier_occurrence[supplier] == 1:
        primary_pay_list.append("Y")
    else:
        if supplier_seen[supplier] == 1:
            primary_pay_list.append("Y")
        else:
            primary_pay_list.append("N")

df["PRIMARY_PAY"] = primary_pay_list

def format_taxpayer_id(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""

    # Keep only digits
    digits = re.sub(r"\D", "", str(value))

    # Format as XX-XXXXXXX
    if len(digits) == 9:
        return f"{digits[:2]}-{digits[2:]}"
    else:
        return digits  # Return as-is if not exactly 9 digits

# ==========================================================
# SHEET 1
# POZ_SUPPLIERS_INT
# ==========================================================

suppliers_df = pd.DataFrame()

suppliers_df["Import Action"] = "CREATE"
suppliers_df["DataConversion Name"] = df["DataConversion Name"]
suppliers_df["Tax Organization Type"] = "Corporation"
suppliers_df["DataConversion Type"] = "DataConversion"
suppliers_df["Business Relationship"] = "SPEND_AUTHORIZED"
suppliers_df["Alias"] = df["Alias"]
suppliers_df["Taxpayer Country"] = "US"
#suppliers_df["Taxpayer ID"] = df["Taxpayer ID"]
suppliers_df["Taxpayer ID"] = df["Taxpayer ID"].apply(format_taxpayer_id)
suppliers_df["Federal reportable"] = df["Track 1099"]
suppliers_df["Federal Type"] = df["Federal Type"]
suppliers_df["Payment Method"]="Check"

# ==========================================================
# SHEET 2
# POZ_SUPPLIER_ADDRESSES_INT
# ==========================================================

address_rows = []

for _, row in df.iterrows():

    phone_cc, phone_area, phone_num = split_phone(row["Phone"])
    fax_cc, fax_area, fax_num = split_fax(row["Fax"])

    address_rows.append({
        "Import Action": "CREATE",
        "DataConversion Name": row["DataConversion Name"],
        "Address Name": row["ADDRESS_NAME"],
        "Country": row["Country"],
        "Address Line 1": row["Address Line 1"],
        "Address Line 2": row["Address Line 2"],
        "Address Line 3": row["Address Line 3"],
        "Address Line 4": row["Address Line 4"],
        "City": row["City"],
        "State": row["State"],
        "Postal Code": row["Postal code"],
        "Phone Country Code": phone_cc,
        "Phone Area Code": phone_area,
        "Phone": phone_num,
        "Fax Country Code": fax_cc,
        "Fax Area Code": fax_area,
        "Fax": fax_num,
        "Ordering": "Y",
        "Pay": "Y",
        "E-Mail": row["E-Mail"],
        "E-Mail2":row["E-Mail2"]
    })

addresses_df = pd.DataFrame(address_rows)

# ==========================================================
# SHEET 3
# POZ_SUPPLIER_SITES_INT
# ==========================================================

site_rows = []

for _, row in df.iterrows():

    communication_method = ""

    if str(row["E-Mail"]).strip():
        communication_method = "EMAIL"

    elif str(row["Fax"]).strip():
        communication_method = "FAX"

    fax_cc = ""
    fax_area = ""
    fax_num = ""

    if communication_method == "FAX":
        fax_cc, fax_area, fax_num = split_fax(row["Fax"])

    site_rows.append({
        "Import Action": "CREATE",
        "DataConversion Name": row["DataConversion Name"],
        "Procurement BU": Bussiness_Unit,
        "Address Name": row["ADDRESS_NAME"],
        "DataConversion Site": row["ADDRESS_NAME"],
        "Sourcing Only": "N",
        "Purchasing": "Y",
        "Pay": "Y",
        "Primary Pay": row["PRIMARY_PAY"],
        "Communication Method": communication_method,
        "E-Mail": row["E-Mail"] if communication_method == "EMAIL" else "",
        "Fax Country Code": fax_cc,
        "Fax Area Code": fax_area,
        "Fax": fax_num,
        "Create Debit Memo From Return": "Y",
        "Payment Terms": row["Payment Terms"],
        "Tms Cd": row["Tms Cd"],
        "ATTRIBUTE1": row["legacy no"],
        "ATTRIBUTE2": row["Attribute 2"]
    })

sites_df = pd.DataFrame(site_rows)

# ==========================================================
# SHEET 4
# POZ_SITE_ASSIGNMENTS_INT
# ==========================================================

assignments_df = pd.DataFrame()

assignments_df["Import Action"] = "CREATE"
assignments_df["DataConversion Name"] = df["DataConversion Name"]
assignments_df["DataConversion Site"] = df["ADDRESS_NAME"]
assignments_df["Procurement BU"] = Bussiness_Unit
assignments_df["Client BU"] = Bussiness_Unit
assignments_df["Bill-to BU"] = Bussiness_Unit

# ==========================================================
# SHEET 5
# POZ_SUP_CONTACTS
# ==========================================================

contact_rows = []
prevalidation_rows = []

for _, row in df.iterrows():

    first_name = str(row["First Name"]).strip()
    last_name = str(row["Last Name"]).strip()

    if first_name == "" or last_name == "":

        prevalidation_rows.append({
            **row.to_dict(),
            "Remark": "DataConversion dont have contact details"
        })

        continue

    phone_cc, phone_area, phone_num = split_phone(row["Phone"])
    phone2_cc, phone2_area, phone2_num = split_phone(row["Phone2"])

    fax_cc, fax_area, fax_num = split_fax(row["Fax"])

    contact_rows.append({
        "Import Action": "CREATE",
        "DataConversion Name": row["DataConversion Name"],
        "First Name": first_name,
        "Last Name": last_name,
        "E-Mail": row["E-Mail"],
        "Email2": row["E-Mail2"],
        "Phone Country Code": phone_cc,
        "Phone Area Code": phone_area,
        "Phone": phone_num,
        "Fax Country Code": fax_cc,
        "Fax Area Code": fax_area,
        "Fax": fax_num,
        "Phone2 Country Code": phone2_cc,
        "Phone2 Area Code": phone2_area,
        "Phone2": phone2_num
    })

contacts_df = pd.DataFrame(contact_rows)

# ==========================================================
# SHEET 6
# POZ_SUPP_CONTACT_ADDRESSES_INT
# ==========================================================

contact_address_rows = []

for _, row in df.iterrows():

    first_name = str(row["First Name"]).strip()
    last_name = str(row["Last Name"]).strip()

    if first_name == "" or last_name == "":
        continue

    contact_address_rows.append({
        "Import Action": "CREATE",
        "DataConversion Name": row["DataConversion Name"],
        "Address Name": row["ADDRESS_NAME"],
        "First Name": first_name,
        "Last Name": last_name,
        "E-Mail": row["E-Mail"],
        "Email2": row["E-Mail2"]
    })

contact_addresses_df = pd.DataFrame(contact_address_rows)

prevalidation_df = pd.DataFrame(prevalidation_rows)


# ==========================================================
# DATA CLEANUP / FILTERING
# ==========================================================

suppliers_df = suppliers_df.drop_duplicates(subset=["DataConversion Name"], keep="first").reset_index(drop=True)

duplicate_addresses_df = addresses_df[addresses_df.duplicated(
    subset=["DataConversion Name","Address Line 1"], keep=False
)].copy()

addresses_df = addresses_df.drop_duplicates(
    subset=["DataConversion Name","Address Line 1"], keep="first"
).reset_index(drop=True)

valid_addresses = addresses_df[["DataConversion Name","Address Name"]].drop_duplicates()

sites_df = sites_df.merge(
    valid_addresses,
    on=["DataConversion Name","Address Name"],
    how="inner"
)

valid_sites = sites_df[["DataConversion Name","DataConversion Site"]].drop_duplicates()

assignments_df = assignments_df.merge(
    valid_sites,
    on=["DataConversion Name","DataConversion Site"],
    how="inner"
)

# ==========================================================
# WRITE OUTPUT
# ==========================================================

with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
) as writer:

    suppliers_df.to_excel(
        writer,
        sheet_name="POZ_SUPPLIERS_INT",
        index=False
    )

    addresses_df.to_excel(
        writer,
        sheet_name="POZ_SUPPLIER_ADDRESSES_INT",
        index=False
    )

    sites_df.to_excel(
        writer,
        sheet_name="POZ_SUPPLIER_SITES_INT",
        index=False
    )

    assignments_df.to_excel(
        writer,
        sheet_name="POZ_SITE_ASSIGNMENTS_INT",
        index=False
    )

    contacts_df.to_excel(
        writer,
        sheet_name="POZ_SUP_CONTACTS",
        index=False
    )

    contact_addresses_df.to_excel(
        writer,
        sheet_name="POZ_SUPP_CONTACT_ADDRESSES_INT",
        index=False
    )

    duplicate_addresses_df.to_excel(
        writer,
        sheet_name="Duplicate Address",
        index=False
    )

    prevalidation_df.to_excel(
        writer,
        sheet_name="PREVALIDATION",
        index=False
    )

print("DataConversion FBDI file created successfully.")
print(OUTPUT_FILE)
