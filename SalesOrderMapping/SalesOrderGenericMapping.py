import re
import warnings
import pandas as pd
from difflib import SequenceMatcher

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =========================================================
# CONFIG
# =========================================================
INPUT_FILE = r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\SalesOrderMapping\DEV1 AT\AiretechCustomerMapping.xlsx"

SOURCE_SHEET = "Source"
FUSION_SHEET = "Fusion"
GENERIC_SHEET = "Generic"

OUTPUT_FILE = INPUT_FILE.replace(".xlsx", "Mapping.xlsx")

# SOURCE
SOURCE_IDENTIFIER_COL = "Identifier"
SOURCE_CUSTOMER_COL = "Customer"
SOURCE_SHIP_ADDR1 = "ShipToAddress1"
SOURCE_SHIP_ADDR2 = "ShipToAddress1.1"

# FUSION
FUSION_PARTY_NAME_COL = "PARTY_NAME"
FUSION_PARTY_NUMBER_COL = "PARTY_NUMBER"
FUSION_ACCOUNT_NUMBER_COL = "ACCOUNT_NUMBER"
FUSION_PARTY_SITE_ID_COL = "PARTY_SITE_ID"
FUSION_PARTY_SITE_NUMBER_COL = "PARTY_SITE_NUMBER"
FUSION_SITE_USE_ID_COL = "SITE_USE_ID"
FUSION_SITE_USE_CODE_COL = "SITE_USE_CODE"
FUSION_PRIMARY_FLAG_COL = "PRIMARY_FLAG"
FUSION_ADDRESS_COL = "ADDRESS1"

FUZZY_THRESHOLD = 65
ADDRESS_THRESHOLD = 65


# =========================================================
# HELPERS
# =========================================================
def clean_text(val):
    return "" if pd.isna(val) else str(val).strip()


def normalize_customer_name(name):
    name = clean_text(name).upper()
    name = name.replace("&", " AND ")
    name = re.sub(r"\s*-\s*.*$", "", name)
    name = re.sub(r"[^A-Z0-9 ]", " ", name)

    remove_words = {"INC", "LLC", "CORP", "CORPORATION", "CO", "COMPANY", "LTD", "LIMITED", "GROUP", "THE"}
    tokens = [t for t in name.split() if t not in remove_words]

    return re.sub(r"\s+", " ", " ".join(tokens)).strip().lower()


def normalize_address(val):
    val = clean_text(val).upper()

    # handle ATTN notations
    val = re.sub(r"ATTN.*$", "", val)

    val = re.sub(r"\b[A-Z0-9]{5,}\b$", "", val)

    val = val.replace("SUITE", "STE").replace("APARTMENT", "APT").replace("BUILDING", "BLDG")
    val = re.sub(r"[^A-Z0-9 ]", " ", val)

    replacements = {
        "STREET": "ST", "ROAD": "RD", "AVENUE": "AVE",
        "DRIVE": "DR", "LANE": "LN", "BOULEVARD": "BLVD"
    }

    for k, v in replacements.items():
        val = val.replace(k, v)

    return re.sub(r"\s+", " ", val).strip()


def similarity(a, b):
    return round(SequenceMatcher(None, a, b).ratio() * 100, 2)


def has_number(txt):
    return bool(re.search(r"\d+", txt))


def is_bad_generic(addr):
    bad_words = ["ATTN", "PAYABLE", "ACCOUNTS"]
    return any(word in addr.upper() for word in bad_words)


def address_score(a, b):
    base = similarity(a, b)

    if has_number(a) and has_number(b):
        if re.findall(r"\d+", a) == re.findall(r"\d+", b):
            base += 10

    if has_number(a) != has_number(b):
        base -= 20

    if a in b or b in a:
        base = max(base, 90)

    return max(0, min(base, 100))


def normalize_site(val):
    v = clean_text(val).upper().replace(" ", "").replace("-", "")
    if v == "BILLTO": return "BILL_TO"
    if v == "SHIPTO": return "SHIP_TO"
    return clean_text(val).upper()


def priority(val):
    return 1 if clean_text(val).upper() == "Y" else 0


# =========================================================
# LOAD
# =========================================================
source_df = pd.read_excel(INPUT_FILE, sheet_name=SOURCE_SHEET, dtype=str).fillna("")
fusion_df = pd.read_excel(INPUT_FILE, sheet_name=FUSION_SHEET, dtype=str).fillna("")
generic_df = pd.read_excel(INPUT_FILE, sheet_name=GENERIC_SHEET, dtype=str).fillna("")

source_df.columns = source_df.columns.str.strip()
fusion_df.columns = fusion_df.columns.str.strip()
generic_df.columns = generic_df.columns.str.strip()

# =========================================================
# NORMALIZE
# =========================================================
source_df["CUST_NORM"] = source_df[SOURCE_CUSTOMER_COL].apply(normalize_customer_name)
source_df["ADDR1_NORM"] = source_df[SOURCE_SHIP_ADDR1].apply(normalize_address)
source_df["ADDR2_NORM"] = source_df[SOURCE_SHIP_ADDR2].apply(normalize_address)

fusion_df["CUST_NORM"] = fusion_df[FUSION_PARTY_NAME_COL].apply(normalize_customer_name)
fusion_df["ADDR_NORM"] = fusion_df[FUSION_ADDRESS_COL].apply(normalize_address)
fusion_df["SITE_NORM"] = fusion_df[FUSION_SITE_USE_CODE_COL].apply(normalize_site)

generic_df["ADDR_NORM"] = generic_df[FUSION_ADDRESS_COL].apply(normalize_address)

# =========================================================
# GROUPING
# =========================================================
group_raw = {k.upper(): g for k, g in fusion_df.groupby(FUSION_PARTY_NAME_COL)}
group_norm = {k: g for k, g in fusion_df.groupby("CUST_NORM")}

fusion_customer_groups = []
for party_name, grp in fusion_df.groupby(FUSION_PARTY_NAME_COL):
    fusion_customer_groups.append({
        "party_name": clean_text(party_name),
        "party_name_norm": normalize_customer_name(party_name),
        "rows": grp.copy()
    })

# =========================================================
# PROCESS
# =========================================================
results = []

for _, row in source_df.iterrows():

    notes = ""  # ✅ NEW COLUMN

    identifier = clean_text(row[SOURCE_IDENTIFIER_COL])
    cust = clean_text(row[SOURCE_CUSTOMER_COL])
    cust_norm = row["CUST_NORM"]

    addr1 = row["ADDR1_NORM"]
    addr2 = row["ADDR2_NORM"]

    match_rows = None
    match_type = "NOT_FOUND"
    score = 0

    # CUSTOMER MATCH
    if cust.upper() in group_raw:
        match_rows = group_raw[cust.upper()]
        match_type = "EXACT"
        score = 100

    elif cust_norm in group_norm:
        match_rows = group_norm[cust_norm]
        match_type = "NORMALIZED"
        score = 100

    else:
        best_score = 0
        best_group = None

        for item in fusion_customer_groups:
            s = similarity(cust_norm, item["party_name_norm"])

            if cust_norm in item["party_name_norm"] or item["party_name_norm"] in cust_norm:
                s = min(100, s + 10)

            if s > best_score:
                best_score = s
                best_group = item

        if best_group is not None and best_score >= FUZZY_THRESHOLD:
            match_rows = best_group["rows"]
            match_type = "FUZZY"
            score = best_score

    # BILL TO
    bill_party = bill_account = bill_site_id = bill_site_use = bill_site_num = ""
    fusion_party_name = ""

    if match_rows is not None:
        fusion_party_name = clean_text(match_rows.iloc[0][FUSION_PARTY_NAME_COL])

        temp = match_rows.copy()
        temp["_PRI"] = temp[FUSION_PRIMARY_FLAG_COL].apply(priority)

        master = temp.sort_values(by="_PRI", ascending=False).iloc[0]
        bill_party = master[FUSION_PARTY_NUMBER_COL]
        bill_account = master[FUSION_ACCOUNT_NUMBER_COL]

        bill_rows = temp[temp["SITE_NORM"] == "BILL_TO"]
        if not bill_rows.empty:
            b = bill_rows.sort_values(by="_PRI", ascending=False).iloc[0]
            bill_site_use = b[FUSION_SITE_USE_ID_COL]
            bill_site_id = b[FUSION_PARTY_SITE_ID_COL]
            bill_site_num = b[FUSION_PARTY_SITE_NUMBER_COL]


    # SHIP TO
    def find_ship(address):
        if not address:
            return None, "", "", ""

        if match_rows is not None:
            temp = match_rows[match_rows["SITE_NORM"] == "SHIP_TO"].copy()
            if not temp.empty:
                temp["SCORE"] = temp["ADDR_NORM"].apply(lambda x: address_score(address, x))
                best = temp.sort_values(by="SCORE", ascending=False).iloc[0]
                if best["SCORE"] >= ADDRESS_THRESHOLD:
                    return best, "FUSION", best["SCORE"], best[FUSION_ADDRESS_COL]

        temp = generic_df.copy()
        temp["SCORE"] = temp["ADDR_NORM"].apply(lambda x: address_score(address, x))
        temp = temp.sort_values(by="SCORE", ascending=False)

        for _, best in temp.iterrows():
            if best["SCORE"] < ADDRESS_THRESHOLD:
                break
            if is_bad_generic(best[FUSION_ADDRESS_COL]):
                continue
            return best, "GENERIC", best["SCORE"], best[FUSION_ADDRESS_COL]

        return None, "", 0, ""


    ship_row, ship_source, ship_score, matched_addr = find_ship(addr1)
    addr_used = "ADDR1"

    if ship_row is None:
        ship_row, ship_source, ship_score, matched_addr = find_ship(addr2)
        addr_used = "ADDR2"

    # ✅ NOTES (NO LOGIC CHANGE)
    if match_rows is not None:
        if not addr1 and not addr2:
            notes = "No address provided"
        elif ship_row is None:
            notes = "Address not matched to SHIP_TO"
        else:
            if ship_source == "FUSION":
                notes = "Matched with Fusion SHIP_TO"
            elif ship_source == "GENERIC":
                notes = "Matched with Generic"

    ship_party = ship_account = ship_site_id = ship_site_use = ship_site_num = ""

    if ship_row is not None:
        ship_party = clean_text(ship_row.get(FUSION_PARTY_NUMBER_COL, ""))
        ship_account = clean_text(ship_row.get(FUSION_ACCOUNT_NUMBER_COL, ""))
        ship_site_use = clean_text(ship_row.get(FUSION_SITE_USE_ID_COL, ""))
        ship_site_id = clean_text(ship_row.get(FUSION_PARTY_SITE_ID_COL, ""))
        ship_site_num = clean_text(ship_row.get(FUSION_PARTY_SITE_NUMBER_COL, ""))

    # RESULT
    results.append({
        "Identifier": identifier,
        "Source_Customer": cust,
        "Source_ShipToAddress1": row[SOURCE_SHIP_ADDR1],
        "Source_ShipToAddress2": row[SOURCE_SHIP_ADDR2],
        "Fusion_PARTY_NAME": fusion_party_name,
        "Match_Type": match_type,
        "Match_Percent": score,
        "BillTo_PARTY_NUMBER": bill_party,
        "BillTo_ACCOUNT_NUMBER": bill_account,
        "BillTo_SITE_USE_ID": bill_site_use,
        "BillTo_PARTY_SITE_ID": bill_site_id,
        "BillTo_SITE_NUMBER": bill_site_num,
        "ShipTo_PARTY_NUMBER": ship_party,
        "ShipTo_ACCOUNT_NUMBER": ship_account,
        "ShipTo_SITE_USE_ID": ship_site_use,
        "ShipTo_PARTY_SITE_ID": ship_site_id,
        "ShipTo_SITE_NUMBER": ship_site_num,
        "ShipTo_Source": ship_source,
        "ShipTo_Match_Score": ship_score,
        "ShipTo_Matched_Address": matched_addr,
        "Address1_Used": addr_used,
        "Notes": notes  # ✅ ADDED ONLY
    })

# SAVE
pd.DataFrame(results).to_excel(OUTPUT_FILE, index=False)

print("✅ Done:", OUTPUT_FILE)