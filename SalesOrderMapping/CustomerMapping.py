import re
import warnings
import pandas as pd
from difflib import SequenceMatcher

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =========================================================
# CONFIG
# =========================================================
INPUT_FILE = r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\SalesOrderMapping\EtairosCustomerAR.xlsx"
SOURCE_SHEET = "Source"
FUSION_SHEET = "Fusion"

OUTPUT_FILE = INPUT_FILE.replace(".xlsx", "_Mapped_Output.xlsx")

SOURCE_IDENTIFIER_COL = "Identifier"
SOURCE_CUSTOMER_COL   = "Customer"

FUSION_PARTY_NAME_COL        = "PARTY_NAME"
FUSION_PARTY_NUMBER_COL      = "PARTY_NUMBER"
FUSION_PARTY_SITE_ID_COL     = "PARTY_SITE_ID"
FUSION_PARTY_SITE_NUMBER_COL = "PARTY_SITE_NUMBER"
FUSION_SITE_USE_ID_COL       = "SITE_USE_ID"
FUSION_SITE_USE_CODE_COL     = "SITE_USE_CODE"
FUSION_PRIMARY_FLAG_COL      = "PRIMARY_FLAG"

FUZZY_THRESHOLD = 75

# =========================================================
# HELPERS
# =========================================================
def clean_text(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

def normalize_customer_name(name):
    """
    Example:
    REMCO - Allentown -> remco
    """
    name = clean_text(name).upper()

    name = name.replace("&", " AND ")

    # remove anything after hyphen
    name = re.sub(r"\s*-\s*.*$", "", name)

    # remove punctuation
    name = re.sub(r"[^A-Z0-9 ]", " ", name)

    # remove common suffixes
    remove_words = {
        "INC", "LLC", "L", "C", "CORP", "CORPORATION", "CO", "COMPANY",
        "LTD", "LIMITED", "GROUP", "THE", "PC", "PLC", "LP", "LLP"
    }

    tokens = [t for t in name.split() if t not in remove_words]
    name = " ".join(tokens)
    name = re.sub(r"\s+", " ", name).strip()

    return name.lower()

def similarity(a, b):
    return round(SequenceMatcher(None, a, b).ratio() * 100, 2)

def normalize_site_use_code(val):
    v = clean_text(val).upper().replace(" ", "").replace("-", "").replace("_", "")
    if v == "BILLTO":
        return "BILL_TO"
    if v == "SHIPTO":
        return "SHIP_TO"
    return clean_text(val).upper()

def priority_flag(val):
    return 1 if clean_text(val).upper() == "Y" else 0

def pick_preferred_row(df, site_use_code):
    """
    Pick one row for the matched customer and given site use code.
    Preference:
    1. matching site use code
    2. PRIMARY_FLAG = Y
    3. first available row
    """
    if df.empty:
        return None

    temp = df.copy()
    temp["_SITE_USE_CODE_NORM"] = temp[FUSION_SITE_USE_CODE_COL].apply(normalize_site_use_code)
    temp["_PRIMARY_SORT"] = temp[FUSION_PRIMARY_FLAG_COL].apply(priority_flag)

    temp = temp[temp["_SITE_USE_CODE_NORM"] == site_use_code].copy()
    if temp.empty:
        return None

    temp = temp.sort_values(
        by=["_PRIMARY_SORT", FUSION_PARTY_SITE_NUMBER_COL],
        ascending=[False, True]
    )

    return temp.iloc[0]

# =========================================================
# LOAD FILES
# =========================================================
source_df = pd.read_excel(INPUT_FILE, sheet_name=SOURCE_SHEET, dtype=str)
fusion_df = pd.read_excel(INPUT_FILE, sheet_name=FUSION_SHEET, dtype=str)

source_df.columns = source_df.columns.str.strip()
fusion_df.columns = fusion_df.columns.str.strip()

required_source_cols = [
    SOURCE_IDENTIFIER_COL,
    SOURCE_CUSTOMER_COL
]

required_fusion_cols = [
    FUSION_PARTY_NAME_COL,
    FUSION_PARTY_NUMBER_COL,
    FUSION_PARTY_SITE_ID_COL,
    FUSION_PARTY_SITE_NUMBER_COL,
    FUSION_SITE_USE_ID_COL,
    FUSION_SITE_USE_CODE_COL,
    FUSION_PRIMARY_FLAG_COL
]

missing_source = [c for c in required_source_cols if c not in source_df.columns]
missing_fusion = [c for c in required_fusion_cols if c not in fusion_df.columns]

if missing_source:
    raise KeyError(f"Missing source columns: {missing_source}")

if missing_fusion:
    raise KeyError(f"Missing fusion columns: {missing_fusion}")

# =========================================================
# PREPARE DATA
# =========================================================
source_df[SOURCE_CUSTOMER_COL] = source_df[SOURCE_CUSTOMER_COL].fillna("").astype(str).str.strip()

for col in required_fusion_cols:
    fusion_df[col] = fusion_df[col].fillna("").astype(str).str.strip()

fusion_df = fusion_df[fusion_df[FUSION_PARTY_NAME_COL] != ""].copy()

source_df["SOURCE_CUSTOMER_NORMALIZED"] = source_df[SOURCE_CUSTOMER_COL].apply(normalize_customer_name)
fusion_df["FUSION_PARTY_NAME_NORMALIZED"] = fusion_df[FUSION_PARTY_NAME_COL].apply(normalize_customer_name)
fusion_df["SITE_USE_CODE_NORMALIZED"] = fusion_df[FUSION_SITE_USE_CODE_COL].apply(normalize_site_use_code)

# =========================================================
# BUILD CUSTOMER LOOKUPS
# Keep grouped rows because we need site data from same customer
# =========================================================
grouped_by_raw = {}
grouped_by_normalized = {}

for party_name, grp in fusion_df.groupby(FUSION_PARTY_NAME_COL, dropna=False):
    key = clean_text(party_name).upper()
    if key and key not in grouped_by_raw:
        grouped_by_raw[key] = grp.copy()

for norm_name, grp in fusion_df.groupby("FUSION_PARTY_NAME_NORMALIZED", dropna=False):
    key = clean_text(norm_name)
    if key and key not in grouped_by_normalized:
        grouped_by_normalized[key] = grp.copy()

fusion_customer_groups = []
for party_name, grp in fusion_df.groupby(FUSION_PARTY_NAME_COL, dropna=False):
    fusion_customer_groups.append({
        "party_name": clean_text(party_name),
        "party_name_upper": clean_text(party_name).upper(),
        "party_name_norm": normalize_customer_name(party_name),
        "rows": grp.copy()
    })

# =========================================================
# MATCHING + SITE PICKING
# =========================================================
results = []

for _, srow in source_df.iterrows():
    identifier = clean_text(srow[SOURCE_IDENTIFIER_COL])
    src_customer = clean_text(srow[SOURCE_CUSTOMER_COL])
    src_norm = clean_text(srow["SOURCE_CUSTOMER_NORMALIZED"])

    matched_party_name = ""
    matched_party_number = ""
    match_percent = 0
    match_type = "NOT_FOUND"

    bill_to_site_use_id = ""
    bill_to_party_site_id = ""
    bill_to_party_site_number = ""

    ship_to_party_site_id = ""
    ship_to_party_site_number = ""
    ship_to_site_use_id = ""

    matched_customer_rows = None

    # -----------------------------------------------------
    # 1. Exact raw match
    # -----------------------------------------------------
    raw_key = src_customer.upper()
    if raw_key in grouped_by_raw:
        matched_customer_rows = grouped_by_raw[raw_key].copy()
        matched_party_name = clean_text(matched_customer_rows.iloc[0][FUSION_PARTY_NAME_COL])
        match_percent = 100
        match_type = "EXACT"

    # -----------------------------------------------------
    # 2. Exact normalized match
    # -----------------------------------------------------
    elif src_norm in grouped_by_normalized and src_norm != "":
        matched_customer_rows = grouped_by_normalized[src_norm].copy()
        matched_party_name = clean_text(matched_customer_rows.iloc[0][FUSION_PARTY_NAME_COL])
        match_percent = 100
        match_type = "NORMALIZED"

    # -----------------------------------------------------
    # 3. Fuzzy match on customer level
    # -----------------------------------------------------
    else:
        best_score = 0
        best_group = None

        for item in fusion_customer_groups:
            fusion_norm = item["party_name_norm"]
            if not fusion_norm:
                continue

            score = similarity(src_norm, fusion_norm)

            if src_norm and fusion_norm and (src_norm in fusion_norm or fusion_norm in src_norm):
                score = min(100, score + 10)

            if score > best_score:
                best_score = score
                best_group = item

        if best_group is not None and best_score >= FUZZY_THRESHOLD:
            matched_customer_rows = best_group["rows"].copy()
            matched_party_name = best_group["party_name"]
            match_percent = best_score
            match_type = "FUZZY"

    # -----------------------------------------------------
    # If matched, pick site rows only from that matched customer
    # -----------------------------------------------------
    if matched_customer_rows is not None and not matched_customer_rows.empty:
        # party number from same matched customer
        matched_party_number = clean_text(matched_customer_rows.iloc[0][FUSION_PARTY_NUMBER_COL])

        # BILL_TO row
        bill_row = pick_preferred_row(matched_customer_rows, "BILL_TO")
        if bill_row is not None:
            bill_to_site_use_id = clean_text(bill_row[FUSION_SITE_USE_ID_COL])
            bill_to_party_site_id = clean_text(bill_row[FUSION_PARTY_SITE_ID_COL])
            bill_to_party_site_number = clean_text(bill_row[FUSION_PARTY_SITE_NUMBER_COL])

        # SHIP_TO row
        ship_row = pick_preferred_row(matched_customer_rows, "SHIP_TO")
        if ship_row is not None:
            ship_to_party_site_id = clean_text(ship_row[FUSION_PARTY_SITE_ID_COL])
            ship_to_party_site_number = clean_text(ship_row[FUSION_PARTY_SITE_NUMBER_COL])
            ship_to_site_use_id = clean_text(ship_row[FUSION_SITE_USE_ID_COL])

    results.append({
        "Identifier": identifier,
        "Source_Customer": src_customer,
        "Source_Customer_Normalized": src_norm,
        "Matched_PARTY_NAME": matched_party_name,
        "PARTY_NUMBER": matched_party_number,
        "Match_Percent": match_percent,
        "Match_Type": match_type,

        "BILL_TO_SITE_USE_ID": bill_to_site_use_id,
        "BILL_TO_PARTY_SITE_ID": bill_to_party_site_id,
        "BILL_TO_PARTY_SITE_NUMBER": bill_to_party_site_number,

        "SHIP_TO_PARTY_SITE_ID": ship_to_party_site_id,
        "SHIP_TO_PARTY_SITE_NUMBER": ship_to_party_site_number,
        "SHIP_TO_SITE_USE_ID": ship_to_site_use_id
    })

result_df = pd.DataFrame(results)

mapped_df = result_df[result_df["Match_Type"] != "NOT_FOUND"].copy()
not_found_df = result_df[result_df["Match_Type"] == "NOT_FOUND"].copy()

# =========================================================
# SAVE OUTPUT
# =========================================================
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    result_df.to_excel(writer, sheet_name="Source_Mapped", index=False)
    mapped_df.to_excel(writer, sheet_name="Mapped_Only", index=False)
    not_found_df.to_excel(writer, sheet_name="Not_Found", index=False)
    fusion_df.to_excel(writer, sheet_name="Fusion_Normalized", index=False)

print(f"Output saved to: {OUTPUT_FILE}")
print(f"Total Source Records : {len(result_df)}")
print(f"Mapped Records       : {len(mapped_df)}")
print(f"Not Found Records    : {len(not_found_df)}")