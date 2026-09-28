import pandas as pd
from openpyxl import Workbook
from datetime import datetime

# ==========================================================
# INPUTS
# ==========================================================

SOURCE_FILE = r"C:\Users\Shruti Pawar\Workspace\DataConversion\APInvoiceInputFile.xlsx"
OUTPUT_FILE = r"C:\Users\Shruti Pawar\Workspace\1D_Load\AP\Etairos\OutPut\Etairos_ApInvoiceOutput1.xlsx"

BUSINESS_UNIT = input("Enter Business Unit: ").strip()
START_INVOICE_ID = int(input("Enter starting Invoice ID: "))

# ==========================================================
# READ SOURCE FILE
# ==========================================================

source_df = pd.read_excel(
    SOURCE_FILE,
    sheet_name="SourceData",
    dtype=str
)

supplier_df = pd.read_excel(
    SOURCE_FILE,
    sheet_name="SupplierData",
    dtype=str
)

source_df = source_df.fillna("")
supplier_df = supplier_df.fillna("")

# ==========================================================
# DATE FORMATTER
# ==========================================================

def format_date(value):

    if value in ("", None):
        return ""

    try:
        return pd.to_datetime(value).strftime("%Y/%m/%d")
    except:
        return ""

# ==========================================================
# SUPPLIER LOOKUP
# ==========================================================

# ==========================================================
# HEADER DATA
# ==========================================================

header_rows = []

invoice_id_map = {}

current_invoice_id = START_INVOICE_ID

unique_invoices = source_df.drop_duplicates(subset=["Invoice Number"])

for _, row in unique_invoices.iterrows():

    # match supplier from supplier data sheet
    supplier_matches = supplier_df[supplier_df["SourceSupplier"] == row["Supplier Name"]]
    supplier = supplier_matches.iloc[0].to_dict() if not supplier_matches.empty else {}

    invoice_id_map[row["Invoice Number"]] = current_invoice_id

    header_rows.append({

        "Invoice ID": current_invoice_id,

        "Business Unit": BUSINESS_UNIT,

        "Source": "Conversion",

        "Invoice Number": row["Invoice Number"],

        "Invoice Amount": row["Invoice Amount"],

        "Invoice Date": format_date(row["Invoice Date"]),

        "Supplier Name": supplier.get("Supplier Name", ""),

        "Supplier Number": supplier.get("SupplierNumber", ""),

        "Supplier Site": supplier.get("SupplierSite", ""),

        "Invoice Currency": "USD",

        "Payment Currency": "USD",

        "Description": row["Description"],

        "Import Set": BUSINESS_UNIT + "_Invoice_Load",

        "Invoice Type": row["Invoice Type"],

        "Payment Terms": row["Payment Terms"],

        "Terms Date": format_date(row["Terms Date"]),

        "Accounting Date": "2026/07/31"

    })

    current_invoice_id += 1

header_df = pd.DataFrame(header_rows)

# ==========================================================
# LINE DATA
# ==========================================================

line_rows = []

invoice_counter = {}

for _, row in source_df.iterrows():

    invoice_number = row["Invoice Number"]

    invoice_id = invoice_id_map[invoice_number]

    line_number = str(row["Line Number"]).strip()

    if line_number == "" or line_number.lower() == "nan":

        invoice_counter.setdefault(invoice_number, 0)

        invoice_counter[invoice_number] += 1

        line_number = invoice_counter[invoice_number]

    line_rows.append({

        "Invoice ID": invoice_id,

        "Line Number": line_number,

        "Line Type": "ITEM",

        "Amount": row["Invoice Amount"],

        "Description": BUSINESS_UNIT + "_Line_1",
        "DescriptionWithPO": row["PONumber"],

        "Distribution Combination": row["Distribution Combination"],

        "Accounting Date": "2026/07/31",

        "Ship-to Location": row["Ship-to Location"]

    })

line_df = pd.DataFrame(line_rows)

# ==========================================================
# WRITE OUTPUT
# ==========================================================

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

    header_df.to_excel(
        writer,
        sheet_name="AP_INVOICES_INTERFACE",
        index=False
    )

    line_df.to_excel(
        writer,
        sheet_name="AP_INVOICE_LINES_INTERFACE",
        index=False
    )

print("\n=======================================")
print("AP Invoice Interface Created Successfully")
print("=======================================")
print(f"Headers : {len(header_df)}")
print(f"Lines   : {len(line_df)}")
print(f"Output  : {OUTPUT_FILE}")