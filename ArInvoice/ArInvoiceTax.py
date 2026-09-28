import pandas as pd
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from openpyxl import load_workbook


# ==========================================================
# INPUT / OUTPUT FILES
# ==========================================================

INPUT_FILE = r"C:\Users\Shruti Pawar\Workspace\1D_Load\AR\Etairos\SourceFile\ARInvoiceInputFileFor_TAXConversion.xlsx"

OUTPUT_FILE = r"C:\Users\Shruti Pawar\Workspace\1D_Load\AR\Etairos\OUTPUT\ETA_FBDIOutput4.xlsx"


# ==========================================================
# CONSTANTS
# ==========================================================

SOURCE_SHEET = "SourceData"
CUSTOMER_SHEET = "CustomerData"
SALESPERSON_SHEET = "SalesPersonData"
PAYMENT_TERM_SHEET = "PaymentTerm"


# ==========================================================
# HELPER FUNCTIONS
# ==========================================================

def normalize(value):
    """
    Trim spaces and convert to uppercase for comparison.
    """
    if pd.isna(value):
        return ""

    return str(value).strip().upper()


def clean_value(value):
    """
    Return blank for NaN otherwise trimmed string/value.
    """
    if pd.isna(value):
        return ""

    if isinstance(value, str):
        return value.strip()

    return value


def format_date(value):
    """
    Convert date into YYYY/MM/DD.
    Source date is expected to be MM-DD-YYYY,
    but function also handles Excel datetime/date values.
    """

    if pd.isna(value) or str(value).strip() == "":
        return ""

    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%Y/%m/%d")

    value = str(value).strip()

    date_formats = [
        "%m-%d-%Y",
        "%m/%d/%Y",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
    ]

    for fmt in date_formats:
        try:
            return datetime.strptime(value, fmt).strftime("%Y/%m/%d")
        except ValueError:
            continue

    try:
        return pd.to_datetime(value).strftime("%Y/%m/%d")
    except Exception:
        return value


def is_zero_tax(value):
    """
    Returns True when Tax Amount is zero or blank.

    Blank/NaN is treated as no tax because there is
    no tax amount from which to create a TAX record.
    """

    if pd.isna(value):
        return True

    if str(value).strip() == "":
        return True

    try:
        return float(value) == 0
    except Exception:
        return False


def convert_number(value):
    """
    Convert numeric values cleanly.
    """
    if pd.isna(value) or str(value).strip() == "":
        return ""

    try:
        number = float(value)

        if number.is_integer():
            return int(number)

        return number

    except Exception:
        return value


def split_account(account):
    """
    Split 9-segment account into Segment 1 to Segment 9.

    Example:
    1270.0000.000.000.199999.0000.000.0000.0000
    """

    account = clean_value(account)

    if account == "":
        return [""] * 9

    segments = str(account).split(".")

    segments = segments[:9]

    while len(segments) < 9:
        segments.append("")

    return segments


# ==========================================================
# CROSS REFERENCE GENERATION
# ==========================================================

def generate_invoice_cross_references(source_df):
    """
    Cross Reference is invoice-level.

    If one Cross Reference belongs to only one Transaction Number:
        031541

    If same Cross Reference belongs to multiple Transaction Numbers:
        INV001 -> 031541_1
        INV002 -> 031541_2
        INV003 -> 031541_3

    All lines belonging to the same Transaction Number
    receive the same generated Cross Reference.
    """

    transaction_cross_reference = {}

    for _, row in source_df.iterrows():

        trx_number = (
            ""
            if pd.isna(row["Transaction Number"])
            else str(row["Transaction Number"]).strip()
        )

        cross_reference = (
            ""
            if pd.isna(row["Cross Reference"])
            else str(row["Cross Reference"]).strip()
        )

        trx_key = normalize(trx_number)

        if not trx_key:
            continue

        if trx_key not in transaction_cross_reference and cross_reference:
            transaction_cross_reference[trx_key] = cross_reference

    # ------------------------------------------------------
    # Find how many different invoices use each Cross Ref
    # ------------------------------------------------------

    cross_reference_transactions = defaultdict(set)

    for trx_key, cross_reference in transaction_cross_reference.items():

        cross_reference_key = normalize(cross_reference)

        cross_reference_transactions[cross_reference_key].add(trx_key)

    # ------------------------------------------------------
    # Generate invoice-level Cross Reference
    # ------------------------------------------------------

    cross_reference_sequence = defaultdict(int)

    transaction_reference = {}

    for trx_key, cross_reference in transaction_cross_reference.items():

        cross_reference_key = normalize(cross_reference)

        if len(cross_reference_transactions[cross_reference_key]) > 1:

            cross_reference_sequence[cross_reference_key] += 1

            transaction_reference[trx_key] = (
                f"{cross_reference}_"
                f"{cross_reference_sequence[cross_reference_key]}"
            )

        else:

            transaction_reference[trx_key] = cross_reference

    # ------------------------------------------------------
    # Return reference for every source row
    # ------------------------------------------------------

    result = []

    for _, row in source_df.iterrows():

        trx_number = (
            ""
            if pd.isna(row["Transaction Number"])
            else str(row["Transaction Number"]).strip()
        )

        trx_key = normalize(trx_number)

        result.append(
            transaction_reference.get(trx_key, "")
        )

    return result


# ==========================================================
# LINE NUMBER GENERATION
# ==========================================================

def generate_line_numbers(source_df):
    """
    Keep existing Line No.

    If Line No is blank, generate sequential line number
    based on occurrence within Transaction Number.
    """

    occurrence = defaultdict(int)

    result = []

    for _, row in source_df.iterrows():

        trx_number = (
            ""
            if pd.isna(row["Transaction Number"])
            else str(row["Transaction Number"]).strip()
        )

        trx_key = normalize(trx_number)

        line_no = row["Line No"]

        if pd.isna(line_no) or str(line_no).strip() == "":

            occurrence[trx_key] += 1

            result.append(occurrence[trx_key])

        else:

            cleaned = clean_value(line_no)

            try:
                number = float(cleaned)

                if number.is_integer():
                    cleaned = int(number)

            except Exception:
                pass

            result.append(cleaned)

            # Keep generated occurrence in sync
            try:
                numeric_line = int(float(cleaned))

                if numeric_line > occurrence[trx_key]:
                    occurrence[trx_key] = numeric_line

            except Exception:
                pass

    return result


# ==========================================================
# LOAD INPUT FILE
# ==========================================================

print("\nReading input Excel file...")

source_df = pd.read_excel(
    INPUT_FILE,
    sheet_name=SOURCE_SHEET
)

customer_df = pd.read_excel(
    INPUT_FILE,
    sheet_name=CUSTOMER_SHEET
)

salesperson_df = pd.read_excel(
    INPUT_FILE,
    sheet_name=SALESPERSON_SHEET
)

payment_term_df = pd.read_excel(
    INPUT_FILE,
    sheet_name=PAYMENT_TERM_SHEET
)


# ==========================================================
# NORMALIZE COLUMN NAMES
# ==========================================================

source_df.columns = source_df.columns.str.strip()
customer_df.columns = customer_df.columns.str.strip()
salesperson_df.columns = salesperson_df.columns.str.strip()
payment_term_df.columns = payment_term_df.columns.str.strip()


# ==========================================================
# USER INPUT
# ==========================================================

print("\n" + "=" * 60)
print("AR INVOICE FBDI GENERATION")
print("=" * 60)

BU_NAME = input("\nEnter Business Unit Name: ").strip()

accounting_date_input = input(
    "Enter Accounting Date: "
).strip()

ACCOUNTING_DATE = format_date(accounting_date_input)

REV_ACCOUNT = input(
    "Enter REV Account: "
).strip()

REC_ACCOUNT = input(
    "Enter REC Account: "
).strip()

TAX_ACCOUNT = input(
    "Enter TAX Account: "
).strip()


# ==========================================================
# SPLIT ACCOUNTS
# ==========================================================

REV_SEGMENTS = split_account(REV_ACCOUNT)
REC_SEGMENTS = split_account(REC_ACCOUNT)
TAX_SEGMENTS = split_account(TAX_ACCOUNT)


# ==========================================================
# CREATE LOOKUP DICTIONARIES
# ==========================================================

print("\nPreparing lookup data...")


# ----------------------------------------------------------
# Customer Lookup
# ----------------------------------------------------------

bill_to_lookup = {}
ship_to_lookup = {}

for _, row in customer_df.iterrows():

    source_customer = normalize(
        row.get("SourceCustomer", "")
    )

    site_use_code = normalize(
        row.get("Site UseCode", "")
    )

    account_number = clean_value(
        row.get("ACCOUNT_NUMBER", "")
    )

    site_number = clean_value(
        row.get("SITE_NUMBER", "")
    )

    if not source_customer:
        continue

    if site_use_code == "BILL_TO":

        bill_to_lookup[source_customer] = (
            account_number,
            site_number
        )

    elif site_use_code == "SHIP_TO":

        ship_to_lookup[source_customer] = (
            account_number,
            site_number
        )


# ----------------------------------------------------------
# Salesperson Lookup
# ----------------------------------------------------------

salesperson_lookup = {}

for _, row in salesperson_df.iterrows():

    source_salesperson = normalize(
        row.get("SourceSalesPerson", "")
    )

    salesperson_number = clean_value(
        row.get("Salesperson_NO", "")
    )

    if source_salesperson:

        salesperson_lookup[
            source_salesperson
        ] = salesperson_number


# ----------------------------------------------------------
# Payment Term Lookup
# ----------------------------------------------------------

payment_term_lookup = {}

for _, row in payment_term_df.iterrows():

    source_payment_term = normalize(
        row.get("SourcePaymentTerm", "")
    )

    fusion_payment_term = clean_value(
        row.get("FusionPaymentTerm", "")
    )

    if source_payment_term:

        payment_term_lookup[
            source_payment_term
        ] = fusion_payment_term


# ==========================================================
# GENERATE CROSS REFERENCE
# ==========================================================

source_df["_CrossReference"] = (
    generate_invoice_cross_references(source_df)
)


# ==========================================================
# GENERATE LINE NUMBERS
# ==========================================================

source_df["_OutputLineNo"] = (
    generate_line_numbers(source_df)
)


# ==========================================================
# TRANSACTION KEY
# ==========================================================

source_df["_TransactionKey"] = (
    source_df["Transaction Number"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)


# ==========================================================
# OUTPUT DATA LISTS
# ==========================================================

lines_output = []
distribution_output = []
salescredit_output = []
exclude_output = []


# ==========================================================
# OUTPUT COLUMN DEFINITIONS
# ==========================================================

LINES_COLUMNS = [
    "Business Unit Name",
    "Transaction Batch Source Name",
    "Transaction Type Name",
    "Payment Terms",
    "Transaction Date",
    "Accounting Date",
    "Transaction Number",
    "Bill-to Customer Account Number",
    "Bill-to Customer Site Number",
    "Ship-to Customer Account Number",
    "Ship-to Customer Site Number",
    "Transaction Line Type",
    "Transaction Line Description",
    "Currency Code",
    "Currency Conversion Type",
    "Currency Conversion Date",
    "Currency Conversion Rate",
    "Transaction Line Amount",
    "Transaction Line Quantity",
    "Unit Selling Price",
    "Line Transactions Flexfield Context",
    "Invoice Transactions Flexfield Segment 1",
    "Invoice Transactions Flexfield Segment 2",
    "Invoice Transactions Flexfield Segment 3",
    "Primary Salesperson Number",
    "Memo Line Name",
]


DISTRIBUTION_COLUMNS = [
    "Business Unit Name",
    "Account Class",
    "Amount",
    "Percent",
    "Line Transactions Flexfield Context",
    "Invoice Transactions Flexfield Segment 1",
    "Invoice Transactions Flexfield Segment 2",
    "Invoice Transactions Flexfield Segment 3",
    "Accounting Flexfield Segment 1",
    "Accounting Flexfield Segment 2",
    "Accounting Flexfield Segment 3",
    "Accounting Flexfield Segment 4",
    "Accounting Flexfield Segment 5",
    "Accounting Flexfield Segment 6",
    "Accounting Flexfield Segment 7",
    "Accounting Flexfield Segment 8",
    "Accounting Flexfield Segment 9",
]


SALESCREDIT_COLUMNS = [
    "Business Unit Name",
    "Transaction Number",
    "Line Number",
    "Salesperson Number",
    "Sales Credit Amount Split",
    "Sales Credit Percentage Split",
    "Name of Sales Credit Type",
    "Line Transactions Flexfield Context",
    "Invoice Transactions Flexfield Segment 1",
    "Invoice Transactions Flexfield Segment 2",
    "Invoice Transactions Flexfield Segment 3",
]


# ==========================================================
# SALESPERSON COLUMNS
# ==========================================================

salesperson_columns = []

for i in range(1, 7):

    salesperson_columns.append(
        (
            f"SalesPerson{i}",
            f"SalesPerson{i}%"
        )
    )


# ==========================================================
# GROUP SOURCE DATA BY INVOICE
# ==========================================================

grouped_invoices = source_df.groupby(
    "_TransactionKey",
    sort=False
)


# ==========================================================
# PROCESS INVOICE BY INVOICE
# ==========================================================

print("\nProcessing invoices...")


for transaction_key, invoice_df in grouped_invoices:

    # ------------------------------------------------------
    # Track first valid source row.
    #
    # This row is used for invoice-level TAX data.
    # ------------------------------------------------------

    first_valid_row = None

    transaction_number = clean_value(
        invoice_df.iloc[0]["Transaction Number"]
    )

    invoice_cross_reference = clean_value(
        invoice_df.iloc[0]["_CrossReference"]
    )

    # ======================================================
    # FIRST: PROCESS ALL NORMAL LINE DATA
    # ======================================================

    for _, row in invoice_df.iterrows():

        customer_name = clean_value(
            row["Customer Name"]
        )

        customer_key = normalize(
            customer_name
        )

        # --------------------------------------------------
        # Customer lookup
        # --------------------------------------------------

        bill_to = bill_to_lookup.get(
            customer_key
        )

        ship_to = ship_to_lookup.get(
            customer_key
        )

        bill_to_account = (
            bill_to[0]
            if bill_to
            else ""
        )

        bill_to_site = (
            bill_to[1]
            if bill_to
            else ""
        )

        ship_to_account = (
            ship_to[0]
            if ship_to
            else ""
        )

        ship_to_site = (
            ship_to[1]
            if ship_to
            else ""
        )

        # --------------------------------------------------
        # Validate customer
        # --------------------------------------------------

        exclude_reason = []

        if not bill_to_account:

            exclude_reason.append(
                "Bill-to Customer Account Number not found"
            )

        if not bill_to_site:

            exclude_reason.append(
                "Bill-to Customer Site Number not found"
            )

        if not ship_to_account:

            exclude_reason.append(
                "Ship-to Customer Account Number not found"
            )

        if not ship_to_site:

            exclude_reason.append(
                "Ship-to Customer Site Number not found"
            )

        # --------------------------------------------------
        # Exclude invalid records
        # --------------------------------------------------

        if exclude_reason:

            exclude_row = row.drop(
                labels=[
                    "_CrossReference",
                    "_OutputLineNo",
                    "_TransactionKey"
                ],
                errors="ignore"
            ).to_dict()

            exclude_row[
                "Exclude Reason"
            ] = "; ".join(exclude_reason)

            exclude_output.append(
                exclude_row
            )

            continue

        # --------------------------------------------------
        # First valid row for TAX
        # --------------------------------------------------

        if first_valid_row is None:

            first_valid_row = row

        # --------------------------------------------------
        # Common values
        # --------------------------------------------------

        transaction_type = clean_value(
            row["Transaction Type"]
        )

        transaction_date = format_date(
            row["Transaction Date"]
        )

        transaction_line_amount = convert_number(
            row["Transaction Line Amount"]
        )

        line_no = clean_value(
            row["_OutputLineNo"]
        )

        description = clean_value(
            row["Transaction Line Description"]
        )

        if description == "":
            description = transaction_number

        payment_term = payment_term_lookup.get(
            normalize(row["Payment Term"]),
            ""
        )

        # --------------------------------------------------
        # Primary salesperson
        # --------------------------------------------------

        primary_salesperson = salesperson_lookup.get(
            normalize(row["SalesPerson1"]),
            ""
        )

        # ==================================================
        # RA_INTERFACE_LINES_ALL
        # ==================================================

        line_data = {
            "Business Unit Name": BU_NAME,
            "Transaction Batch Source Name": "Conversion",
            "Transaction Type Name": transaction_type,
            "Payment Terms": payment_term,
            "Transaction Date": transaction_date,
            "Accounting Date": ACCOUNTING_DATE,
            "Transaction Number": transaction_number,

            "Bill-to Customer Account Number":
                bill_to_account,

            "Bill-to Customer Site Number":
                bill_to_site,

            "Ship-to Customer Account Number":
                ship_to_account,

            "Ship-to Customer Site Number":
                ship_to_site,

            "Transaction Line Type": "LINE",

            "Transaction Line Description":
                description,

            "Currency Code": "USD",

            "Currency Conversion Type": "User",

            "Currency Conversion Date":
                ACCOUNTING_DATE,

            "Currency Conversion Rate": 1,

            "Transaction Line Amount":
                transaction_line_amount,

            "Transaction Line Quantity": 1,

            "Unit Selling Price":
                transaction_line_amount,

            "Line Transactions Flexfield Context":
                "Conversion",

            "Invoice Transactions Flexfield Segment 1":
                invoice_cross_reference,

            "Invoice Transactions Flexfield Segment 2":
                line_no,

            "Invoice Transactions Flexfield Segment 3":
                invoice_cross_reference,

            "Primary Salesperson Number":
                primary_salesperson,

            "Memo Line Name": "",
        }

        lines_output.append(
            line_data
        )

        # ==================================================
        # RA_INTERFACE_DISTRIBUTIONS_ALL
        #
        # Normal line distributions only.
        # TAX will be appended after ALL invoice lines.
        # ==================================================

        for salesperson_col, percent_col in salesperson_columns:

            salesperson_name = clean_value(
                row[salesperson_col]
            )

            salesperson_percent = clean_value(
                row[percent_col]
            )

            if normalize(salesperson_name) == "":
                continue

            # ----------------------------------------------
            # REV
            # ----------------------------------------------

            rev_distribution = {
                "Business Unit Name": BU_NAME,
                "Account Class": "REV",
                "Amount": transaction_line_amount,
                "Percent": salesperson_percent,

                "Line Transactions Flexfield Context":
                    "Conversion",

                "Invoice Transactions Flexfield Segment 1":
                    invoice_cross_reference,

                "Invoice Transactions Flexfield Segment 2":
                    line_no,

                "Invoice Transactions Flexfield Segment 3":
                    invoice_cross_reference,
            }

            rev_distribution.update({
                f"Accounting Flexfield Segment {i + 1}":
                    REV_SEGMENTS[i]
                for i in range(9)
            })

            distribution_output.append(
                rev_distribution
            )

            # ----------------------------------------------
            # REC
            # ----------------------------------------------

            rec_distribution = {
                "Business Unit Name": BU_NAME,
                "Account Class": "REC",
                "Amount": transaction_line_amount,
                "Percent": salesperson_percent,

                "Line Transactions Flexfield Context":
                    "Conversion",

                "Invoice Transactions Flexfield Segment 1":
                    invoice_cross_reference,

                "Invoice Transactions Flexfield Segment 2":
                    line_no,

                "Invoice Transactions Flexfield Segment 3":
                    invoice_cross_reference,
            }

            rec_distribution.update({
                f"Accounting Flexfield Segment {i + 1}":
                    REC_SEGMENTS[i]
                for i in range(9)
            })

            distribution_output.append(
                rec_distribution
            )

        # ==================================================
        # RA_INTERFACE_SALESCREDITS_ALL
        #
        # Normal line sales credits only.
        # TAX will be appended after ALL invoice lines.
        # ==================================================

        for salesperson_col, percent_col in salesperson_columns:

            salesperson_name = clean_value(
                row[salesperson_col]
            )

            salesperson_percent = clean_value(
                row[percent_col]
            )

            if normalize(salesperson_name) == "":
                continue

            salesperson_number = salesperson_lookup.get(
                normalize(salesperson_name),
                ""
            )

            if salesperson_number == "":
                continue

            sales_credit_data = {
                "Business Unit Name": BU_NAME,

                "Transaction Number":
                    transaction_number,

                "Line Number":
                    line_no,

                "Salesperson Number":
                    salesperson_number,

                # Required to remain blank
                "Sales Credit Amount Split": "",

                "Sales Credit Percentage Split":
                    salesperson_percent,

                "Name of Sales Credit Type":
                    "Quota Sales Credit",

                "Line Transactions Flexfield Context":
                    "Conversion",

                "Invoice Transactions Flexfield Segment 1":
                    invoice_cross_reference,

                "Invoice Transactions Flexfield Segment 2":
                    line_no,

                "Invoice Transactions Flexfield Segment 3":
                    invoice_cross_reference,
            }

            salescredit_output.append(
                sales_credit_data
            )

    # ======================================================
    # SECOND: PROCESS TAX DATA
    #
    # IMPORTANT:
    # TAX IS CREATED ONLY IF TAX AMOUNT != 0.
    #
    # TAX IS APPENDED AFTER ALL NORMAL LINE DATA
    # FOR THIS INVOICE.
    # ======================================================

    if first_valid_row is None:
        continue

    tax_amount = first_valid_row["Tax Amount"]

    # ------------------------------------------------------
    # NO TAX
    # ------------------------------------------------------

    if is_zero_tax(tax_amount):

        continue

    # ------------------------------------------------------
    # TAX EXISTS
    # ------------------------------------------------------

    tax_amount = convert_number(
        tax_amount
    )

    tax_transaction_date = format_date(
        first_valid_row["Transaction Date"]
    )

    tax_payment_term = payment_term_lookup.get(
        normalize(first_valid_row["Payment Term"]),
        ""
    )

    customer_name = clean_value(
        first_valid_row["Customer Name"]
    )

    customer_key = normalize(
        customer_name
    )

    bill_to = bill_to_lookup.get(
        customer_key
    )

    ship_to = ship_to_lookup.get(
        customer_key
    )

    bill_to_account = (
        bill_to[0]
        if bill_to
        else ""
    )

    bill_to_site = (
        bill_to[1]
        if bill_to
        else ""
    )

    ship_to_account = (
        ship_to[0]
        if ship_to
        else ""
    )

    ship_to_site = (
        ship_to[1]
        if ship_to
        else ""
    )

    primary_salesperson = salesperson_lookup.get(
        normalize(first_valid_row["SalesPerson1"]),
        ""
    )

    # ------------------------------------------------------
    # TAX LINE NUMBER
    #
    # Keeping the current logic for Segment 2.
    # This can be changed later.
    # ------------------------------------------------------

    tax_line_no = 1

    # ======================================================
    # RA_INTERFACE_LINES_ALL - TAX
    # ======================================================

    tax_line = {
        "Business Unit Name": BU_NAME,

        "Transaction Batch Source Name":
            "Conversion",

        "Transaction Type Name":
            clean_value(first_valid_row["Transaction Type"]),

        "Payment Terms":
            tax_payment_term,

        "Transaction Date":
            tax_transaction_date,

        "Accounting Date":
            ACCOUNTING_DATE,

        "Transaction Number":
            transaction_number,

        "Bill-to Customer Account Number":
            bill_to_account,

        "Bill-to Customer Site Number":
            bill_to_site,

        "Ship-to Customer Account Number":
            ship_to_account,

        "Ship-to Customer Site Number":
            ship_to_site,

        "Transaction Line Type":
            "LINE",

        "Transaction Line Description":
            "Conversion Tax",

        "Currency Code":
            "USD",

        "Currency Conversion Type":
            "User",

        "Currency Conversion Date":
            ACCOUNTING_DATE,

        "Currency Conversion Rate":
            1,

        "Transaction Line Amount":
            tax_amount,

        "Transaction Line Quantity":
            1,

        "Unit Selling Price":
            tax_amount,

        "Line Transactions Flexfield Context":
            "Conversion",

        "Invoice Transactions Flexfield Segment 1":
            invoice_cross_reference,

        # KEEPING CURRENT LOGIC FOR NOW
        "Invoice Transactions Flexfield Segment 2":
            tax_line_no,

        "Invoice Transactions Flexfield Segment 3":
            invoice_cross_reference,

        "Primary Salesperson Number":
            primary_salesperson,

        "Memo Line Name":
            "Conversion Tax",
    }

    lines_output.append(
        tax_line
    )

    # ======================================================
    # RA_INTERFACE_DISTRIBUTIONS_ALL - TAX
    # ======================================================

    tax_distribution = {
        "Business Unit Name":
            BU_NAME,

        "Account Class":
            "TAX",

        "Amount":
            tax_amount,

        "Percent":
            100,

        "Line Transactions Flexfield Context":
            "Conversion",

        "Invoice Transactions Flexfield Segment 1":
            invoice_cross_reference,

        # KEEPING CURRENT LOGIC FOR NOW
        "Invoice Transactions Flexfield Segment 2":
            tax_line_no,

        "Invoice Transactions Flexfield Segment 3":
            invoice_cross_reference,
    }

    tax_distribution.update({
        f"Accounting Flexfield Segment {i + 1}":
            TAX_SEGMENTS[i]
        for i in range(9)
    })

    distribution_output.append(
        tax_distribution
    )

    # ======================================================
    # RA_INTERFACE_SALESCREDITS_ALL - TAX
    # ======================================================

    tax_sales_credit = {
        "Business Unit Name":
            BU_NAME,

        "Transaction Number":
            transaction_number,

        "Line Number":
            tax_line_no,

        "Salesperson Number":
            primary_salesperson,

        # Required to remain blank
        "Sales Credit Amount Split":
            "",

        "Sales Credit Percentage Split":
            100,

        "Name of Sales Credit Type":
            "Quota Sales Credit",

        "Line Transactions Flexfield Context":
            "Conversion",

        "Invoice Transactions Flexfield Segment 1":
            invoice_cross_reference,

        # KEEPING CURRENT LOGIC FOR NOW
        "Invoice Transactions Flexfield Segment 2":
            tax_line_no,

        "Invoice Transactions Flexfield Segment 3":
            invoice_cross_reference,
    }

    salescredit_output.append(
        tax_sales_credit
    )


# ==========================================================
# CREATE DATAFRAMES
# ==========================================================

lines_df = pd.DataFrame(
    lines_output,
    columns=LINES_COLUMNS
)

distribution_df = pd.DataFrame(
    distribution_output,
    columns=DISTRIBUTION_COLUMNS
)

salescredit_df = pd.DataFrame(
    salescredit_output,
    columns=SALESCREDIT_COLUMNS
)


# ==========================================================
# EXCLUDE DATAFRAME
# ==========================================================

if exclude_output:

    exclude_df = pd.DataFrame(
        exclude_output
    )

else:

    exclude_df = pd.DataFrame(
        columns=list(source_df.columns)
        + ["Exclude Reason"]
    )

# Remove internal processing columns if present
exclude_df = exclude_df.drop(
    columns=[
        "_CrossReference",
        "_OutputLineNo",
        "_TransactionKey"
    ],
    errors="ignore"
)


# ==========================================================
# CREATE OUTPUT DIRECTORY
# ==========================================================

output_path = Path(
    OUTPUT_FILE
)

output_path.parent.mkdir(
    parents=True,
    exist_ok=True
)


# ==========================================================
# WRITE EXCEL FILE
# ==========================================================

print("\nWriting output Excel file...")

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    lines_df.to_excel(
        writer,
        sheet_name="RA_INTERFACE_LINES_ALL",
        index=False
    )

    distribution_df.to_excel(
        writer,
        sheet_name="RA_INTERFACE_DISTRIBUTIONS_ALL",
        index=False
    )

    salescredit_df.to_excel(
        writer,
        sheet_name="RA_INTERFACE_SALESCREDITS_ALL",
        index=False
    )

    exclude_df.to_excel(
        writer,
        sheet_name="Exclude",
        index=False
    )


# ==========================================================
# AUTO SIZE EXCEL COLUMNS
# ==========================================================

print("Formatting Excel file...")

workbook = load_workbook(
    OUTPUT_FILE
)

for worksheet in workbook.worksheets:

    for column_cells in worksheet.columns:

        max_length = 0

        column_letter = column_cells[0].column_letter

        for cell in column_cells:

            try:

                cell_length = len(
                    str(cell.value)
                )

                if cell_length > max_length:
                    max_length = cell_length

            except Exception:
                pass

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            50
        )


workbook.save(
    OUTPUT_FILE
)


# ==========================================================
# SUMMARY
# ==========================================================

print("\n" + "=" * 60)
print("AR FBDI GENERATION COMPLETED")
print("=" * 60)

print(
    f"\nOutput File:\n{OUTPUT_FILE}"
)

print(
    f"\nRA_INTERFACE_LINES_ALL Records: "
    f"{len(lines_df)}"
)

print(
    f"RA_INTERFACE_DISTRIBUTIONS_ALL Records: "
    f"{len(distribution_df)}"
)

print(
    f"RA_INTERFACE_SALESCREDITS_ALL Records: "
    f"{len(salescredit_df)}"
)

print(
    f"Exclude Records: "
    f"{len(exclude_df)}"
)

print("\nDone.")