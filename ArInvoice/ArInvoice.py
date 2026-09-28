"""
AR Invoice Converter Skeleton
NOTE:
This is a starter implementation matching the requested mapping structure.
Fill column names if they differ in your workbook.
"""
import pandas as pd
from collections import defaultdict

INPUT=r"C:\Users\Shruti Pawar\Workspace\DataConversion\ARInvoiceInputFile.xlsx"
OUTPUT=r"C:\Users\Shruti Pawar\Workspace\MOC2(DEV3)\GlobalAR\Global_ARInvoiceOutputFile.xlsx"

bu=input("Enter Business Unit: ")
rev=input("Enter Revenue Account (9 segments separated by .): ")
rev_seg=rev.split(".")
if len(rev_seg)!=9:
    raise ValueError("Revenue account must contain 9 segments.")

src=pd.read_excel(INPUT,sheet_name="SourceData")
sales=pd.read_excel(INPUT,sheet_name="SalespersonData")
cust=pd.read_excel(INPUT,sheet_name="CustomerData")

sales=sales.fillna("")
cust=cust.fillna("")
src=src.fillna("")

cnt=defaultdict(int)
def unique_ref(v,inv):
    key=v if str(v).strip() else str(inv)
    cnt[key]+=1
    return key if cnt[key]==1 else f"{key}_{cnt[key]-1}"

cust_map={(str(r.get("SourcefileCustomer","")).strip(),str(r.get("SITE_USE_CODE","")).strip()):r for _,r in cust.iterrows()}
sales_map={str(r.get("Invoice Number","")).strip():r for _,r in sales.iterrows()}

lines=[]
dist=[]
credits=[]

for _,r in src.iterrows():
    ref=unique_ref(r.get("CrossRefrence",""),r["Invoice Number"])
    ship=cust_map.get((r["SourceCustomer"],"SHIP_TO"),{})
    bill=cust_map.get((r["SourceCustomer"],"BILL_TO"),{})
    lines.append({
        "Business Unit Name":bu,
        "Transaction Batch Source Name":"Conversion",
        "Transaction Type Name":r["Transaction Type"],
        "Payment Terms":r["Payment Terms"],
        "Transaction Date":r["Invoice Date"],
        "Accounting Date":r["Accounting Date"],
        "Transaction Number":r["Invoice Number"],
        "Bill-to Customer Account Number":ship.get("Account Number",""),
        "Bill-to Customer Site Number":ship.get("Site Number",""),
        "Ship-to Customer Account Number":bill.get("Account Number",""),
        "Ship-to Customer Site Number":bill.get("Site Number",""),
        "Transaction Line Description":r["Header Line Description"],
        "Currency Code":"USD",
        "Currency Conversion Type":"User",
        "Currency Conversion Date":r["Accounting Date"],
        "Currency Conversion Rate":1,
        "Transaction Line Amount":r["Amount"],
        "Transaction Line Quantity":1,
        "Unit Selling Price":r["Amount"],
        "Line Transactions Flexfield Context":"Conversion",
        "Line Transactions Flexfield Segment 1":ref,
        "Line Transactions Flexfield Segment 2":1,
        "Line Transactions Flexfield Segment 3":ref,
        "Primary Salesperson Number":sales_map.get(r["Invoice Number"],{}).get("SalesPerson1Number","")
    })
    s=sales_map.get(r["Invoice Number"],{})
    for i in range(1,8):
        if str(s.get(f"SalesPerson{i}","")).strip():
            pct=s.get(f"Salesperson{i}%","")
            num=s.get(f"SalesPerson{i}Number","")
            for cls in ["REV","REC"]:
                row={"Business Unit Name":bu,"Account Class":cls,"Amount":r["Amount"],"Percent":pct,
                "Line Transactions Flexfield Context":"Conversion",
                "Line Transactions Flexfield Segment 1":ref,
                "Line Transactions Flexfield Segment 2":1,
                "Line Transactions Flexfield Segment 3":ref}
                if cls=="REV":
                    for j,v in enumerate(rev_seg,1):
                        row[f"Accounting Flexfield Segment {j}"]=v
                else:
                    names=["Oracle BU","Oracle Branch","Oracle LOB","Oracle Dept","Oracle Account","Oracle IC Partner","Oracle Add Back","Oracle Prod Line","Oracle Future"]
                    for j,n in enumerate(names,1):
                        row[f"Accounting Flexfield Segment {j}"]=r[n]
                dist.append(row)
            credits.append({
                "Business Unit Name":bu,
                "Salesperson Number":num,
                "Name of Sales Credit Type":"Quota Sales Credit",
                "Sales Credit Percentage Split":pct,
                "Line Transactions Flexfield Context":"Conversion",
                "Line Transactions Flexfield Segment 1":ref,
                "Line Transactions Flexfield Segment 2":1,
                "Line Transactions Flexfield Segment 3":ref
            })

with pd.ExcelWriter(OUTPUT,engine="openpyxl") as w:
    pd.DataFrame(lines).to_excel(w,sheet_name="RA_INTERFACE_LINES_ALL",index=False)
    pd.DataFrame(dist).to_excel(w,sheet_name="RA_INTERFACE_DISTRIBUTIONS_ALL",index=False)
    pd.DataFrame(credits).to_excel(w,sheet_name="RA_INTERFACE_SALESCREDITS_ALL",index=False)

print("Done")
