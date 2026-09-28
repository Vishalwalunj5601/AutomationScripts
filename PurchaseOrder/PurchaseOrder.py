# Updated Oracle Fusion PO FBDI Generator
import pandas as pd
from pathlib import Path

input_file = r"C:\Users\Shruti Pawar\Workspace\DataConversion\PO_InputFile.xlsx"
output_folder = Path(r"C:\Users\Shruti Pawar\Workspace\1D_Load\PO\Output")
output_folder.mkdir(parents=True, exist_ok=True)

HeaderKey=int(input("Enter Header key Number start from: "))
BuName=input("Enter Business Unit: ")
BatchId=int(input("Enter Batch ID: "))

hdr=pd.read_excel(input_file,sheet_name="Header",dtype={"PO Number":str})
lines=pd.read_excel(input_file,sheet_name="Lines",dtype={"PO Number":str})
sup=pd.read_excel(input_file,sheet_name="Supplier")
buyer=pd.read_excel(input_file,sheet_name="Buyer")
item=pd.read_excel(input_file,sheet_name="Item")

def clean_po(v):
    if pd.isna(v):
        return ""
    s=str(v).strip()
    if s.endswith(".0"):
        s=s[:-2]
    return s

hdr["PO Number"]=hdr["PO Number"].apply(clean_po)
lines["PO Number"]=lines["PO Number"].apply(clean_po)

sup_map=sup.set_index("Supplier")
buyer_map=buyer.set_index("Buyer Name")
item_map=item.set_index("Item")

headers=[];header_keys={};excluded=[];neg=[]
for i,(po,row) in enumerate(hdr.drop_duplicates("PO Number").set_index("PO Number").iterrows(),1):
    hk=f"{BuName}_{i}"
    header_keys[po]=hk
    b=buyer_map.loc[row["Buyer Name"],"FusionBuyer"] if row["Buyer Name"] in buyer_map.index else f'{row["Buyer Name"]} Name Not Found'
    if row["Supplier"] in sup_map.index:
        s=sup_map.loc[row["Supplier"]]; fs,fsn,fss=s["Fusion Supplier"],s["Supplier Number"],s["Supplier Site"]
    else:
        fs=fsn=fss=f'{row["Supplier"]} Not Found'
    headers.append({"Interface Header Key":hk,"Action":"ORIGINAL","Batch ID":BatchId,"Approval Action":"BYPASS","Order":po,"Document Type Code":"STANDARD","Style":"Purchase Order","Procurement BU":BuName,"Requisitioning BU":BuName,"Bill-to BU":BuName,"Buyer":b,"Currency Code":"USD","Bill-to Location":row["ShipToLocation"],"Ship-to Location":row["ShipToLocation"],"Supplier":fs,"Supplier Number":fsn,"Supplier Site":fss,"Payment Terms":row["Payment Term"],"Initiating Party":"BUYER","Required Acknowledgment":"N","ATTRIBUTE4":"Parts Purchases"})
po_lines=[];locs=[];dists=[];count={}
for _,r in lines.iterrows():
    po=clean_po(r["PO Number"])
    if po not in header_keys:
        rr=r.copy();rr["Note"]="Excluded PO";excluded.append(rr);continue
    try:
        qty_val = float(r["Open Qty"])
    except:
        qty_val = 0.0

    if qty_val <= 0:
        rr=r.copy();rr["Note"]="Exclude this record due to negative Qty.";neg.append(rr);continue
    count[po]=count.get(po,0)+1
    occ=count[po]
    lk=f"{po}_L{occ}"; ll=f"{po}_LL{occ}"; dk=f"{po}_D{occ}"
    it=item_map.loc[r["Item"],"Fusion Item"] if r["Item"] in item_map.index else f'{r["Item"]} is not found'
    po_lines.append({"Interface Line Key":lk,"Interface Header Key":header_keys[po],"Action":"ADD","Line":occ,"Line Type":"Goods","Item":it,"Quantity":r["Open Qty"],"UOM":"Each","Price":r["Cost"]})
    h=hdr[hdr["PO Number"]==po].iloc[0]
    need_by = pd.to_datetime(h.get("Need By Date")).strftime("%Y/%m/%d") if pd.notna(h.get("Need By Date")) else ""
    promised = pd.to_datetime(h.get("Promised_Date")).strftime("%Y/%m/%d") if pd.notna(h.get("Promised_Date")) else ""
    locs.append({"Interface Line Location Key":ll,"Interface Line Key":lk,"Schedule":occ,"Ship-to Location":h["ShipToLocation"],"Quantity":r["Open Qty"],"Need-by Date":need_by,"Promised Date":promised,"Destination Type Code":"INVENTORY"})
    dists.append({"Interface Distribution Key":dk,"Interface Line Location Key":ll,"Distribution":occ,"Deliver-to Location":h["ShipToLocation"],"Quantity":r["Open Qty"]})
out=output_folder/"PO_FBDI_Output2.xlsx"
with pd.ExcelWriter(out,engine="openpyxl") as w:
    pd.DataFrame(headers).to_excel(
        excel_writer=w,
        sheet_name="PO_HEADERS_INTERFACE",
        index=False
    )

    pd.DataFrame(po_lines).to_excel(
        excel_writer=w,
        sheet_name="PO_LINES_INTERFACE",
        index=False
    )

    pd.DataFrame(locs).to_excel(
        excel_writer=w,
        sheet_name="PO_LINE_LOCATIONS_INTERFACE",
        index=False
    )

    pd.DataFrame(dists).to_excel(
        excel_writer=w,
        sheet_name="PO_DISTRIBUTIONS_INTERFACE",
        index=False
    )

    pd.DataFrame(excluded).to_excel(
        excel_writer=w,
        sheet_name="Excluded PO",
        index=False
    )

    pd.DataFrame(neg).to_excel(
        excel_writer=w,
        sheet_name="Negative Quantity",
        index=False
    )
print(out)