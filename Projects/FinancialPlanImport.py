import os
import pandas as pd
import warnings
from datetime import datetime
from dateutil.relativedelta import relativedelta

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ================= CONFIG =================
BASE_DIR = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\03 Project Financial Workbooks for active projects"
OUTPUT_FILE = os.path.join(BASE_DIR, "FINANCIAL_PLAN-V5.xlsx")

ACTIVE_FILE = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\CJBS Active Project List_06_25_26.xlsx"

# NEW: static resource mapping file (with Resource_Mapping sheet)
RESOURCE_MAP_FILE = r"C:\AIR_CONTROL_CONCEPTS\DEV3-PROD_1C\Projects\OneDrive_2026-06-26\Pre\FinancialPlanMapping.xlsx"

ESTIMATE_KEYWORDS = [
    "booking estimate",
    "co1 estimate",
    "co2 estimate",
    "co3",
    "co4",
    "co5",
    "change order"
]

TURNOVER_SHEET_KEYWORD = "turnover information"
PROJECT_SCHEDULE_SHEET_KEYWORD = "project schedule"

# ================= HELPERS =================
def normalize(val):
    return str(val).strip().lower()

def to_number(val):
    if pd.isna(val):
        return None
    try:
        val = str(val).replace(",", "").replace("$", "").strip()
        num = float(val)
        return num if num != 0 else None
    except:
        return None

def format_date_yyyymmdd(value):
    """Return YYYY/MM/DD string or None."""
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y/%m/%d")
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%Y/%m/%d")

# =========== STATIC RESOURCE/TASK MAPPING (from Excel) ===========

# mapping_cache will be: (source_task_norm, source_resource_norm) -> {"Final Task": ..., "FusionResource": ...}
mapping_cache = {}

if os.path.exists(RESOURCE_MAP_FILE):
    try:
        rm_xls = pd.ExcelFile(RESOURCE_MAP_FILE)
        # try to find sheet with name like "Resource_Mapping"
        sheet_name = None
        for s in rm_xls.sheet_names:
            if "resource_mapping" in normalize(s):
                sheet_name = s
                break
        if sheet_name is None:
            # fallback – assume sheet is exactly "Resource_Mapping" or first sheet
            if "Resource_Mapping" in rm_xls.sheet_names:
                sheet_name = "Resource_Mapping"
            else:
                sheet_name = rm_xls.sheet_names[0]

        rm_df = pd.read_excel(rm_xls, sheet_name=sheet_name)
        rm_df.columns = [c.strip() for c in rm_df.columns]

        required_cols = {"SourceTask", "Source Resource", "Final Task", "FusionResource"}
        missing = required_cols - set(rm_df.columns)
        if missing:
            print(f"❌ Static mapping file is missing columns: {missing}")
        else:
            for _, r in rm_df.iterrows():
                s_task = normalize(r["SourceTask"])
                s_res = normalize(r["Source Resource"])
                final_task = str(r["Final Task"]).strip()
                fusion_res = str(r["FusionResource"]).strip()
                mapping_cache[(s_task, s_res)] = {
                    "Final Task": final_task,
                    "FusionResource": fusion_res
                }
            print(f"✅ Loaded static mapping rows: {len(mapping_cache)}")

    except Exception as e:
        print(f"❌ Error loading static mapping file: {e}")
else:
    print(f"❌ Resource mapping file not found: {RESOURCE_MAP_FILE}")

def resolve_mapping(base_task, resource_name):
    """
    Static mapping:
      - Use (SourceTask, Source Resource) from Resource_Mapping sheet.
      - If not found, fall back to (Final Task = base_task, FusionResource = base_task).
    """
    key = (normalize(base_task), normalize(resource_name))
    if key in mapping_cache:
        m = mapping_cache[key]
        return m["Final Task"], m["FusionResource"]

    # Optional: try task-only mapping if you ever put blank Source Resource in mapping
    key_task_only = (normalize(base_task), "")
    if key_task_only in mapping_cache:
        m = mapping_cache[key_task_only]
        return m["Final Task"], m["FusionResource"]

    # If still not found, use simple fallback and log
    print(f"⚠️ No static mapping for Task='{base_task}' Resource='{resource_name}'. "
          f"Using Final Task = '{base_task}', FusionResource = '{base_task}'.")
    return base_task, base_task

# ================= LOAD ACTIVE PROJECTS (Warranty End Date) =================
active_map = {}
today_dt = datetime.today()
today_date = today_dt.date()

if os.path.exists(ACTIVE_FILE):
    a_df = pd.read_excel(ACTIVE_FILE)
    a_df.columns = [c.strip() for c in a_df.columns]
    if "Project Number" in a_df.columns:
        a_df["Project Number Key"] = a_df["Project Number"].astype(str).str.strip()
        a_df = a_df.drop_duplicates("Project Number Key", keep="last").set_index("Project Number Key")
        active_map = a_df.to_dict(orient="index")

def get_active_row(project_number):
    return active_map.get(str(project_number).strip())

# ================= PROJECT CORE DATA (Start/Finish) =================
def get_project_core_dates(xls):
    project_number = None
    start_date = None
    duration_months = None

    for sheet in xls.sheet_names:
        if TURNOVER_SHEET_KEYWORD in normalize(sheet):
            df = pd.read_excel(xls, sheet_name=sheet, header=None)

            for _, row in df.iterrows():
                label = normalize(row.iloc[0])

                if label in {"project number:", "project number"}:
                    if len(row) > 1 and pd.notna(row.iloc[1]):
                        project_number = str(row.iloc[1]).strip()

                if label in {"project start date", "project start date:"}:
                    if len(row) > 1 and pd.notna(row.iloc[1]):
                        start_date = format_date_yyyymmdd(row.iloc[1])

                if "project duration" in label and "month" in label:
                    values = row.iloc[1:].dropna().tolist()
                    if values:
                        try:
                            duration_months = int(float(str(values[0]).strip()))
                        except:
                            duration_months = None
            break

    finish_date = None
    if start_date:
        sdt = pd.to_datetime(start_date, errors="coerce")
        if not pd.isna(sdt):
            if duration_months is not None:
                planned_end = sdt + relativedelta(months=duration_months)

                # NEW LOGIC:
                # Planned End Date <= 2026/06/30, then Project Finish Date = 2026/09/30
                if planned_end <= pd.Timestamp("2026-06-30"):
                    finish_date = "2026/09/30"
                else:
                    finish_date = planned_end.strftime("%Y/%m/%d")
            else:
                if sdt.year < today_dt.year:
                    finish_date = (today_dt + relativedelta(months=3)).strftime("%Y/%m/%d")
                else:
                    finish_date = sdt.strftime("%Y/%m/%d")

    return project_number, start_date, finish_date

def get_task_dates(task_name, project_start, project_finish, project_number):
    planned_start = project_start
    planned_end = project_finish

    is_milestone_for_dates = task_name in {"As Built", "Software Engineering", "P&R", "Warranty"}

    # ===== WARRANTY SPECIAL LOGIC =====
    if task_name == "Warranty":
        active_row = get_active_row(project_number)
        warranty_end = None

        if active_row is not None:
            warranty_end = active_row.get("Warranty End Date")

        warr_date = pd.to_datetime(warranty_end, errors="coerce")

        # If warranty end exists → use it
        if not pd.isna(warr_date):
            planned_end = warr_date.strftime("%Y/%m/%d")

        # Start date should ALWAYS be Project Start
        planned_start = project_start

        return planned_start, planned_end
    # ===================================

    # Existing milestone logic for other tasks (UNCHANGED)
    if is_milestone_for_dates and planned_end is not None:
        planned_start = planned_end

    return planned_start, planned_end

# ============== PROJECT SCHEDULE TOTALS (FOR VALIDATION) ==============
def get_project_schedule_totals(xls):
    totals = {}

    for sheet in xls.sheet_names:
        if PROJECT_SCHEDULE_SHEET_KEYWORD not in normalize(sheet):
            continue

        df = pd.read_excel(xls, sheet_name=sheet, header=None)

        for _, row in df.iterrows():
            if len(row) < 2:
                continue

            if pd.isna(row.iloc[0]) or pd.isna(row.iloc[1]):
                continue

            desc = normalize(row.iloc[0])
            cost = to_number(row.iloc[1])

            if cost is None:
                continue

            if "task description" in desc or "booked" in desc:
                continue

            totals[desc] = totals.get(desc, 0.0) + cost

        break

    return totals

def schedule_cost_for_task(project_number, task_name, schedule_by_project):
    """
    Map Task to the appropriate Project Schedule cost.

    For Material: we want Material + Material Tax + Material Shipping.
    For Subcontracting: Subcontracting + Subcontractor + Graphics.
    """
    sched = schedule_by_project.get(str(project_number).strip())
    if not sched:
        return None

    tnorm = normalize(task_name)

    # Subcontracting = subcontracting + subcontractor + graphics
    if tnorm in {"subcontracting", "subcontractor"}:
        keys = ["subcontracting", "subcontractor", "graphics"]
        total = 0.0
        found = False
        for k in keys:
            if k in sched:
                total += sched[k]
                found = True
        return total if found else None

    # Material = material + material tax + material shipping
    if tnorm == "material":
        total = 0.0
        found = False
        for k in ["material", "material tax", "material shipping"]:
            if k in sched:
                total += sched[k]
                found = True
        return total if found else None

    key_map = {
        "labor": "labor",
        "mileage": "mileage",
        "p&r": "p&r",
        "warranty": "warranty",
        "admin and oh": "admin and oh",
        "misc costs": "misc costs"
    }

    key = key_map.get(tnorm)
    if key and key in sched:
        return sched[key]

    # fall back to direct name if exists
    return sched.get(tnorm)

# ================= MAIN =================
final_rows = []
project_id_map = {}
project_counter = 1
schedule_by_project = {}

print("🔍 Starting extraction...\n")

for file in os.listdir(BASE_DIR):
    if not file.lower().endswith((".xlsx", ".xlsm")):
        continue
    if "consolidated_cost_output" in file.lower():
        continue

    print(f"📂 Processing file: {file}")
    path = os.path.join(BASE_DIR, file)

    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        print(f"❌ Cannot open file: {e}")
        continue

    project_number, proj_start, proj_finish = get_project_core_dates(xls)
    if not project_number:
        print("⚠️  Project number not found, skipping.")
        continue

    if project_number not in project_id_map:
        project_id_map[project_number] = project_counter
        project_counter += 1
    project_id = project_id_map[project_number]

    schedule_totals = get_project_schedule_totals(xls)
    schedule_by_project[str(project_number).strip()] = schedule_totals

    print(f"   🧾 Project Number: {project_number} | Project ID: {project_id}")
    print(f"   📅 Project Start: {proj_start} | Project Finish: {proj_finish}")

    # ========== ESTIMATE SHEETS ==========
    for sheet in xls.sheet_names:
        sheet_norm = normalize(sheet)
        if not any(k in sheet_norm for k in ESTIMATE_KEYWORDS):
            continue

        print(f"   📄 Using estimate sheet: {sheet}")
        df = pd.read_excel(xls, sheet_name=sheet, header=None)

        # ---------- MATERIAL ----------
        print("   ➤ Extracting MATERIAL...")
        inside = False
        for _, row in df.iterrows():
            row_text = " ".join(normalize(x) for x in row.values if pd.notna(x))

            if "control product cost" in row_text:
                inside = True
                continue

            if inside and "total material" in row_text:
                break

            if not inside:
                continue

            res = row.iloc[0]
            cost = to_number(row.iloc[1]) if len(row) > 1 else None

            if pd.isna(res) or normalize(res).startswith("total"):
                continue
            if cost is None:
                continue

            base_task = "Material"
            task_name, fusion_res = resolve_mapping(base_task, res)
            start_date, end_date = get_task_dates(task_name, proj_start, proj_finish, project_number)

            final_rows.append({
                "Project ID": project_id,
                "Project Number": project_number,
                "File Name": file,
                "Task": task_name,
                "Resource": res,
                "FusionResource": fusion_res,
                "Start Date": start_date,
                "End Date": end_date,
                "Cost": cost
            })

        # ---------- LABOR ----------
        print("   ➤ Extracting LABOR...")
        inside = False
        cost_col = None

        for _, row in df.iterrows():
            row_text = " ".join(normalize(x) for x in row.values if pd.notna(x))

            if "labor cost" in row_text:
                inside = True
                cost_col = None
                continue

            if not inside:
                continue

            if cost_col is None and "total cost" in row_text:
                for i, cell in enumerate(row):
                    if isinstance(cell, str) and "total cost" in normalize(cell):
                        cost_col = i
                        print(f"     🔹 Labor Total Cost column index: {cost_col}")
                        break
                continue

            if "total labor" in row_text:
                break

            if cost_col is None:
                continue

            res = row.iloc[0]
            if pd.isna(res) or normalize(res).startswith("total"):
                continue

            cost = to_number(row.iloc[cost_col]) if len(row) > cost_col else None
            if cost is None:
                continue

            base_task = "Labor"
            task_name, fusion_res = resolve_mapping(base_task, res)
            start_date, end_date = get_task_dates(task_name, proj_start, proj_finish, project_number)

            final_rows.append({
                "Project ID": project_id,
                "Project Number": project_number,
                "File Name": file,
                "Task": task_name,
                "Resource": res,
                "FusionResource": fusion_res,
                "Start Date": start_date,
                "End Date": end_date,
                "Cost": cost
            })

        # ---------- SUBCONTRACTOR ----------
        print("   ➤ Extracting SUBCONTRACTOR...")
        inside = False

        for _, row in df.iterrows():
            row_text = " ".join(normalize(x) for x in row.values if pd.notna(x))

            if "subcontractor cost" in row_text:
                inside = True
                continue

            if inside and "total subcontractor" in row_text:
                break

            if not inside:
                continue

            res = row.iloc[0]
            cost = to_number(row.iloc[1]) if len(row) > 1 else None

            if pd.isna(res) or normalize(res).startswith("total"):
                continue
            if cost is None:
                continue

            base_task = "Subcontracting"
            task_name, fusion_res = resolve_mapping(base_task, res)
            start_date, end_date = get_task_dates(task_name, proj_start, proj_finish, project_number)

            final_rows.append({
                "Project ID": project_id,
                "Project Number": project_number,
                "File Name": file,
                "Task": task_name,
                "Resource": res,
                "FusionResource": fusion_res,
                "Start Date": start_date,
                "End Date": end_date,
                "Cost": cost
            })

        # ---------- OVERHEAD & ADMIN ----------
        print("   ➤ Extracting OVERHEAD & ADMIN...")
        inside = False

        for _, row in df.iterrows():
            row_text = " ".join(normalize(x) for x in row.values if pd.notna(x))

            if ("description" in row_text and "adjustment" in row_text and "estimate cost" in row_text):
                inside = True
                continue

            if not inside:
                continue

            desc = row.iloc[0] if len(row) > 0 else None
            adj = row.iloc[1] if len(row) > 1 else None
            cost = to_number(row.iloc[3]) if len(row) > 3 else None

            if pd.isna(desc):
                continue

            desc_text = str(desc).strip()

            if normalize(desc_text).startswith("total estimated cost"):
                break

            if desc_text in ["Estimate Cost", "Estimate Margin %", "Target Sales Price"]:
                continue

            if pd.isna(adj) or cost is None:
                continue

            base_task = "Admin and OH"
            task_name, fusion_res = resolve_mapping(base_task, desc_text)
            start_date, end_date = get_task_dates(task_name, proj_start, proj_finish, project_number)

            final_rows.append({
                "Project ID": project_id,
                "Project Number": project_number,
                "File Name": file,
                "Task": task_name,
                "Resource": desc_text,
                "FusionResource": fusion_res,
                "Start Date": start_date,
                "End Date": end_date,
                "Cost": cost
            })

    print("--------------------------------------------------")

# ================= DETAIL (FINANCIAL PLAN) OUTPUT =================
detail_df = pd.DataFrame(final_rows)

def merge_resources(series):
    """
    If multiple source resources roll into the same FusionResource,
    return them as [res1, res2, ...].
    If only one, just return the single resource name.
    """
    uniq = sorted({str(v).strip() for v in series if pd.notna(v)})
    if not uniq:
        return ""
    if len(uniq) == 1:
        return uniq[0]
    return "[" + ", ".join(uniq) + "]"

# One row per Project + Task + FusionResource (duplicates summed)
detail_df = (
    detail_df
    .groupby(
        [
            "Project ID", "Project Number", "File Name",
            "Task", "FusionResource",
            "Start Date", "End Date"
        ],
        as_index=False
    )
    .agg({
        "Cost": "sum",
        "Resource": merge_resources
    })
)

# ===== TASK-LEVEL SUMMARY CHECK (sum of resources vs schedule) =====
summary_df = (
    detail_df
    .groupby(
        ["Project ID", "Project Number", "File Name", "Task"],
        as_index=False
    )["Cost"]
    .sum()
    .rename(columns={"Cost": "Estimate Cost"})
)

schedule_costs = []
diffs = []

for _, row in summary_df.iterrows():
    proj_num = row["Project Number"]
    task_name = row["Task"]
    est = row["Estimate Cost"]

    sched_cost = schedule_cost_for_task(proj_num, task_name, schedule_by_project)
    schedule_costs.append(sched_cost)

    if sched_cost is not None:
        diffs.append(est - sched_cost)
    else:
        diffs.append(None)

summary_df["Schedule Cost"] = schedule_costs
summary_df["Difference (Estimate - Schedule)"] = diffs

# NEW: simple match flag (are we matching the project schedule total?)
def match_flag(diff, sched_cost):
    if sched_cost is None:
        return "NO_SCHEDULE"
    if diff is None:
        return "NO_SCHEDULE"
    return "MATCH" if abs(diff) < 0.01 else "MISMATCH"

summary_df["Match"] = summary_df.apply(
    lambda r: match_flag(r["Difference (Estimate - Schedule)"], r["Schedule Cost"]),
    axis=1
)

# ================= SAVE – ONLY FINANCIAL PLAN + TASK SUMMARY =================
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    detail_df.to_excel(writer, index=False, sheet_name="Financial_Plan")
    summary_df.to_excel(writer, index=False, sheet_name="Task_Summary")

print("🎉 FINANCIAL PLAN + VALIDATION COMPLETE")
print(f"📁 Output file: {OUTPUT_FILE}")
print(f"📊 Financial plan rows: {len(detail_df)}")
print(f"📊 Task summary rows: {len(summary_df)}")