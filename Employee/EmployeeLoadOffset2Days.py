import pandas as pd
import re
from datetime import timedelta
import os


# =====================
# HELPERS
# =====================
def format_date(date_str, days_offset=0):
    try:
        dt = pd.to_datetime(date_str, dayfirst=True) + timedelta(days=days_offset)
        return dt.strftime('%Y/%m/%d')
    except (ValueError, TypeError):
        return '2010/09/07'

def get_initials(bu):
    clean_bu = bu.replace("LLC", "").replace("JV", "").replace(",", "").strip()
    clean_bu = re.sub(r"[^a-zA-Z0-9 ]", "", clean_bu)
    return "".join(w[0].upper() for w in clean_bu.split() if w)


# =====================
# MAIN FUNCTION
# =====================
def generate_worker_dat(source_path, ref_job_path, output_path='Worker-MOCK-FINAL-2.dat'):
    # ============================================================
    # REMOVE EMPLOYEE LOGIC (Exclude employees present in UAT file)
    # ============================================================
    EXCLUDE_EMPLOYEE_FILE = r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\Employee\Airetech\PreRequisite\UAT Expense Account V2.csv"

    # ---------- Load Source ----------
    if source_path.lower().endswith(('.xlsx', '.xls')):
        df_master = pd.read_excel(source_path, dtype=str)
    else:
        df_master = pd.read_csv(source_path, dtype=str, encoding='utf-8')

    df_master = df_master.fillna('')

    # ---------- Load Exclude File ----------
    df_exclude = pd.read_csv(EXCLUDE_EMPLOYEE_FILE, dtype=str, encoding='utf-8').fillna('')

    # Normalize IDs
    df_master['External Employee Id'] = df_master['External Employee Id'].astype(str).str.strip()
    df_exclude['External Employee Id'] = df_exclude['External Employee Id'].astype(str).str.strip()

    exclude_ids = set(df_exclude['External Employee Id'])

    # Recon - excluded records
    df_excluded_from_master = df_master[df_master['External Employee Id'].isin(exclude_ids)].copy()

    # Final df used for HDL
    df = df_master[~df_master['External Employee Id'].isin(exclude_ids)].copy()

    print(f"✅ Total employees in master        : {len(df_master)}")
    print(f"❌ Employees excluded (UAT match)  : {len(df_excluded_from_master)}")
    print(f"✅ Employees sent to HDL            : {len(df)}")

    if not df_excluded_from_master.empty:
        df_excluded_from_master.to_csv("Employees_Excluded_From_Master.csv", index=False)
        print("📄 Recon created: Employees_Excluded_From_Master.csv")

    if not df_exclude.empty:
        df_exclude.to_csv("Employees_Present_In_UAT_Expense.csv", index=False)
        print("📄 Recon created: Employees_Present_In_UAT_Expense.csv")

    # ---------- Load Job Reference ----------
    ref_jobs = pd.read_csv(ref_job_path, dtype=str, encoding='utf-8').fillna('')
    ref_job_map = dict(zip(
        ref_jobs['job_title'].str.strip(),
        ref_jobs['job_code'].str.strip()
    ))

    # ---------- Allowed Business Units ----------
    allowed_business_units = [
        "Airetech Corporation",
        "Etairos HVAC JV, LLC"
    ]
    df = df[df['Business Unit'].isin(allowed_business_units)]

    # ---------- Manager Maps ----------
    manager_map = {}
    hire_date_map = {}

    for _, row in df.iterrows():
        pn = str(row['External Employee Id']).strip()
        fname = str(row['First Name']).strip()
        lname = str(row['Last Name']).strip()
        email = str(row['Work Email']).strip()

        manager_map[f"{lname},{fname}"] = pn
        if email:
            manager_map[email] = pn

        hire_date_map[pn] = format_date(row['Hire Date'], days_offset=-2)

    failed_manager_assignments = []
    manager_hire_date_issues = []
    missing_job_codes = []

    # =====================
    # HDL HEADERS
    # =====================
    worker_lines = ["METADATA|Worker|PersonNumber|StartDate|EffectiveStartDate|ActionCode"]
    personname_lines = ["METADATA|PersonName|PersonNumber|EffectiveStartDate|NameType|FirstName|LastName|KnownAs"]
    personemail_lines = [
        "METADATA|PersonEmail|SourceSystemOwner|SourceSystemId|PersonNumber|DateFrom|DateTo|EmailType|EmailAddress|PrimaryFlag"]
    personphone_lines = [
        "METADATA|PersonPhone|SourceSystemOwner|SourceSystemId|PersonNumber|PhoneNumber|PhoneType|PrimaryFlag|DateFrom|DateTo"]

    personaddress_lines = [
        "METADATA|PersonAddress|PersonNumber|SourceSystemOwner|SourceSystemId|"
        "AddressType|EffectiveStartDate|EffectiveEndDate|"
        "AddressLine1|AddressLine2|AddressLine3|TownOrCity|Region1|Region2|Country|PostalCode|PrimaryFlag"
    ]

    workrelationship_lines = [
        "METADATA|WorkRelationship|PersonNumber|DateStart|PrimaryFlag|LegalEmployerName|WorkerType"]
    workterms_lines = [
        "METADATA|WorkTerms|PersonNumber|AssignmentNumber|LegalEmployerName|DateStart|EffectiveStartDate|EffectiveSequence|EffectiveLatestChange|ActionCode|PersonTypeCode|WorkerType|BusinessUnitShortCode|SystemPersonType|AssignmentName|AssignmentType|AssignmentStatusTypeCode|DefaultExpenseAccount"]
    assignment_lines = [
        "METADATA|Assignment|PersonNumber|AssignmentNumber|WorkTermsNumber|LegalEmployerName|DateStart|EffectiveStartDate|EffectiveSequence|EffectiveLatestChange|ActionCode|PersonTypeCode|BusinessUnitShortCode|WorkerType|SystemPersonType|AssignmentName|AssignmentType|AssignmentStatusTypeCode|DefaultExpenseAccount|LocationCode|OrganizationId|DepartmentName|JobCode"]
    assignmentsupervisor_lines = [
        "METADATA|AssignmentSupervisor|SourceSystemOwner|SourceSystemId|EffectiveStartDate|EffectiveEndDate|AssignmentNumber|ManagerType|ManagerAssignmentNumber|PrimaryFlag"]

    workerextrainfo_lines = [
        "METADATA|WorkerExtraInfo|FLEX:PER_PERSON_EIT_EFF|EFF_CATEGORY_CODE|"
        "etc(PER_PERSON_EIT_EFF=Air EFF)|isExempt(PER_PERSON_EIT_EFF=Air EFF)|"
        "EffectiveStartDate|EffectiveEndDate|PersonNumber|InformationType|"
        "PeiInformationCategory|SourceSystemOwner|SourceSystemId"
    ]

    # =====================
    # PROCESS RECORDS
    # =====================
    for _, row in df.iterrows():

        person_number = str(row['External Employee Id']).strip()
        first_name = str(row['First Name']).strip()
        last_name = str(row['Last Name']).strip()
        known_as = str(row['Preferred Name']).strip()
        email = str(row['Work Email']).strip()
        work_phone = str(row['Work Phone']).strip()
        mobile_phone = str(row['Mobile Phone']).strip()
        job_title = str(row['Job Title']).strip()
        department = str(row['department']).strip()
        legal_employer = str(row['Business Unit']).strip()
        business_unit = f"{legal_employer} BU"
        expense_account = str(row['ExpenseAccount']).strip()

        address1 = str(row.get("Address 1", "")).strip()
        address2 = str(row.get("Address 2", "")).strip()
        city = str(row.get("city", "")).strip()
        county = str(row.get("county", "")).strip()
        state = str(row.get("state", "")).strip()
        country = str(row.get("country", "")).strip()
        zip_code = str(row.get("zip", "")).strip()

        employee_hire_date = format_date(row['Hire Date'], days_offset=-2)
        end_date = "4712/12/31"
        job_code = ref_job_map.get(job_title)
        if not job_code:
            missing_job_codes.append({
                "Employee Name": f"{first_name} {last_name}",
                "Employee ID": person_number,
                "Job Title": job_title,
                "Business Unit": legal_employer
            })
            job_code = job_title

        # ---------- Worker ----------
        worker_lines.append(f"MERGE|Worker|{person_number}|{employee_hire_date}|{employee_hire_date}|HIRE")

        # ---------- Person Name ----------
        personname_lines.append(
            f"MERGE|PersonName|{person_number}|{employee_hire_date}|GLOBAL|{first_name}|{last_name}|{known_as}"
        )

        # ---------- Email ----------
        if email:
            personemail_lines.append(
                f"MERGE|PersonEmail|Employee_Navigator|EMAIL_{person_number}|{person_number}|"
                f"{employee_hire_date}|{end_date}|W1|{email}|Y"
            )

        # ---------- Phones ----------
        if work_phone:
            personphone_lines.append(
                f"MERGE|PersonPhone|Employee_Navigator|PHONE_WORK_{person_number}|{person_number}|"
                f"{work_phone}|W1|Y|{employee_hire_date}|{end_date}"
            )

        if mobile_phone:
            personphone_lines.append(
                f"MERGE|PersonPhone|Employee_Navigator|PHONE_MOBILE_{person_number}|{person_number}|"
                f"{mobile_phone}|WM|Y|{employee_hire_date}|{end_date}"
            )

        # ---------- Address ----------
        if address1 or address2 or city or country or zip_code:
            personaddress_lines.append(
                f"MERGE|PersonAddress|{person_number}|Employee_Navigator|ADDR_{person_number}|"
                f"HOME|{employee_hire_date}|{end_date}|"
                f"{address1}|{address2}||{city}|{county}|{state}|{country}|{zip_code}|Y"
            )

        # ---------- Work Relationship ----------
        workrelationship_lines.append(
            f"MERGE|WorkRelationship|{person_number}|{employee_hire_date}|Y|{legal_employer}|E"
        )

        # ---------- Work Terms ----------
        workterms_lines.append(
            f"MERGE|WorkTerms|{person_number}|ET_{person_number}|{legal_employer}|"
            f"{employee_hire_date}|{employee_hire_date}|1|Y|HIRE|Employee|E|"
            f"{business_unit}|EMP|ET_{person_number}|ET|ACTIVE_PROCESS|{expense_account}"
        )

        # ---------- Assignment ----------
        assignment_lines.append(
            f"MERGE|Assignment|{person_number}|E_{person_number}|ET_{person_number}|"
            f"{legal_employer}|{employee_hire_date}|{employee_hire_date}|1|Y|HIRE|Employee|"
            f"{business_unit}|E|EMP|E_{person_number}|E|ACTIVE_PROCESS|"
            f"{expense_account}|||{department}|{job_code}"
        )

        # ---------- Assignment Supervisor ----------
        manager_name = str(row['Manager']).strip()
        manager_email = str(row['Manager Email']).strip()
        manager_pn = manager_map.get(manager_name) or manager_map.get(manager_email)

        if not manager_pn:
            failed_manager_assignments.append({
                "Employee Name": f"{first_name} {last_name}",
                "Employee ID": person_number,
                "Manager Name": manager_name,
                "Manager Email": manager_email,
                "Business Unit": legal_employer
            })
        else:
            manager_hire_date = hire_date_map.get(manager_pn)
            effective_start = employee_hire_date

            if manager_hire_date and manager_hire_date > employee_hire_date:
                effective_start = manager_hire_date
                manager_hire_date_issues.append({
                    "Employee Name": f"{first_name} {last_name}",
                    "Employee ID": person_number,
                    "Employee Hire Date": employee_hire_date,
                    "Manager ID": manager_pn,
                    "Manager Hire Date": manager_hire_date,
                    "Supervisor Start Date Used": effective_start,
                    "Business Unit": legal_employer
                })

            assignmentsupervisor_lines.append(
                f"MERGE|AssignmentSupervisor|Employee_Navigator|SUP_{person_number}|"
                f"{effective_start}|{end_date}|E_{person_number}|"
                f"LINE_MANAGER|E_{manager_pn}|Y"
            )

        # ---------- Worker Extra Info ----------
        eta = str(row.get("Eta", "")).strip()
        exempt = str(row.get("Is Exempt", "")).strip()

        if eta or exempt:
            workerextrainfo_lines.append(
                f"MERGE|WorkerExtraInfo|Air EFF|PER_EIT|{eta}|{exempt}|"
                f"{employee_hire_date}|{end_date}|{person_number}|Air EFF|Air EFF|"
                f"Employee_Navigator|EFF_{person_number}"
            )

    # =====================
    # WRITE DAT
    # =====================
    all_lines = (
            worker_lines + [""] +
            personname_lines + [""] +
            personemail_lines + [""] +
            personphone_lines + [""] +
            personaddress_lines + [""] +
            workrelationship_lines + [""] +
            workterms_lines + [""] +
            assignment_lines + [""] +
            assignmentsupervisor_lines + [""] +
            workerextrainfo_lines + [""]
    )

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(all_lines))

    print(f"✅ Worker DAT created: {output_path}")

    # =====================
    # RECON FILES
    # =====================
    if failed_manager_assignments:
        pd.DataFrame(failed_manager_assignments).to_csv("Missing_Manager_Assignments.csv", index=False)
        print("❌ Missing_Manager_Assignments.csv created")

    if manager_hire_date_issues:
        pd.DataFrame(manager_hire_date_issues).to_csv("Manager_Hire_Date_After_Employee.csv", index=False)
        print("⚠️ Manager_Hire_Date_After_Employee.csv created")
    else:
        print("✅ No manager hire date issues found")

    if missing_job_codes:
        pd.DataFrame(missing_job_codes).to_csv("Missing_Job_Codes.csv", index=False)
        print("⚠️ Missing_Job_Codes.csv created")
    else:
        print("✅ All job titles matched reference file")


# EXECUTION
generate_worker_dat(
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\Employee\Gas\SourceFile\GAS_Employees.csv",
    r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\Employee\Airetech\PreRequisite\Reference_Job_PROD.csv"
)