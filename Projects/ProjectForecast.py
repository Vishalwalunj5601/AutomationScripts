import json
import time
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

# =========================================================
# CONFIG
# =========================================================
OracleURL="https://air-ibyqjb-dev1.fa.ocs.oraclecloud.com"
BASE_URL = f"{OracleURL}:443/fscmRestApi/resources/11.13.18.05/projectForecasts"

USERNAME = "Conversion"
PASSWORD = "Welcome@12345"

OUTPUT_FILE = r"C:\Users\vishal walunj\Vishal\1D_AIR_CONTROL_CONCEPTS\Projects\dataLoad_New\Cost\PjcTxnXfaceStageAll_misc\PjcTxnXfaceStageAll_misc.xlsx"

REQUEST_TIMEOUT = 120
SLEEP_BETWEEN_CALLS = 0.3

HEADERS = {
    "Content-Type": "application/vnd.oracle.adf.resourceitem+json",
    "Accept": "application/json"
}

# =========================================================
# PROJECT LIST
# =========================================================
PROJECT_NUMBERS = ["0000087889",
"0000087912",
"0000087913",
"0000087925",
"0000087930",
"0000087937",
"0000088016",
"0000088054",
"0000088101",
"0000088136",
"0000088166",
"0000088203",
"0000088204",
"0000088205",
]

# =========================================================
# HELPERS
# =========================================================
def build_payload(project_number: str) -> dict:
    return {
        "ProjectNumber": project_number,
        "FinancialPlanType": "AIR - Financial Plan Type - Forecast",
        "ForecastCreationMethod": "GENERATE",
        "ForecastGenerationSource": "PROJECT_PLAN",
        "PlanVersionStatus": "Current Working",
        "SourcePlanVersionStatus": "Current Working"
    }

def extract_error_text(response):
    try:
        data = response.json()
        if isinstance(data, dict):
            # Common Oracle REST error fields
            for key in ["title", "detail", "o:errorDetails", "errorDetails", "message"]:
                if key in data:
                    return str(data[key])

            # Sometimes nested structures are returned
            return json.dumps(data, ensure_ascii=False)
        return str(data)
    except Exception:
        return response.text.strip()

def create_project_forecast(project_number: str):
    payload = build_payload(project_number)

    try:
        response = requests.post(
            BASE_URL,
            headers=HEADERS,
            auth=HTTPBasicAuth(USERNAME, PASSWORD),
            json=payload,
            timeout=REQUEST_TIMEOUT
        )

        result = {
            "ProjectNumber": project_number,
            "StatusCode": response.status_code,
            "Status": "SUCCESS" if response.ok else "FAILED",
            "Response": "",
            "Payload": json.dumps(payload, ensure_ascii=False)
        }

        if response.ok:
            try:
                resp_json = response.json()
                result["Response"] = json.dumps(resp_json, ensure_ascii=False)
            except Exception:
                result["Response"] = response.text.strip()
        else:
            result["Response"] = extract_error_text(response)

        return result

    except requests.exceptions.RequestException as e:
        return {
            "ProjectNumber": project_number,
            "StatusCode": "REQUEST_ERROR",
            "Status": "FAILED",
            "Response": str(e),
            "Payload": json.dumps(payload, ensure_ascii=False)
        }

# =========================================================
# MAIN
# =========================================================
def main():
    results = []

    print(f"Starting forecast creation for {len(PROJECT_NUMBERS)} projects...\n")

    for i, project_number in enumerate(PROJECT_NUMBERS, start=1):
        print(f"[{i}/{len(PROJECT_NUMBERS)}] Processing Project: {project_number}")
        result = create_project_forecast(project_number)
        results.append(result)
        print(f"   -> StatusCode: {result['StatusCode']} | Status: {result['Status']}")
        time.sleep(SLEEP_BETWEEN_CALLS)

    df = pd.DataFrame(results)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="ForecastResults", index=False)

        summary = pd.DataFrame({
            "Metric": [
                "Total Projects",
                "Success Count",
                "Failed Count"
            ],
            "Value": [
                len(df),
                (df["Status"] == "SUCCESS").sum(),
                (df["Status"] == "FAILED").sum()
            ]
        })
        summary.to_excel(writer, sheet_name="Summary", index=False)

    print("\nDone.")
    print(f"Result file saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()