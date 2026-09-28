import requests
from requests.auth import HTTPBasicAuth

BASE_URL = "https://air-ibyqjb-dev1.fa.ocs.oraclecloud.com/hcmRestApi/resources/11.13.18.05/userAccounts"

response = requests.get(
    f"{BASE_URL}?limit=1000",
    auth=HTTPBasicAuth("Conversion", "Welcome@12345"),
    headers={"Accept": "application/json"}
)

response.raise_for_status()

data = response.json()

print("Items Returned:", len(data["items"]))

for user in data["items"]:
    print(user["Username"])