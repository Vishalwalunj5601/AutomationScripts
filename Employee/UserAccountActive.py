import requests
from requests.auth import HTTPBasicAuth

# ==========================================================
# Configuration
# ==========================================================
BASE_URL = "https://air-ibyqjb-dev1.fa.ocs.oraclecloud.com/hcmRestApi/resources/11.13.18.05/userAccounts"

FUSION_USERNAME = "Conversion"
FUSION_PASSWORD = "Welcome@12345"

# ==========================================================
# If USERNAMES has values -> Process only these users
# If USERNAMES = [] -> Activate ALL suspended users
# ==========================================================
USERNAMES = [
    "jaldrich@airetechcorp.com",
    # "john.doe@company.com",
]

headers = {
    "Content-Type": "application/vnd.oracle.adf.resourceitem+json",
    "Accept": "application/json"
}


# ==========================================================
# Get all users (handles pagination)
# ==========================================================
def get_all_users():
    offset = 0
    limit = 1000

    while True:

        url = f"{BASE_URL}?limit={limit}&offset={offset}"

        response = requests.get(
            url,
            headers=headers,
            auth=HTTPBasicAuth(FUSION_USERNAME, FUSION_PASSWORD)
        )

        response.raise_for_status()

        data = response.json()
        users = data.get("items", [])

        if not users:
            break

        for user in users:
            yield user

        if len(users) < limit:
            break

        offset += limit


# ==========================================================
# Find a user by Username
# ==========================================================
def find_user(username):

    for user in get_all_users():

        if user.get("Username", "").strip().lower() == username.strip().lower():
            return user

    return None


# ==========================================================
# Activate user
# ==========================================================
def activate_user(user):

    username = user.get("Username")

    if user.get("SuspendedFlag") is False:
        print(f"ℹ {username} : Already Active")
        return

    # Get self URL
    self_url = None

    for link in user.get("links", []):
        if link.get("rel") == "self":
            self_url = link.get("href")
            break

    if self_url is None:
        print(f"❌ {username} : Self URL not found")
        return

    payload = {
        "SuspendedFlag": False
    }

    response = requests.patch(
        self_url,
        headers=headers,
        auth=HTTPBasicAuth(FUSION_USERNAME, FUSION_PASSWORD),
        json=payload
    )

    if response.status_code in (200, 204):
        print(f"✅ {username} : Activated Successfully")
    else:
        print(f"❌ {username} : Activation Failed")
        print(response.text)


# ==========================================================
# Main
# ==========================================================
def main():

    if USERNAMES:

        print(f"\nProcessing {len(USERNAMES)} user(s)...\n")

        for username in USERNAMES:

            print(f"Searching : {username}")

            user = find_user(username)

            if user:
                activate_user(user)
            else:
                print(f"❌ {username} : User Not Found")

    else:

        print("\nNo usernames supplied.")
        print("Processing ALL suspended users...\n")

        scanned = 0
        activated = 0

        for user in get_all_users():

            scanned += 1

            if user.get("SuspendedFlag") is True:
                activate_user(user)
                activated += 1

        print("\n================================")
        print(f"Users Scanned   : {scanned}")
        print(f"Users Activated : {activated}")
        print("================================")


if __name__ == "__main__":
    main()