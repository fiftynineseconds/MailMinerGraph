import json
import requests
import pandas as pd
from msal import ConfidentialClientApplication

# Load credentials from config.json
with open("config.json") as f:
    config = json.load(f)

CLIENT_ID = config["client_id"]
CLIENT_SECRET = config["client_secret"]
TENANT_ID = config["tenant_id"]
EMAIL = config["email"]  # Load email dynamically

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# Create an MSAL authentication app
app = ConfidentialClientApplication(CLIENT_ID, CLIENT_SECRET, AUTHORITY)

# Try to get a cached token first
token_response = app.acquire_token_silent(SCOPE, account=None)

# If no cached token, request a new one
if not token_response:
    token_response = app.acquire_token_for_client(SCOPE)

if "access_token" in token_response:
    access_token = token_response["access_token"]
    print("✅ Access token obtained successfully!\n")
else:
    print("❌ Failed to obtain access token:", token_response.get("error_description", token_response))
    exit()

# Headers for API requests
headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

# Ask user for folder (or use default Inbox)
folder_name = input("📂 Enter folder name (default: inbox): ") or "inbox"
print(f"\n🔍 Fetching emails from '{folder_name}' folder for {EMAIL}...\n")

# Use EMAIL from config.json in API request
url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders/{folder_name}/messages?$top=100"

email_data = []

while url:
    response = requests.get(url, headers=headers)
    data = response.json()

    # Print raw response for debugging
    print("\n📜 RAW API RESPONSE:\n", json.dumps(data, indent=2), "\n")

    # Handle errors
    if "error" in data:
        print(f"❌ API Error: {data['error']['message']}")
        exit()

    # Extract emails
    if "value" in data:
        for email in data["value"]:
            email_data.append({
                "EmailID": email["id"],
                "InternetMessageID": email.get("internetMessageId", ""),
                "ConversationID": email.get("conversationId", ""),
                "Subject": email["subject"],
                "From": email.get("from", {}).get("emailAddress", {}).get("address", ""),
                "To": "; ".join([recipient["emailAddress"]["address"] for recipient in email.get("toRecipients", [])]),
                "Date": email["receivedDateTime"],
                "FolderID": email.get("parentFolderId", ""),
                "Importance": email["importance"],
                "IsRead": email["isRead"],
                "HasAttachments": email["hasAttachments"]
            })

    # Handle pagination (fetch next batch of emails if available)
    url = data.get("@odata.nextLink")

# Save emails to CSV
df = pd.DataFrame(email_data)
csv_filename = "email_metadata.csv"
df.to_csv(csv_filename, index=False)
print(f"\n✅ Email metadata saved to {csv_filename} 🎉")
