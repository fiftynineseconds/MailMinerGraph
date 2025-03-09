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

# Authenticate with Microsoft Graph API
app = ConfidentialClientApplication(CLIENT_ID, CLIENT_SECRET, AUTHORITY)
token_response = app.acquire_token_silent(SCOPE, account=None)

if not token_response:
    token_response = app.acquire_token_for_client(SCOPE)

if "access_token" in token_response:
    access_token = token_response["access_token"]
    print("✅ Access token obtained successfully!\n")
else:
    print("❌ Failed to obtain access token:", token_response.get("error_description", token_response))
    exit()

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

# 🔹 Fetch all mail folders for lookup
print("📂 Fetching mail folders...")
folder_lookup = {}
parent_folder_lookup = {}

folder_url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders?$top=100"
while folder_url:
    folder_response = requests.get(folder_url, headers=headers).json()
    
    for folder in folder_response.get("value", []):
        folder_lookup[folder["id"]] = folder["displayName"]
        parent_folder_lookup[folder["id"]] = folder.get("parentFolderId", "")

    folder_url = folder_response.get("@odata.nextLink")

print(f"✅ Retrieved {len(folder_lookup)} folders.\n")

# 🔹 Ask user for folder name (default: Inbox)
folder_name = input("📂 Enter folder name (default: inbox): ") or "inbox"
print(f"\n🔍 Fetching emails from '{folder_name}' folder for {EMAIL}...\n")

# Fetch emails
url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders/{folder_name}/messages?$top=100"
email_data = []

while url:
    response = requests.get(url, headers=headers)
    data = response.json()

    if "error" in data:
        print(f"❌ API Error: {data['error']['message']}")
        exit()

    for email in data.get("value", []):
        folder_id = email.get("parentFolderId", "")
        folder_name = folder_lookup.get(folder_id, "Unknown Folder")
        parent_folder_name = folder_lookup.get(parent_folder_lookup.get(folder_id, ""), "Unknown Parent")

        email_data.append({
            "EmailID": email["id"],
            "InternetMessageID": email.get("internetMessageId", ""),
            "ConversationID": email.get("conversationId", ""),
            "Subject": email["subject"],
            "From": email.get("from", {}).get("emailAddress", {}).get("address", ""),
            "To": "; ".join([recipient["emailAddress"]["address"] for recipient in email.get("toRecipients", [])]),
            "Date": email["receivedDateTime"],
            "FolderName": folder_name,
            "ParentFolderName": parent_folder_name,
            "Importance": email["importance"],
            "IsRead": email["isRead"],
            "HasAttachments": email["hasAttachments"],
            "Categories": ", ".join(email.get("categories", [])),
            "Preview": email.get("bodyPreview", "").replace("\n", " ")[:100]  # Limit to 100 chars
        })

    url = data.get("@odata.nextLink")

# Save emails to CSV
df = pd.DataFrame(email_data)
csv_filename = "email_metadata.csv"
df.to_csv(csv_filename, index=False)
print(f"\n✅ Email metadata saved to {csv_filename} 🎉")
