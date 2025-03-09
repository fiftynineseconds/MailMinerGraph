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
EMAIL = config["email"]

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

# 🔹 Step 1: Fetch all mail folders
print("📂 Fetching all mail folders...")
folder_lookup = {}
parent_folder_lookup = {}
total_emails_estimated = 0  # Store the estimated total email count

folder_url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders?$top=200"
while folder_url:
    folder_response = requests.get(folder_url, headers=headers).json()

    for folder in folder_response.get("value", []):
        folder_id = folder["id"]
        folder_lookup[folder_id] = folder["displayName"]
        parent_folder_lookup[folder_id] = folder.get("parentFolderId")

        # 🔹 Get estimated email count for each folder
        count_url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders/{folder_id}/messages/$count"
        count_headers = {**headers, "ConsistencyLevel": "eventual"}
        count_response = requests.get(count_url, headers=count_headers)

        if count_response.status_code == 200:
            folder_email_count = int(count_response.text)
            total_emails_estimated += folder_email_count
        else:
            folder_email_count = "Unknown"

        print(f"📂 {folder['displayName']} ({folder_id}) - Estimated Emails: {folder_email_count}")

    folder_url = folder_response.get("@odata.nextLink")

print(f"\n📊 Estimated Total Emails to Process: {total_emails_estimated}\n")
print("📨 Fetching ALL emails from all folders...\n")

# 🔹 Step 2: Fetch all emails (with real-time progress)
email_data = []
email_count = 0

for folder_id, folder_name in folder_lookup.items():
    print(f"📂 Processing folder: {folder_name} ({folder_id})")

    url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders/{folder_id}/messages?$top=100"

    while url:
        response = requests.get(url, headers=headers)
        data = response.json()

        if "error" in data:
            print(f"❌ API Error: {data['error']['message']}")
            break

        for email in data.get("value", []):
            parent_folder_id = parent_folder_lookup.get(folder_id)
            parent_folder_name = folder_lookup.get(parent_folder_id, "Root Folder" if parent_folder_id is None else "Unknown Parent")

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
            })

            email_count += 1

            # 🔹 Print progress every 100 emails
            if email_count % 100 == 0:
                print(f"📊 Processed {email_count} emails...")

        url = data.get("@odata.nextLink")  # Handle pagination

print(f"\n✅ Finished fetching emails! Total Retrieved: {email_count}\n")

# 🔹 Step 3: Save emails to CSV
df = pd.DataFrame(email_data)
csv_filename = "email_metadata.csv"
df.to_csv(csv_filename, index=False)
print(f"✅ Email metadata saved to {csv_filename} 🎉")
