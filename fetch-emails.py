import json
import requests
import pandas as pd
import time
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
access_token = None
token_expiration = 0  

def get_access_token():
    """Gets a new access token and refreshes it if expired"""
    global access_token, token_expiration

    if access_token and time.time() < token_expiration - 60:
        return access_token

    print("🔄 Refreshing access token...")

    token_response = app.acquire_token_silent(SCOPE, account=None)
    if not token_response:
        token_response = app.acquire_token_for_client(SCOPE)

    if "access_token" in token_response:
        access_token = token_response["access_token"]
        token_expiration = time.time() + 3600  
        return access_token
    else:
        print("❌ Failed to refresh access token:", token_response.get("error_description", token_response))
        exit()

headers = {
    "Authorization": f"Bearer {get_access_token()}",
    "Content-Type": "application/json"
}

# 🔹 Step 1: Fetch all folders (INCLUDING SUBFOLDERS)
print("📂 Fetching all mail folders, including subfolders...")
folder_lookup = {}
parent_folder_lookup = {}

def fetch_folders(url):
    """Recursively fetch all folders and subfolders"""
    while url:
        headers["Authorization"] = f"Bearer {get_access_token()}"
        response = requests.get(url, headers=headers).json()

        for folder in response.get("value", []):
            folder_lookup[folder["id"]] = folder["displayName"]
            parent_folder_lookup[folder["id"]] = folder.get("parentFolderId")

            # If the folder has subfolders, fetch them recursively
            subfolder_url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders/{folder['id']}/childFolders"
            fetch_folders(subfolder_url)

        url = response.get("@odata.nextLink")

fetch_folders(f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders?$top=200")

print(f"✅ Retrieved {len(folder_lookup)} folders (including subfolders).\n")

# 🔹 Step 2: Fetch all emails (including subfolders)
email_data = []
email_count = 0

print("📨 Fetching ALL emails from all folders and subfolders...\n")
for folder_id, folder_name in folder_lookup.items():
    print(f"📂 Processing folder: {folder_name} ({folder_id})")

    url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders/{folder_id}/messages?$top=100"

    while url:
        headers["Authorization"] = f"Bearer {get_access_token()}"
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
            if email_count % 100 == 0:
                print(f"📊 Processed {email_count} emails...")

        url = data.get("@odata.nextLink")  

print(f"\n✅ Finished fetching emails! Total Retrieved: {email_count}\n")

# 🔹 Step 3: Save emails to CSV
df = pd.DataFrame(email_data)
csv_filename = "email_metadata.csv"
df.to_csv(csv_filename, index=False)
print(f"✅ Email metadata saved to {csv_filename} 🎉")
