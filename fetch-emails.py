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
total_email_estimate = 0  # Store estimated total email count

def fetch_folders(url):
    """Recursively fetch all folders and subfolders"""
    global total_email_estimate
    while url:
        headers["Authorization"] = f"Bearer {get_access_token()}"
        response = requests.get(url, headers=headers).json()

        for folder in response.get("value", []):
            folder_lookup[folder["id"]] = folder["displayName"]
            parent_folder_lookup[folder["id"]] = folder.get("parentFolderId")

            # Estimate email count for each folder
            count_url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders/{folder['id']}/messages/$count"
            count_headers = {**headers, "ConsistencyLevel": "eventual"}
            count_response = requests.get(count_url, headers=count_headers)

            if count_response.status_code == 200:
                folder_email_count = int(count_response.text)
                total_email_estimate += folder_email_count
            else:
                folder_email_count = "Unknown"

            print(f"📂 {folder['displayName']} ({folder['id']}) - Estimated Emails: {folder_email_count}")

            # Recursively fetch subfolders
            subfolder_url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders/{folder['id']}/childFolders"
            fetch_folders(subfolder_url)

        url = response.get("@odata.nextLink")

fetch_folders(f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders?$top=200")

print(f"\n📊 Estimated Total Emails to Process: {total_email_estimate}\n")

# 🔹 Step 2: Fetch all emails (WRITE DIRECTLY TO CSV)
csv_filename = "email_metadata.csv"
first_write = True  # To track if it's the first write (for headers)

print("📨 Fetching ALL emails from all folders...\n")
email_count = 0
start_time = time.time()

for folder_id, folder_name in folder_lookup.items():
    print(f"📂 Processing folder: {folder_name} ({folder_id})")

    url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders/{folder_id}/messages?$top=100"

    while url:
        headers["Authorization"] = f"Bearer {get_access_token()}"
        response = requests.get(url, headers=headers)
        data = response.json()

        if "error" in data:
            print(f"❌ API Error in {folder_name}: {data['error']['message']}")
            break  

        email_batch = []
        for email in data.get("value", []):
            parent_folder_id = parent_folder_lookup.get(folder_id)
            parent_folder_name = folder_lookup.get(parent_folder_id, "Root Folder" if parent_folder_id is None else "Unknown Parent")

            email_metadata = {
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
            }

            email_batch.append(email_metadata)
            email_count += 1

            # 🔹 Print progress indicator
            if email_count % 100 == 0:
                elapsed_time = time.time() - start_time
                speed = email_count / elapsed_time  # Emails per second
                estimated_time_remaining = (total_email_estimate - email_count) / speed if speed > 0 else 0

                progress = (email_count / total_email_estimate) * 100 if total_email_estimate > 0 else 0
                print(f"📊 Processed {email_count}/{total_email_estimate} emails ({progress:.2f}%) | Speed: {speed:.2f} emails/sec | ETA: {estimated_time_remaining:.2f} sec")

        # Write batch to CSV immediately
        df = pd.DataFrame(email_batch)
        df.to_csv(csv_filename, mode='a', index=False, header=first_write)
        first_write = False  # Ensure header is only written once

        url = data.get("@odata.nextLink")  

print(f"\n✅ Finished fetching emails! Total Retrieved: {email_count}\n")
print(f"✅ Email metadata saved to {csv_filename} 🎉")
