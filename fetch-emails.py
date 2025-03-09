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

# Error log file
ERROR_LOG_FILE = "errors.log"

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

# 🔹 Rate-Limited Request Handler
def make_request_with_backoff(url, headers, max_retries=5):
    """Handles rate limiting with exponential backoff."""
    retry_count = 0
    
    while retry_count < max_retries:
        response = requests.get(url, headers=headers)

        if response.status_code == 429:  # Too Many Requests
            retry_after = int(response.headers.get("Retry-After", 5))  # Default to 5 sec
            print(f"⚠️ Rate limit hit! Retrying in {retry_after} sec...")
            time.sleep(retry_after)  # Wait before retrying
            retry_count += 1
        else:
            return response  # Success!

    log_error("Max retries reached for URL", url)
    return None

# 🔹 Error Logging
def log_error(message, detail=""):
    """Logs errors to a file for later review."""
    with open(ERROR_LOG_FILE, "a") as error_log:
        error_log.write(f"{message}: {detail}\n")
    print(f"⚠️ Logged error: {message}")

# 🔹 Step 1: Fetch all folders and estimate total emails
print("📂 Fetching all mail folders and estimating total emails...")
folder_lookup = {}
parent_folder_lookup = {}
total_email_estimate = 0

def fetch_folders(url):
    """Recursively fetch all folders and estimate total emails"""
    global total_email_estimate

    while url:
        print(f"🔄 Fetching folder list: {url}")

        headers["Authorization"] = f"Bearer {get_access_token()}"
        response = make_request_with_backoff(url, headers)

        if response is None:
            return  

        print(f"✅ Folder list response received")  
        data = response.json()

        for folder in data.get("value", []):
            folder_lookup[folder["id"]] = folder["displayName"]
            parent_folder_lookup[folder["id"]] = folder.get("parentFolderId")

            count_url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders/{folder['id']}/messages/$count"
            count_headers = {**headers, "ConsistencyLevel": "eventual"}
            count_response = make_request_with_backoff(count_url, count_headers)

            if count_response and count_response.status_code == 200:
                folder_email_count = int(count_response.text)
                total_email_estimate += folder_email_count
            else:
                folder_email_count = ""
                log_error("Failed to estimate email count", folder["displayName"])

            print(f"📂 {folder['displayName']} - Estimated Emails: {folder_email_count}")

        url = data.get("@odata.nextLink")

fetch_folders(f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders?$top=200")

print(f"\n📊 Estimated Total Emails to Process: {total_email_estimate}\n")

# 🔹 Step 2: Fetch all emails (WRITE DIRECTLY TO CSV)
csv_filename = "email_metadata.csv"
first_write = True  
email_count = 0
start_time = time.time()

def fetch_emails_from_folder(folder_id, folder_name):
    """Fetch all emails from a specific folder, following pagination correctly"""
    global first_write, email_count  

    url = f"https://graph.microsoft.com/v1.0/users/{EMAIL}/mailFolders/{folder_id}/messages?$top=100"

    while url:
        print(f"🔄 Fetching emails from folder: {folder_name}")

        headers["Authorization"] = f"Bearer {get_access_token()}"
        response = make_request_with_backoff(url, headers)

        if response is None:
            log_error("Skipping folder due to API failures", folder_name)
            break  

        print(f"✅ Response received for folder {folder_name}")  
        data = response.json()

        if "error" in data:
            log_error(f"API Error in {folder_name}", data['error']['message'])
            break  

        email_batch = []
        for email in data.get("value", []):
            try:
                email_metadata = {
                    "EmailID": email.get("id", ""),
                    "InternetMessageID": email.get("internetMessageId", ""),
                    "ConversationID": email.get("conversationId", ""),
                    "Subject": email.get("subject", ""),
                    "From": email.get("from", {}).get("emailAddress", {}).get("address", ""),
                    "To": "; ".join([recipient.get("emailAddress", {}).get("address", "") for recipient in email.get("toRecipients", [])]),
                    "Cc": "; ".join([recipient.get("emailAddress", {}).get("address", "") for recipient in email.get("ccRecipients", [])]),
                    "Bcc": "; ".join([recipient.get("emailAddress", {}).get("address", "") for recipient in email.get("bccRecipients", [])]),
                    "ReceivedDateTime": email.get("receivedDateTime", ""),
                    "SentDateTime": email.get("sentDateTime", ""),
                    "FolderName": folder_name,
                    "ParentFolderName": folder_lookup.get(parent_folder_lookup.get(folder_id, ""), ""),
                    "Importance": email.get("importance", ""),
                    "IsRead": email.get("isRead", ""),
                    "HasAttachments": email.get("hasAttachments", ""),
                    "Categories": ", ".join(email.get("categories", [])),
                }

                email_batch.append(email_metadata)
                email_count += 1

                # 🔹 Progress & ETA Calculation
                if email_count % 100 == 0:
                    elapsed_time = time.time() - start_time
                    speed = email_count / elapsed_time if elapsed_time > 0 else 0
                    remaining = total_email_estimate - email_count
                    estimated_seconds_remaining = remaining / speed if speed > 0 else 0

                    hours = int(estimated_seconds_remaining // 3600)
                    minutes = int((estimated_seconds_remaining % 3600) // 60)
                    seconds = int(estimated_seconds_remaining % 60)
                    eta_formatted = f"{hours}h {minutes}m {seconds}s" if hours > 0 else f"{minutes}m {seconds}s"

                    print(f"📊 Processed {email_count}/{total_email_estimate} emails ({(email_count/total_email_estimate)*100:.2f}%) | Speed: {speed:.2f} emails/sec | ETA: {eta_formatted}")

            except Exception as e:
                log_error(f"Skipping problematic email in {folder_name}", str(e))

        df = pd.DataFrame(email_batch)
        df.to_csv(csv_filename, mode='a', index=False, header=first_write)
        first_write = False  

        url = data.get("@odata.nextLink")

# 🔹 Step 3: Process Emails in Each Folder
print("\n📨 Fetching ALL emails from all folders...\n")

for folder_id, folder_name in folder_lookup.items():
    print(f"📂 Processing folder: {folder_name} ({folder_id})")
    fetch_emails_from_folder(folder_id, folder_name)

print(f"\n✅ Finished fetching emails! Total Retrieved: {email_count}\n")
print(f"✅ Email metadata saved to {csv_filename} 🎉")
print(f"⚠️ Errors logged in {ERROR_LOG_FILE}, check for skipped emails.")
