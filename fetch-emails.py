import requests
import json
import pandas as pd
from msal import ConfidentialClientApplication

# Load Azure credentials
with open("config.json") as f:
    config = json.load(f)

CLIENT_ID = config["client_id"]
CLIENT_SECRET = config["client_secret"]
TENANT_ID = config["tenant_id"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# Authenticate and get an access token
app = ConfidentialClientApplication(CLIENT_ID, CLIENT_SECRET, AUTHORITY)
token_response = app.acquire_token_for_client(SCOPE)
access_token = token_response["access_token"]

# Headers for Graph API
headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

# Fetch emails from Microsoft Graph API
url = "https://graph.microsoft.com/v1.0/me/messages?$top=100"
response = requests.get(url, headers=headers)
emails = response.json().get("value", [])

# Extract metadata
email_data = []
for email in emails:
    email_data.append({
        "EmailID": email["id"],
        "InternetMessageID": email.get("internetMessageId", ""),
        "ConversationID": email.get("conversationId", ""),
        "Subject": email["subject"],
        "From": email["from"]["emailAddress"]["address"] if "from" in email else "",
        "To": "; ".join([recipient["emailAddress"]["address"] for recipient in email.get("toRecipients", [])]),
        "Date": email["receivedDateTime"],
        "FolderID": email.get("parentFolderId", ""),
        "Importance": email["importance"],
        "IsRead": email["isRead"],
        "HasAttachments": email["hasAttachments"]
    })

# Convert to DataFrame
df = pd.DataFrame(email_data)

# Save to CSV
df.to_csv("email_metadata.csv", index=False)

print("✅ Email metadata saved to email_metadata.csv")
