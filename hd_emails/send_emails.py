import os
import json
import base64
import requests
from msal import PublicClientApplication, SerializableTokenCache
from dotenv import load_dotenv

load_dotenv()

# === CONFIGURATION ===
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.Send"]
TOKEN_CACHE_PATH = "msal_token_cache.json"

TO_EMAILS = ["panderson@americanbathgroup.com"]
SUBJECT = "TESTTESTTEST2"
BODY = """1
This message was sent using Microsoft Graph API.
2
This is a test email with an attachment.
3
The only recipient is panderson@americanbathgroup.com.
4
5"""

ATTACHMENT_PATH = r"C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports/HD_Update_Needed_2025-05-08.xlsx"

# === TOKEN CACHE SETUP ===
cache = SerializableTokenCache()
if os.path.exists(TOKEN_CACHE_PATH):
    cache.deserialize(open(TOKEN_CACHE_PATH, "r").read())

app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

# === TRY TO ACQUIRE TOKEN FROM CACHE FIRST ===
accounts = app.get_accounts()
token = None
if accounts:
    token = app.acquire_token_silent(SCOPES, account=accounts[0])

# === FALLBACK TO DEVICE CODE FLOW ===
if not token:
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception("Failed to initiate device flow")
    print(flow["message"])  # Prompt user to sign in
    token = app.acquire_token_by_device_flow(flow)

# === SAVE UPDATED CACHE IF NEEDED ===
if cache.has_state_changed:
    with open(TOKEN_CACHE_PATH, "w") as f:
        f.write(cache.serialize())

if "access_token" not in token:
    raise Exception(f"Could not obtain token.\n{token.get('error_description', token)}")

# === LOAD AND ENCODE ATTACHMENT ===
if not os.path.exists(ATTACHMENT_PATH):
    raise FileNotFoundError(f"Attachment file not found: {ATTACHMENT_PATH}")

with open(ATTACHMENT_PATH, "rb") as f:
    attachment_content = base64.b64encode(f.read()).decode('utf-8')

attachment = {
    "@odata.type": "#microsoft.graph.fileAttachment",
    "name": os.path.basename(ATTACHMENT_PATH),
    "contentBytes": attachment_content,
    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
}

# === BUILD EMAIL MESSAGE ===
email_msg = {
    "message": {
        "subject": SUBJECT,
        "body": {
            "contentType": "Text",
            "content": BODY
        },
        "toRecipients": [{"emailAddress": {"address": addr}} for addr in TO_EMAILS],
        "attachments": [attachment]
    },
    "saveToSentItems": "true"
}

# === SEND EMAIL VIA GRAPH API ===
response = requests.post(
    "https://graph.microsoft.com/v1.0/me/sendMail",
    headers={
        "Authorization": f"Bearer {token['access_token']}",
        "Content-Type": "application/json"
    },
    json=email_msg
)

# === REPORT RESULT ===
if response.status_code == 202:
    print("✅ Email sent successfully.")
else:
    print(f"❌ Failed to send email: {response.status_code}\n{response.text}")
