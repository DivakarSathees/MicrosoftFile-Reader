import os
import json
import requests
from datetime import datetime
from dotenv import load_dotenv
from msal import PublicClientApplication

# -----------------------------
# Load environment variables
# -----------------------------
load_dotenv()
CLIENT_ID = ""
TENANT_ID = ""
SCOPES = os.getenv("SCOPES1").split(",")  # ['Chat.ReadWrite', 'User.Read']

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# -----------------------------
# MSAL - Device Code Flow
# -----------------------------
app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

# Try to use cached account
accounts = app.get_accounts()
if accounts:
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
else:
    result = None

# If no token, start device code flow
if not result:
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception("Failed to create device flow")
    print(flow["message"])  # instruct user to authenticate
    result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    raise Exception("Failed to obtain token:", result.get("error_description"))

ACCESS_TOKEN = result["access_token"]
headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}
print(ACCESS_TOKEN)
print("✅ Access token acquired successfully!")

# -----------------------------
# Send Teams Message
# -----------------------------
CHAT_ID = "<your-chat-id>"  # e.g., 19:xxxx@thread.v2
GRAPH_URL = f"https://graph.microsoft.com/v1.0/chats/{CHAT_ID}/messages"

now = datetime.now().strftime("%Y-%m-%d %H:%M")
html_content = f"""
<b>Automated Message at {now}</b><br>
<p>This is a test message sent via MSAL device code flow.</p>
"""

payload = {
    "body": {
        "contentType": "html",
        "content": html_content
    }
}

response = requests.post(GRAPH_URL, headers=headers, json=payload)

if response.status_code == 201:
    print("✅ Message sent successfully!")
else:
    print(f"❌ Failed to send message: {response.status_code}")
    print(response.text)
