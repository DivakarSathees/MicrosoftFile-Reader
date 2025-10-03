import requests
import msal
import json
from datetime import datetime
import os
from dotenv import load_dotenv
load_dotenv()

# ----------------------------
# CONFIG
# ----------------------------
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
# CLIENT_SECRET = os.getenv("CLIENT_SECRET")  # Not needed for public client
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "http://localhost:8000/redirect"
# SCOPES = ["Chat.ReadWrite", "ChatMessage.Send", "User.Read"]
SCOPES = ["https://graph.microsoft.com/.default"]  # Required for app-only auth


# ----------------------------
# MSAL - INTERACTIVE LOGIN
# ----------------------------
app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

# Try silent token first
accounts = app.get_accounts()
print("Accounts:", accounts)
result = None
if accounts:
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result:
        print("Using cached account:", accounts[0]["username"])

# If no token, use interactive login (opens browser)
if not result:
    result = app.acquire_token_interactive(scopes=SCOPES)

if "access_token" not in result:
    raise Exception("Authentication failed:", result.get("error_description"))

token = result["access_token"]
print("Access token acquired ✅")
headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

# ----------------------------
# BUILD MESSAGE
# ----------------------------
chat_id = "19:3e85c11c1ba54f409d6a2829e19c9fe3@thread.v2"
now = datetime.now().strftime("%Y-%m-%d %H:%M")

html_content = f"""
<b>Reminder at {now}</b><br>
<table border='1' style='border-collapse:collapse'>
<tr><th>Time</th><th>Assessment</th><th>Main Resource</th><th>Additional Resource</th></tr>
<tr><td>10:00 AM</td><td>Dotnet FS — MS2 Mock 01</td><td>Dhayananth D</td><td>Hari Haran N R, Johnson Joy</td></tr>
<tr><td>4:00 PM</td><td>Java FS — MS2 Mock 02</td><td>Preethika</td><td>Kiruthika, Vikram, Pradeep</td></tr>
</table>
<p>📅 Plan accordingly.</p>
"""

payload = {
    "body": {
        "contentType": "html",
        "content": html_content
    }
}

# ----------------------------
# POST TO TEAMS CHAT
# ----------------------------
url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
resp = requests.post(url, headers=headers, json=payload)

if resp.status_code == 201:
    print("✅ Message sent successfully")
else:
    print("❌ Failed:", resp.status_code, resp.text)
