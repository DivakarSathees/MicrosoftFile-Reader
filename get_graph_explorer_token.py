import os
import json
import msal
import requests

# === Configuration ===
CLIENT_ID = ""
TENANT_ID = ""

# Delegated Graph API permissions
SCOPES = [
    "Chat.ReadWrite",
    "ChatMessage.Send",
    "User.Read"
]

TOKEN_FILE = "token_cache.json"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Replace with your group chat ID
GROUP_CHAT_ID = ""

# === Token Cache Functions ===
def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_FILE):
        cache.deserialize(open(TOKEN_FILE, "r").read())
    return cache

def save_cache(cache):
    if cache.has_state_changed:
        with open(TOKEN_FILE, "w") as f:
            f.write(cache.serialize())

# === Acquire Access Token ===
def get_access_token():
    cache = load_cache()
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )

    # Try silent token acquisition
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            save_cache(cache)
            print("✅ Token acquired silently.")
            return result["access_token"]

    # Device code flow (no server/port needed)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise ValueError("Failed to create device flow")

    print(f"🔑 Go to {flow['verification_uri']} and enter code: {flow['user_code']}")
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        save_cache(cache)
        print("✅ Token acquired via device code login.")
        return result["access_token"]
    else:
        raise Exception(result.get("error_description"))

# === Send Message to Group Chat ===
def send_group_chat_message(access_token, message_text, mentions=None):
    url = f"https://graph.microsoft.com/v1.0/chats/{GROUP_CHAT_ID}/messages"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    body = {"body": {"content": message_text}}

    if mentions:
        body["mentions"] = mentions

    resp = requests.post(url, headers=headers, json=body)
    if resp.status_code == 201:
        print("✅ Message sent successfully!")
    else:
        print("❌ Failed to send message:", resp.status_code, resp.text)

# === Example Usage ===
if __name__ == "__main__":
    token = get_access_token()

    # Example message with @mention
    # Replace USER_ID with the actual Azure AD object ID of the user
    mentions = [
        {
            "id": 0,
            "mentionText": "Alice",
            "mentioned": {
                "user": {
                    "id": "USER_ID_GUID",
                    "displayName": "Alice"
                }
            }
        }
    ]
    message_content = "Hello <at id='0'>Alice</at>! This is a test message from Python."

    send_group_chat_message(token, message_content, mentions)
