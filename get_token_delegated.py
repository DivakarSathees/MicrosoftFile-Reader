

import os
import json
import msal

# === Configuration ===
CLIENT_ID = ""
TENANT_ID = ""

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = [
    "User.Read",
    "Chat.ReadWrite",
    "ChatMessage.Send",
    "Files.ReadWrite.All"
]


TOKEN_FILE = "token_cache.json"
REDIRECT_URI = "http://localhost:8000"


# === Token Cache Helpers ===
def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache


def save_cache(cache):
    if cache.has_state_changed:
        with open(TOKEN_FILE, "w") as f:
            f.write(cache.serialize())


# === Token Retrieval ===
def get_access_token():
    cache = load_cache()
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )

    # 1️⃣ Try silent token
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result:
            print("✅ Token refreshed silently.")
            save_cache(cache)
            return result["access_token"]

    # 2️⃣ Fallback: Interactive login via browser
    print("🌐 No cached token found. Launching browser for login...")
    result = app.acquire_token_interactive(
        scopes=SCOPES,
        # redirect_uri=REDIRECT_URI
    )

    if "access_token" in result:
        save_cache(cache)
        print("✅ Access token acquired successfully.")
        return result["access_token"]
    else:
        raise Exception(f"Login failed: {result.get('error_description')}")


# === Main ===
if __name__ == "__main__":
    token = get_access_token()
    with open("access_token.txt", "w") as f:
        f.write(token)
    print("🔑 Access token saved to access_token.txt")
