import os
import json
import msal
import requests

# === Configuration ===
CLIENT_ID = "YOUR_CLIENT_ID"
TENANT_ID = "YOUR_TENANT_ID"
SCOPES = [
    "User.Read",
    "Mail.ReadWrite",
    "Chat.ReadWrite",
    "Files.ReadWrite.All"
]

TOKEN_FILE = "token_cache.json"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"


def load_cache():
    if os.path.exists(TOKEN_FILE):
        cache = msal.SerializableTokenCache()
        cache.deserialize(open(TOKEN_FILE, "r").read())
        return cache
    return msal.SerializableTokenCache()


def save_cache(cache):
    if cache.has_state_changed:
        with open(TOKEN_FILE, "w") as f:
            f.write(cache.serialize())


def get_access_token():
    cache = load_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    # Try silent token refresh
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result:
            save_cache(cache)
            return result["access_token"]

    # Otherwise, use device code flow
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception("Failed to create device flow")

    print(f"🔐 Go to {flow['verification_uri']} and enter code: {flow['user_code']}")
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        save_cache(cache)
        return result["access_token"]
    else:
        raise Exception(result.get("error_description"))


if __name__ == "__main__":
    token = get_access_token()
    print("✅ Access token acquired successfully.")
    with open("access_token.txt", "w") as f:
        f.write(token)
