import os
import requests
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REFRESH_TOKEN = os.getenv("REFRESH_TOKEN")

token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

data = {
    "grant_type": "refresh_token",
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "refresh_token": REFRESH_TOKEN,
    "scope": "offline_access Chat.ReadWrite User.Read Files.Read.All"
}

resp = requests.post(token_url, data=data)
print(resp)


# import requests
# import webbrowser
# import os
# from urllib.parse import urlencode
# from dotenv import load_dotenv

# load_dotenv()

# TENANT_ID = os.getenv("TENANT_ID")
# CLIENT_ID = os.getenv("CLIENT_ID")
# REDIRECT_URI = "http://localhost:5000"

# # ✅ Correct scopes — no '.default'
# SCOPES = "offline_access Chat.ReadWrite User.Read Files.Read.All"

# # Step 1: Ask user to log in
# params = {
#     "client_id": CLIENT_ID,
#     "response_type": "code",
#     "redirect_uri": REDIRECT_URI,
#     "response_mode": "query",
#     "scope": SCOPES,
# }
# url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize?{urlencode(params)}"
# print("Open this URL to sign in:")
# print(url)
# webbrowser.open(url)

# # Step 2: After login, paste the "code" from redirect URL
# auth_code = input("Paste the 'code' from the redirected URL: ").strip()

# # Step 3: Exchange for tokens
# token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
# data = {
#     "client_id": CLIENT_ID,
#     "grant_type": "authorization_code",
#     "code": auth_code,
#     "redirect_uri": REDIRECT_URI,
#     "scope": SCOPES,
# }
# response = requests.post(token_url, data=data)

# print(response.json())
