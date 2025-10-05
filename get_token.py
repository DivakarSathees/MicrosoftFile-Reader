import requests
import os
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

data = {
    "grant_type": "client_credentials",
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "scope": "https://graph.microsoft.com/.default"
}

response = requests.post(TOKEN_URL, data=data)
token_data = response.json()

if "access_token" in token_data:
    print("✅ Token fetched successfully!")
    print(token_data["access_token"][:200] + "...")
    with open("access_token.txt", "w") as f:
        f.write(token_data["access_token"])
else:
    print("❌ Error fetching token:", token_data)
