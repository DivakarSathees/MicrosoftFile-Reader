import requests
import json
from datetime import datetime
import os
from dotenv import load_dotenv
load_dotenv()


# -----------------------------
# CONFIG
# -----------------------------
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
CHAT_ID = "19:8b58d438bf7241828ee4009c477a13c6@thread.v2"
JSON_FILE = "filtered_today.json"

GRAPH_URL = f"https://graph.microsoft.com/beta/chats/{CHAT_ID}/messages"
GRAPH_USER_URL = "https://graph.microsoft.com/v1.0/users"

# -----------------------------
# HELPER FUNCTIONS
# -----------------------------

headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

def get_user_id_by_email_or_name(email_or_name):
    """Fetch the Teams user ID by email or display name using Graph API"""
    # Try to search by email first
    print(f"Searching for user: {email_or_name}")
    # params = {"$filter": f"mail eq '{email_or_name}'"}
    # response = requests.get(GRAPH_USER_URL, headers=headers, params=params)
    # if response.status_code == 200:
    #     data = response.json()
    #     if data.get("value"):
    #         return data["value"][0]["id"]

    # Fallback: search by displayName
    params = {"$filter": f"startswith(displayName, '{email_or_name}')"}

    response = requests.get(GRAPH_USER_URL, headers=headers, params=params)
    print(f"Response Status: {response.status_code}, Response Text: {response.text}")
    if response.status_code == 200:
        data = response.json()
        if data.get("value"):
            return data["value"][0]["id"]

    return None  # Not found

def build_mentions_html_and_list(resources):
    """
    For a list of resource names, return HTML content with <at> tags
    and a list of mention objects for Graph API.
    """
    html_parts = []
    mentions = []
    for idx, res in enumerate(resources):
        res = res.strip()
        if not res:
            continue
        user_id = get_user_id_by_email_or_name(res)
        print(f"Resource: {res}, User ID: {user_id}")
        if not user_id:
            html_parts.append(res)  # fallback: just text
            continue

        html_parts.append(f"<at id='{idx}'>{res}</at>")
        mentions.append({
            "id": idx,
            "mentionText": res,
            "mentioned": {
                "user": {
                    "id": user_id,
                    "displayName": res
                }
            }
        })
    return ", ".join(html_parts), mentions

# -----------------------------
# LOAD FILTERED JSON
# -----------------------------

with open(JSON_FILE, "r", encoding="utf-8") as f:
    rows = json.load(f)

# -----------------------------
# BUILD HTML TABLE CONTENT
# -----------------------------

html_content = "<b>This is an Automatically triggered message.</b><br><br>"
html_content += ""
html_content += "<p><b>📌 Reminder for Result Analysis - {}</b></p>".format(datetime.today().strftime("%b %d"))
html_content += ""
html_content += "<table border='1' style='border-collapse:collapse'>"
html_content += "<tr><th>Track</th><th>Time</th><th>Assessment</th><th>Main Resource</th><th>Additional Resource</th></tr>"

all_mentions = []

for row in rows:
    track = row.get("Track", "")
    time = row.get("Time", "")
    assessment = row.get("Test", "")
    main_res = row.get("Stark Corp Resource", "")
    add_res_str = row.get("Additional resource(If failures are high)", "")
    add_res_list = [r.strip() for r in add_res_str.split(",") if r.strip()]

    # Build mentions for additional resources
    add_res_html, mentions = build_mentions_html_and_list(add_res_list)
    all_mentions.extend(mentions)

    html_content += f"<tr><td>{track}</td><td>{time}</td><td>{assessment}</td><td>{main_res}</td><td>{add_res_html}</td></tr>"

html_content += "</table><p>📅 Plan accordingly.</p>"

# -----------------------------
# POST MESSAGE TO TEAMS
# -----------------------------

payload = {
    "body": {
        "contentType": "html",
        "content": html_content
    }
}

if all_mentions:
    payload["mentions"] = all_mentions

print("Payload to be sent:")    
print(json.dumps(payload, indent=2))

response = requests.post(GRAPH_URL, headers=headers, json=payload)

if response.status_code == 201:
    print("Message sent successfully!")
else:
    print(f"Failed to send message: {response.status_code}")
    print(response.text)
