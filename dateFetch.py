# import requests
# import json
# from datetime import datetime, timedelta
# # from logincred import get_access_token   # import function from auth.py
# # load from env
# import os
# from dotenv import load_dotenv
# load_dotenv()



# # -----------------------------
# # CONFIG
# # -----------------------------
# # token = get_access_token()  # get fresh token

# # ACCESS_TOKEN = get_access_token()
# ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
# print("Access Token:", ACCESS_TOKEN)

# OUTPUT_FILE = "filtered_today.json"

# def excel_date_to_datetime(excel_serial):
#     """Convert Excel serial date to datetime."""
#     if isinstance(excel_serial, (int, float)):
#         return datetime(1899, 12, 30) + timedelta(days=excel_serial)
#     return None

# def excel_time_to_str(excel_time):
#     """Convert Excel fraction of day to HH:MM string, or return as-is if already a string."""
#     if isinstance(excel_time, (int, float)):
#         hours = int(excel_time * 24)
#         minutes = int(round((excel_time * 24 - hours) * 60))
#         return f"{hours:02d}:{minutes:02d}"
#     elif isinstance(excel_time, str):
#         return excel_time.strip()
#     return ""

# # -----------------------------
# # FETCH ALL WORKSHEETS
# # -----------------------------

# WORKBOOK_URL = "https://graph.microsoft.com/v1.0/drives/b!qFZ6QUc4VEi3I3L4f0VhuBtcxM__oQVCufiwqYslGCloiNIthSJwSJvZ6FviPX5U/items/01MTOP4SZMUAJ5IMFDEBB3QJTXXJ4HLDEC/workbook/worksheets"

# headers = {
#     "Authorization": f"Bearer {ACCESS_TOKEN}",
#     "Content-Type": "application/json"
# }

# resp = requests.get(WORKBOOK_URL, headers=headers)
# if resp.status_code != 200:
#     raise Exception(f"Error fetching worksheets: {resp.status_code}, {resp.text}")

# worksheets = resp.json().get("value", [])

# # Filter visible sheets ending with "Batches"
# target_sheets = [
#     sheet for sheet in worksheets
#     if sheet.get("visibility", "Visible") == "Visible" and sheet.get("name", "").endswith("Batches")
# ]

# if not target_sheets:
#     print("No visible sheets ending with 'Batches' found.")
#     exit()

# # -----------------------------
# # PROCESS EACH SHEET
# # -----------------------------

# today = datetime.today().date()
# all_filtered_rows = []

# for sheet in target_sheets:
#     sheet_name = sheet["name"]
#     print(f"Processing sheet: {sheet_name}")

#     GRAPH_URL = f"https://graph.microsoft.com/v1.0/drives/b!qFZ6QUc4VEi3I3L4f0VhuBtcxM__oQVCufiwqYslGCloiNIthSJwSJvZ6FviPX5U/items/01MTOP4SZMUAJ5IMFDEBB3QJTXXJ4HLDEC/workbook/worksheets/{sheet_name}/usedRange(valuesOnly=true)"
    
#     response = requests.get(GRAPH_URL, headers=headers)
#     if response.status_code != 200:
#         print(f"Failed to fetch data for {sheet_name}: {response.status_code}, {response.text}")
#         continue
    
#     data = response.json()
#     values = data.get("values", [])
#     if not values:
#         continue

#     # Fill merged cells
#     header = values[0]
#     date_col_index = header.index("Date")
#     track_col_index = header.index("Track")
#     time_col_index = header.index("Time")

#     filled_rows = []
#     last_track_value = None

#     for row in values[1:]:
#         if row[track_col_index]:
#             last_track_value = row[track_col_index]
#         else:
#             row[track_col_index] = last_track_value
#         filled_rows.append(row)

#     # Filter by today's date
#     for row in filled_rows:
#         excel_date = row[date_col_index]
#         row_date = excel_date_to_datetime(excel_date)
#         if row_date and row_date.date() == today:
#             row_dict = {}
#             for i in range(len(header)):
#                 key = header[i]
#                 value = row[i]
#                 if i == date_col_index:
#                     value = row_date.strftime("%Y-%m-%d")
#                 elif i == time_col_index:
#                     value = excel_time_to_str(value)
#                 row_dict[key] = value
#             all_filtered_rows.append(row_dict)

# # -----------------------------
# # SAVE TO JSON
# # -----------------------------

# with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
#     json.dump(all_filtered_rows, f, ensure_ascii=False, indent=4)

# print(f"Processed {len(target_sheets)} sheet(s). Filtered rows saved to {OUTPUT_FILE}.")


# # CHAT_ID = "19:8b58d438bf7241828ee4009c477a13c6@thread.v2" # demo test chat
# # CHAT_ID = "19:3e85c11c1ba54f409d6a2829e19c9fe3@thread.v2" # demo test chat
# CHAT_ID = "19:db4e52f07e324ac984588ef6d2346d93@thread.v2" # AI Project Generatoe chat
# JSON_FILE = "filtered_today.json"

# GRAPH_URL = f"https://graph.microsoft.com/beta/chats/{CHAT_ID}/messages"
# GRAPH_USER_URL = "https://graph.microsoft.com/v1.0/users"


# def get_user_id_by_email_or_name(email_or_name):
#     """Fetch the Teams user ID by email or display name using Graph API"""
#     # Dynamic mapping: name to email
#     name_email_map = {
#         "dhayananth": "dhayananth.d@iamneo.ai",
#         "pradeep": "pradeep.s@iamneo.ai",
#         "sriram": "sriramkumar.ramesh@iamneo.ai",
#         "sri ram": "sriramkumar.ramesh@iamneo.ai"
#     }

#     # Normalize input for case-insensitive matching
#     key = email_or_name.strip().lower()
#     email = name_email_map.get(key, email_or_name)
#     print(f"Searching for user: {email} (original: {email_or_name})")

#     # Try searching by email first
#     params = {"$filter": f"mail eq '{email}'"}
#     response = requests.get(GRAPH_USER_URL, headers=headers, params=params)
#     print(f"Response Status: {response.status_code}, Response Text: {response.text}")
#     if response.status_code == 200:
#         data = response.json()
#         if data.get("value") and len(data["value"]) == 1:
#             return data["value"][0]["id"]
#         # If not found by email, try by displayName
#         params = {"$filter": f"startswith(displayName, '{email_or_name}')"}
#         response = requests.get(GRAPH_USER_URL, headers=headers, params=params)
#         print(f"DisplayName search Status: {response.status_code}, Text: {response.text}")
#         if response.status_code == 200:
#             data = response.json()
#             if not data.get("value"):
#                 print("No users found matching the filter.")
#                 return None
#             if len(data["value"]) > 1:
#                 print("Multiple users found, returning None for safety.")
#                 return None
#             return data["value"][0]["id"]
#     return None  # Not found

# def build_mentions_html_and_list(resources, start_idx=0):
#     """
#     For a list of resource names, return HTML content with <at> tags
#     and a list of mention objects for Graph API.
#     start_idx ensures unique mention IDs across rows.
#     """
#     html_parts = []
#     mentions = []
#     idx = start_idx
#     for res in resources:
#         res = res.strip()
#         if not res:
#             continue
#         user_id = get_user_id_by_email_or_name(res)
#         print(f"Resource: {res}, User ID: {user_id}")
#         if not user_id:
#             html_parts.append(res)  # fallback: just text
#             continue

#         html_parts.append(f"<at id='{idx}'>{res}</at>")
#         mentions.append({
#             "id": idx,
#             "mentionText": res,
#             "mentioned": {
#                 "user": {
#                     "id": user_id,
#                     "displayName": res
#                 }
#             }
#         })
#         idx += 1
#     return ", ".join(html_parts), mentions, idx


# with open(JSON_FILE, "r", encoding="utf-8") as f:
#     rows = json.load(f)

# # -----------------------------
# # BUILD HTML TABLE CONTENT
# # -----------------------------

# html_content = "<b>This is an Automatically triggered message at {}</b><br>".format(datetime.now().strftime("%Y-%m-%d %H:%M"))
# html_content += "<p><b>📌 Reminder for Result Analysis - {}</b></p><br>".format(datetime.today().strftime("%b %d"))
# html_content += "<table border='1' style='border-collapse:collapse'>"
# html_content += "<tr><th>Track</th><th>Time</th><th>Assessment</th><th>Main Resource</th><th>Additional Resource</th></tr>"

# all_mentions = []
# mention_idx = 0

# for row in rows:
#     track = row.get("Track", "")
#     time = row.get("Time", "")
#     assessment = row.get("Test", "")
#     main_res_str = row.get("Stark Corp Resource", "")
#     add_res_str = row.get("Additional resource(If failures are high)", "")

#     # Ensure both are lists for uniform handling
#     main_res_list = [r.strip() for r in main_res_str.split(",") if r.strip()]
#     add_res_list = [r.strip() for r in add_res_str.split(",") if r.strip()]

#     # Mentions for main resource(s)
#     main_res_html, mentions_main, mention_idx = build_mentions_html_and_list(main_res_list, mention_idx)
#     all_mentions.extend(mentions_main)

#     # Mentions for additional resources
#     add_res_html, mentions_add, mention_idx = build_mentions_html_and_list(add_res_list, mention_idx)
#     all_mentions.extend(mentions_add)

#     html_content += f"<tr><td>{track}</td><td>{time}</td><td>{assessment}</td><td>{main_res_html}</td><td>{add_res_html}</td></tr>"

# html_content += "</table><br><p>📅 Plan accordingly.</p>"

# # -----------------------------
# # POST MESSAGE TO TEAMS
# # -----------------------------

# # display the html content in seperate html file for verification
# with open("message_preview.html", "w", encoding="utf-8") as f:
#     f.write(html_content)



# payload = {
#     "body": {
#         "contentType": "html",
#         "content": html_content
#     }
# }

# if all_mentions:
#     payload["mentions"] = all_mentions


# # print("Payload to be sent:")
# # print(json.dumps(payload, indent=2))

# response = requests.post(GRAPH_URL, headers=headers, json=payload)

# if response.status_code == 201:
#     print("Message sent successfully!")
# else:
#     print(f"Failed to send message: {response.status_code}")
#     print(response.text)


#!/usr/bin/env python3
import requests
import json
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv
import logging
import sys

# -----------------------------
# LOAD ENVIRONMENT VARIABLES
# -----------------------------
load_dotenv()
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
if not ACCESS_TOKEN:
    raise Exception("ACCESS_TOKEN not set in environment variables")

# -----------------------------
# CONFIG
# -----------------------------
OUTPUT_FILE = "filtered_today.json"  # absolute path
MESSAGE_PREVIEW_FILE = "message_preview.html"

WORKBOOK_ID = "01MTOP4SZMUAJ5IMFDEBB3QJTXXJ4HLDEC"
DRIVE_ID = "b!qFZ6QUc4VEi3I3L4f0VhuBtcxM__oQVCufiwqYslGCloiNIthSJwSJvZ6FviPX5U"
WORKBOOK_URL = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{WORKBOOK_ID}/workbook/worksheets"

CHAT_ID = "19:db4e52f07e324ac984588ef6d2346d93@thread.v2"
GRAPH_CHAT_URL = f"https://graph.microsoft.com/beta/chats/{CHAT_ID}/messages"
GRAPH_USER_URL = "https://graph.microsoft.com/v1.0/users"

# -----------------------------
# LOGGING
# -----------------------------
logging.basicConfig(
    filename="daily_excel_processor.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def excel_date_to_datetime(excel_serial):
    if isinstance(excel_serial, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=excel_serial)
    return None

def excel_time_to_str(excel_time):
    if isinstance(excel_time, (int, float)):
        hours = int(excel_time * 24)
        minutes = int(round((excel_time * 24 - hours) * 60))
        return f"{hours:02d}:{minutes:02d}"
    elif isinstance(excel_time, str):
        return excel_time.strip()
    return ""

def get_user_id_by_email_or_name(email_or_name):
    """Fetch Teams user ID from Graph API by email or name."""
    name_email_map = {
        "dhayananth": "dhayananth.d@iamneo.ai",
        "pradeep": "pradeep.s@iamneo.ai",
        "sriram": "sriramkumar.ramesh@iamneo.ai",
        "sri ram": "sriramkumar.ramesh@iamneo.ai"
    }
    key = email_or_name.strip().lower()
    email = name_email_map.get(key, email_or_name)

    params = {"$filter": f"mail eq '{email}'"}
    response = requests.get(GRAPH_USER_URL, headers={"Authorization": f"Bearer {ACCESS_TOKEN}"}, params=params)
    if response.status_code == 200:
        data = response.json()
        if data.get("value") and len(data["value"]) == 1:
            return data["value"][0]["id"]
        # fallback: search by displayName
        params = {"$filter": f"startswith(displayName, '{email_or_name}')"}
        response = requests.get(GRAPH_USER_URL, headers={"Authorization": f"Bearer {ACCESS_TOKEN}"}, params=params)
        if response.status_code == 200:
            data = response.json()
            if data.get("value") and len(data["value"]) == 1:
                return data["value"][0]["id"]
    return None

def build_mentions_html_and_list(resources, start_idx=0):
    """Build HTML <at> mentions and Graph mention objects"""
    html_parts = []
    mentions = []
    idx = start_idx
    for res in resources:
        res = res.strip()
        if not res:
            continue
        user_id = get_user_id_by_email_or_name(res)
        if not user_id:
            html_parts.append(res)
            continue
        html_parts.append(f"<at id='{idx}'>{res}</at>")
        mentions.append({
            "id": idx,
            "mentionText": res,
            "mentioned": {"user": {"id": user_id, "displayName": res}}
        })
        idx += 1
    return ", ".join(html_parts), mentions, idx

# -----------------------------
# MAIN FUNCTION
# -----------------------------
def main():
    try:
        headers = {
            "Authorization": f"Bearer {ACCESS_TOKEN}",
            "Content-Type": "application/json"
        }

        # FETCH ALL WORKSHEETS
        resp = requests.get(WORKBOOK_URL, headers=headers)
        if resp.status_code != 200:
            logging.error(f"Error fetching worksheets: {resp.status_code}, {resp.text}")
            return
        worksheets = resp.json().get("value", [])
        target_sheets = [s for s in worksheets if s.get("visibility", "Visible") == "Visible" and s.get("name", "").endswith("Batches")]

        if not target_sheets:
            logging.info("No visible sheets ending with 'Batches' found.")
            return

        today = datetime.today().date()
        all_filtered_rows = []

        # PROCESS EACH SHEET
        for sheet in target_sheets:
            sheet_name = sheet["name"]
            logging.info(f"Processing sheet: {sheet_name}")
            GRAPH_URL = f"{WORKBOOK_URL}/{sheet_name}/usedRange(valuesOnly=true)"
            response = requests.get(GRAPH_URL, headers=headers)
            if response.status_code != 200:
                logging.warning(f"Failed fetching sheet {sheet_name}: {response.status_code}, {response.text}")
                continue

            values = response.json().get("values", [])
            if not values:
                continue

            header = values[0]
            date_col_index = header.index("Date")
            track_col_index = header.index("Track")
            time_col_index = header.index("Time")

            filled_rows = []
            last_track_value = None
            for row in values[1:]:
                if row[track_col_index]:
                    last_track_value = row[track_col_index]
                else:
                    row[track_col_index] = last_track_value
                filled_rows.append(row)

            # FILTER BY TODAY
            for row in filled_rows:
                excel_date = row[date_col_index]
                row_date = excel_date_to_datetime(excel_date)
                if row_date and row_date.date() == today:
                    row_dict = {}
                    for i in range(len(header)):
                        key = header[i]
                        value = row[i]
                        if i == date_col_index:
                            value = row_date.strftime("%Y-%m-%d")
                        elif i == time_col_index:
                            value = excel_time_to_str(value)
                        row_dict[key] = value
                    all_filtered_rows.append(row_dict)

        # SAVE JSON
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(all_filtered_rows, f, ensure_ascii=False, indent=4)

        # -----------------------------
        # BUILD HTML MESSAGE
        # -----------------------------
        html_content = "<b>Automatically triggered message at {}</b><br>".format(datetime.now().strftime("%Y-%m-%d %H:%M"))
        html_content += "<p><b>📌 Reminder for Result Analysis - {}</b></p><br>".format(datetime.today().strftime("%b %d"))
        html_content += "<table border='1' style='border-collapse:collapse'>"
        html_content += "<tr><th>Track</th><th>Time</th><th>Assessment</th><th>Main Resource</th><th>Additional Resource</th></tr>"

        all_mentions = []
        mention_idx = 0
        for row in all_filtered_rows:
            track = row.get("Track", "")
            time = row.get("Time", "")
            assessment = row.get("Test", "")
            main_res_list = [r.strip() for r in row.get("Stark Corp Resource", "").split(",") if r.strip()]
            add_res_list = [r.strip() for r in row.get("Additional resource(If failures are high)", "").split(",") if r.strip()]

            main_html, mentions_main, mention_idx = build_mentions_html_and_list(main_res_list, mention_idx)
            add_html, mentions_add, mention_idx = build_mentions_html_and_list(add_res_list, mention_idx)
            all_mentions.extend(mentions_main + mentions_add)

            html_content += f"<tr><td>{track}</td><td>{time}</td><td>{assessment}</td><td>{main_html}</td><td>{add_html}</td></tr>"

        html_content += "</table><br><p>📅 Plan accordingly.</p>"

        # SAVE PREVIEW HTML
        with open(MESSAGE_PREVIEW_FILE, "w", encoding="utf-8") as f:
            f.write(html_content)

        # -----------------------------
        # POST TO TEAMS
        # -----------------------------
        payload = {"body": {"contentType": "html", "content": html_content}}
        if all_mentions:
            payload["mentions"] = all_mentions

        response = requests.post(GRAPH_CHAT_URL, headers=headers, json=payload)
        if response.status_code == 201:
            logging.info("Message sent successfully!")
            print("Message sent successfully!")
        else:
            logging.error(f"Failed to send message: {response.status_code}, {response.text}")
            print(f"Failed to send message: {response.status_code}, response logged")

    except Exception as e:
        logging.exception(f"Error in main process: {str(e)}")
        print(f"Error occurred: {e}")

# -----------------------------
# ENTRY POINT
# -----------------------------
if __name__ == "__main__":
    main()
