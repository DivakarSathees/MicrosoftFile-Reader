from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from datetime import datetime, timedelta
import requests
import json
import os
import logging

app = FastAPI(title="Teams Daily Excel Processor with Frontend")

# -----------------------------
# CONFIG
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "filtered_today.json")
MESSAGE_PREVIEW_FILE = os.path.join(BASE_DIR, "message_preview.html")
LOG_FILE = os.path.join(BASE_DIR, "daily_excel_processor.log")

WORKBOOK_ID = "01MTOP4SZMUAJ5IMFDEBB3QJTXXJ4HLDEC"
DRIVE_ID = "b!qFZ6QUc4VEi3I3L4f0VhuBtcxM__oQVCufiwqYslGCloiNIthSJwSJvZ6FviPX5U"
WORKBOOK_URL = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{WORKBOOK_ID}/workbook/worksheets"

CHAT_ID = "19:3e85c11c1ba54f409d6a2829e19c9fe3@thread.v2"  # testing chat
GRAPH_CHAT_URL = f"https://graph.microsoft.com/beta/chats/{CHAT_ID}/messages"
GRAPH_USER_URL = "https://graph.microsoft.com/v1.0/users"

# -----------------------------
# LOGGING
# -----------------------------
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# -----------------------------
# FRONTEND TEMPLATES
# -----------------------------
templates = Jinja2Templates(directory="templates")

# -----------------------------
# HELPER FUNCTIONS (same as before)
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

def get_user_id_by_email_or_name(email_or_name, access_token):
    name_email_map = {
        "dhayananth": "dhayananth.d@iamneo.ai",
        "pradeep": "pradeep.s@iamneo.ai",
        "sriram": "sriramkumar.ramesh@iamneo.ai",
        "sri ram": "sriramkumar.ramesh@iamneo.ai",
        "hariharan": "nrhariharan@iamneo.ai",
        "jai chandru": "jaichandru.ss@iamneo.ai"
    }
    key = email_or_name.strip().lower()
    email = name_email_map.get(key, email_or_name)

    params = {"$filter": f"mail eq '{email}'"}
    response = requests.get(GRAPH_USER_URL, headers={"Authorization": f"Bearer {access_token}"}, params=params)
    if response.status_code == 200:
        data = response.json()
        if data.get("value") and len(data["value"]) == 1:
            return data["value"][0]["id"]
        params = {"$filter": f"startswith(displayName, '{email_or_name}')"}
        response = requests.get(GRAPH_USER_URL, headers={"Authorization": f"Bearer {access_token}"}, params=params)
        if response.status_code == 200:
            data = response.json()
            if data.get("value") and len(data["value"]) == 1:
                return data["value"][0]["id"]
    return None

def build_mentions_html_and_list(resources, access_token, start_idx=0):
    html_parts = []
    mentions = []
    idx = start_idx
    for res in resources:
        res = res.strip()
        if not res:
            continue
        user_id = get_user_id_by_email_or_name(res, access_token)
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

def process_excel_and_post(access_token):
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

    resp = requests.get(WORKBOOK_URL, headers=headers)
    if resp.status_code != 200:
        return f"Error fetching worksheets: {resp.text}"

    worksheets = resp.json().get("value", [])
    target_sheets = [s for s in worksheets if s.get("visibility", "Visible") == "Visible" and s.get("name", "").endswith("Batches")]

    if not target_sheets:
        return "No visible sheets ending with 'Batches' found."

    today = datetime.today().date()
    all_filtered_rows = []

    for sheet in target_sheets:
        sheet_name = sheet["name"]
        GRAPH_URL = f"{WORKBOOK_URL}/{sheet_name}/usedRange(valuesOnly=true)"
        response = requests.get(GRAPH_URL, headers=headers)
        if response.status_code != 200:
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

    # Build HTML
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

        main_html, mentions_main, mention_idx = build_mentions_html_and_list(main_res_list, access_token, mention_idx)
        add_html, mentions_add, mention_idx = build_mentions_html_and_list(add_res_list, access_token, mention_idx)
        all_mentions.extend(mentions_main + mentions_add)

        html_content += f"<tr><td>{track}</td><td>{time}</td><td>{assessment}</td><td>{main_html}</td><td>{add_html}</td></tr>"

    html_content += "</table><br><p>📅 Plan accordingly.</p>"

    with open(MESSAGE_PREVIEW_FILE, "w", encoding="utf-8") as f:
        f.write(html_content)

    payload = {"body": {"contentType": "html", "content": html_content}}
    if all_mentions:
        payload["mentions"] = all_mentions

    response = requests.post(GRAPH_CHAT_URL, headers=headers, json=payload)
    if response.status_code == 201:
        return "Message sent successfully!"
    else:
        return f"Failed to send message: {response.text}"

# -----------------------------
# FRONTEND ROUTE
# -----------------------------
@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/process", response_class=HTMLResponse)
def process(request: Request, access_token: str = Form(...)):
    result = process_excel_and_post(access_token)
    return templates.TemplateResponse("index.html", {"request": request, "result": result})

@app.post("/webhook/lifecycle")
async def lifecycle_handler(request: Request):
    data = await request.json()
    logging.info(f"Lifecycle event: {data}")
    return {"status": "ok"}

