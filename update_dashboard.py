"""
One-off script to add a colour legend + polish to the Dashboard tab.
Run this once; future sync_to_sheet.py runs maintain it automatically.
"""

import os
import re
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime

SHEET_ID  = "1VptcDahrMKwqgo3wWmxpoVzXEyY5je2U73mvx4rOYnU"
CREDS_FILE = (
    os.environ.get("GOOGLE_CREDS_FILE")
    or r"C:\Users\Sahil Gaur\Downloads\replacement-and-refund-8a7f8939a062.json"
)
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

LEGEND = [
    # label,          hex-ish RGB (0-1 scale),                          meaning
    ("Refund",       {"red": 0.23, "green": 0.47, "blue": 0.85}, "Order cancelled & refund raised"),
    ("Returned",     {"red": 0.96, "green": 0.60, "blue": 0.07}, "Delivered but customer returned"),
    ("RTO",          {"red": 0.83, "green": 0.18, "blue": 0.18}, "Return to Origin — undeliverable"),
    ("Undelivered",  {"red": 0.42, "green": 0.65, "blue": 0.31}, "In transit, not yet delivered"),
    ("Replacement",  {"red": 0.60, "green": 0.30, "blue": 0.70}, "Replacement order raised"),
]


def add_legend_and_polish():
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    gc      = gspread.authorize(creds)
    sh      = gc.open_by_key(SHEET_ID)
    service = build("sheets", "v4", credentials=creds)

    # Get Dashboard sheet id
    try:
        ws_dash = sh.worksheet("Dashboard")
    except gspread.exceptions.WorksheetNotFound:
        ws_dash = sh.add_worksheet(title="Dashboard", rows=60, cols=20)

    dash_id = ws_dash.id

    # ── Write legend text into cells (col P onward, row 2) ───────────────────
    legend_start_col = 15   # column P (0-based)
    legend_start_row = 1    # row 2 (0-based)

    # Header
    ws_dash.update(
        values=[["Order Type", "Meaning"]],
        range_name="P2",
        value_input_option="USER_ENTERED"
    )
    for i, (label, _, meaning) in enumerate(LEGEND):
        ws_dash.update(
            values=[[label, meaning]],
            range_name=f"P{3 + i}",
            value_input_option="USER_ENTERED"
        )

    # ── Format legend via batchUpdate ─────────────────────────────────────────
    requests = []

    # Title row bold + dark bg
    requests.append({"repeatCell": {
        "range": {
            "sheetId": dash_id,
            "startRowIndex": 1, "endRowIndex": 2,
            "startColumnIndex": legend_start_col, "endColumnIndex": legend_start_col + 2,
        },
        "cell": {"userEnteredFormat": {
            "backgroundColor": {"red": 0.13, "green": 0.13, "blue": 0.13},
            "textFormat": {
                "bold": True,
                "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                "fontSize": 10,
            },
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat)",
    }})

    # Legend box title above
    requests.append({"repeatCell": {
        "range": {
            "sheetId": dash_id,
            "startRowIndex": 0, "endRowIndex": 1,
            "startColumnIndex": legend_start_col, "endColumnIndex": legend_start_col + 2,
        },
        "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True, "fontSize": 12},
        }},
        "fields": "userEnteredFormat.textFormat",
    }})

    # Colour each label row with its chart colour
    for i, (label, color, _) in enumerate(LEGEND):
        row_idx = 2 + i   # 0-based, row 3 onward
        # Colour swatch cell (col P)
        requests.append({"repeatCell": {
            "range": {
                "sheetId": dash_id,
                "startRowIndex": row_idx, "endRowIndex": row_idx + 1,
                "startColumnIndex": legend_start_col, "endColumnIndex": legend_start_col + 1,
            },
            "cell": {"userEnteredFormat": {
                "backgroundColor": color,
                "textFormat": {
                    "bold": True,
                    "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                },
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }})
        # Meaning cell (col Q) — light tint of same colour
        tint = {k: min(1.0, v * 0.25 + 0.75) for k, v in color.items()}
        requests.append({"repeatCell": {
            "range": {
                "sheetId": dash_id,
                "startRowIndex": row_idx, "endRowIndex": row_idx + 1,
                "startColumnIndex": legend_start_col + 1, "endColumnIndex": legend_start_col + 2,
            },
            "cell": {"userEnteredFormat": {
                "backgroundColor": tint,
                "textFormat": {"fontSize": 10},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }})

    # Border around entire legend block
    requests.append({"updateBorders": {
        "range": {
            "sheetId": dash_id,
            "startRowIndex": 1, "endRowIndex": 2 + len(LEGEND),
            "startColumnIndex": legend_start_col, "endColumnIndex": legend_start_col + 2,
        },
        "top":    {"style": "SOLID_MEDIUM", "color": {"red": 0.4, "green": 0.4, "blue": 0.4}},
        "bottom": {"style": "SOLID_MEDIUM", "color": {"red": 0.4, "green": 0.4, "blue": 0.4}},
        "left":   {"style": "SOLID_MEDIUM", "color": {"red": 0.4, "green": 0.4, "blue": 0.4}},
        "right":  {"style": "SOLID_MEDIUM", "color": {"red": 0.4, "green": 0.4, "blue": 0.4}},
        "innerHorizontal": {"style": "SOLID", "color": {"red": 0.8, "green": 0.8, "blue": 0.8}},
    }})

    # Auto-resize legend columns
    requests.append({"autoResizeDimensions": {
        "dimensions": {
            "sheetId": dash_id,
            "dimension": "COLUMNS",
            "startIndex": legend_start_col,
            "endIndex": legend_start_col + 2,
        }
    }})

    # Make Dashboard the first tab
    month_pattern = re.compile(
        r"^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$"
    )
    worksheets = sh.worksheets()

    def tab_sort_key(ws):
        if ws.title == "Dashboard":  return (0, "")
        if ws.title == "All Data":   return (2, "")
        if ws.title == "_ChartData": return (3, "")
        if month_pattern.match(ws.title):
            try:    return (1, datetime.strptime(ws.title, "%B %Y").strftime("%Y%m"))
            except: return (1, ws.title)
        return (4, ws.title)

    for idx, ws in enumerate(sorted(worksheets, key=tab_sort_key)):
        requests.append({"updateSheetProperties": {
            "properties": {"sheetId": ws.id, "index": idx},
            "fields": "index",
        }})

    # Hide _ChartData
    cd_ws = next((w for w in worksheets if w.title == "_ChartData"), None)
    if cd_ws:
        requests.append({"updateSheetProperties": {
            "properties": {"sheetId": cd_ws.id, "hidden": True},
            "fields": "hidden",
        }})

    service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": requests},
    ).execute()

    print("Done — legend added, Dashboard moved to first tab, _ChartData hidden.")
    print(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")


if __name__ == "__main__":
    add_legend_and_polish()
