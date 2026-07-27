"""
Swiss Beauty – Replacement / Refund / RTO monthly sheet sync.

Usage:
  python sync_to_sheet.py              # dumps the CURRENT calendar month
  python sync_to_sheet.py 2026-07      # dumps a specific YYYY-MM

Flow each run:
  1. Fetch full calendar month from Shopify
  2. Write / overwrite "Month YYYY" tab
  3. Rebuild "All Data" tab (all months combined)
  4. Create Dashboard charts if they don't exist yet (auto-update thereafter)
"""

import sys
import os
import re
import calendar
import requests
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime, timezone, timedelta
from collections import defaultdict

# ── Config ────────────────────────────────────────────────────────────────────
SHOPIFY_STORE   = "swiss-beauty-dev.myshopify.com"
SHOPIFY_TOKEN   = os.environ.get("SHOPIFY_TOKEN", "")
SHOPIFY_API_VER = "2024-01"
SHEET_ID        = "1VptcDahrMKwqgo3wWmxpoVzXEyY5je2U73mvx4rOYnU"
CREDS_FILE      = (
    os.environ.get("GOOGLE_CREDS_FILE")
    or r"C:\Users\Sahil Gaur\Downloads\replacement-and-refund-8a7f8939a062.json"
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

TARGET_TAGS = {
    "Refund_initiated", "Refund_Initiated",
    "Refund_credited",
    "Returned", "RTO", "Undelivered", "Replacement",
    "Partial_replacement", "Full_replacement", "refund_given",
}

MONTH_HEADERS = [
    "Order Number", "Order Date (IST)", "Order Type", "Cancel Reason",
    "Financial Status", "Fulfillment Status",
    "Actions Taken", "Notes",                               # ← new cols (index 6 & 7)
    "Customer Name", "Customer Phone", "City", "State",
    "Payment Method", "Order Value (₹)", "Refunded Amount (₹)", "Refund Status",
    "Product 1", "SKU 1", "Qty 1",
    "Product 2", "SKU 2", "Qty 2",
    "Product 3", "SKU 3", "Qty 3",
]

ALL_DATA_HEADERS = ["Month"] + MONTH_HEADERS

GQL_QUERY = """
query FetchOrders($query: String!, $after: String) {
  orders(first: 100, query: $query, after: $after, sortKey: CREATED_AT, reverse: true) {
    pageInfo { hasNextPage endCursor }
    edges {
      node {
        name createdAt tags
        displayFinancialStatus displayFulfillmentStatus
        totalPriceSet    { shopMoney { amount } }
        totalRefundedSet { shopMoney { amount } }
        customer { firstName lastName phone }
        shippingAddress { city province }
        lineItems(first: 3) { edges { node { name sku quantity } } }
      }
    }
  }
}
"""

# ── Shopify ───────────────────────────────────────────────────────────────────

def shopify_gql(variables):
    url  = f"https://{SHOPIFY_STORE}/admin/api/{SHOPIFY_API_VER}/graphql.json"
    hdrs = {"X-Shopify-Access-Token": SHOPIFY_TOKEN, "Content-Type": "application/json"}
    r = requests.post(url, headers=hdrs, json={"query": GQL_QUERY, "variables": variables})
    r.raise_for_status()
    body = r.json()
    if "errors" in body:
        raise RuntimeError(body["errors"])
    return body["data"]["orders"]


def fetch_month_orders(year: int, month: int):
    month_start = datetime(year, month, 1, tzinfo=timezone.utc) - timedelta(hours=5, minutes=30)
    last_day    = calendar.monthrange(year, month)[1]
    month_end   = datetime(year, month, last_day, 23, 59, 59, tzinfo=timezone.utc) - timedelta(hours=5, minutes=30)

    tag_filter = " OR ".join(f"tag:{t}" for t in sorted(TARGET_TAGS))
    q = (
        f"created_at:>={month_start.strftime('%Y-%m-%dT%H:%M:%SZ')} "
        f"created_at:<={month_end.strftime('%Y-%m-%dT%H:%M:%SZ')} "
        f"({tag_filter})"
    )

    orders, cursor, page = [], None, 0
    while True:
        page += 1
        print(f"  Page {page}...", end=" ", flush=True)
        result = shopify_gql({"query": q, "after": cursor})
        for e in result["edges"]:
            node = e["node"]
            tags = set(node.get("tags", []))
            if "Undelivered" in tags and "Delivered" in tags:
                continue
            orders.append(node)
        print(f"{len(result['edges'])} fetched (kept running: {len(orders)})")
        if not result["pageInfo"]["hasNextPage"]:
            break
        cursor = result["pageInfo"]["endCursor"]
    return orders

# ── Row builders ──────────────────────────────────────────────────────────────

def order_type(tags):
    t = set(tags)
    if "Replacement" in t:                                              return "Replacement"
    if t & {"Refund_initiated", "Refund_Initiated", "Refund_credited"}: return "Refund"
    if "RTO"         in t:                                              return "RTO"
    if "Returned"    in t:                                              return "Returned"
    if "Undelivered" in t:                                              return "Undelivered"
    return "Other"

def cancel_reason(tags):
    tl = {t.lower() for t in tags}
    if "customer-cancel"        in tl: return "Customer Cancel"
    if "ivr_cancel"             in tl: return "IVR Cancel"
    if "nsz-cancel"             in tl: return "NSZ Cancel"
    if "order_cancelled_lc_bot" in tl: return "LC Bot Cancel"
    return ""

def payment_method(tags):
    MAP = {
        "UPI": "UPI", "Cards": "Cards", "COD": "COD",
        "PPCOD-UPI": "PPCOD-UPI", "Snapmint": "Snapmint",
        "Wallets": "Wallets", "bnpl-lazypay": "LazyPay", "pay-later": "Pay Later",
    }
    for tag in tags:
        if tag in MAP:
            return MAP[tag]
    return ""

def actions_taken(tags):
    t = set(tags)
    actions = []
    if "Partial_replacement" in t: actions.append("Partial Replacement")
    if "Full_replacement"    in t: actions.append("Full Replacement")
    if "refund_given"        in t: actions.append("Refund Given")
    return ", ".join(actions)

def refund_status(tags):
    t = set(tags)
    if "Refund_credited" in t:                     return "Credited"
    if t & {"Refund_initiated", "Refund_Initiated"}: return "Initiated"
    return "Pending"

def to_ist(utc_str):
    try:
        dt  = datetime.fromisoformat(utc_str.replace("Z", "+00:00"))
        ist = dt + timedelta(hours=5, minutes=30)
        return ist.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return utc_str

def build_row(node):
    tags     = node.get("tags", [])
    customer = node.get("customer") or {}
    addr     = node.get("shippingAddress") or {}
    items    = [e["node"] for e in node.get("lineItems", {}).get("edges", [])]
    while len(items) < 3:
        items.append({"name": "", "sku": "", "quantity": ""})

    return [
        node.get("name", ""),
        to_ist(node.get("createdAt", "")),
        order_type(tags),
        cancel_reason(tags),
        node.get("displayFinancialStatus", ""),
        node.get("displayFulfillmentStatus", ""),
        actions_taken(tags),                                 # col 6 – Actions Taken
        "",                                                  # col 7 – Notes (manual fill)
        f"{customer.get('firstName') or ''} {customer.get('lastName') or ''}".strip(),
        customer.get("phone", "") or "",
        addr.get("city", "")     or "",
        addr.get("province", "") or "",
        payment_method(tags),
        node.get("totalPriceSet",    {}).get("shopMoney", {}).get("amount", "0"),
        node.get("totalRefundedSet", {}).get("shopMoney", {}).get("amount", "0"),
        refund_status(tags),
        items[0]["name"], items[0]["sku"], str(items[0]["quantity"]),
        items[1]["name"], items[1]["sku"], str(items[1]["quantity"]),
        items[2]["name"], items[2]["sku"], str(items[2]["quantity"]),
    ]

# ── Google Sheets helpers ─────────────────────────────────────────────────────

def get_or_create_tab(sh, name, rows=10000, cols=30):
    try:
        return sh.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        return sh.add_worksheet(title=name, rows=rows, cols=cols)


def write_tab(ws, headers, rows):
    needed = len(rows) + 10
    if ws.row_count < needed:
        ws.resize(rows=needed + 500)
    ws.clear()
    all_rows = [headers] + rows
    BATCH = 500
    for i in range(0, len(all_rows), BATCH):
        chunk     = all_rows[i:i + BATCH]
        start_row = i + 1
        ws.update(values=chunk, range_name=f"A{start_row}", value_input_option="USER_ENTERED")
    last_col = chr(ord("A") + len(headers) - 1)
    ws.format(f"A1:{last_col}1", {
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        "backgroundColor": {"red": 0.13, "green": 0.13, "blue": 0.13},
    })


def rebuild_all_data(sh):
    """Combine every 'Month YYYY' tab into a single All Data tab."""
    month_pattern = re.compile(r"^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$")
    combined = []
    for ws in sh.worksheets():
        if not month_pattern.match(ws.title):
            continue
        print(f"  Reading tab '{ws.title}'...")
        records = ws.get_all_values()
        if len(records) < 2:
            continue
        for row in records[1:]:   # skip header
            combined.append([ws.title] + row)

    # Sort by order date (col index 2 = "Order Date (IST)" after prepending Month)
    combined.sort(key=lambda r: r[2] if len(r) > 2 else "", reverse=True)

    ws_all = get_or_create_tab(sh, "All Data", rows=max(len(combined) + 10, 1000))
    write_tab(ws_all, ALL_DATA_HEADERS, combined)
    print(f"  All Data tab rebuilt — {len(combined)} rows.")
    return ws_all


# ── Dashboard charts ──────────────────────────────────────────────────────────

def _col(name):
    """Return 0-based column index in ALL_DATA_HEADERS."""
    return ALL_DATA_HEADERS.index(name)


def build_summary_data(sh):
    """
    Write a hidden _ChartData tab with two pivot tables:
      A) MoM summary  — Month | Refund | Returned | RTO | Undelivered | Replacement | Total
      B) Top 20 SKUs  — SKU | Product | Refund | Returned | RTO | Undelivered | Total
    Returns (ws_chart, mom_rows, sku_rows)
    """
    ws_all   = sh.worksheet("All Data")
    records  = ws_all.get_all_values()
    if len(records) < 2:
        return None, [], []

    header  = records[0]
    col_mon  = header.index("Month")
    col_type = header.index("Order Type")
    col_sku1 = header.index("SKU 1")
    col_prd1 = header.index("Product 1")
    col_sku2 = header.index("SKU 2")
    col_prd2 = header.index("Product 2")
    col_sku3 = header.index("SKU 3")
    col_prd3 = header.index("Product 3")

    ORDER_TYPES = ["Refund", "Returned", "RTO", "Undelivered", "Replacement"]

    # ── MoM pivot ────────────────────────────────────────────────────────────
    mom = defaultdict(lambda: defaultdict(int))
    for row in records[1:]:
        m = row[col_mon]  if len(row) > col_mon  else ""
        t = row[col_type] if len(row) > col_type else ""
        if m and t:
            mom[m][t] += 1

    # Sort months chronologically
    def month_sort_key(m):
        try:
            return datetime.strptime(m, "%B %Y")
        except Exception:
            return datetime.min

    sorted_months = sorted(mom.keys(), key=month_sort_key)
    mom_header = ["Month"] + ORDER_TYPES + ["Total"]
    mom_rows   = []
    for m in sorted_months:
        counts = [mom[m].get(t, 0) for t in ORDER_TYPES]
        mom_rows.append([m] + counts + [sum(counts)])

    # ── SKU pivot ────────────────────────────────────────────────────────────
    sku_data = defaultdict(lambda: {"product": "", **{t: 0 for t in ORDER_TYPES}})
    for row in records[1:]:
        t = row[col_type] if len(row) > col_type else ""
        if t not in ORDER_TYPES:
            continue
        for sku_col, prd_col in [(col_sku1, col_prd1), (col_sku2, col_prd2), (col_sku3, col_prd3)]:
            sku = row[sku_col] if len(row) > sku_col else ""
            prd = row[prd_col] if len(row) > prd_col else ""
            if sku:
                sku_data[sku]["product"] = prd or sku_data[sku]["product"]
                sku_data[sku][t] += 1

    sku_header = ["SKU", "Product"] + ORDER_TYPES + ["Total"]
    sku_rows   = []
    for sku, d in sku_data.items():
        counts = [d.get(t, 0) for t in ORDER_TYPES]
        sku_rows.append([sku, d["product"]] + counts + [sum(counts)])
    sku_rows.sort(key=lambda r: r[-1], reverse=True)
    sku_rows = sku_rows[:25]  # top 25 SKUs

    # ── Write _ChartData tab ─────────────────────────────────────────────────
    ws_cd = get_or_create_tab(sh, "_ChartData", rows=200, cols=20)
    ws_cd.clear()

    # MoM block starting at A1
    ws_cd.update(values=[mom_header] + mom_rows, range_name="A1", value_input_option="USER_ENTERED")

    # SKU block starting at row len(mom_rows)+4  (leave a gap)
    sku_start = len(mom_rows) + 4
    ws_cd.update(values=[sku_header] + sku_rows, range_name=f"A{sku_start}", value_input_option="USER_ENTERED")

    print(f"  _ChartData written — {len(mom_rows)} months, {len(sku_rows)} SKUs.")
    return ws_cd, mom_rows, sku_rows, sku_start


def create_dashboard_charts(service, sh, ws_cd_id, mom_rows, sku_rows, sku_start):
    """
    Create charts on the Dashboard tab using the Sheets API.
    Skips creation if charts already exist on that sheet.
    """
    # Get or create Dashboard tab via gspread
    try:
        ws_dash = sh.worksheet("Dashboard")
    except gspread.exceptions.WorksheetNotFound:
        ws_dash = sh.add_worksheet(title="Dashboard", rows=50, cols=20)

    dash_id = ws_dash.id

    # Check if charts already exist on Dashboard — skip if so
    spreadsheet = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    existing_charts = [
        c for sheet in spreadsheet.get("sheets", [])
        for c in sheet.get("charts", [])
        if sheet["properties"]["sheetId"] == dash_id
    ]
    if existing_charts:
        print(f"  Dashboard already has {len(existing_charts)} chart(s) — skipping creation.")
        return

    n_months = len(mom_rows)
    n_skus   = len(sku_rows)
    cd_id    = ws_cd_id

    # Helper: column range spec
    def col_range(sheet_id, start_row, end_row, col_idx):
        return {
            "sheetId": sheet_id,
            "startRowIndex": start_row,
            "endRowIndex":   end_row,
            "startColumnIndex": col_idx,
            "endColumnIndex":   col_idx + 1,
        }

    ORDER_TYPES  = ["Refund", "Returned", "RTO", "Undelivered", "Replacement"]
    COLORS = [
        {"red": 0.23, "green": 0.47, "blue": 0.85},  # blue   – Refund
        {"red": 0.96, "green": 0.60, "blue": 0.07},  # amber  – Returned
        {"red": 0.83, "green": 0.18, "blue": 0.18},  # red    – RTO
        {"red": 0.42, "green": 0.65, "blue": 0.31},  # green  – Undelivered
        {"red": 0.60, "green": 0.30, "blue": 0.70},  # purple – Replacement
    ]

    requests_list = []

    # ── Chart 1: MoM Stacked Column ──────────────────────────────────────────
    mom_data_rows = (1, 1 + n_months)   # rows 1..n_months+1 (0-based: 1 to n_months)
    series_mom = []
    for i, (otype, color) in enumerate(zip(ORDER_TYPES, COLORS)):
        series_mom.append({
            "series": {"sourceRange": {"sources": [col_range(cd_id, *mom_data_rows, i + 1)]}},
            "targetAxis": "LEFT_AXIS",
            "color": color,
        })

    requests_list.append({"addChart": {"chart": {
        "spec": {
            "title": "Month-on-Month Order Concerns",
            "basicChart": {
                "chartType": "COLUMN",
                "stackedType": "STACKED",
                "legendPosition": "BOTTOM_LEGEND",
                "domains": [{"domain": {"sourceRange": {"sources": [col_range(cd_id, *mom_data_rows, 0)]}}}],
                "series": series_mom,
                "headerCount": 0,
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": dash_id, "rowIndex": 1, "columnIndex": 0},
            "widthPixels": 620, "heightPixels": 380,
        }},
    }}})

    # ── Chart 2: Order Type Donut ─────────────────────────────────────────────
    # Aggregate totals per type from mom_rows
    type_totals = [sum(r[i + 1] for r in mom_rows) for i in range(len(ORDER_TYPES))]
    # Write a tiny pivot at column J (index 9) in _ChartData for the donut
    donut_data = [["Order Type", "Count"]] + [[t, v] for t, v in zip(ORDER_TYPES, type_totals)]
    donut_start_col = 9  # column J
    sh.worksheet("_ChartData").update(
        values=donut_data,
        range_name=f"J1",
        value_input_option="USER_ENTERED"
    )

    requests_list.append({"addChart": {"chart": {
        "spec": {
            "title": "Overall Order Type Breakdown",
            "pieChart": {
                "legendPosition": "RIGHT_LEGEND",
                "pieHole": 0.4,
                "domain": {"sourceRange": {"sources": [col_range(cd_id, 0, len(donut_data), donut_start_col)]}},
                "series": {"sourceRange": {"sources": [col_range(cd_id, 0, len(donut_data), donut_start_col + 1)]}},
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": dash_id, "rowIndex": 1, "columnIndex": 7},
            "widthPixels": 420, "heightPixels": 380,
        }},
    }}})

    # ── Chart 3: Top 25 SKUs Bar ──────────────────────────────────────────────
    sku_data_rows = (sku_start, sku_start + n_skus)   # 0-based row indices in sheet
    requests_list.append({"addChart": {"chart": {
        "spec": {
            "title": "Top 25 SKUs by Total Concerns",
            "basicChart": {
                "chartType": "BAR",
                "stackedType": "STACKED",
                "legendPosition": "BOTTOM_LEGEND",
                "domains": [{"domain": {"sourceRange": {"sources": [col_range(cd_id, *sku_data_rows, 0)]}}}],
                "series": [
                    {
                        "series": {"sourceRange": {"sources": [col_range(cd_id, *sku_data_rows, i + 2)]}},
                        "targetAxis": "BOTTOM_AXIS",
                        "color": COLORS[i],
                    }
                    for i in range(len(ORDER_TYPES))
                ],
                "headerCount": 0,
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": dash_id, "rowIndex": 23, "columnIndex": 0},
            "widthPixels": 1050, "heightPixels": 520,
        }},
    }}})

    service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": requests_list},
    ).execute()
    print(f"  Dashboard charts created (MoM column + donut + SKU bar).")


# ── Sheet polish ─────────────────────────────────────────────────────────────

def polish_sheet(service, sh):
    month_pattern = re.compile(r"^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$")
    worksheets    = sh.worksheets()
    requests_list = []

    # Tab order: Dashboard first, then months chronologically, then All Data, _ChartData last
    def tab_sort_key(ws):
        if ws.title == "Dashboard":   return (0, "")
        if ws.title == "All Data":    return (2, "")
        if ws.title == "_ChartData":  return (3, "")
        if month_pattern.match(ws.title):
            try:    return (1, datetime.strptime(ws.title, "%B %Y").strftime("%Y%m"))
            except: return (1, ws.title)
        return (4, ws.title)

    sorted_tabs = sorted(worksheets, key=tab_sort_key)
    for idx, ws in enumerate(sorted_tabs):
        requests_list.append({"updateSheetProperties": {
            "properties": {"sheetId": ws.id, "index": idx},
            "fields": "index",
        }})

    # Hide _ChartData
    cd_ws = next((w for w in worksheets if w.title == "_ChartData"), None)
    if cd_ws:
        requests_list.append({"updateSheetProperties": {
            "properties": {"sheetId": cd_ws.id, "hidden": True},
            "fields": "hidden",
        }})

    # For each data tab: freeze header row, auto-resize columns, banded rows
    data_tabs = [w for w in worksheets if w.title != "Dashboard"]
    GREY_LIGHT = {"red": 0.95, "green": 0.95, "blue": 0.95}
    WHITE      = {"red": 1.0,  "green": 1.0,  "blue": 1.0}

    for ws in data_tabs:
        sid = ws.id
        col_count = len(MONTH_HEADERS) if month_pattern.match(ws.title) else len(ALL_DATA_HEADERS)

        # Freeze row 1
        requests_list.append({"updateSheetProperties": {
            "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1}},
            "fields": "gridProperties.frozenRowCount",
        }})

        # Auto-resize all columns
        requests_list.append({"autoResizeDimensions": {
            "dimensions": {
                "sheetId": sid,
                "dimension": "COLUMNS",
                "startIndex": 0,
                "endIndex": col_count,
            }
        }})

        # Alternating row band (skip if one already exists)
        requests_list.append({"addBanding": {
            "bandedRange": {
                "range": {
                    "sheetId": sid,
                    "startRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": col_count,
                },
                "rowProperties": {
                    "headerColor":     {"red": 0.13, "green": 0.13, "blue": 0.13},
                    "firstBandColor":  WHITE,
                    "secondBandColor": GREY_LIGHT,
                },
            }
        }})

    # Execute — ignore "already exists" errors for banding
    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={"requests": requests_list},
        ).execute()
        print("  Tabs reordered, headers frozen, columns resized, banding applied.")
    except Exception as e:
        # On any banding conflict, retry without addBanding requests
        if "banding" in str(e).lower() or "alternating" in str(e).lower():
            clean_requests = [r for r in requests_list if "addBanding" not in r]
            service.spreadsheets().batchUpdate(
                spreadsheetId=SHEET_ID,
                body={"requests": clean_requests},
            ).execute()
            print("  Tabs reordered, headers frozen, columns resized (banding already applied).")
        else:
            print(f"  Polish warning: {e}")


# ── Entry point ───────────────────────────────────────────────────────────────

def parse_month_arg(arg: str):
    try:
        dt = datetime.strptime(arg, "%Y-%m")
        return dt.year, dt.month
    except ValueError:
        raise SystemExit(f"Bad month argument '{arg}'. Expected format: YYYY-MM (e.g. 2026-07)")


def main():
    if len(sys.argv) > 1:
        year, month = parse_month_arg(sys.argv[1])
    else:
        now   = datetime.now(timezone.utc) + timedelta(hours=5, minutes=30)
        year  = now.year
        month = now.month

    month_name = datetime(year, month, 1).strftime("%B %Y")

    print(f"=== Swiss Beauty Returns & Refund Tracker — {month_name} ===")

    # ── 1. Fetch from Shopify ─────────────────────────────────────────────────
    print(f"\nFetching {month_name} orders from Shopify...")
    orders = fetch_month_orders(year, month)
    print(f"Total: {len(orders)} orders\n")

    rows = [build_row(o) for o in orders]

    # ── 2. Connect to Sheets ──────────────────────────────────────────────────
    print("Connecting to Google Sheets...")
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    gc      = gspread.authorize(creds)
    sh      = gc.open_by_key(SHEET_ID)
    service = build("sheets", "v4", credentials=creds)

    # ── 3. Write monthly tab ──────────────────────────────────────────────────
    print(f"Writing '{month_name}' tab...")
    ws_month = get_or_create_tab(sh, month_name)
    write_tab(ws_month, MONTH_HEADERS, rows)
    print(f"  {len(rows)} rows written.")

    # ── 4. Rebuild All Data ───────────────────────────────────────────────────
    print("\nRebuilding 'All Data' tab...")
    rebuild_all_data(sh)

    # ── 5. Rebuild _ChartData + create Dashboard charts ───────────────────────
    print("\nUpdating chart data...")
    ws_cd, mom_rows, sku_rows, sku_start = build_summary_data(sh)
    if ws_cd and mom_rows:
        create_dashboard_charts(service, sh, ws_cd.id, mom_rows, sku_rows, sku_start)

    # ── 6. Polish the sheet ───────────────────────────────────────────────────
    print("\nPolishing sheet...")
    polish_sheet(service, sh)

    print(f"\nDone! https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")


if __name__ == "__main__":
    main()
