"""
Swiss Beauty – Replacement / Refund / RTO monthly sheet sync.

Usage:
  python sync_to_sheet.py              # dumps the CURRENT calendar month
  python sync_to_sheet.py 2026-07      # dumps a specific YYYY-MM

Each run creates / overwrites a tab named "July 2026" (or whichever month)
inside the Google Sheet. Other tabs are left untouched.
"""

import sys
import json
import calendar
import requests
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone, timedelta

# ── Config ────────────────────────────────────────────────────────────────────
import os

SHOPIFY_STORE   = "swiss-beauty-dev.myshopify.com"
SHOPIFY_TOKEN   = os.environ.get("SHOPIFY_TOKEN", "")
SHOPIFY_API_VER = "2024-01"
SHEET_ID        = "1VptcDahrMKwqgo3wWmxpoVzXEyY5je2U73mvx4rOYnU"
CREDS_FILE = (
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
}

HEADERS = [
    "Order Number", "Order Date (IST)", "Order Type", "Cancel Reason",
    "Financial Status", "Fulfillment Status",
    "Customer Name", "Customer Phone", "City", "State",
    "Payment Method", "Order Value (₹)", "Refunded Amount (₹)", "Refund Status",
    "Product 1", "SKU 1", "Qty 1",
    "Product 2", "SKU 2", "Qty 2",
    "Product 3", "SKU 3", "Qty 3",
]

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

# ── Shopify helpers ───────────────────────────────────────────────────────────

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
    """Fetch all tagged orders created within a given calendar month (IST → UTC)."""
    # Month boundaries in UTC (IST is UTC+5:30, so subtract to get UTC)
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
        result   = shopify_gql({"query": q, "after": cursor})
        edges    = result["edges"]
        for e in edges:
            node = e["node"]
            tags = set(node.get("tags", []))
            # Skip orders that were temporarily Undelivered but ultimately Delivered
            if "Undelivered" in tags and "Delivered" in tags:
                continue
            orders.append(node)
        print(f"{len(edges)} orders (running total: {len(orders)})")
        if not result["pageInfo"]["hasNextPage"]:
            break
        cursor = result["pageInfo"]["endCursor"]

    return orders

# ── Row builders ──────────────────────────────────────────────────────────────

def order_type(tags):
    t = set(tags)
    if "Replacement"  in t:                                        return "Replacement"
    if t & {"Refund_initiated", "Refund_Initiated", "Refund_credited"}: return "Refund"
    if "RTO"          in t:                                        return "RTO"
    if "Returned"     in t:                                        return "Returned"
    if "Undelivered"  in t:                                        return "Undelivered"
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

def refund_status(tags):
    t = set(tags)
    if "Refund_credited" in t:                                     return "Credited"
    if t & {"Refund_initiated", "Refund_Initiated"}:               return "Initiated"
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

# ── Google Sheets ─────────────────────────────────────────────────────────────

def get_or_create_tab(sh, tab_name: str):
    """Return worksheet; create it if it doesn't exist."""
    try:
        return sh.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        return sh.add_worksheet(title=tab_name, rows=5000, cols=len(HEADERS))


def write_to_sheet(rows, tab_name: str):
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(SHEET_ID)
    ws    = get_or_create_tab(sh, tab_name)

    ws.clear()
    # Write in batches of 500 rows to stay within API payload limits
    all_rows = [HEADERS] + rows
    BATCH = 500
    for i in range(0, len(all_rows), BATCH):
        chunk = all_rows[i:i + BATCH]
        start_row = i + 1
        ws.update(
            values=chunk,
            range_name=f"A{start_row}",
            value_input_option="USER_ENTERED",
        )

    ws.format(f"A1:{chr(ord('A') + len(HEADERS) - 1)}1", {
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        "backgroundColor": {"red": 0.13, "green": 0.13, "blue": 0.13},
    })
    print(f"  Written to tab '{tab_name}' — {len(rows)} rows.")

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
        now   = datetime.now(timezone.utc) + timedelta(hours=5, minutes=30)  # IST
        year  = now.year
        month = now.month

    month_name = datetime(year, month, 1).strftime("%B %Y")   # e.g. "July 2026"
    tab_name   = month_name

    print(f"=== Swiss Beauty Refund/RTO Sync — {month_name} ===")
    print(f"Fetching orders for {month_name} from Shopify...")
    orders = fetch_month_orders(year, month)
    print(f"Total matching orders: {len(orders)}\n")

    rows = [build_row(o) for o in orders]

    print(f"Writing to Google Sheet tab '{tab_name}'...")
    write_to_sheet(rows, tab_name)

    print(f"\nDone! Sheet: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")


if __name__ == "__main__":
    main()
