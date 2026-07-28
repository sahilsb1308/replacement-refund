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
    "Actions Taken", "Notes",
    "Customer Name", "Customer Phone", "City", "State",
    "Payment Method", "Order Value (₹)", "Refunded Amount (₹)", "Refund Status",
    "Product 1", "SKU 1", "Qty 1",
    "Product 2", "SKU 2", "Qty 2",
    "Product 3", "SKU 3", "Qty 3",
    "Clone MRP (₹)", "Order Loss (₹)",                      # ← loss cols
]

ALL_DATA_HEADERS = ["Month"] + MONTH_HEADERS

GQL_QUERY = """
query FetchOrders($query: String!, $after: String) {
  orders(first: 100, query: $query, after: $after, sortKey: CREATED_AT, reverse: true) {
    pageInfo { hasNextPage endCursor }
    edges {
      node {
        name createdAt tags note
        displayFinancialStatus displayFulfillmentStatus
        totalPriceSet    { shopMoney { amount } }
        totalRefundedSet { shopMoney { amount } }
        customer { firstName lastName phone }
        shippingAddress { city province }
        lineItems(first: 3) { edges { node { name sku quantity originalUnitPriceSet { shopMoney { amount } } } } }
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

def _clone_mrp_value(tags, items):
    """Sum originalUnitPrice × qty for replacement order items.
    Returns None when MRP can't be determined (price ≤ 1 = clone token price)."""
    if "Replacement" not in set(tags):
        return None
    total = sum(
        float(((item.get("originalUnitPriceSet") or {}).get("shopMoney") or {}).get("amount", 0) or 0)
        * int(item.get("quantity") or 0)
        for item in items
    )
    return round(total, 2) if total >= 1.0 else None


def _order_loss(tags, items, refunded_str):
    """Return the financial loss value for this order, or '' if unknown/not applicable.

    Rules (per business logic):
      Full replacement only   → Clone MRP + ₹200 shipping
      Partial replacement only→ Clone MRP + ₹100 shipping
      Replacement + refund    → Clone MRP + Refunded Amount + ₹200/100 shipping
      Refund only             → Refunded Amount + ₹100 shipping
      Clone MRP unknown       → '' (null – tracked separately as unknown)
    """
    t = set(tags)
    try:    refunded = float(refunded_str or 0)
    except: refunded = 0.0

    is_replacement  = "Replacement"   in t
    is_refund_given = "refund_given"  in t  # refund also issued alongside replacement
    is_refund_only  = bool(t & {"Refund_initiated", "Refund_Initiated", "Refund_credited"}) and not is_replacement
    is_full         = "Full_replacement"    in t
    ship            = 200 if is_full else 100

    if is_replacement:
        mrp = _clone_mrp_value(tags, items)
        if mrp is None:
            return ""   # null — MRP unknown
        if is_refund_given and refunded > 0:
            return round(mrp + refunded + ship, 2)   # combined
        return round(mrp + ship, 2)                   # replacement only

    if is_refund_only and refunded > 0:
        return round(refunded + 100, 2)               # refund + ₹100 shipping

    return ""


def build_row(node):
    tags     = node.get("tags", [])
    customer = node.get("customer") or {}
    addr     = node.get("shippingAddress") or {}
    items    = [e["node"] for e in node.get("lineItems", {}).get("edges", [])]
    while len(items) < 3:
        items.append({"name": "", "sku": "", "quantity": "", "originalUnitPriceSet": None})

    refunded_str = node.get("totalRefundedSet", {}).get("shopMoney", {}).get("amount", "0")
    mrp          = _clone_mrp_value(tags, items)
    loss         = _order_loss(tags, items, refunded_str)

    return [
        node.get("name", ""),
        to_ist(node.get("createdAt", "")),
        order_type(tags),
        cancel_reason(tags),
        node.get("displayFinancialStatus", ""),
        node.get("displayFulfillmentStatus", ""),
        actions_taken(tags),
        node.get("note") or "",
        f"{customer.get('firstName') or ''} {customer.get('lastName') or ''}".strip(),
        customer.get("phone", "") or "",
        addr.get("city", "")     or "",
        addr.get("province", "") or "",
        payment_method(tags),
        node.get("totalPriceSet",    {}).get("shopMoney", {}).get("amount", "0"),
        refunded_str,
        refund_status(tags),
        items[0]["name"], items[0]["sku"], str(items[0]["quantity"]),
        items[1]["name"], items[1]["sku"], str(items[1]["quantity"]),
        items[2]["name"], items[2]["sku"], str(items[2]["quantity"]),
        str(mrp)  if mrp  is not None else "",   # Clone MRP (₹)
        str(loss) if loss != ""       else "",   # Order Loss (₹)
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
    last_col = _col_letter(len(headers) - 1)   # handles >26 cols (AA, AB, …)
    ws.format(f"A1:{last_col}1", {
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        "backgroundColor": {"red": 0.13, "green": 0.13, "blue": 0.13},
    })
    # Wrap Notes column so the 3-line template displays correctly
    # Notes is col H in month tabs, col I in All Data (which has "Month" prepended)
    notes_col = "I" if "Month" in headers else "H"
    if rows:
        ws.format(f"{notes_col}2:{notes_col}{len(rows) + 1}", {"wrapStrategy": "WRAP"})


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


SKU_COL_START    = 12   # col M (0-indexed) — SKU pivot lives in same rows as MoM, cols M onward
REASON_LABELS    = ["Damaged", "Missing", "Wrong Item", "Used", "Other"]
TOP_DAMAGE_SKUS  = 10


def _col_letter(n):
    """0-indexed column number → A1-notation letter (A=0, Z=25, AA=26 …)."""
    result, n = '', n + 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _parse_reason(note):
    """Return normalised reason string from a Shopify order note, or None."""
    for line in note.split('\n'):
        m = re.match(r'reason[-:\s]+(.+)', line.strip(), re.IGNORECASE)
        if m:
            r = m.group(1).strip().lower()
            if 'damage'              in r: return 'Damaged'
            if 'missing'             in r: return 'Missing'
            if 'wrong' in r or 'incorrect' in r: return 'Wrong Item'
            if 'used'                in r: return 'Used'
            return 'Other'
    return None


def _parse_damage_sku(note):
    """Return the affected SKU from a note (first SKU on the sku- / ReplacementSKU: line).
    Requires a hyphen so free-text values like 'ALL' or 'missing' are ignored."""
    for line in note.split('\n'):
        line = line.strip()
        m = re.match(r'(?:replacement\s*)?sku[-:\s=]+(.+)', line, re.IGNORECASE)
        if m:
            raw = m.group(1).strip()
            sku = re.split(r'[\s,;\n]+', raw)[0].strip()
            if sku and '-' in sku:   # all real SB SKUs have at least one hyphen
                return sku.upper()
    return None


def build_summary_data(sh):
    """
    Write _ChartData so a single slicer on col A controls all five Dashboard charts.
    All data fits in rows 1..n (one row per month):

      Cols A–G  (rows 0..n): MoM table — Month | Refund | Returned | RTO | Undelivered | Replacement | Total
      Cols J–K  (rows 0..5): Donut     — Order Type | =SUBTOTAL() recalcs on slicer filter
      Cols M+   (rows 0..n): SKU pivot — top-15 SKU codes as column headers, counts as data
      Cols AD+  (rows 0..n): Damage reason pivot — Damaged | Missing | Wrong Item | Used | Other
      Cols AK+  (rows 0..n): Damage SKU pivot — top-10 SKUs from damage/missing notes
    """
    ws_all   = sh.worksheet("All Data")
    records  = ws_all.get_all_values()
    if len(records) < 2:
        return None, [], 0

    header   = records[0]
    col_mon  = header.index("Month")
    col_type = header.index("Order Type")
    col_sku1 = header.index("SKU 1")
    col_sku2 = header.index("SKU 2")
    col_sku3 = header.index("SKU 3")

    ORDER_TYPES = ["Refund", "Returned", "RTO", "Undelivered", "Replacement"]

    def month_sort_key(m):
        try:    return datetime.strptime(m, "%B %Y")
        except: return datetime.min

    # ── MoM pivot ────────────────────────────────────────────────────────────
    mom = defaultdict(lambda: defaultdict(int))
    for row in records[1:]:
        m = row[col_mon]  if len(row) > col_mon  else ""
        t = row[col_type] if len(row) > col_type else ""
        if m and t:
            mom[m][t] += 1

    sorted_months = sorted(mom.keys(), key=month_sort_key)
    mom_header = ["Month"] + ORDER_TYPES + ["Total"]
    mom_rows   = []
    for m in sorted_months:
        counts = [mom[m].get(t, 0) for t in ORDER_TYPES]
        mom_rows.append([m] + counts + [sum(counts)])

    n = len(mom_rows)

    # ── SKU wide pivot — same rows as MoM, cols M onward ─────────────────────
    all_sku_counts = defaultdict(int)
    for row in records[1:]:
        t = row[col_type] if len(row) > col_type else ""
        if t not in ORDER_TYPES:
            continue
        for sku_col in [col_sku1, col_sku2, col_sku3]:
            sku = row[sku_col] if len(row) > sku_col else ""
            if sku:
                all_sku_counts[sku] += 1

    top_skus = [s for s, _ in sorted(all_sku_counts.items(), key=lambda x: -x[1])[:15]]

    sku_month_map = defaultdict(lambda: defaultdict(int))
    for row in records[1:]:
        m = row[col_mon]  if len(row) > col_mon  else ""
        t = row[col_type] if len(row) > col_type else ""
        if not m or t not in ORDER_TYPES:
            continue
        for sku_col in [col_sku1, col_sku2, col_sku3]:
            sku = row[sku_col] if len(row) > sku_col else ""
            if sku:
                sku_month_map[m][sku] += 1

    sku_pivot_data = [
        [sku_month_map[m].get(sku, 0) for sku in top_skus]
        for m in sorted_months
    ]

    # ── Damage / Missing pivot from order notes ───────────────────────────────
    col_notes = header.index("Notes") if "Notes" in header else None
    damage_reason_by_month:  dict = defaultdict(lambda: defaultdict(int))
    damaged_sku_by_month:    dict = defaultdict(lambda: defaultdict(int))
    missing_sku_by_month:    dict = defaultdict(lambda: defaultdict(int))

    if col_notes is not None:
        for row in records[1:]:
            m    = row[col_mon]   if len(row) > col_mon   else ""
            note = row[col_notes] if len(row) > col_notes else ""
            if not m or not note.strip():
                continue
            reason = _parse_reason(note)
            if reason:
                damage_reason_by_month[m][reason] += 1
                sku = _parse_damage_sku(note)
                if sku:
                    if reason == 'Damaged':
                        damaged_sku_by_month[m][sku] += 1
                    elif reason == 'Missing':
                        missing_sku_by_month[m][sku] += 1

    reason_pivot_data = [
        [damage_reason_by_month[m].get(r, 0) for r in REASON_LABELS]
        for m in sorted_months
    ]

    def _top_skus_for(by_month, n=TOP_DAMAGE_SKUS):
        totals: dict = defaultdict(int)
        for m_data in by_month.values():
            for sku, cnt in m_data.items():
                totals[sku] += cnt
        return [s for s, _ in sorted(totals.items(), key=lambda x: -x[1])[:n]]

    top_damaged_skus = _top_skus_for(damaged_sku_by_month)
    top_missing_skus = _top_skus_for(missing_sku_by_month)

    damaged_sku_pivot_data = [
        [damaged_sku_by_month[m].get(sku, 0) for sku in top_damaged_skus]
        for m in sorted_months
    ]
    missing_sku_pivot_data = [
        [missing_sku_by_month[m].get(sku, 0) for sku in top_missing_skus]
        for m in sorted_months
    ]

    # Column layout (0-indexed):
    #   M(12)..AB(27)  = top-15 SKU pivot
    #   AD(29)..AH(33) = damage reason pivot   (gap of 1 col)
    #   AJ(35)..AS(44) = damaged SKU pivot     (gap of 1 col)
    #   AU(46)..BD(55) = missing SKU pivot     (gap of 1 col)
    damage_reason_col = SKU_COL_START + len(top_skus) + 2
    damaged_sku_col   = damage_reason_col + len(REASON_LABELS) + 2
    missing_sku_col   = damaged_sku_col   + TOP_DAMAGE_SKUS    + 2

    # ── Write _ChartData ──────────────────────────────────────────────────────
    ws_cd = get_or_create_tab(sh, "_ChartData", rows=max(n + 10, 20), cols=70)
    ws_cd.clear()
    if ws_cd.col_count < 70:
        ws_cd.resize(rows=max(n + 10, 20), cols=70)

    # A) MoM cols A–G, rows 1..n+1
    ws_cd.update(values=[mom_header] + mom_rows, range_name="A1", value_input_option="USER_ENTERED")

    # B) Donut cols J–K — SUBTOTAL recalcs when slicer hides MoM rows
    donut_data = [["Order Type", "Count"]] + [
        [t, f"=SUBTOTAL(9,{col}2:{col}{n + 1})"]
        for t, col in zip(ORDER_TYPES, ["B", "C", "D", "E", "F"])
    ]
    ws_cd.update(values=donut_data, range_name="J1", value_input_option="USER_ENTERED")

    # C) SKU pivot cols M+, SAME rows as MoM — header in row 1, data in rows 2..n+1
    ws_cd.update(values=[top_skus],      range_name="M1", value_input_option="USER_ENTERED")
    ws_cd.update(values=sku_pivot_data,  range_name="M2", value_input_option="USER_ENTERED")

    # D) Damage reason pivot — same row structure as MoM
    dr_letter = _col_letter(damage_reason_col)
    ws_cd.update(values=[REASON_LABELS],    range_name=f"{dr_letter}1", value_input_option="USER_ENTERED")
    if reason_pivot_data:
        ws_cd.update(values=reason_pivot_data, range_name=f"{dr_letter}2", value_input_option="USER_ENTERED")

    # E) Damaged SKU pivot (reason == 'Damaged' only)
    if top_damaged_skus:
        de_letter = _col_letter(damaged_sku_col)
        ws_cd.update(values=[top_damaged_skus],     range_name=f"{de_letter}1", value_input_option="USER_ENTERED")
        ws_cd.update(values=damaged_sku_pivot_data, range_name=f"{de_letter}2", value_input_option="USER_ENTERED")

    # F) Missing SKU pivot (reason == 'Missing' only)
    if top_missing_skus:
        mf_letter = _col_letter(missing_sku_col)
        ws_cd.update(values=[top_missing_skus],     range_name=f"{mf_letter}1", value_input_option="USER_ENTERED")
        ws_cd.update(values=missing_sku_pivot_data, range_name=f"{mf_letter}2", value_input_option="USER_ENTERED")

    # ── Loss summary pivot ────────────────────────────────────────────────────
    col_actions   = header.index("Actions Taken")       if "Actions Taken"       in header else None
    col_refunded  = header.index("Refunded Amount (₹)") if "Refunded Amount (₹)" in header else None
    col_clone_mrp = header.index("Clone MRP (₹)")       if "Clone MRP (₹)"       in header else None

    def _sf(v):
        try:    return float(v or 0)
        except: return 0.0

    loss_replacement: dict = defaultdict(float)
    loss_refund:      dict = defaultdict(float)
    loss_combined:    dict = defaultdict(float)
    loss_null_count:  dict = defaultdict(int)

    for row in records[1:]:
        m     = row[col_mon]  if len(row) > col_mon  else ""
        otype = row[col_type] if len(row) > col_type else ""
        if not m:
            continue
        actions   = (row[col_actions]   if col_actions   and len(row) > col_actions   else "") or ""
        refunded  = _sf(row[col_refunded]  if col_refunded  and len(row) > col_refunded  else 0)
        clone_mrp_val = _sf(row[col_clone_mrp] if col_clone_mrp and len(row) > col_clone_mrp else 0)

        is_full      = "Full Replacement"    in actions
        is_combined  = "Refund Given"        in actions  # replacement + refund on same order
        ship         = 200 if is_full else 100

        if otype == "Replacement":
            if is_combined:
                if clone_mrp_val >= 1:
                    loss_combined[m] += clone_mrp_val + refunded + ship
                else:
                    loss_null_count[m] += 1
            else:
                if clone_mrp_val >= 1:
                    loss_replacement[m] += clone_mrp_val + ship
                else:
                    loss_null_count[m] += 1
        elif otype == "Refund" and refunded > 0:
            loss_refund[m] += refunded + 100   # +₹100 return shipping

    LOSS_LABELS = ["Replacement Loss (₹)", "Refund Loss (₹)", "Combined Loss (₹)", "Null MRP Orders"]
    loss_pivot_data = [
        [
            round(loss_replacement[m], 2),
            round(loss_refund[m],      2),
            round(loss_combined[m],    2),
            loss_null_count[m],
        ]
        for m in sorted_months
    ]

    loss_col = missing_sku_col + TOP_DAMAGE_SKUS + 2
    ll_letter = _col_letter(loss_col)
    ws_cd.update(values=[LOSS_LABELS],    range_name=f"{ll_letter}1", value_input_option="USER_ENTERED")
    if loss_pivot_data:
        ws_cd.update(values=loss_pivot_data, range_name=f"{ll_letter}2", value_input_option="USER_ENTERED")

    total_loss = sum(loss_replacement.values()) + sum(loss_refund.values()) + sum(loss_combined.values())
    total_notes = sum(sum(d.values()) for d in damage_reason_by_month.values())
    print(f"  _ChartData: {n} months, top {len(top_skus)} SKUs, {total_notes} damage notes, ₹{total_loss:,.0f} total loss ({sum(loss_null_count.values())} null-MRP orders).")
    return ws_cd, mom_rows, len(top_skus), damage_reason_col, damaged_sku_col, len(top_damaged_skus), missing_sku_col, len(top_missing_skus), loss_col


def create_dashboard_charts(service, sh, ws_cd_id, mom_rows, n_sku_series,
                             damage_reason_col,
                             damaged_sku_col, n_damaged_skus,
                             missing_sku_col,  n_missing_skus,
                             loss_col):
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

    # Ensure Dashboard has enough rows for all charts (new damage charts anchor at row 50+)
    if ws_dash.row_count < 120:
        ws_dash.resize(rows=120)

    # Delete existing charts on Dashboard so we always recreate with fresh ranges
    spreadsheet = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    existing_charts = [
        c for sheet in spreadsheet.get("sheets", [])
        for c in sheet.get("charts", [])
        if sheet["properties"]["sheetId"] == dash_id
    ]
    if existing_charts:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={"requests": [
                {"deleteEmbeddedObject": {"objectId": c["chartId"]}}
                for c in existing_charts
            ]},
        ).execute()
        print(f"  Deleted {len(existing_charts)} existing chart(s) — rebuilding.")

    n_months = len(mom_rows)
    n_skus   = n_months   # SKU pivot has same number of rows as MoM
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
    # Data is in J:K of _ChartData — SUBTOTAL formulas written by build_summary_data
    # They recalculate automatically when the slicer hides/shows MoM rows
    n_types         = len(ORDER_TYPES)
    donut_start_col = 9   # col J (0-indexed)
    requests_list.append({"addChart": {"chart": {
        "spec": {
            "title": "Order Type Breakdown",
            "pieChart": {
                "legendPosition": "RIGHT_LEGEND",
                "pieHole": 0.4,
                # J2:J6 = labels, K2:K6 = SUBTOTAL values
                "domain": {"sourceRange": {"sources": [col_range(cd_id, 1, 1 + n_types, donut_start_col)]}},
                "series": {"sourceRange": {"sources": [col_range(cd_id, 1, 1 + n_types, donut_start_col + 1)]}},
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": dash_id, "rowIndex": 1, "columnIndex": 7},
            "widthPixels": 420, "heightPixels": 380,
        }},
    }}})

    # ── Chart 3: Top-15 SKUs — same rows as MoM, cols M onward ──────────────
    # Domain = col A (Month), same rows 1..n_months.  Series = cols M, N, … (one per SKU).
    # Row 0 col M+ has SKU names as series labels (headerCount=1).
    # Slicer hides entire rows → MoM chart, donut, AND this SKU chart all update together.
    sku_data_end = n_months + 1  # 0-indexed exclusive (rows 0..n_months inclusive)
    requests_list.append({"addChart": {"chart": {
        "spec": {
            "title": "Top 15 SKUs — Share of Issue Orders by Month (%)",
            "basicChart": {
                "chartType": "BAR",
                "stackedType": "PERCENT_STACKED",
                "legendPosition": "RIGHT_LEGEND",
                # Both domain and series start at row 0 with headerCount=1.
                # Row 0 = header ("Month" in col A, SKU names in cols M+).
                # Rows 1..n = May/June/July data — all three bars appear.
                "domains": [{"domain": {"sourceRange": {"sources": [col_range(cd_id, 0, sku_data_end, 0)]}}}],
                "series": [
                    {
                        "series": {"sourceRange": {"sources": [col_range(cd_id, 0, sku_data_end, SKU_COL_START + i)]}},
                        "targetAxis": "BOTTOM_AXIS",
                    }
                    for i in range(n_sku_series)
                ],
                "headerCount": 1,
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": dash_id, "rowIndex": 23, "columnIndex": 0},
            "widthPixels": 1050, "heightPixels": 520,
        }},
    }}})

    # ── Chart 4: Damage / Missing — Reason Breakdown (stacked column, absolute counts) ──
    n_reasons    = len(REASON_LABELS)
    reason_end   = n_months + 1
    REASON_COLORS = [
        {"red": 0.83, "green": 0.18, "blue": 0.18},  # red    – Damaged
        {"red": 0.23, "green": 0.47, "blue": 0.85},  # blue   – Missing
        {"red": 0.96, "green": 0.60, "blue": 0.07},  # amber  – Wrong Item
        {"red": 0.42, "green": 0.65, "blue": 0.31},  # green  – Used
        {"red": 0.60, "green": 0.30, "blue": 0.70},  # purple – Other
    ]
    requests_list.append({"addChart": {"chart": {
        "spec": {
            "title": "Damage & Missing — Reason Breakdown",
            "basicChart": {
                "chartType": "COLUMN",
                "stackedType": "STACKED",
                "legendPosition": "BOTTOM_LEGEND",
                "domains": [{"domain": {"sourceRange": {"sources": [
                    col_range(cd_id, 0, reason_end, 0)
                ]}}}],
                "series": [
                    {
                        "series": {"sourceRange": {"sources": [
                            col_range(cd_id, 0, reason_end, damage_reason_col + i)
                        ]}},
                        "targetAxis": "LEFT_AXIS",
                        "color": REASON_COLORS[i],
                    }
                    for i in range(n_reasons)
                ],
                "headerCount": 1,
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": dash_id, "rowIndex": 50, "columnIndex": 0},
            "widthPixels": 560, "heightPixels": 380,
        }},
    }}})

    # ── Chart 5: Top Damaged SKUs — stacked bar (absolute counts, not %) ────────
    # STACKED (not PERCENT_STACKED) so months with 0 damage still show an empty bar
    if n_damaged_skus > 0:
        damaged_end = n_months + 1
        requests_list.append({"addChart": {"chart": {
            "spec": {
                "title": f"Top {n_damaged_skus} SKUs — Damaged Issues by Month",
                "basicChart": {
                    "chartType": "BAR",
                    "stackedType": "STACKED",
                    "legendPosition": "RIGHT_LEGEND",
                    "domains": [{"domain": {"sourceRange": {"sources": [
                        col_range(cd_id, 0, damaged_end, 0)
                    ]}}}],
                    "series": [
                        {
                            "series": {"sourceRange": {"sources": [
                                col_range(cd_id, 0, damaged_end, damaged_sku_col + i)
                            ]}},
                            "targetAxis": "BOTTOM_AXIS",
                        }
                        for i in range(n_damaged_skus)
                    ],
                    "headerCount": 1,
                },
            },
            "position": {"overlayPosition": {
                "anchorCell": {"sheetId": dash_id, "rowIndex": 70, "columnIndex": 0},
                "widthPixels": 560, "heightPixels": 380,
            }},
        }}})

    # ── Chart 6: Top Missing SKUs — stacked bar ───────────────────────────────
    if n_missing_skus > 0:
        missing_end = n_months + 1
        requests_list.append({"addChart": {"chart": {
            "spec": {
                "title": f"Top {n_missing_skus} SKUs — Missing Issues by Month",
                "basicChart": {
                    "chartType": "BAR",
                    "stackedType": "STACKED",
                    "legendPosition": "RIGHT_LEGEND",
                    "domains": [{"domain": {"sourceRange": {"sources": [
                        col_range(cd_id, 0, missing_end, 0)
                    ]}}}],
                    "series": [
                        {
                            "series": {"sourceRange": {"sources": [
                                col_range(cd_id, 0, missing_end, missing_sku_col + i)
                            ]}},
                            "targetAxis": "BOTTOM_AXIS",
                        }
                        for i in range(n_missing_skus)
                    ],
                    "headerCount": 1,
                },
            },
            "position": {"overlayPosition": {
                "anchorCell": {"sheetId": dash_id, "rowIndex": 70, "columnIndex": 9},
                "widthPixels": 560, "heightPixels": 380,
            }},
        }}})

    # ── Chart 7: Monthly Loss Breakdown (₹) — stacked column ─────────────────
    LOSS_COLORS = [
        {"red": 0.83, "green": 0.18, "blue": 0.18},  # red   – Replacement
        {"red": 0.23, "green": 0.47, "blue": 0.85},  # blue  – Refund
        {"red": 0.96, "green": 0.60, "blue": 0.07},  # amber – Combined
    ]
    LOSS_NAMES = ["Replacement Loss (₹)", "Refund Loss (₹)", "Combined Loss (₹)"]
    loss_end = n_months + 1
    requests_list.append({"addChart": {"chart": {
        "spec": {
            "title": "Monthly Loss Breakdown (₹)",
            "basicChart": {
                "chartType": "COLUMN",
                "stackedType": "STACKED",
                "legendPosition": "BOTTOM_LEGEND",
                "domains": [{"domain": {"sourceRange": {"sources": [
                    col_range(cd_id, 0, loss_end, 0)
                ]}}}],
                "series": [
                    {
                        "series": {"sourceRange": {"sources": [
                            col_range(cd_id, 0, loss_end, loss_col + i)
                        ]}},
                        "targetAxis": "LEFT_AXIS",
                        "color": LOSS_COLORS[i],
                    }
                    for i in range(3)  # 3 loss types; skip col +3 (null count)
                ],
                "headerCount": 1,
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": dash_id, "rowIndex": 92, "columnIndex": 0},
            "widthPixels": 700, "heightPixels": 400,
        }},
    }}})

    service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": requests_list},
    ).execute()
    n_charts = 4 + (1 if n_damaged_skus > 0 else 0) + (1 if n_missing_skus > 0 else 0) + 1
    print(f"  Dashboard charts created ({n_charts} total: MoM + donut + SKU + reason + damaged SKU + missing SKU + loss).")


# ── Dashboard slicer ─────────────────────────────────────────────────────────

def create_month_slicer(service, sh, ws_cd_id, n_months):
    """Add (or refresh) a Month slicer placed ON _ChartData (same sheet as data).
    Placing the slicer on the same sheet as its dataRange is the only way Google Sheets
    populates the filter-value dropdown. Dashboard charts still reference _ChartData rows
    and update automatically when the slicer hides rows."""
    try:
        ws_dash = sh.worksheet("Dashboard")
        dash_id = ws_dash.id
    except gspread.exceptions.WorksheetNotFound:
        dash_id = None

    # Delete existing slicers on BOTH Dashboard (old location) and _ChartData
    spreadsheet = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    requests = []
    deleted = 0
    for sheet in spreadsheet.get("sheets", []):
        sid = sheet["properties"]["sheetId"]
        if sid in (dash_id, ws_cd_id):
            for s in sheet.get("slicers", []):
                requests.append({"deleteEmbeddedObject": {"objectId": s["slicerId"]}})
                deleted += 1

    # Unhide _ChartData so it's visible and usable for filtering
    requests.append({"updateSheetProperties": {
        "properties": {"sheetId": ws_cd_id, "hidden": False},
        "fields": "hidden",
    }})

    # Slicer placed ON _ChartData — same sheet as dataRange — so dropdown shows months.
    # Cross-sheet slicers (slicer on Dashboard, data on _ChartData) don't populate values.
    requests.append({"addSlicer": {
        "slicer": {
            "spec": {
                "dataRange": {
                    "sheetId":          ws_cd_id,
                    "startRowIndex":    1,           # skip header row
                    "endRowIndex":      n_months + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex":   7,
                },
                "columnIndex": 0,
                "title": "Filter by Month",
                "backgroundColor": {"red": 0.13, "green": 0.13, "blue": 0.13},
                "textFormat": {
                    "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "bold": True,
                    "fontSize": 10,
                },
            },
            "position": {
                "overlayPosition": {
                    "anchorCell": {
                        "sheetId":     ws_cd_id,
                        "rowIndex":    0,
                        "columnIndex": 9,   # col J — next to donut data
                    },
                    "widthPixels":  260,
                    "heightPixels": 56,
                }
            }
        }
    }})

    service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": requests},
    ).execute()
    print(f"  Month slicer on _ChartData (deleted {deleted} old slicer(s)).")


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

    # Keep _ChartData visible so the Dashboard slicer can read month values
    cd_ws = next((w for w in worksheets if w.title == "_ChartData"), None)
    if cd_ws:
        requests_list.append({"updateSheetProperties": {
            "properties": {"sheetId": cd_ws.id, "hidden": False},
            "fields": "hidden",
        }})

    # For each data tab: freeze header row, auto-resize columns, banded rows
    data_tabs = [w for w in worksheets if w.title != "Dashboard"]
    GREY_LIGHT = {"red": 0.95, "green": 0.95, "blue": 0.95}
    WHITE      = {"red": 1.0,  "green": 1.0,  "blue": 1.0}

    for ws in data_tabs:
        sid = ws.id
        col_count = len(MONTH_HEADERS) if month_pattern.match(ws.title) else len(ALL_DATA_HEADERS)

        # Clear all existing cell borders
        requests_list.append({"updateBorders": {
            "range": {
                "sheetId": sid,
                "startRowIndex": 0,
                "endRowIndex": 10000,
                "startColumnIndex": 0,
                "endColumnIndex": col_count,
            },
            "top":             {"style": "NONE"},
            "bottom":          {"style": "NONE"},
            "left":            {"style": "NONE"},
            "right":           {"style": "NONE"},
            "innerHorizontal": {"style": "NONE"},
            "innerVertical":   {"style": "NONE"},
        }})

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
    ws_cd, mom_rows, n_sku_series, damage_reason_col, \
        damaged_sku_col, n_damaged_skus, missing_sku_col, n_missing_skus, \
        loss_col = build_summary_data(sh)
    if ws_cd and mom_rows:
        create_dashboard_charts(service, sh, ws_cd.id, mom_rows, n_sku_series,
                                 damage_reason_col,
                                 damaged_sku_col, n_damaged_skus,
                                 missing_sku_col,  n_missing_skus,
                                 loss_col)
        create_month_slicer(service, sh, ws_cd.id, len(mom_rows))

    # ── 6. Polish the sheet ───────────────────────────────────────────────────
    print("\nPolishing sheet...")
    polish_sheet(service, sh)

    print(f"\nDone! https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")


if __name__ == "__main__":
    main()
