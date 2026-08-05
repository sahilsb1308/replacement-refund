"""
Update only the SKU Report tab using data already in the sheet.
No Shopify fetch — reads from 'All Data' tab directly.

Usage:
  python update_sku_report.py
"""

import os
import sys
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Import all logic from the main sync script
sys.path.insert(0, os.path.dirname(__file__))
from sync_to_sheet import (
    SHEET_ID, CREDS_FILE, SCOPES,
    build_summary_data, write_sku_report_tab,
    create_dashboard_charts, create_month_slicer,
)


def main():
    print("Connecting to Google Sheets...")
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    gc      = gspread.authorize(creds)
    sh      = gc.open_by_key(SHEET_ID)

    print("Building SKU data from All Data tab...")
    result = build_summary_data(sh)
    if result[0] is None:
        print("No data found in All Data tab. Run a full sync first.")
        return

    (ws_cd, mom_rows, n_sku_series, damage_reason_col,
     damaged_sku_col, n_damaged_skus,
     missing_sku_col, n_missing_skus,
     used_sku_col,    n_used_skus,
     wrong_sku_col,   n_wrong_skus,
     loss_col,
     sorted_months,
     top_skus,        sku_pivot_data,
     top_damaged_skus, damaged_sku_pivot_data,
     top_missing_skus, missing_sku_pivot_data,
     top_used_skus,   used_sku_pivot_data,
     top_wrong_skus,  wrong_sku_pivot_data,
     loss_labels,     loss_pivot_data) = result

    service = build("sheets", "v4", credentials=creds)

    print("Rebuilding Dashboard charts...")
    create_dashboard_charts(service, sh, ws_cd.id, mom_rows, n_sku_series,
                             damage_reason_col,
                             damaged_sku_col, n_damaged_skus,
                             missing_sku_col,  n_missing_skus,
                             used_sku_col,    n_used_skus,
                             wrong_sku_col,   n_wrong_skus,
                             loss_col)
    create_month_slicer(service, sh, ws_cd.id, len(mom_rows))

    print("Writing SKU Report tab...")
    write_sku_report_tab(
        sh, gc, sorted_months,
        top_skus, sku_pivot_data,
        top_damaged_skus, damaged_sku_pivot_data,
        top_missing_skus, missing_sku_pivot_data,
        top_used_skus,   used_sku_pivot_data,
        top_wrong_skus,  wrong_sku_pivot_data,
        loss_labels,     loss_pivot_data,
    )

    print(f"\nDone! https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")


if __name__ == "__main__":
    main()
