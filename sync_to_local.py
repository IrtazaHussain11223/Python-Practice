import os
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
import openpyxl

# --- Configuration ---
SPREADSHEET_ID = os.environ.get("RESEARCH_SHEET_ID", "1afcXQNMrfAHvZWV0yf0qrBKczmxAgqsIfWKmOq6rE_E")
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "research-tracker-sync-515210-f4571a7380d2.json")
LOCAL_XLSX_PATH = os.environ.get("LOCAL_TRACKER_XLSX", "Research_Opportunity_Tracker.xlsx")
SHEET_TAB_NAME = "Outreach Tracker"
START_ROW = 5  # Row 4 is headers; data begins at row 5

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


def fetch_sheet_rows(spreadsheet_id: str, tab_name: str, service_account_file: str) -> list[list]:
    creds = service_account.Credentials.from_service_account_file(
        service_account_file, scopes=SCOPES
    )
    service = build("sheets", "v4", credentials=creds)
    range_name = f"'{tab_name}'!A{START_ROW}:P200"
    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=range_name)
        .execute()
    )
    return result.get("values", [])


def parse_date_if_possible(val):
    if not val or not isinstance(val, str):
        return val
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(val.strip(), fmt).date()
        except ValueError:
            continue
    return val


def sync_xlsx(rows: list[list], xlsx_path: str) -> int:
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb[SHEET_TAB_NAME]

    total_rows_to_clear = max(len(rows), ws.max_row - START_ROW + 1)
    updated_count = 0

    for idx in range(total_rows_to_clear):
        r = START_ROW + idx
        row_vals = rows[idx] if idx < len(rows) else []

        # Pad to 16 columns (A:P)
        padded = list(row_vals) + [None] * (16 - len(row_vals))

        # Check if row has any meaningful content (Prof Name or other fields)
        has_data = any(
            padded[c] is not None and str(padded[c]).strip() != ""
            for c in range(1, 14)  # B through M
        ) or (padded[15] is not None and str(padded[15]).strip() != "")

        if has_data:
            updated_count += 1

        for col_idx in range(1, 17):
            # Col 15 (O) is "Days Since Contact" formula — preserve formula so it stays dynamic
            if col_idx == 15:
                if not ws.cell(row=r, column=15).value:
                    ws.cell(
                        row=r,
                        column=15,
                        value=f'=IF(N{r}<>"", TODAY()-N{r}, IF(M{r}<>"", TODAY()-M{r}, ""))',
                    )
                continue

            val = padded[col_idx - 1]
            # Convert empty strings to None for clean Excel cells
            if isinstance(val, str) and val.strip() == "":
                val = None

            # Parse date columns: Col 13 (M: Initial Email Date) and Col 14 (N: Follow-Up 1 Date)
            if col_idx in (13, 14) and val is not None:
                val = parse_date_if_possible(val)

            # Default status on blank pre-formatted rows
            if col_idx == 12 and not has_data:
                val = "Not Started"

            cell = ws.cell(row=r, column=col_idx, value=val)

            # Re-attach clickable hyperlinks for Email (Col 8) and URLs (Col 9, Col 10)
            if isinstance(val, str):
                s = val.strip()
                if col_idx == 8 and "@" in s:
                    cell.hyperlink = f"mailto:{s}"
                elif col_idx in (9, 10) and s.startswith("http"):
                    cell.hyperlink = s

    wb.save(xlsx_path)
    return updated_count


def main():
    if SPREADSHEET_ID == "YOUR_GOOGLE_SHEET_ID_HERE":
        raise ValueError("Set RESEARCH_SHEET_ID environment variable or edit SPREADSHEET_ID in the script.")
    rows = fetch_sheet_rows(SPREADSHEET_ID, SHEET_TAB_NAME, SERVICE_ACCOUNT_FILE)
    count = sync_xlsx(rows, LOCAL_XLSX_PATH)
    print(f"[{datetime.now().isoformat(timespec='seconds')}] Synced {count} active entries to {LOCAL_XLSX_PATH}")


if __name__ == "__main__":
    main()