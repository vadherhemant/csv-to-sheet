import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials as GoogleCredentials
import copy

# CONFIG
CSV_URL = "https://raw.githubusercontent.com/vadherhemant/csv-to-sheet/refs/heads/main/source.csv"
SPREADSHEET_ID = "1gpV5T5Ol45VqmS8nI6Xk2MXWEeJMiXU1yoFUDMODi6g"
CREDENTIALS_FILE = "creds.json"

# Constants
START_COL = 9  # Column J (0-indexed)
BLOCK_WIDTH = 9  # 7 data + 2 gap

def convert_cell(val):
    try:
        f = float(val)
        if f.is_integer():
            return int(f)
        return f
    except:
        return str(val).strip()

try:
    # Setup Google Sheets APIs
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
    scoped_creds = GoogleCredentials.from_service_account_file(CREDENTIALS_FILE, scopes=scope)
    service = build("sheets", "v4", credentials=scoped_creds)

    # Read the CSV (assuming it has one block only)
    raw = pd.read_csv(CSV_URL, header=None)
    if raw.shape[1] < 8 or raw.shape[0] < 2:
        raise ValueError("CSV block is incomplete or malformed.")

    # Parse top label and headers
    date_label = str(raw.iloc[1, 0]).strip()
    header_row = [str(raw.iloc[0, i]).strip() for i in range(1, 8)]
    
    # Parse data
    data_rows = []
    for i in range(1, raw.shape[0]):
        if pd.isna(raw.iloc[i, 0]):
            break
        row = [convert_cell(raw.iloc[i, j]) for j in range(1, 8)]
        data_rows.append(row)

    max_height = len(data_rows)
    
    # Read existing sheet content
    existing_data = sheet.get_all_values()
    while len(existing_data) < max_height + 2:
        existing_data.append([])

    # Extend all rows to match existing + new block width
    for i in range(len(existing_data)):
        while len(existing_data[i]) < START_COL:
            existing_data[i].append("")

    # Shift existing content to the right
    for i in range(len(existing_data)):
        row = existing_data[i]
        old_tail = row[START_COL:]
        gap = [""] * BLOCK_WIDTH
        row[START_COL:] = gap + old_tail

    # Insert top label row
    if len(existing_data[0]) < START_COL + 7:
        existing_data[0] += [""] * (START_COL + 7 - len(existing_data[0]) + 2)
    existing_data[0][START_COL] = date_label

    # Insert header row
    for j in range(7):
        existing_data[1][START_COL + j] = header_row[j]

    # Insert data
    for r in range(max_height):
        for c in range(7):
            existing_data[r + 2][START_COL + c] = data_rows[r][c]

    # Update sheet without clearing
    sheet.update("A1", existing_data)

    # Merge date header
    requests = [{
        "mergeCells": {
            "range": {
                "sheetId": sheet._properties["sheetId"],
                "startRowIndex": 0,
                "endRowIndex": 1,
                "startColumnIndex": START_COL,
                "endColumnIndex": START_COL + 7
            },
            "mergeType": "MERGE_ALL"
        }
    }]

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests}
    ).execute()



#FORMATTING------------------------

# Add conditional formatting for B vs I, C vs J, D vs K (rows 3–17)
try:
    for col_offset, pair in enumerate([("B", "I"), ("C", "J"), ("D", "K")]):
        col_left, col_right = pair
        row_start = 3
        row_end = 18  # 3 to 17 inclusive
        col_left_idx = ord(col_left) - ord("A")
        col_right_idx = ord(col_right) - ord("A")

        # Red background for mismatch
        red_rule = ConditionalFormatRule(
            ranges=[GridRange(
                sheetId=sheet_id,
                startRowIndex=row_start - 1,
                endRowIndex=row_end,
                startColumnIndex=col_left_idx,
                endColumnIndex=col_left_idx + 1
            )],
            booleanRule=BooleanRule(
                condition=BooleanCondition(
                    type='CUSTOM_FORMULA',
                    values=[{'userEnteredValue': f'=${col_left}{row_start}<>${col_right}{row_start}'}]
                ),
                format=CellFormat(backgroundColor=color(1, 0.8, 0.8))
            )
        )

        # Green background for match
        green_rule = ConditionalFormatRule(
            ranges=[GridRange(
                sheetId=sheet_id,
                startRowIndex=row_start - 1,
                endRowIndex=row_end,
                startColumnIndex=col_left_idx,
                endColumnIndex=col_left_idx + 1
            )],
            booleanRule=BooleanRule(
                condition=BooleanCondition(
                    type='CUSTOM_FORMULA',
                    values=[{'userEnteredValue': f'=${col_left}{row_start}=${col_right}{row_start}'}]
                ),
                format=CellFormat(backgroundColor=color(0.8, 1, 0.8))
            )
        )

        rules.append(red_rule)
        rules.append(green_rule)

    rules.save()

except Exception as e:
    print(f"Error during conditional formatting: {e}")


# 5. Resize columns I–N to 50% of current width
# First: get current widths (we'll just set a fixed scaled width)
# Example: if default is 100px -> set 50px
set_column_widths(ws, [("I:N", 50)])  # adjust scale value

print("✔️ Conditional formatting and resizing applied!")

    print(f"✅ New block inserted at column {chr(START_COL + 65)} successfully.")

except Exception as e:
    print(f"❌ ERROR: {e}")
    raise
