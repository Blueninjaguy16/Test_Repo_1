import os
from dotenv import load_dotenv
import smartsheet
import pandas as pd
from datetime import datetime, timedelta

# === LOAD API TOKEN FROM .env FILE ===
load_dotenv()
ACCESS_TOKEN = os.getenv('SMARTSHEET_API_TOKEN')

# === INITIALIZE SMARTSHEET CLIENT ===
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

# === CONFIGURATION ===
SHEET_ID = 1473435297337220
DEALER_COLUMN_NAME = "Dealer"
EXPORT_FOLDER = r"C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports"

# === FETCH SHEET ===
sheet = smartsheet_client.Sheets.get_sheet(SHEET_ID)

# === IDENTIFY COLUMN IDS ===
column_map = {col.title: col.id for col in sheet.columns}
dealer_col_id = column_map.get(DEALER_COLUMN_NAME)
date_requested_col_id = column_map.get("Date Requested")

if not dealer_col_id or not date_requested_col_id:
    raise Exception("Missing required column(s): 'Dealer' or 'Date Requested'")

# === FILTER ROWS WHERE 'Dealer' IS 'Home Depot' ===
matching_rows = []
for row in sheet.rows:
    for cell in row.cells:
        if cell.column_id == dealer_col_id and cell.value == "Home Depot":
            matching_rows.append(row)
            break

print(f"✅ Found {len(matching_rows)} rows for 'Home Depot'")

# === MAP COLUMN IDS TO TITLES ===
column_id_map = {col.id: col.title for col in sheet.columns}

# === BUILD LIST OF DICTS FOR DATAFRAME, FILTERING BY STATUS & DATE ===
data = []
cutoff_date = datetime.now() - timedelta(days=14)

for row in matching_rows:
    row_dict = {column_id_map.get(cell.column_id, f"Col_{cell.column_id}"): cell.value for cell in row.cells}

    # === FILTER BY STATUS ===
    status = str(row_dict.get("Status") or "").strip().lower()
    if status in ["complete", "cancelled", "submission error"]:
        print(f"⏭️ Row {row.id} skipped — status: {status}")
        continue

    # === FILTER BY DATE REQUESTED ===
    date_requested = None
    for cell in row.cells:
        if cell.column_id == date_requested_col_id:
            date_requested = cell.display_value or cell.value
            break

    try:
        if isinstance(date_requested, str):
            try:
                request_date = datetime.strptime(date_requested.strip(), "%m/%d/%Y")
            except ValueError:
                request_date = datetime.strptime(date_requested.strip(), "%Y-%m-%d")
        elif isinstance(date_requested, datetime):
            request_date = date_requested
        elif hasattr(date_requested, 'isoformat'):
            request_date = datetime.combine(date_requested, datetime.min.time())
        else:
            print(f"⏭️ Row {row.id} skipped — unknown date format: {date_requested}")
            continue

    except Exception as e:
        print(f"⏭️ Row {row.id} skipped — date parsing error: {e}")
        continue

    # === FETCH ALL COMMENTS WITH AUTHOR ===
    all_comments = []
    latest_timestamp_raw = None
    latest_author = ""

    try:
        discussions = smartsheet_client.Discussions.get_row_discussions(
            sheet_id=SHEET_ID,
            row_id=row.id,
            include="comments"
        ).data

        for discussion in discussions:
            for comment in discussion.comments:
                author = comment.created_by.name if comment.created_by else "Unknown"
                comment_time = comment.created_at
                comment_text = comment.text.strip().replace('\n', ' ')

                # Track latest timestamp
                if not latest_timestamp_raw or comment_time > latest_timestamp_raw:
                    latest_timestamp_raw = comment_time
                    latest_author = author

                all_comments.append(f"{author} [{comment_time.strftime('%Y-%m-%d %H:%M:%S')}]: {comment_text}")


        if not latest_timestamp_raw:
            print(f"⏭️ Row {row.id} skipped — no comments found")
            continue

        latest_timestamp = latest_timestamp_raw.strftime('%Y-%m-%d %H:%M:%S')
        all_comments_str = "; ".join(all_comments)

    except Exception as e:
        print(f"⏭️ Row {row.id} skipped — error fetching comments: {e}")
        continue



    row_dict["All Comments"] = all_comments_str
    row_dict["Comment Author"] = latest_author
    row_dict["Comment Date"] = latest_timestamp


    print(f"✅ Row {row.id} added — comment date: {latest_timestamp}")
    data.append(row_dict)

# === EXPORT TO EXCEL ===
df = pd.DataFrame(data)
today_str = datetime.now().strftime("%Y-%m-%d")
filename = f"TEST_{today_str}.xlsx"
output_path = os.path.join(EXPORT_FOLDER, filename)

if not data:
    print("\n⚠️ No rows passed all filters — Excel file will be empty.")

df.to_excel(output_path, index=False)
print(f"\n✅ Excel file saved: {output_path}")
