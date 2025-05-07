import os
from dotenv import load_dotenv
import smartsheet
import csv
from datetime import datetime

# --- CONFIGURATION ---
load_dotenv()
ACCESS_TOKEN = os.getenv('SMARTSHEET_API_TOKEN')
SHEET_ID = 1473435297337220  # Replace with your actual Sheet ID
OUTPUT_CSV = 'C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports/PCM_Requests_Summary.csv'

# --- INIT CLIENT ---
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)

# --- GET SHEET ---
sheet = smartsheet_client.Sheets.get_sheet(SHEET_ID)

# --- FIND COLUMN IDS FOR NEEDED FIELDS ---
column_map = {col.title: col.id for col in sheet.columns}
project_id_col_id = column_map.get("Project ID")
date_requested_col_id = column_map.get("Date Requested")
status_col_id = column_map.get("Status")
request_description_col_id = column_map.get("Request Description")

# --- BUILD ROW DATA WITH COMMENTS ---
output_rows = []

for row in sheet.rows:
    def get_cell_value(col_id):
        return next((cell.display_value for cell in row.cells if cell.column_id == col_id), '')

    project_id = get_cell_value(project_id_col_id)
    if len(row.cells) > 3:
        cell = row.cells[3]
        date_requested = cell.display_value or cell.value or ''
        if isinstance(date_requested, (datetime, )):
            date_requested = date_requested.strftime('%Y-%m-%d')
    else:
        date_requested = ''
    status = get_cell_value(status_col_id)
    request_description = get_cell_value(request_description_col_id)

    if status in ("Complete","Cancelled","submission error","Submissions Error - Channel","Submission Error"):
        continue  # Skip rows with status "Complete" or "Cancelled"

    # Inside the for row in sheet.rows loop:
    print(f"Row {row.id} — Project ID: {project_id}")

    try:
        discussions = smartsheet_client.Discussions.get_row_discussions(
            sheet_id=SHEET_ID,
            row_id=row.id,
            include="comments"
        ).data
    except Exception as e:
        print(f"Error getting discussions for row {row.id}: {e}")
        discussions = []

    if discussions:
        print(f"  Found {len(discussions)} discussions")
    else:
        print("  No discussions found")


    # Find most recent comment if exists
    latest_comment = ''
    latest_author = ''
    latest_timestamp = ''

    for discussion in discussions:
        for comment in discussion.comments:
            if not latest_timestamp or comment.created_at > datetime.fromisoformat(latest_timestamp):
                latest_comment = comment.text
                latest_author = comment.created_by.name
                latest_timestamp = comment.created_at.isoformat()

    if latest_timestamp:
        latest_timestamp = datetime.fromisoformat(latest_timestamp).strftime('%Y-%m-%d %H:%M:%S')

    output_rows.append([
        project_id,
        date_requested,
        status,
        request_description,
        latest_comment,
        latest_author,
        latest_timestamp
    ])

# --- WRITE TO CSV ---
with open(OUTPUT_CSV, mode='w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow([
        'Project ID',
        'Date Requested',
        'Status',
        'Request Description',
        'Most Recent Comment',
        'Comment Author',
        'Comment Timestamp'
    ])
    writer.writerows(output_rows)

print(f"✅ Export complete: {OUTPUT_CSV}")
