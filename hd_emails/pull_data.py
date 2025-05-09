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
            continue

        if request_date > cutoff_date:
            continue  # skip if too recent
    except Exception:
        continue  # skip invalid date formats

    data.append(row_dict)

# === EXPORT TO EXCEL ===
df = pd.DataFrame(data)
today_str = datetime.now().strftime("%Y-%m-%d")
filename = f"HD_Update_Needed_{today_str}.xlsx"
output_path = os.path.join(EXPORT_FOLDER, filename)
df.to_excel(output_path, index=False)
print(f"\n✅ Excel file saved: {output_path}")
