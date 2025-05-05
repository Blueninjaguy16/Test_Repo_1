import os
from dotenv import load_dotenv
import smartsheet

# Load environment variables from .env file
load_dotenv()
ACCESS_TOKEN = os.getenv('SMARTSHEET_API_TOKEN')
# Initialize client
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

# Get the sheet
sheet_id = 1473435297337220  # Replace with your actual sheet ID

sheet = smartsheet_client.Sheets.get_sheet(sheet_id)

# Find the column ID for "Project ID"
project_id_column = next(
    (col for col in sheet.columns if col.title == "Project ID"),
    None
)

if project_id_column is None:
    raise ValueError("Column 'Project ID' not found in sheet.")

project_col_id = project_id_column.id
bottom_rows = sheet.rows[-10:]  # Get the last 10 rows of the sheet

# Extract only "Project ID" values
for row in bottom_rows:
    for cell in row.cells:
        if cell.column_id == project_col_id:
            print(cell.value)
