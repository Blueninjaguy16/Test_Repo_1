import os
from dotenv import load_dotenv
import smartsheet
import pandas as pd
from openpyxl.utils import get_column_letter

# === LOAD ENV VARIABLES ===
load_dotenv()
ACCESS_TOKEN = os.getenv("SMARTSHEET_API_TOKEN")
PCM_PROCESSING_ID = int(os.getenv("PCM_PROCESSING_ID"))
PCM_REQUESTS_ID = int(os.getenv("PCM_REQUESTS_ID"))

# === DESIRED COLUMNS FOR EXPORT ===
DESIRED_COLUMNS = {
    "PCM Processing": ["Project ID", "Distributing to Retailer SKU Count", "Data Contract Name", "Retailer", "Assigned To", "Request Type", "Status", "Distributing Brand", "Priority Level"],
    "PCM Requests": ["Project ID", "SKU Count", "Request Description", "Dealer", "Assigned to PIM", "Assigned to PCM", "Request Type", "Status", "Brand", "High Priority", "Date Requested"]
}

# === INIT CLIENT ===
smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

# === FUNCTION TO FETCH AND FILTER ROWS ===
def get_filtered_df(sheet_id, sheet_name):
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    column_map = {col.id: col.title for col in sheet.columns}
    status_col_id = next((col.id for col in sheet.columns if col.title == "Status"), None)

    if not status_col_id:
        raise Exception(f"'Status' column not found in sheet '{sheet_name}'")

    rows_data = []
    for row in sheet.rows:
        status_cell = next((cell for cell in row.cells if cell.column_id == status_col_id), None)
        status = str(status_cell.value).strip().lower() if status_cell and status_cell.value else ""
        if status not in ("complete", "canceled", "cancelled", "submission error", "submissions error - channel"):
            row_dict = {column_map[cell.column_id]: cell.value for cell in row.cells}
            row_dict["Source Sheet"] = sheet_name
            rows_data.append(row_dict)

    df = pd.DataFrame(rows_data)

    # === CUSTOM FILTERING ===
    if sheet_name == "PCM Processing":
        if "Retailer" in df.columns:
            allowed_retailers = ["Rona", "Home Hardware", "Wayfair CAN"]
            df = df[df["Retailer"].astype(str).str.strip().isin(allowed_retailers)]

    elif sheet_name == "PCM Requests":
        if "Dealer" in df.columns:
            allowed_dealers = ["Home Depot CA", "Home Hardware", "Rona"]
            df = df[df["Dealer"].astype(str).str.strip().isin(allowed_dealers)]

    # === COLUMN TRIMMING ===
    desired_columns = DESIRED_COLUMNS.get(sheet_name, [])
    df = df[[col for col in desired_columns if col in df.columns]]


    return df

# === FETCH DATA FROM BOTH SHEETS ===
df_processing = get_filtered_df(PCM_PROCESSING_ID, "PCM Processing")
df_requests = get_filtered_df(PCM_REQUESTS_ID, "PCM Requests")

# === COMBINE AND EXPORT TO EXCEL (SEPARATE TABS) ===
output_path = "C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports/CAN-open-projects.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for sheet_name, df in [("PCM Processing", df_processing), ("Digital Commerce", df_requests)]:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        for col_idx, column in enumerate(df.columns, 1):
            worksheet.column_dimensions[get_column_letter(col_idx)].width = 20


print(f"Exported {len(df_processing)} PCM Processing rows and {len(df_requests)} PCM Requests rows to {output_path}")
