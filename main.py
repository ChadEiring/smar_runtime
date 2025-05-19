import smartsheet
from datetime import datetime, timedelta
import os
import dotenv
dotenv.load_dotenv()
tokenId=os.getenv("TOKEN")
sheetId=os.getenv("SHEET_ID")
if tokenId and "|" in tokenId:
    token = tokenId.split("|")[0]
    sheetId = tokenId.split("|")[1]

# Create a Smartsheet client.
ss_client = smartsheet.Smartsheet(token)

# Get the sheet
sheet = ss_client.Sheets.get_sheet(sheetId)

# Create a column map for easy access to column IDs
column_map = {col.title: col.id for col in sheet.columns}

# Find any rows with column Needs to be moved checked
rows_to_move = []
for row in sheet.rows:
    for cell in row.cells:
        if cell.column_id == column_map['Needs to be moved']:
            if cell.value == True:
                rows_to_move.append(row)

# Move each row under the parent row based on the value in the column named Area matching the column named Customer
for row in rows_to_move:
    # Get the value of the Area column
    area_value = None
    for cell in row.cells:
        if cell.column_id == column_map['Area']:
            area_value = cell.value
            break
    # Find the parent row based on the Area value
    parent_row = None
    for r in sheet.rows:
        for cell in r.cells:
            if cell.column_id == column_map['Customer'] and cell.value == area_value:
                parent_row = r
                break
        if parent_row:
            break
    # If the parent row is found, move the row under it
    # Get the value of the Customer column
    customer_value = None
    for cell in row.cells:
        if cell.column_id == column_map['Customer']:
            customer_value = cell.value
            break
    # If the customer value is not found, skip this row
    if customer_value is None:
        print(f"Customer value not found for row {row.id}. Skipping this row.")
        continue
    # If the area value is not found, skip this row
    if area_value is None:
        print(f"Area value not found for row {row.id}. Skipping this row.")
        continue
    # If the parent row is not found, skip this row
    if parent_row is None:
        print(f"Parent row not found for row {row.id} with area {area_value}. Skipping this row.")
        continue
    # Print the row and parent row details
    print(f"Row ID: {row.id}, Customer: {customer_value}, Area: {area_value}")
    print(f"Parent Row ID: {parent_row.id}, Customer: {customer_value}, Area: {area_value}")

    # Move the row under the parent row
    # Create a new row object with the same values as the original row
    new_row = smartsheet.models.Row()
    new_row.id = row.id
    new_row.parent_id = parent_row.id
    new_row.to_bottom = True
    
    updated_row = ss_client.Sheets.update_rows(
    sheetId,      # sheet_id
    [new_row])

