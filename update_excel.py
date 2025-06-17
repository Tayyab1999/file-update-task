import requests
import pandas as pd
import os
import openpyxl
from copy import copy

def download_large_file(file_id, destination):
    """
    Download a large file from Google Drive
    """
    url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t"
    response = requests.get(url, stream=True)
    
    if response.status_code == 200:
        with open(destination, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        print(f"File successfully downloaded to {destination}")
        return True
    else:
        print(f"Failed to download file. Status code: {response.status_code}")
        return False

def update_excel_data(file_path, updates):
    """
    Update Excel file with new values in the AMS NFL sheet
    """
    try:
        # Load the workbook first
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook['AMS NFL']
        
        # Find the ASIN column
        asin_col = None
        asin_row = None
        target_asin = updates.get('asin')
        
        # Find headers first (usually in row 2 based on the file structure)
        headers = {}
        for cell in sheet[2]:  # Row 2
            if cell.value:
                headers[cell.value] = cell.column_letter
        
        if 'ASIN' not in headers:
            print("ASIN column not found in headers")
            return False
            
        # Find the row with our target ASIN
        asin_col = headers['ASIN']
        for row in range(3, sheet.max_row + 1):  # Start from row 3
            if sheet[f"{asin_col}{row}"].value == target_asin:
                asin_row = row
                break
        
        if not asin_row:
            print(f"ASIN {target_asin} not found in sheet")
            return False
            
        print(f"\nFound ASIN {target_asin} in row {asin_row}")
        
        # Map of fields to update
        field_mapping = {
            'FBA INV': 'FBA INV',
            'OTS QOH': 'OTS QOH',
            'QOHQTY': 'QOHQTY',
            'AMZ VC INV': 'AMZ VC INV',
            'WIP QTY': 'WIP QTY',
            'WIP ETA': 'WIP ETA'
        }
        
        # Update each field
        for field, column_name in field_mapping.items():
            if field in updates and column_name in headers:
                col_letter = headers[column_name]
                cell = sheet[f"{col_letter}{asin_row}"]
                cell.value = updates[field]
                print(f"Updated {column_name} to {updates[field]}")
        
        # Save the workbook
        workbook.save(file_path)
        print(f"\nSaved updates to file: {file_path}")
        return True
            
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        return False

def main():
    # Your file ID from Google Drive
    file_id = "11xfaf0nGgscOpYpE9U3tfRelyCxsDoOJ"
    excel_file = "ICERWORKSHEET.xlsx"
    
    # Updates for specific ASIN
    updates = {
        'asin': 'B084TRRKBY',
        'FBA INV': 98,
        'OTS QOH': 123,
        'QOHQTY': 456,
        'AMZ VC INV': 789,
        'WIP QTY': 10,
        'WIP ETA': '2024-07-01'
    }
    
    # Download and update the file
    if download_large_file(file_id, excel_file):
        update_excel_data(excel_file, updates)

if __name__ == "__main__":
    main() 