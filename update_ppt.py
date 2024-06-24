import os
import sys
import requests
from msal import ConfidentialClientApplication
from openpyxl import load_workbook
from pptx import Presentation

config = {
    "client_id": "3acd75e1-dbf0-4df0-88aa-2c7a4bd5ee8b",
    "authority": "https://login.microsoftonline.com/7f65e0c2-5159-471c-9af9-e57501d53752",
    "client_secret": "MlC8Q~XZ_vLrsVb4E_afMEwZVKjQBk41PjIhObS0",
    "scopes": ["https://graph.microsoft.com/.default"]
}

def get_token():
    app = ConfidentialClientApplication(
        config["client_id"], authority=config["authority"], client_credential=config["client_secret"]
    )
    result = app.acquire_token_for_client(scopes=config["scopes"])
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Could not obtain access token")

def download_file(access_token, site_id, item_id, file_path):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/content"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }
    print(f"Downloading file from URL: {url}")
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    with open(file_path, "wb") as file:
        file.write(response.content)

def upload_file(access_token, site_id, item_id, file_path):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/content"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    }
    with open(file_path, "rb") as file:
        print(f"Uploading file to URL: {url}")
        response = requests.put(url, headers=headers, data=file)
    response.raise_for_status()

def read_excel_data(file_path, sheet_name, start_row, end_row):
    workbook = load_workbook(filename=file_path, data_only=True)
    sheet = workbook[sheet_name]
    data = []
    for row in sheet.iter_rows(min_row=start_row, max_row=end_row, values_only=True):
        data.append(row)
    return data

def update_powerpoint(ppt_path, data):
    presentation = Presentation(ppt_path)
    slide = presentation.slides[5]  # Slide 6 (0-indexed)

    table = None
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            break

    if not table:
        raise Exception("No table found on slide 6")

    # Clear existing table data
    for row in table.rows[1:]:  # Assuming first row is header
        for cell in row.cells:
            cell.text = ""

    # Fill in new data
    for i, row_data in enumerate(data, start=1):
        row = table.rows[i]
        for j, cell_data in enumerate(row_data):
            row.cells[j].text = str(cell_data)

    presentation.save(ppt_path)

def main():
    if len(sys.argv) != 3:
        print("Usage: python update_ppt.py <sourceFileId> <destinationFileId>")
        sys.exit(1)

    source_file_id = sys.argv[1]
    destination_file_id = sys.argv[2]

    access_token = get_token()
    site_id = "motohaus.sharepoint.com,2c54175f-ca53-4ca4-8eab-1530b7f64072,07a87623-8134-4e67-b860-0ff98346cfc2"
    
    excel_path = "source.xlsx"
    ppt_path = "destination.pptx"

    # Debug print statements
    print(f"Access Token: {access_token}")
    print(f"Site ID: {site_id}")
    print(f"Source File ID: {source_file_id}")
    print(f"Destination File ID: {destination_file_id}")

    # Download files from SharePoint
    download_file(access_token, site_id, source_file_id, excel_path)
    download_file(access_token, site_id, destination_file_id, ppt_path)

    # Read data from Excel
    excel_data = read_excel_data(excel_path, "For Monthly Reports", 13, 21)

    # Update PowerPoint with Excel data
    update_powerpoint(ppt_path, excel_data)

    # Upload updated PowerPoint back to SharePoint
    upload_file(access_token, site_id, destination_file_id, ppt_path)

if __name__ == "__main__":
    main()
