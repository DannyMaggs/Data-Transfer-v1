import warnings
from urllib3.exceptions import InsecureRequestWarning

warnings.simplefilter('ignore', InsecureRequestWarning)

import os
import sys
import requests
from msal import ConfidentialClientApplication
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from dotenv import load_dotenv

load_dotenv()

config = {
    "client_id": os.getenv("CLIENT_ID"),
    "authority": os.getenv("AUTHORITY"),
    "client_secret": os.getenv("CLIENT_SECRET"),
    "scopes": [os.getenv("SCOPES")]
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

from pptx.util import Inches, Pt

def update_powerpoint(ppt_path, data):
    # Trim data to fit within 10 rows and 9 columns
    trimmed_data = [row[:9] for row in data[:10]]

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
    for i, row in enumerate(table.rows):
        if i == 0:
            continue  # Skip header row
        for cell in row.cells:
            cell.text = ""

    # Fill in new data
    for i, row_data in enumerate(trimmed_data):
        for j, cell_data in enumerate(row_data):
            cell = table.cell(i + 1, j)
            cell.text = str(cell_data)

    # Adjust table size
    table_shape = table._graphic_frame
    table_shape.width = Inches(9)  # Adjust width as needed
    table_shape.height = Inches(5)  # Adjust height as needed

    # Set font size for all cells
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(12)

    presentation.save(ppt_path)




def main():
    if len(sys.argv) != 3:
        print("Usage: python update_ppt.py <sourceFileId> <destinationFileId>")
        sys.exit(1)

    source_file_id = sys.argv[1]
    destination_file_id = sys.argv[2]

    access_token = get_token()
    site_id = os.getenv("SITE_ID")
    
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
