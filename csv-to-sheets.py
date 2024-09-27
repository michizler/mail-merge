# Import required libraries

import os
import json
from google.oauth2 import service_account
import csv
import gspread

# Load JSON key file
creds = service_account.Credentials.from_service_account_file(
    'C:/Users/BRIGHT UZOSIKE/Downloads/rtls-242692-129b9a57503a.json',
    scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
)

# Authenticate with Google Sheets API
client = gspread.authorize(creds)

# Open the CSV file
with open('C:/Users/BRIGHT UZOSIKE/Downloads/Sample-Spreadsheet-10-rows.csv', 'r') as csvfile:
    reader = csv.reader(csvfile)
    data = list(reader)

# Create a new Google Sheet or open an existing one
spreadsheet = client.create('mail-merge-test')  # or client.open('Existing Sheet').sheet1

spreadsheet.share('brightuzosike@gmail.com', perm_type='user', role='writer')

spreadsheet_instance = client.open('mail-merge-test')

# Create a new sheet
worksheet = spreadsheet.add_worksheet(title='mail-merge', rows=100, cols=100)

# Convert data to list of Cell objects
cell_list = []
for row, row_data in enumerate(data, start=1):
    for col, value in enumerate(row_data, start=1):
        cell_list.append(gspread.Cell(row, col, value))

# Write the CSV data to the Google Sheet
worksheet.update_cells(cell_list)

print(f"Google Sheet URL: {spreadsheet_instance.url}")