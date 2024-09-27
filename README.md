# Walkthrough

It requores two major phases;

1. Setting up the code snippet locally
2. Using Mail Merge

## Setting Up Code Snippet

1. **Import required python libraries**

```
import os
import json
from google.oauth2 import service_account
import csv
import gspread
```

Use `pip install <librarytobedownloaded>` to download the required libraries. You can also use `pip show *<library>*` to check if a library is accessible to you.

2. **Load JSON key file**

```
creds = service_account.Credentials.from_service_account_file(
    'C:/Users/path/to/jsonkey.json',
    scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
)
```

The Scope attribute denotes that google sheets and google drive apis will be used specifically

- Head over to google cloud console and enable the Google sheets api.
  _Note: You would be required to have a project opened on your google cloud console account. If you do not have a project or want a seperate project specific to this use-case, you can create one_.

- Under the API/Service details, create a service account credential and download the json key file to your computer's local directory. [Watch](https://www.youtube.com/watch?v=rWcLDax-VmM)

- Next, enable the google drive api and ensure it has the same service account as the google sheets api.

3. **Authenticate with Google Sheets API**

Using the gspread library, you can create a client object with properties and attributes useful to this project.

`client = gspread.authorize(creds)`

4. **Open the CSV file**

Here we use the `with` function to open and read the csv file to be converted

```
with open('C:/Users/BRIGHT UZOSIKE/Downloads/Sample-Spreadsheet-10-rows.csv', 'r') as csvfile:
    reader = csv.reader(csvfile)
    data = list(reader)
```

5. **Create a new Google Sheet or open an existing one**

- As suggested by the title of the step, we now call on the create property of the client object to instantiate a new spreadsheet

`spreadsheet = client.create('name of new sheet')`

Or we can simply open an existing one if you so desire

`spreadsheet = client.open('Existing Sheet').sheet1`

'sheet1' denotes the worksheet in the google sheets file where the csv contents will be written to. If you have more than one, you can access it using the dot notation.

- You could also manually configure the access permissions to the sheet and share to for confirmation purposes.

`spreadsheet.share('youremail@gmail.com', perm_type='user', role='writer')`

I also created an instance that'll be used for further confirmation purposes.

`spreadsheet_instance = client.open('name of new sheet')`

6. **Create new sheet (optional)**

If you created a new spreadsheet in the previous step, you can create a customized worksheet with the snippet;

`worksheet = spreadsheet.add_worksheet(title='mail-merge', rows=100, cols=100)`

7. **Convert data to list of cell objects**

The data object is converted to a list of row and column objects to enable easy writing to the worksheet

```
cell_list = []
for row, row_data in enumerate(data, start=1):
    for col, value in enumerate(row_data, start=1):
        cell_list.append(gspread.Cell(row, col, value))
```

8. **Write the CSV data to the Google Sheet**

The worksheet can finally be written to usin the `update_cells()` property

`worksheet.update_cells(cell_list)`

9. **Print confirmation url to terminal**

To corfirm complete execution of the process, you can generate a confirmatory url of the created sheet in your drive account

`print(f"Google Sheet URL: {spreadsheet_instance.url}")`

## Using Mail Merge

Navigate to the now created gogle sheet in your drive where you want to perform the the mail merge. The mail merge add-on by Quicklution is simple and efficient in this regard to get the job done.

[Watch the tutorial guide and implement likewise](https://www.youtube.com/watch?v=jcmiK5Q2VEA&list=PLp7ziVJy63Sx2rXagKmncLeMu4fGN5Z4K&index=4)

Thank you for following through.
