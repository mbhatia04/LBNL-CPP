import time
import gspread
import argparse
from oauth2client.service_account import ServiceAccountCredentials

# Returns list of mac addresses currently in gsheet
def gsheet_mac_addr() :
    return worksheet.col_values(3)

# Read data from the file & turn mac addresses into list
def log_file_mac_addr():
    with open(data_file_path, 'r') as file:
        adds_list = [line.strip().split('|')[2].strip() for line in file if '|' in line]
    return adds_list

# Delete rows in gsheet that are not in log
def delete_rows():
    rows_to_delete = [i + 1 for i, mac in enumerate(gsheet_mac_addr()) if mac.strip() not in log_file_mac_addr()]

    # Delete rows
    for row in reversed(rows_to_delete):
        worksheet.delete_rows(row)

# TODO: check for dupes in source file
# TODO: exception handling (e.g. machine was removed from denyboot log and then added back)

def append_new_rows_to_sheet():

    # Get the number of rows with data in the sheet
    num_existing_rows = len(worksheet.get_all_values())

    # Read data from the log file starting from the next index after the last row with data
    with open(data_file_path, 'r') as file:
        log_data = file.readlines()[num_existing_rows:]

    # Loop through the data and insert into the sheet
    batch_values = []
    for line in log_data:
        values = line.strip().split('|')
        batch_values.append(values)

        # Submit a batch of 10 rows
        if len(batch_values) >= 10:
            worksheet.append_rows(batch_values)
            batch_values = []
            time.sleep(1)  # Add a 1-second delay between requests as Google imposes rate limiting

    # Insert any remaining rows
    if batch_values:
        worksheet.append_rows(batch_values)

def reload_sheet():
    worksheet.clear()

    # Read data from the log file starting from the next index after the last row with data
    with open(data_file_path, 'r') as file:
        log_data = file.readlines()

    # Loop through the data and insert into the sheet
    batch_values = []
    for line in log_data:
        values = line.strip().split('|')
        batch_values.append(values)

        # Submit a batch of 10 rows
        if len(batch_values) >= 10:
            worksheet.append_rows(batch_values)
            batch_values = []
            time.sleep(1)

    # Insert any remaining rows
    if batch_values:
        worksheet.append_rows(batch_values)
    
    print('Sheet reinitialized')

if __name__ == "__main__":
    # Location of .json credentials file
    JSON_KEY_FILE = '/Users/meghanbhatia/Desktop/LBL/dboot/eastern-crawler-397602-0d1b818309ec.json'
    SHEET_ID = '1Om6xaAbWHTzyGGbcvelS3e9NoPzhJzoLwrwN-9bA5Lw'

    # Authenticate using the service account and JSON key file
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    client = gspread.authorize(credentials)

    # Define sheet & worksheet
    sheet = client.open_by_key(SHEET_ID)
    worksheet = sheet.get_worksheet(0)

    data_file_path = '/Users/meghanbhatia/Desktop/LBL/dboot/DENYBOOT.why'

    delete_rows()
    append_new_rows_to_sheet()

    parser = argparse.ArgumentParser(description='Command line options to edit sheet accordingly')
    parser.add_argument('--reload', action='store_true', help='Reinitialize GSheet')
    args = parser.parse_args()

    reload_sheet() if args.reload else None