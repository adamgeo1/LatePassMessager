from google.oauth2 import service_account
from googleapiclient.discovery import build
import datetime
import re
import platform
from dotenv import load_dotenv
import os

load_dotenv()

# Path to your service account key
SERVICE_ACCOUNT_FILE = os.getenv('GOOGLE_CREDS_PATH')

# Scopes needed for Drive access
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Authenticate
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

service = build('sheets', 'v4', credentials=creds)

RESPONSES_ID = os.getenv('RESPONSES_ID')
RESPONSES_SHEET = 'Form Responses 1'
LATE_PASSES_ID = os.getenv('LATE_PASSES_ID')
LATE_PASSES_SHEET = 'roster'
LATE_PASSES_MESSAGE = 'message'

def scrape_spreadsheet(id, sheet):
    result = service.spreadsheets().values().get(
        spreadsheetId=id, range=sheet
    ).execute()

    return result.get('values', [])

def main():
    response_values = scrape_spreadsheet(RESPONSES_ID, RESPONSES_SHEET)

    #yesterday = datetime.date.today() - datetime.timedelta(days=1)
    yesterday = datetime.date(2025, 5, 24) - datetime.timedelta(days=1)
    month_day_str = yesterday.strftime("%B %#d") if platform.system() == 'Windows' else yesterday.strftime("%B %-d")
    pattern = rf'\(due {re.escape(month_day_str)}\)'

    headers = response_values[0]
    assignment_column = 'Choose Homework Assignment'
    formatted_responses = [
        dict(zip(headers, row))
        for row in response_values[1:]
        if len(row) == len(headers) and re.search(pattern, row[headers.index(assignment_column)], re.IGNORECASE)
    ]

    late_pass_values = scrape_spreadsheet(LATE_PASSES_ID, LATE_PASSES_SHEET)

    headers = late_pass_values[0]
    formatted_late_passes = [
        dict(zip(headers, row))
        for row in late_pass_values[1:]
        if len(row) == len(headers)
    ]

if __name__ == '__main__':
    main()