from google.oauth2 import service_account
from googleapiclient.discovery import build
import datetime
import re
from dotenv import load_dotenv
import os
import win32com.client
import platform

load_dotenv() # loads .env file

# Path to your service account key
SERVICE_ACCOUNT_FILE = os.getenv('GOOGLE_CREDS_PATH')

# Scopes needed for Drive access
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Authenticate
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

service = build('sheets', 'v4', credentials=creds)

#RESPONSES_ID = os.getenv('RESPONSES_ID') # google form speadsheet ID (between d/ and /edit in the URL
RESPONSES_SHEET = 'Form Responses 1'
#LATE_PASSES_ID = os.getenv('LATE_PASSES_ID') # google sheet late pass usage ID (between d/ and /edit in the URL)
LATE_PASSES_SHEET = 'roster'
LATE_PASSES_MESSAGE = 'message'

def scrape_spreadsheet(id, sheet):
    result = service.spreadsheets().values().get(
        spreadsheetId=id, range=sheet
    ).execute()

    return result.get('values', [])

def format_date(d: datetime.date) -> str:
    day = d.day
    suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return d.strftime(f'%A %B {day}{suffix}')

def update_cell(headers, id, sheet, row, col, value):
    col_letter = chr(ord('A') + headers.index(col))
    cell_range = f"{sheet}!{col_letter}{row}"
    service.spreadsheets().values().update(
        spreadsheetId=id, range=cell_range, valueInputOption="RAW", body={"values": [[value]]}
    ).execute()

def main():
    response_values = scrape_spreadsheet(RESPONSES_ID, RESPONSES_SHEET)
    late_pass_values = scrape_spreadsheet(LATE_PASSES_ID, LATE_PASSES_SHEET)

    today = datetime.date.today()
    last_saturday = today - datetime.timedelta(days=(today.weekday() - 5) % 7)
    month_day_str = last_saturday.strftime("%B %#d") if platform.system() == "Windows" else last_saturday.strftime(
        "%B %-d")
    pattern = rf'\(due {re.escape(month_day_str)}\)'

    response_headers = response_values[0]
    assignment_column = 'Choose Homework Assignment'
    formatted_responses = [
        dict(zip(response_headers, row))
        for row in response_values[1:]
        if len(row) == len(response_headers) and re.search(pattern, row[response_headers.index(assignment_column)],
                                                           re.IGNORECASE)
    ]

    if not formatted_responses:
        print("No responses for yesterday's due date.")
        return

    # Step 1: Get HW number from first matching response
    first_assignment = formatted_responses[0].get("Choose Homework Assignment", "")
    current_hw_match = re.search(r'\bhw(\d+)\b', first_assignment, re.IGNORECASE)
    if not current_hw_match:
        print("Could not extract HW number from yesterday's assignment.")
        return
    current_hw = current_hw_match.group(1)

    # Step 2: Search second row of late pass sheet for "last email: hw{num}"
    last_email_hw = None
    for cell in late_pass_values[1]:  # second row
        match = re.search(r'last email:\s*hw(\d+)', cell, re.IGNORECASE)
        if match:
            last_email_hw = match.group(1)
            break

    # Step 3: Skip if already sent
    if last_email_hw == current_hw:
        print(f"Emails already sent for HW{current_hw}. Skipping main().")
        return

    late_pass_values = scrape_spreadsheet(LATE_PASSES_ID, LATE_PASSES_SHEET)

    headers = late_pass_values[0]
    formatted_late_passes = [
        {**dict(zip(headers, row)), "_row_index": i + 2}
        for i, row in enumerate(late_pass_values[1:])
    ]

    '''print("=== Formatted Responses ===") # for debugging purposes
    for i, entry in enumerate(formatted_responses, start=1):
        print(f"Response #{i}:")
        for key, value in entry.items():
            print(f"  {key}: {value}")
        print()

    print("=== Formatted Late Passes ===")
    for i, entry in enumerate(formatted_late_passes, start=1):
        print(f"Late Pass #{i}:")
        for key, value in entry.items():
            print(f"  {key}: {value}")
        print()'''

    messages = {}

    for response in formatted_responses:
        assignment_text = response.get("Choose Homework Assignment", "")
        match = re.search(r'\bHW(\d+)\b', assignment_text) # gets hw code
        hw_num = match.group(1) if match else None
        hw_code = f"HW{hw_num}" if hw_num else "the assignment"
        hw_code_with_hash = f"HW#{hw_num}" if hw_num else "the assignment"

        for student in formatted_late_passes:
            if response.get("user ID (initials followed by digits, you don't need the \"@drexel.edu\")") == student.get(
                    "email"):

                email = student.get("email")
                subject = "Late Pass Usage Confirmation"

                if student.get("P1") and student.get("P2"):
                    subject = "Late Pass Usage Error"
                    body = (
                        f"We have on record that you have already used your two given late passes on assignments "
                        f"{student.get("P1").upper()} and {student.get("P2").upper()}, therefore "
                        f"there are none remaining. Please speak to your instructor if "
                        f"you believe this is in error."
                    )
                elif student.get("P1") and not student.get("P2"):
                    student["P2"] = hw_code.lower()
                    update_cell(headers, LATE_PASSES_ID, LATE_PASSES_SHEET, student.get("_row_index"), "P2", hw_code.lower())
                    body = (
                        f"You are receiving this email as confirmation of your late "
                        f"pass usage for {hw_code_with_hash}. You may now submit "
                        f"{hw_code_with_hash} by {format_date(today)} at 11:59 PM with no "
                        f"late penalty. This was your last late pass for the quarter, "
                        f"and so any future assignments will be assessed by the "
                        f"standard -10%/day penalty. Be aware that homework submissions "
                        f"are no longer accepted after Tuesday nights, regardless of "
                        f"any late pass use."
                    )
                else:
                    student["P1"] = hw_code.lower()
                    update_cell(headers, LATE_PASSES_ID, LATE_PASSES_SHEET, student.get("_row_index"), "P1", hw_code.lower())
                    body = (
                        f"You are receiving this email as confirmation of your late "
                        f"pass usage for {hw_code_with_hash}. You may now submit "
                        f"{hw_code_with_hash} by {format_date(today)} at 11:59 PM with no "
                        f"late penalty. You have one late pass remaining, which can be "
                        f"used again on this assignment, should you wish to take until "
                        f"Sunday night, or on a future homework. Be aware that homework "
                        f"submissions are no longer accepted after Tuesday nights, "
                        f"regardless of any late pass use."
                    )

                messages[email] = (subject, body)

    receipt = []

    outlook = win32com.client.Dispatch("Outlook.Application") # uses existing Outlook session on user's PC
    for user_id, (subject, content) in messages.items():
        if user_id == "steve.earth":
            email = f"{user_id}@gmail.com"
        elif user_id == "mboady":
            email = "steve.loves.math@gmail.com"
        else:
            email = f"{user_id}@drexel.edu"
        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = subject
        mail.Body = content
        mail.Send()
        print(f"Email sent to {email}")
        receipt.append(f"{email}: {subject}")

    receipt_mail = outlook.CreateItem(0)
    receipt_mail.To = outlook.Session.CurrentUser.Address
    receipt_mail.Subject = f"Late Pass Receipt for {today.strftime('%B %d, %Y')}"
    receipt_mail.Body = "Late pass confirmation/denial emails were sent to the following:\n\n" + "\n".join(receipt)
    receipt_mail.Send()
    print("Receipt email sent")

    update_cell(headers, LATE_PASSES_ID, LATE_PASSES_SHEET, 2, "other notes", f"last email: {hw_code.lower()}")

if __name__ == '__main__':
    main()