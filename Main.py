from google.oauth2 import service_account
from googleapiclient.discovery import build
import datetime
import re
from dotenv import load_dotenv
import os
import win32com.client
import platform
import argparse
from collections import defaultdict

parser = argparse.ArgumentParser()
parser.add_argument("--test", action="store_true", default=False, help="Sets script to testing mode, uses test Google Sheets")
args = parser.parse_args()

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

RESPONSES_ID = os.getenv('RESPONSES_ID') if not args.test else os.getenv('TEST_RESPONSES_ID') # google form speadsheet ID (between d/ and /edit in the URL
RESPONSES_SHEET = 'Form Responses 1'
LATE_PASSES_ID = os.getenv('LATE_PASSES_ID') if not args.test else os.getenv('TEST_LATE_PASSES_ID') # google sheet late pass usage ID (between d/ and /edit in the URL)
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
    if args.test:
        print("Testing Mode")

    response_values = scrape_spreadsheet(RESPONSES_ID, RESPONSES_SHEET)
    late_pass_values = scrape_spreadsheet(LATE_PASSES_ID, LATE_PASSES_SHEET)

    today = datetime.date.today()
    last_friday = today - datetime.timedelta(days=(today.weekday() - 4) % 7)
    month_day_str = last_friday.strftime("%B %#d") if platform.system() == "Windows" else last_friday.strftime(
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

    grouped_responses = defaultdict(list)
    for r in formatted_responses:
        user = r.get("user ID (initials followed by digits, you don't need the \"@drexel.edu\")")
        grouped_responses[user].append(r)

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

    for user_email, responses in grouped_responses.items():
        student = next((s for s in formatted_late_passes if s.get("email") == user_email), None)
        if not student:
            continue

        assignments = [r.get("Choose Homework Assignment", "") for r in responses]
        assignment_counts = defaultdict(int)
        for a in assignments:
            assignment_counts[a] += 1

        assignment_text = assignments[0]
        is_duplicate = assignment_counts[assignment_text] > 1

        match = re.search(r'\bHW(\d+)\b', assignment_text)
        hw_num = match.group(1) if match else None
        hw_code = f"HW{hw_num}" if hw_num else "the assignment"
        hw_code_with_hash = f"HW#{hw_num}" if hw_num else "the assignment"

        email = student.get("email")
        subject = "Late Pass Usage Confirmation"
        due_date = today + datetime.timedelta(days=1) if is_duplicate else today

        p1 = student.get("P1")
        p2 = student.get("P2")

        if p1 and p2:
            subject = "Late Pass Usage Error"
            body = (
                f"We have on record that you have already used your two given late passes on assignments "
                f"{p1.upper()} and {p2.upper()}, therefore there are none remaining. Please speak to your instructor if "
                f"you believe this is in error."
            )
        elif p1 and not p2:
            student["P2"] = hw_code.lower()
            update_cell(headers, LATE_PASSES_ID, LATE_PASSES_SHEET, student.get("_row_index"), "P2", hw_code.lower())
            if is_duplicate:
                body = (
                    f"You are receiving this email as confirmation of your late pass usage for {hw_code_with_hash}.\n\n"
                    f"You have attempted to use 2 late passes on this assignment, however you only had 1 available, "
                    f"so you have only received a single-day extension on the assignment. If you believe this is in "
                    f"error, please contact your instructor.\n\nYou may now submit {hw_code_with_hash} by "
                    f"{format_date(today)} at 11:59 PM with no late penalty. This was your last late pass for the "
                    f"quarter, and so any future assignments will be assessed by the standard -10%/day penalty. Be "
                    f"aware that homework submissions are no longer accepted after Tuesday nights, regardless of any "
                    f"late pass use."
                )
            else:
                body = (
                    f"You are receiving this email as confirmation of your late pass usage for {hw_code_with_hash}. You may now submit "
                    f"{hw_code_with_hash} by {format_date(due_date)} at 11:59 PM with no late penalty. This was your last late pass for the "
                    f"quarter, and so any future assignments will be assessed by the standard -10%/day penalty. Be aware that homework submissions "
                    f"are no longer accepted after Tuesday nights, regardless of any late pass use."
                )
        else:
            student["P1"] = hw_code.lower()
            update_cell(headers, LATE_PASSES_ID, LATE_PASSES_SHEET, student.get("_row_index"), "P1", hw_code.lower())
            if is_duplicate:
                body = (
                    f"You are receiving this email as confirmation of your late pass usage for {hw_code_with_hash}.\n\n"
                    f"You have used both of your late passes for the quarter on this assignment. If you believe this is in error, please contact your instructor.\n\n"
                    f"You may now submit {hw_code_with_hash} by {format_date(due_date)} at 11:59 PM with no late penalty. This was your last late pass for the "
                    f"quarter, and so any future assignments will be assessed by the standard -10%/day penalty. Be aware that homework submissions "
                    f"are no longer accepted after Tuesday nights, regardless of any late pass use."
                )
            else:
                body = (
                    f"You are receiving this email as confirmation of your late pass usage for {hw_code_with_hash}. You may now submit "
                    f"{hw_code_with_hash} by {format_date(due_date)} at 11:59 PM with no late penalty. You have one late pass remaining, which can be "
                    f"used again on this assignment, should you wish to take until Sunday night, or on a future homework. Be aware that homework "
                    f"submissions are no longer accepted after Tuesday nights, regardless of any late pass use."
                )

        messages[email] = (subject, body)

    receipt = []

    if platform.system() == "Windows":
        outlook = win32com.client.Dispatch("Outlook.Application") # uses existing Outlook session on user's PC
    for user_id, (subject, content) in messages.items():
        if user_id == "steve.earth":
            email = f"{user_id}@gmail.com"
        elif user_id == "mboady":
            email = "steve.loves.math@gmail.com"
        else:
            email = f"{user_id}@drexel.edu"
        if platform.system() == "Windows":
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = subject
            mail.Body = content
            mail.Send()
        else:
            print(f"Email would be sent to {email} with:")
            print(f"\tSubject: {subject}")
            print(f"\tBody: {content}")
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