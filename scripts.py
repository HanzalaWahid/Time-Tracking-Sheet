import pandas as pd
from datetime import datetime
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build


file_path = "Project.xlsx"


if os.path.exists(file_path):
    df = pd.read_excel(file_path)
else:
    df = pd.DataFrame(columns=[
        'Clients', 'Contractors', 'Date', 'Start Time', 'Stop Time',
        'No of hours', 'Whole number Decimal Time', 'Activity Memo/Task Description'
    ])


def time_diff_in_decimal(start, stop):
    time_format = "%H:%M"
    start_dt = datetime.strptime(start, time_format)
    stop_dt = datetime.strptime(stop, time_format)
    delta = stop_dt - start_dt
    return round(delta.total_seconds() / 3600, 2)

contractor = input("Enter contractor name: ")
client = input("Enter client name: ")
date = input("Enter date (YYYY-MM-DD): ")
start_time = input("Enter start time (HH:MM): ")
stop_time = input("Enter stop time (HH:MM): ")
activity = input("Enter task description (optional, press Enter to use default): ") or "General Task"

hours_decimal = time_diff_in_decimal(start_time, stop_time)

new_row = {
    'Clients': client,
    'Contractors': contractor,
    'Date': date,
    'Start Time': start_time,
    'Stop Time': stop_time,
    'No of hours': hours_decimal,
    'Whole number Decimal Time': hours_decimal,
    'Activity Memo/Task Description': activity
}

df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
df.to_excel(file_path, index=False)
print("✅ Entry saved to 'Project.xlsx'")



SERVICE_ACCOUNT_FILE = 'burnished-ember-404109-41a9d1bf0785.json'  
SPREADSHEET_ID = '1jvb8yTe1w-_dAwAiYHx5DwL9jPkyT3vP-AYNIYXvorQ'     
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]


creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
sheets_service = build('sheets', 'v4', credentials=creds)


values = [[
    client, contractor, date, start_time, stop_time,
    hours_decimal, hours_decimal, activity
]]


sheets_service.spreadsheets().values().append(
    spreadsheetId=SPREADSHEET_ID,
    range='Sheet1!A1',
    valueInputOption='USER_ENTERED',
    insertDataOption='INSERT_ROWS',
    body={'values': values}
).execute()

print("✅ Entry also added to Google Sheet.")
