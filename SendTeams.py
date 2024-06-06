import win32com.client as win32
import pandas as pd
import pyautogui
import time

# Connect to Outlook
outlook = win32.Dispatch('Outlook.Application')

# Open the Excel file containing the meeting details
df = pd.read_excel(r'C:\Test.xlsx', sheet_name='TeamsList')

# Ensure the date columns are parsed correctly
df['Meeting Start Time'] = pd.to_datetime(df['Meeting Start Time'], errors='coerce')
df['Meeting End Time'] = pd.to_datetime(df['Meeting End Time'], errors='coerce')

# Check for any issues with date parsing
if df['Meeting Start Time'].isnull().any() or df['Meeting End Time'].isnull().any():
    raise ValueError("There are invalid date formats in the input data")

# Loop through each row of the Excel file and create a Teams meeting request
for i in range(len(df)):
    # Create a new meeting request
    meeting = outlook.CreateItem(1)  # 1 refers to an AppointmentItem

    # Set the meeting details
    meeting.Subject = df.iloc[i]['Subject']
    start_time = pd.to_datetime(df.iloc[i]['Meeting Start Time']).tz_localize('Europe/London')
    end_time = pd.to_datetime(df.iloc[i]['Meeting End Time']).tz_localize('Europe/London')
    meeting.Location = "Microsoft Teams Meeting"
    meeting.Body = df.iloc[i]['Description']
    start_time_utc = start_time.tz_convert('Europe/London')
    end_time_utc = end_time.tz_convert('Europe/London')
    meeting.Start = start_time
    meeting.End = end_time

    # Set the meeting recipients
    to_recipients = df.iloc[i]['To'].split(';')
    cc_recipients = df.iloc[i]['CC'].split(';') if pd.notna(df.iloc[i]['CC']) else []

    for recipient in to_recipients:
        meeting.Recipients.Add(recipient.strip())

    for recipient in cc_recipients:
        meeting.Recipients.Add(recipient.strip())

    # Display the meeting request form to ensure it gets created properly
    meeting.Display()

    # Wait for the Teams meeting link to be generated and inserted
    time.sleep(5)  # Adjusted to give enough time for the link to generate

    # Use PyAutoGUI to click the "Join Microsoft Teams Meeting" button
    pyautogui.hotkey('fn', 'f10')
    time.sleep(2)
    pyautogui.press('6')
    time.sleep(3)

    # Send the meeting request
    meeting.Send()

    # Release the meeting object
    del meeting

# Release the Outlook object
del outlook
