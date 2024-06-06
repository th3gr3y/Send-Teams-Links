# Send-Teams-Links
This Python script automates the process of creating and sending Microsoft Teams meeting requests using the data from an Excel file. The details of each meeting (subject, start and end time, description, attendees etc.) are stored in an Excel file, and the script reads these details to create and send the meeting.

Here is what each section does:

First, it imports necessary libraries: win32com.client to create a COM client for Outlook, pandas to read the Excel file, pyautogui to interact with the GUI, and time to introduce delays.

Then, it uses win32.Dispatch('Outlook.Application') to create a COM object for Microsoft Outlook. This allows Python to interact with the Outlook application.

It uses pd.read_excel() to load the Excel file that contains the meeting details.

Next, it processes the date columns to ensure they are in datetimes using pd.to_datetime(). If there are any issues with the date formatting, it will raise a ValueError.

It then starts a loop that iterates through each row of the DataFrame (df) returned by pd.read_excel(). For each row, it creates a new meeting request using outlook.CreateItem(1), where number 1 corresponds to an AppointmentItem in Outlook.

The meeting details (subject, start and end time, location, and description) are set from corresponding columns of the DataFrame.

The list of attendees (recipients) for each meeting are derived from ‘To’ and ‘CC’ columns of DataFrame.

The meeting.Display() function opens the meeting request form in Outlook to verify that it has been created correctly.

The script then waits for 5 seconds to give time for the Teams meeting link to be generated before using pyautogui.hotkey() and pyautogui.press() functions to interact with the GUI and click the "Join Microsoft Teams Meeting" button.

Finally, it sends the meeting request and releases the Outlook and meeting objects.

Please note that this script assumes that all meeting recipients are valid and that Outlook is set up to automatically include a Teams meeting link when creating a new meeting request. Also, it highly relies on the GUI interactions, meaning it might behave differently based on the actual user interface of the system. Moreover, the file path and some other aspects are hardcoded, limiting the reusability of the code without modifications.

The Excel File should have the following data;

Name,	To,	CC,	Location,	Subject,	Meeting Start Time,	Meeting End Time,	Description
