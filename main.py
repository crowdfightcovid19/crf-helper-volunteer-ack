import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
import time


def get_acknowledgement(sheet, url, language):
    """Return a list of the volunteers who works on the translation in 'language'
    per volunteer return their role, name, email, if they allowed to be mentioned
    and the name that should be display ('name'/Anonymous/Manual check needed)
    """
    if 'id=' not in url:
        return []

    SAMPLE_RANGE_NAME = 'Volunteers!A1:F200'
    SPREADSHEET_ID = url.split('id=')[-1].strip()
    values_input, count = None, 0
    while values_input is None and count < 10:
        try:
            result_input = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
            values_input = result_input.get('values', [])
        except Exception:
            count += 1
            values_input = None
            time.sleep(2)

    if values_input is None:
        print('NONE', language, url)
        return []

    res = []
    for row in values_input[1:]:
        if len(row) < 2:
            continue
        role = row[0]
        name = row[1]
        if not role or not name:
            continue
        email = row[2] if len(row) >= 3 else ''
        allow_mentions = row[4] if len(row) >= 5 else ''
        display_name = "Manual check needed"
        if allow_mentions.lower().strip() == 'yes':
            display_name = name
        elif allow_mentions.lower().strip() == 'no':
            display_name = 'Anonymous'

        res.append([language, role, name, email, allow_mentions, display_name])
    return res


def get_all_volunteers():
    """return a list of all volunteers who works on translations
    """
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    # Add the spreadsheet ID of Coronasurvey Crowdtranslations
    SPREADSHEET_ID_input = ''
    RANGE_NAME = 'A1:C200'

    # Manage the credentials
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)  # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    sheet = service.spreadsheets()

    # read the main spreadsheet with links to sub-spreadsheet for each language
    result_input = sheet.values().get(spreadsheetId=SPREADSHEET_ID_input,
                                      range=RANGE_NAME).execute()

    values_input = result_input.get('values', [])
    all_volunteers = [['Language', 'Role', 'Name', 'Email', 'Agreed to be mentioned', 'Should be acknowledge as']]

    for i, row in enumerate(values_input[1:]):
        if len(row) < 2:
            continue
        language = row[0]
        url = row[1]

        if 'http' in url:
            _ = get_acknowledgement(sheet, url, language)
            print(i, '/', len(values_input[1:]), language, '' if len(row) < 3 else row[2], len(_))
            all_volunteers += _

    return all_volunteers


def main():
    """Save a csv with the list of all volunteers
    """
    all_volunteers = get_all_volunteers()
    df = pd.DataFrame(all_volunteers[1:], columns=all_volunteers[0])
    df.to_csv('all_volunteers.csv', index=False)


main()
