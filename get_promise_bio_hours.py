import datetime
import pandas as pd
import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

def get_events(days_back, calendar_id):
    """Shows basic usage of the Google Calendar API.
    Lists the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('calendar', 'v3', credentials=creds)

    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    past = (datetime.datetime.utcnow() - datetime.timedelta(days=days_back)).isoformat() + 'Z'
    events_result = service.events().list(calendarId=calendar_id, timeMin=past,
                                          maxResults=1000, singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])

    event_list = []
    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        end = event['end'].get('dateTime', event['end'].get('date'))
        duration = (datetime.datetime.fromisoformat(end) - datetime.datetime.fromisoformat(start)).total_seconds() / 60 / 60
        print(f"{event['summary']} - {start} - {duration}")
        event_list.append([event['summary'], start, duration])

    return event_list

def main():
    days_back = 100  # configure this
    calendar_id = 'f443df4629510df4d7a323ab93df431caa5af0fab1f7d41f103cdf552065dd2c@group.calendar.google.com'  # replace with your calendar id
    events = get_events(days_back, calendar_id)
    df = pd.DataFrame(events, columns=['Event Title', 'Start Time', 'Duration'])
    df.to_csv('events.csv', index=False)

if __name__ == '__main__':
    main()
