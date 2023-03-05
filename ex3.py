from googleapiclient.discovery import build
import datetime
import pytz
import os.path
from google.oauth2 import service_account
from googleapiclient.errors import HttpError
from google.oauth2.service_account import Credentials

cid = '27a77b4d5660a6b9844ad5baae9715e78a797267d857041ddf1f36a983cab899@group.calendar.google.com'
creds = Credentials.from_service_account_file('ferrous-freedom-379702-1fb1fcb40ed9.json')
service = build('sheets', 'v4', credentials=creds)
spreadsheet_id = '14k4bZQPr-z84xf15hFMCYTqN17HnCwA3AhhiF86g0Qw'
timezone = 'Asia/Hong_Kong'
service1 = build('calendar', 'v3', credentials=creds)
events_result = service1.events().list(calendarId=cid).execute()
events1 = events_result.get('items', [])
for event in events1:
    service1.events().delete(calendarId=cid, eventId=event['id']).execute()


# Geo
# ICT
# CHEM
# C.HIST
# Econ
# Phy
# Bafs
# 旅款
# M2
# 通識6X
# 通識6Y
# 通識6Z
# Eng
# Chi
# Math
# Math School Hall(1:45-5:35)
# 通識S6
# Chi  3:30-4:30
user = ['Chi','Math School Hall(1:45-5:35)','通識S6','Eng', 'Math','Phy','Geo','M2','通識6Z','Chi  3:30-4:30']

range_name = '6D'
result = service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id, range=range_name).execute()


def exe():
    a = l
    o = a
    lista = [result['values'][l][p],result['values'][l+1][p], result['values'][l+2][p],result['values'][l+3][p],result['values'][l+4][p],result['values'][l+5][p],result['values'][l+6][p],result['values'][l+7][p]]
    for i in lista:
        for u in user:
            if i == u:
                time = result['values'][l][0]
                sh = time[:2]
                sm = time[:-6][-2:]
                eh = time[:-3][-2:]
                em = time[-2:]
                date = result['values'][0][p+1]
                dat = date[:-2]
                month = date[-1:]
                content = result['values'][o][p+1] + " , "+result['values'][o][p+2]
                Sub = result['values'][o][p]
                print(sh, sm, eh, em,dat, month,Sub)
                event = {
                    'summary': Sub,
                    'location': 'PLKCFS',
                    'description': content,
                    'start': {
                        'dateTime': datetime.datetime(2023, int(month), int(dat), int(sh), int(sm), 0).isoformat(),
                        'timeZone': timezone,
                    },
                    'end': {
                        'dateTime': datetime.datetime(2023, int(month), int(dat), int(eh), int(em), 0).isoformat(),
                        'timeZone': timezone,
                    },
                    'reminders': {
                        'useDefault': True,
                    },
                }
                try:
                    event = service1.events().insert(calendarId=cid, body=event).execute()
                    print(f'Event created: {event.get("htmlLink")}')
                except HttpError as error:
                    print(f'An error occurred: {error}')

                
        o = o + 1


for p in range(1,66,3):
    for l in range(3,76,8):
        exe()
