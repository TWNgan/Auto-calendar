# col = ['Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG']
# for i in column_field:
#     for u in range(6,82,8):
#         strt=i + str(u)
#         print(strt)
# list = []
# u = "Q"
# i = 1
# while (ord(u) <= ord("Z")):
#     list.append(i,u)
#     u = chr(ord("Q")+i)
#     i = i + 1 
    
# print(list)
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

from googleapiclient.discovery import build
import datetime
import pytz
import os.path
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from googleapiclient.errors import HttpError
from google.oauth2.service_account import Credentials

col = ['Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG']
# Set up authentication credentials
creds = Credentials.from_service_account_file('ferrous-freedom-379702-1fb1fcb40ed9.json')

# Set up Google Sheets API client
service = build('sheets', 'v4', credentials=creds)

# Set up spreadsheet ID and range for desired cells
spreadsheet_id = '14k4bZQPr-z84xf15hFMCYTqN17HnCwA3AhhiF86g0Qw'
timezone = 'Asia/Hong_Kong'
service1 = build('calendar', 'v3', credentials=creds)
events_result = service1.events().list(calendarId='97c4feb163a23ee5b0b7bc7fa72fb2fa1f53ad15c293445bff834d423d4ca75b@group.calendar.google.com').execute()
events1 = events_result.get('items', [])
for event in events1:
    service1.events().delete(calendarId='97c4feb163a23ee5b0b7bc7fa72fb2fa1f53ad15c293445bff834d423d4ca75b@group.calendar.google.com', eventId=event['id']).execute()


user = [['Chi'],['Math School Hall(1:45-5:35)'],['通識S6'],['Eng'], ['Math'], ['ICT'],['Phy'],['M2'],['通識6Z']]



# Call the Sheets API to get cell data



def exe():
    range_name = '6D!' +p+str(l)+":"+p+str(l+7)
    result = service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id, range=range_name).execute()
    a = l
    o = a
    for i in result['values']:
        for u in user:
            if i == u:
                search = (p+str(o))
                t = "A"+str(a)
                rang = "6D!"+ t +":" + t 
                time = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id, range=rang).execute()
                sh = time['values'][0][0][:2]
                sm = time['values'][0][0][:-6][-2:]
                eh = time['values'][0][0][:-3][-2:]
                em = time['values'][0][0][-2:]
                Date = "6D!"+col[col.index(p)+1]+"3"
                date = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id, range=Date).execute()
                dat = date['values'][0][0][:-2]
                month = date['values'][0][0][-1:]
                content_cell_1 = "6D!"+col[col.index(p)+1]+str(o)
                content_cell_2 = "6D!"+col[col.index(p)+2]+str(o)
                content1 = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id, range=content_cell_1).execute()
                content2 = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id, range=content_cell_2).execute()
                content = content1['values'][0][0] + " , "+content2['values'][0][0]
                Sub = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id, range="6D!"+search).execute()
                event = {
                    'summary': Sub['values'][0][0],
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
                    event = service1.events().insert(calendarId='97c4feb163a23ee5b0b7bc7fa72fb2fa1f53ad15c293445bff834d423d4ca75b@group.calendar.google.com', body=event).execute()
                    print(f'Event created: {event.get("htmlLink")}')
                except HttpError as error:
                    print(f'An error occurred: {error}')

                
        o = o + 1


for p in col[::3]:
    for l in range(6,79,8):
        exe()
