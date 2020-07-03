from __future__ import print_function
import datetime
from datetime import timedelta, datetime, date, time
import dateutil.parser
import pickle
import os.path
from googleapiclient import discovery
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import gspread
from httplib2 import Http
from oauth2client import file, client, tools
import time as time2

def calendar(timefrom, timeto):
    SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
    creds = None
    if os.path.exists('/REPLACE/calendar.pickle'):
        with open('/REPLACE/calendar.pickle', 'rb') as token:
            creds = pickle.load(token)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    '/REPLACE/calcredentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('/REPLACE/calendar.pickle', 'wb') as token:
                pickle.dump(creds, token)
        service = build('calendar', 'v3', credentials=creds)
        events_result = service.events().list(calendarId='REPLACE@group.calendar.google.com', timeMin=timefrom, timeMax=timeto,
                                            maxResults=30, singleEvents=True,
                                            orderBy='startTime').execute()
    events = events_result.get('items', [])
    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
    return events

def maketime(offmin):
    offmax = offmin+1
    from datetime import datetime, timedelta, date, time
    midnight = datetime.combine((date.today()+timedelta(+offmax)), time())
    midnight = midnight+timedelta(hours=7)
    midnight = midnight+timedelta(+offmax)
    midnight = midnight.isoformat()+('Z')
    morning = datetime.combine((date.today()+timedelta(+offmin)), time())
    morning = morning+timedelta(hours=7)
    morning = morning.isoformat()+('Z')
    return morning, midnight

def defineweek(offmin):
    startofwk = pickle.load( open( "/REPLACE/startwk.pickle", "rb" ) )
    offmax = offmin+1
    from datetime import datetime ,timedelta, date, time
    morning = datetime.combine((startofwk+timedelta(offmin)), time())
    morning = morning+timedelta(hours=7)
    morning = morning.isoformat()+('Z')
    midnight = datetime.combine((startofwk+timedelta(offmax)), time())
    midnight = midnight+timedelta(hours=7)
    midnight = midnight.isoformat()+('Z')
    return morning, midnight

def startofwk():
    from datetime import datetime ,timedelta, date, time
    startofwk = date.today()
    return startofwk

def parsecal(events) :
    workers = []
    col0 = []
    col1 = []
    col2 = []
    col3 = []
    col4 = []
    for i in events:
        time = i['start']['dateTime']
        time = dateutil.parser.parse(time)
        time = time.strftime('%m/%d/%Y')
        i['time'] = time
        del i['kind']
        del i['etag']
        del i['id']
        del i['creator']
        del i['status']
        del i['htmlLink']
        del i['updated']
        del i['iCalUID']
        del i['sequence']
        i.pop('reminders')
        i.pop('start')
        i.pop('end')
        del i['organizer']
        del i['created']
    events = sorted(events, key = lambda i: i['summary'])
    for i in events:
        if i['summary'] not in workers:
            if 'call' not in str(i['summary']) and 'Call' not in str(i['summary']):
                workers.append(i['summary'])
    for i in events:
        try:
            if i['summary'] == workers[0]:
                col0.append(i['description'])
            elif i['summary'] == workers[1]:
                col1.append(i['description'])
            elif i['summary'] == workers[2]:
                col2.append(i['description'])
            elif i['summary'] == workers[3]:
                col3.append(i['description'])
            elif i['summary'] == workers[4]:
                col4.append(i['description'])
        except:
            continue
    return workers, events, col0, col1, col2, col3

def appendtosheet(newsheet, day, workers, col0, col1, col2, col3):
    gc = gspread.oauth()
    sh = gc.open_by_key(newsheet)
    newsheet = sh.add_worksheet(title=day, rows="20", cols="5")
    worksheet = sh.worksheet(day)
    worksheet.append_row(workers)
    x=1
    for i in col0:
        x=x+1
        worksheet.update_acell('A'+str(x), i)
    x=1
    for i in col1:
        x=x+1
        worksheet.update_acell('B'+str(x), i)
    x=1
    for i in col2:
        x=x+1
        worksheet.update_acell('C'+str(x), i)
    x=1
    for i in col3:
        x = x+1
        worksheet.update_acell('D'+str(x), i)
    x=1
    sheetId = sh.worksheet(day)._properties['sheetId']
    body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheetId,
                        "dimension": "COLUMNS",
                        "startIndex": 1,
                        "endIndex": 4
                    },
                    "properties": {
                        "pixelSize": 450
                    },
                    "fields": "pixelSize"
                }
            }
        ]
    }
    res = sh.batch_update(body)
    worksheet.format("A1:D1", {"horizontalAlignment": "CENTER", "textFormat": {"bold": True}})
    #worksheet.format("A2:A7", { "backgroundColor": {"red": 5, "green": 15, "blue": 15, "alpha": 0}}) #Equipment
    worksheet.format("B2:B15", { "backgroundColor": {"red": 15, "green": 5, "blue": 15, "alpha": 0}})
    worksheet.format("C2:C15", { "backgroundColor": {"red": 15, "green": 15, "blue": 5, "alpha": 0}})
    worksheet.format("D2:D15", { "backgroundColor": {"red": 5, "green": 15, "blue": 15, "alpha": 0}})

def createsched(morning, midnight, newsheet):
    events = calendar(morning, midnight)
    day = dateutil.parser.parse(morning)
    day = day.strftime('%a-%m/%d/%y')
    workers, events, col0, col1, col2, col3 = parsecal(events)
    appendtosheet(newsheet, day, workers, col0, col1, col2, col3)

def createworksheet(name):
    from oauth2client import file, client, tools
    SCOPES = ['https://spreadsheets.google.com/feeds', 'https://docs.google.com/feeds', 'https://www.googleapis.com/auth/drive.file']
    store = file.Storage('/REPLACE/storage.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('serviceaccount.json', SCOPES)
        creds = tools.run_flow(flow, store)
    drive_service = discovery.build('drive', 'v3', credentials=creds)  # Use "credentials" of "gspread.authorize(credentials)"
    folder_id = 'REPLACE'
    file_metadata = {
        'name': name,
        'parents': [folder_id],
        'mimeType': 'application/vnd.google-apps.spreadsheet'
    }
    file = drive_service.files().create(body=file_metadata,
                                        fields='id').execute()
    return file.get('id')

def evaluate():    # DEPRECATED
    day = dateutil.parser.parse(morning)
    worksheetname = day.strftime('%b-%d/%Y')
    weekday = day.strftime('%w')
    day = day.strftime('%a-%m/%d/%y')
    if weekday == 0 or weekday == '0':
        pickle.dump( startofwk, open( "/REPLACE/startwk.pickle", "wb" ))
        currentsheet = [morning, midnight]
        #pickle.dump( currentsheet, open( "currentsheet.pickle", "wb" ) )
        newsheet = createworksheet(worksheetname)
        events = calendar(morning, midnight)
    else:
        #currentsheet = pickle.load( open( "currentsheet.pickle", "rb" ) )
        events = calendar(currentsheet[0], currentsheet[1])
        workers, events, col0, col1, col2, col3 = parsecal(events)
        appendtosheet(newsheet, day, workers, col0, col1, col2, col3)

sheetexist=False
morning, midnight = maketime(0)
day = dateutil.parser.parse(morning)
weekday = day.strftime('%w')
if weekday == 0 or weekday == '0':
    sheetexist=False
    startofwk = startofwk()
    pickle.dump( startofwk, open( "/REPLACE/startwk.pickle", "wb" ) )
else:
    sheetexist=True
    startofwk = pickle.load( open( "/REPLACE/startwk.pickle", "rb" ) )

worksheetname = startofwk.strftime('%b-%d/%Y')

if sheetexist == False:
    newsheet = createworksheet(worksheetname)
    pickle.dump(newsheet, open('/REPLACE/sheetid.pickle', 'wb'))
elif sheetexist == True:
    newsheet = pickle.load( open( "/REPLACE/sheetid.pickle", "rb" ) )
    gc = gspread.oauth()
    sh = gc.open_by_key(newsheet)
    sheetlist = sh.worksheets()
    del sheetlist[0]
    for i in sheetlist:
        sh.del_worksheet(i)

i = [0,1,2,3,4,5,6]

for x in i:
    #morning, midnight = maketime(x)
    morning, midnight = defineweek(x)
    time2.sleep(100)
    createsched(morning, midnight, newsheet)
