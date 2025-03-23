import os
command = 'pip install pandas flask google-api-python-client google-auth-httplib2 google-auth-oauthlib pdfplumber pysimplegui datetime datefinder'
result = os.popen(command).read()

#importing all required libraries
import datetime
import os.path
from datetime import timedelta
import datefinder
import pandas as pd
import pdfplumber
import PySimpleGUI as sg
from apiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

#function to convert pdf to df and extract data to create event/meet
def dfcon():
    semno = sem
    #extracting pdf into pandas dataframe
    try:
        pdf = pdfplumber.open(pdf_input)
        p_0 = pdf.pages[0]
        table2df = lambda table: pd.DataFrame(table[:-1], columns=table[0])
        d_f = table2df(p_0.extract_table())
        print(d_f)
        branchlist = ['it', 'comp', 'extc', 'etrx', 'mech']

        count = 0
        for i in range(0, len(d_f.columns)):
            if d_f.iloc[i][0] == 'Time/Day':
                break
            else:
                count = count + 1
        d_f = d_f.iloc[count: , :]
        dfcols = list(d_f.columns)
        if dfcols[0] != 'Time/Day':
            d_f.rename(columns = d_f[count])
        print(d_f)
        #cleaning the dataframe
        d_f = d_f.replace('\n',' ', regex=True)
        d_f['Time/Day'] = d_f['Time/Day'].replace(' ', '', regex = True)

        #splitting the time/date column into start and stop columns and deleting that column
        d_f[["start", "stop"]]= d_f["Time/Day"].str.split("-", n = 1, expand = True)
        d_f = d_f.drop(columns = ['Time/Day'])
        d_f = d_f.drop([4, 9])

        #filling data in the cells with null values caused by merged cells of pdf table
        for i in range(0, len(d_f.columns)):
            for j in range(0, len(d_f.index)):
                if pd.isnull(d_f.iloc[j][i]):
                    d_f.iloc[j][i] = d_f.iloc[j-1][i]

        #shifting the start stop columns to front of dataframe
        fcol = d_f.pop('start')
        scol = d_f.pop('stop')
        d_f.insert(0, 'Start', fcol)
        d_f.insert(1, 'Stop', scol)
        try:
            sem_start_date = list(datefinder.find_dates(sem_start))
            start_date = sem_start_date[0].strftime('%Y/%m/%d')
            start = ['8:30', '9:30', '10:30', '11:30', '13:15', '14:15', '15:15', '16:15', '17:30']
            day = 1
            col = datetime.datetime.strptime(start_date, '%Y/%m/%d').weekday() + 2
            for i in range(2, len(d_f.columns)):
                k = 0
                for j in range(0, len(d_f.index)):
                    try:
                        year = str(sem_start_date[0].strftime('%Y'))
                        y_r = int(year)
                        if k >= len(d_f.index):
                            continue
                        else:
                            if d_f.iloc[k][col] == '':
                                k += 1
                                continue
                            elif d_f.iloc[k][col] == d_f.iloc[k+1][col]:
                                start_time = start[k]
                                lec = d_f.iloc[k][col]
                                duration = 2
                                k += 2
                            else:
                                start_time = start[k]
                                lec = d_f.iloc[k][col]
                                duration = 1
                                k += 1
                        start_time = list(datefinder.find_dates(start_time))
                        start_time = str(start_time[0].strftime('%H:%M:%S'))
                        start_time_str = start_date + " " + start_time
                        data_for_mail = lec.lower().split(' ')
                        std = data_for_mail[0]
                        if data_for_mail[0] == 'fy':
                            div = data_for_mail[2]
                            if 'a' in div or 'b' in div:
                                branch = 'comp'
                            elif 'c' in div or 'd' in div:
                                branch = 'extc'
                            elif 'e' in div or 'f' in div:
                                branch = 'etrx'
                            elif 'g' in div or 'h' in div:
                                branch = 'it'
                            if 'i' in div or 'j' in div:
                                branch = 'mech'
                        else:
                            div = data_for_mail[3]
                            branch = data_for_mail[2]
                        if branch not in branchlist:
                            continue
                        if semno == 'Even':
                            y_r = y_r - 1
                        year = str(y_r)
                        mailid = std + branch + '_' + div + '_' + year + '@somaiya.edu'
                        create_event(start_time_str, lec, duration, mailid)
                    except:
                        continue
                col += 1
                if col > 6:
                    col = col % 6 + 1
                    day = 3
                else:
                    day = 1
                start_date = list(datefinder.find_dates(start_date))
                start_date = start_date[0] + timedelta(days = day)
                start_date = start_date.strftime('%Y/%m/%d')
                window['-OUTPUT-'].update('Events Created Successfully!')
        except:
            window['-OUTPUT-'].update('Semester start date format is incorrect.')
    except:
        window['-OUTPUT-'].update('PDF format is incorrect.')


def create_event(start_time_str, summary, duration, mail):
    #establish connection with google calendar api and storing token in a file
    scopes = ['https://www.googleapis.com/auth/calendar']
    flow = InstalledAppFlow.from_client_secrets_file("client_secret_2.json", scopes = scopes)
    if os.path.exists('token.json'):
        credentials = Credentials.from_authorized_user_file('token.json', scopes)
    else:
        credentials = flow.run_local_server(port = 0)
        with open('token.json', 'w') as token:
            token.write(credentials.to_json())
    service  = build("calendar", "v3", credentials = credentials)
    recur = 'RRULE:FREQ=WEEKLY;COUNT=' + week
    time_zone = 'Asia/Kolkata'
    strt_time = list(datefinder.find_dates(start_time_str))
    end_time = strt_time[0] + timedelta(hours = duration)
    event = {
        'summary': summary,
        'start': {
            'dateTime': strt_time[0].strftime("%Y-%m-%dT%H:%M:%S"),
            'timeZone': time_zone,
        },
        'end': {
            'dateTime': end_time.strftime("%Y-%m-%dT%H:%M:%S"),
            'timeZone': time_zone,
        },
        "conferenceData": {
            "createRequest": {
                  "requestId": "xabf56hfsb",
                  "conferenceSolutionKey": {
                        "type": "hangoutsMeet"
                  },
            }
        },
        'recurrence': [
            recur
        ],
        'attendees': [
            {'email': mail},
        ],
        'reminders': {
            'useDefault': False,
            'overrides': [
              {'method': 'email', 'minutes': 30},
            ],
        },
    }
    service.events().insert(calendarId = "primary", body = event, sendNotifications = True, conferenceDataVersion = 1).execute()

#front end using pysimplegui
sg.theme("DarkTeal2")
layout = [[sg.T("")],
          [sg.Text("Choose timetable: "), sg.Input(key = "-IN2-" ,change_submits=True),
          sg.FileBrowse("Browse")],
          [sg.T("")],
          [sg.Text("Sem start date(dd month yyyy): "), sg.InputText(key = "semstart")],
          [sg.T("")],
          [sg.Text("Semester: "), sg.Combo(['Odd', 'Even'], default_value = 'Odd', key = "sem")],
          [sg.T("")],
          [sg.Text("Length of semester(in week nos.): "), sg.InputText(key = "week")],
          [sg.T("")],
          [sg.Text(size=(100,5), key='-OUTPUT-')],
          [sg.Button("Submit"), sg.Button("Exit")],
          [sg.T("")],
          [sg.Text("(NOTE: Event creation takes a few seconds,\
 do not press submit multiple times during this process,\
 to avoid creating multiple meets.)")]]

#Building Window
window = sg.Window('Meet Simplified', layout).Finalize()
window.Maximize()

while True:
    events, values = window.read()
    if events == "Submit":
        pdf_input = values["Browse"]
        sem_start = values["semstart"]
        sem = values["sem"]
        week = values["week"]
        dfcon()
    elif events in (sg.WIN_CLOSED, "Exit"):
        break
