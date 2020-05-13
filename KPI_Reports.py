import requests
import getpass
from datetime import datetime
import time
import pandas as pd
from io import StringIO
from pandas import ExcelWriter
import glob
import os
from openpyxl import load_workbook

now = datetime.now()
start_time = now.strftime("%H:%M:%S")
print("Current Time = ", start_time)

# Login to Siemens RCS API
usern = input("Username: ")
pword = getpass.getpass("Password: ")

startmth = int(input('Data for 2 years start month (1, 2, etc)'))
startday = int(input('Data for 2 years Start day (dd)'))
startyr = int(input('Data for 2 years start year (yyyy)'))
stopmth = int(input('Data for 2 years stop month (1, 2, etc)'))
stopday = int(input('Data for 2 years Stop day (dd)'))
stopyr = int(input('Data for 2 years stop year (yyyy)'))

if startmth < 10:
    mth = str(startmth)
    startmth2 = '0' + mth
else:
    startmth2 = startmth

if stopmth < 10:
    mth = str(stopmth)
    stopmth2 = '0' + mth
else:
    stopmth2 = stopmth

hhyrstart = stopyr
hhmstart = stopmth
hhdstart = '01'
hhyrstop = stopyr
hhmstop = stopmth2
hhdstop = stopday
hhyrstart2 = hhyrstart - 1
hhyrstop2 = hhyrstop - 1

dsdur = int(hhdstop) - int(hhdstart)

# Reports to run ONCE per KPI period

# Vertical Market KPI
print()
print('Pulling Vertical Market KPI')
url = "https://portal.site-controls.net/report-service/rest/v1/clients/56/runreport/114?outputformat=csv"

payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22After%22%2C%22rank%22%3A1%2C" \
          "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + \
          str(hhyrstart2) + '-' + str(hhmstop) + '-' + hhdstart + \
          "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
          "%22Before%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
          "%22before%22%2C%22value%22%3A%22" + str(hhyrstop) + '-' + str(hhmstop) + '-' + str(hhdstop) + \
          "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "
headers = {
    'Content-Type': 'application/x-www-form-urlencoded'
}

response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
print('Status Code:')
print(response.status_code)
df = pd.read_csv(StringIO(response.text))
df.columns = df.columns.str.strip().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
df = df[df.Vertical_Market != 'Health Club']
df = df[df.Vertical_Market != 'Restaurant']
df = df[df.Vertical_Market != 'Other']
df = df[df.Vertical_Market != 'Grocery']
df = df[df.Vertical_Market != 'Financial']
df = df[df.Vertical_Market != 'Manufacturing']
df = df[df.Vertical_Market != 'Demo']
df.columns.str.replace('VerticalMarket', 'Vertical Market')
df.columns = df.columns.str.strip().str.replace('_', ' ').str.replace('(', '').str.replace(')', '')
df.to_excel('Vertical Market.xlsx', index=False, header=True, sheet_name='Vertical Market')
print('Vertical Market.xlsx saved')

# Active Site Count KPI
cont = 0
print()
print('Pulling Active Site Count')
time.sleep(5)
while cont == 0:
    url = "https://portal.site-controls.net/report-service/rest/v1/clients/56/runreport/6?outputformat=csv"

    payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Date%22%2C%22rank%22%3A1%2C" \
              "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22day%22%2C%22value%22%3A%22" + \
              str(stopyr) + '-' + str(stopday) + '-' + str(stopmth2) + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }

    response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
    print('Status Code:')
    print(response.status_code)
    if response.status_code == 504:
        url = "https://portal.site-controls.net/report-service/rest/v1/clients/56/runreport/6?outputformat=csv&" \
                "emailresult=" + usern

        payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Date%22%2C%22rank%22%3A1%2C" \
                  "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22day%22%2C%22value%22%3A%22" + \
                  str(stopyr) + '-' + str(stopday) + '-' + str(stopmth2) + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        print('Active Site Count will be emailed to: ' + usern)
    if response.status_code == 400:
        url = "https://portal.site-controls.net/report-service/rest/v1/clients/56/runreport/6?outputformat=csv&" \
              "emailresult=" + usern

        payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Date%22%2C%22rank%22%3A1%2C" \
                  "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22day%22%2C%22value%22%3A%22" + \
                  str(stopyr) + '-' + str(stopday) + '-' + str(stopmth2) + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        print('Active Site Count will be emailed to: ' + usern)
    if response.status_code == 200:
        file = response.text
        df = pd.read_csv(StringIO(file), sep=',')
        df.to_excel('Active Site Count.xlsx', index=False, header=True, sheet_name='Active Site Count')
        cont = cont + 1
    else:
        print('An error occured when pulling the Active Site Count Report. This report will need to be processed manually.')
        cont = cont + 1

# Run these reports for EACH Client

clientIDarr = ['3', '80', '75', '72', '28', '20', '52', '49', '94', '46', '55', '81']
clientname = 'quit'
clientID = '1'
for clientID in clientIDarr:

    if clientID == '3':
        clientname = 'MSI'
    elif clientID == '80':
        clientname = 'HFT'
    elif clientID == '75':
        clientname = 'TM'
    elif clientID == '72':
        clientname = 'SCVL'
    elif clientID == '28':
        clientname = 'WM'
    elif clientID == '20':
        clientname = 'BL'
    elif clientID == '52':
        clientname = 'CT'
    elif clientID == '49':
        clientname = 'DSW'
    elif clientID == '94':
        clientname = 'SB'
    elif clientID == '46':
        clientname = 'B5'
    elif clientID == '55':
        clientname = '99'
    elif clientID == '81':
        clientname = 'TB'

    print()
    print('Current Client:')
    print(clientname)

    # Monthly KPI for Enterprise
    cont = 0
    time.sleep(5)
    print()
    print('Monthly KPI for Enterprise ' + clientname)
    while cont == 0:
        url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + \
              "/runreport/98?outputformat=csv"

        payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1" \
                  "%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22" \
                  "%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C" \
                  "%22label%22%3A%22Month+Begin%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A" \
                  "%22InputField%22%2C%22name%22%3A%22month_begin%22%2C%22value%22%3A%22" + str(startmth) + \
                  "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22" \
                  "%3A%22Year+Begin%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C" \
                  "%22name%22%3A%22year_begin%22%2C%22value%22%3A%22" + str(startyr) + \
                  "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22" \
                  "%3A%22Month+End%22%2C%22rank%22%3A4%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C" \
                  "%22name%22%3A%22month_end%22%2C%22value%22%3A%22" + str(stopmth) + \
                  "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22" \
                  "%3A%22Year+End%22%2C%22rank%22%3A5%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name" \
                  "%22%3A%22year_end%22%2C%22value%22%3A%22" + str(stopyr) + \
                  "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%5D "
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
        print('Status Code:')
        print(response.status_code)
        if response.status_code == 400:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + \
                  "/runreport/98?outputformat=csv&emailresult=" + usern

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22" \
                      "%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C" \
                      "%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A" \
                      "%22ReportInputField%22%2C%22label%22%3A%22Month+Begin%22%2C%22rank%22%3A2%2C%22inputField%22" \
                      "%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22month_begin%22%2C%22value%22%3A%22" + \
                      str(startmth) + "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A" \
                                      "%22ReportInputField%22%2C%22label%22%3A%22Year+Begin%22%2C%22rank%22%3A3%2C" \
                                      "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                                      "%22year_begin%22%2C%22value%22%3A%22" + str(startyr) + \
                      "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C" \
                      "%22label%22%3A%22Month+End%22%2C%22rank%22%3A4%2C%22inputField%22%3A%7B%22type%22%3A" \
                      "%22InputField%22%2C%22name%22%3A%22month_end%22%2C%22value%22%3A%22" + str(stopmth) + \
                      "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C" \
                      "%22label%22%3A%22Year+End%22%2C%22rank%22%3A5%2C%22inputField%22%3A%7B%22type%22%3A" \
                      "%22InputField%22%2C%22name%22%3A%22year_end%22%2C%22value%22%3A%22" + str(stopyr) + \
                      "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%5D "
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print(clientname + 'Monthly KPI for Enterprise will be emailed to: ' + usern)
        if response.status_code == 504:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + \
                  "/runreport/98?outputformat=csv&emailresult=" + usern

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22" \
                      "%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C" \
                      "%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A" \
                      "%22ReportInputField%22%2C%22label%22%3A%22Month+Begin%22%2C%22rank%22%3A2%2C%22inputField%22" \
                      "%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22month_begin%22%2C%22value%22%3A%22" + \
                      str(startmth) + "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A" \
                                      "%22ReportInputField%22%2C%22label%22%3A%22Year+Begin%22%2C%22rank%22%3A3%2C" \
                                      "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                                      "%22year_begin%22%2C%22value%22%3A%22" + str(startyr) + \
                      "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C" \
                      "%22label%22%3A%22Month+End%22%2C%22rank%22%3A4%2C%22inputField%22%3A%7B%22type%22%3A" \
                      "%22InputField%22%2C%22name%22%3A%22month_end%22%2C%22value%22%3A%22" + str(stopmth) + \
                      "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C" \
                      "%22label%22%3A%22Year+End%22%2C%22rank%22%3A5%2C%22inputField%22%3A%7B%22type%22%3A" \
                      "%22InputField%22%2C%22name%22%3A%22year_end%22%2C%22value%22%3A%22" + str(stopyr) + \
                      "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%5D "
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print(clientname + 'Monthly KPI for Enterprise will be emailed to: ' + usern)
        if response.status_code == 200:
            file = response.text
            df = pd.read_csv(StringIO(file), sep=',')
            df.to_excel('KPI' + clientname + '.xlsx', index=False, header=True, sheet_name='KPI' + clientname)
            cont = cont + 1
        else:
            print('An error occured when pulling the HVAC Health - previous year report for ' + clientname + '. This report will need to be processed manually.')
            cont = cont + 1

# HVAC Health current year

    print()
    print('HVAC Health - current year ' + clientname)
    time.sleep(5)
    cont = 0
    while cont == 0:
        try:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + \
                  "/runreport/187?outputformat=csv"

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1" \
                      "%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22" \
                      "%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C" \
                      "%22label%22%3A%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22" \
                      "%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + str(hhyrstart) + '-' + str(hhmstart) + '-' + \
                      str(hhdstart) + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22" \
                                      "%2C%22label%22%3A%22Before%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22" \
                                      "%3A%22InputField%22%2C%22name%22%3A%22before%22%2C%22value%22%3A%22" + \
                      str(hhyrstop) + '-' + str(hhmstop) + '-' + str(hhdstop) + \
                      "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D"

            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }

            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print()
            print('Status Code:')
            print(response.status_code)
            if response.status_code == 400:
                url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/187" \
                                                                                                      "?outputformat=csv" \
                                                                                                      "&emailresult=" + \
                      usern

                payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22" \
                          "%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C" \
                          "%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A" \
                          "%22ReportInputField%22%2C%22label%22%3A%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B" \
                          "%22type%22%3A%22InputField%22%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + str(hhyrstart) \
                          + '-' + str(hhmstart) + '-' + str(hhdstart) + \
                          "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label" \
                          "%22%3A%22Before%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C" \
                          "%22name%22%3A%22before%22%2C%22value%22%3A%22" + str(hhyrstop) + '-' + str(hhmstop) + '-' + \
                          str(hhdstop) + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "

                headers = {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
                response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
                print(clientname + 'HVAC Health current year will be emailed to: ' + usern)
            if response.status_code == 504:
                url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/187" \
                                                                                                      "?outputformat=csv" \
                                                                                                      "&emailresult=" + \
                      usern

                payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22" \
                          "%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C" \
                          "%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A" \
                          "%22ReportInputField%22%2C%22label%22%3A%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B" \
                          "%22type%22%3A%22InputField%22%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + str(hhyrstart) \
                          + '-' + str(hhmstart) + '-' + str(hhdstart) + \
                          "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label" \
                          "%22%3A%22Before%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C" \
                          "%22name%22%3A%22before%22%2C%22value%22%3A%22" + str(hhyrstop) + '-' + str(hhmstop) + '-' + \
                          str(hhdstop) + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "

                headers = {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
                response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
                print(clientname + 'HVAC Health current year will be emailed to: ' + usern)
            if response.status_code == 200:
                file = response.text
                df = pd.read_csv(StringIO(file), sep=',')
                df.to_excel('HH' + clientname + '20.xlsx', index=False, header=True, sheet_name='HH' + clientname + '20')
                cont = cont + 1
            else:
                print('An error occured when pulling the HVAC Health - previous year report for ' + clientname + '. This report will need to be processed manually.')
                cont = cont + 1
        except:
            print('An error occured when pulling the HVAC Health - previous year report for ' + clientname + '. This report will need to be processed manually.')
            cont = cont + 1

# HVAC Health previous year

    print()
    print('HVAC Health - previous year ' + clientname)
    time.sleep(5)
    cont = 0
    while cont == 0:
        try:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/187" \
                                                                                              "?outputformat=csv "

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1" \
                      "%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22" \
                      "%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C" \
                      "%22label%22%3A%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22" \
                      "%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + str(hhyrstart2) + '-' + str(hhmstart) + '-' + \
                      str(hhdstart) + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22" \
                                      "%2C%22label%22%3A%22Before%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22" \
                                      "%3A%22InputField%22%2C%22name%22%3A%22before%22%2C%22value%22%3A%22" + \
                      str(hhyrstop2) + '-' + str(hhmstop) + '-' + str(hhdstop) + "%22%2C%22inputType%22%3A%22Date%22%7D" \
                                                                                 "%7D%5D "

            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }

            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print('HVAC Health Previous Year')
            print('Status Code:')
            print(response.status_code)
            if response.status_code == 400:
                url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + \
                      "/runreport/187?outputformat=csv&emailresult=" + usern

                payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22" \
                          "%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C" \
                          "%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A" \
                          "%22ReportInputField%22%2C%22label%22%3A%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B" \
                          "%22type%22%3A%22InputField%22%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + \
                          str(hhyrstart2) + '-' + str(hhmstart) + '-' + str(hhdstart) + \
                          "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label" \
                          "%22%3A%22Before%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C" \
                          "%22name%22%3A%22before%22%2C%22value%22%3A%22" + str(hhyrstop2) + '-' + str(hhmstop) + '-' + \
                          str(hhdstop) + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "

                headers = {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
                response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
                print('HVAC Health previous year will be emailed to: ' + usern)
            if response.status_code == 504:
                url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + \
                  "/runreport/187?outputformat=csv&emailresult=" + usern

                payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22" \
                          "%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C" \
                          "%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A" \
                          "%22ReportInputField%22%2C%22label%22%3A%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B" \
                          "%22type%22%3A%22InputField%22%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + \
                          str(hhyrstart2) + '-' + str(hhmstart) + '-' + str(hhdstart) + \
                          "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label" \
                          "%22%3A%22Before%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C" \
                          "%22name%22%3A%22before%22%2C%22value%22%3A%22" + str(hhyrstop2) + '-' + str(hhmstop) + '-' + \
                          str(hhdstop) + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "

                headers = {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
                response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
                print('HVAC Health previous year will be emailed to: ' + usern)
            if response.status_code == 200:
                file = response.text
                df = pd.read_csv(StringIO(file), sep=',')
                df.to_excel('HH' + clientname + '19.xlsx', index=False, header=True, sheet_name='HH' + clientname + '19')
                cont = cont + 1
            else:
                print('An error occured when pulling the HVAC Health - previous year report for ' + clientname + '. This report will need to be processed manually.')
                cont = cont + 1
        except:
            print('An error occured when pulling the HVAC Health - previous year report for ' + clientname + '. This report will need to be processed manually.')
            cont = cont + 1

# Demand Stats - Current Year

    print()
    print('Demand Stats - Current Year ' + clientname)
    time.sleep(5)
    cont = 0
    while cont == 0:
        try:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/110?outputformat=csv"
            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + str(stopyr) + '-' + str(stopmth) + '-' + hhdstart + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Numofdays%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22numofdays%22%2C%22value%22%3A%22" + str(dsdur) + "%22%2C%22inputType%22%3A%22Integer%22%7D%7D%5D"

            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print('Status Code:')
            print(response.status_code)
            if response.status_code == 400:
                url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/110?outputformat=csv&emailresult=" + usern

                payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + str(stopyr) + '-' + str(stopmth) + '-' + hhdstart + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Numofdays%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22numofdays%22%2C%22value%22%3A%22" + str(dsdur) + "%22%2C%22inputType%22%3A%22Integer%22%7D%7D%5D"

                headers = {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
                response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
                print(clientname + 'Demand Stats will be emailed to: ' + usern)
            if response.status_code == 504:
                url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/110?outputformat=csv&emailresult=" + usern

                payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + str(stopyr) + '-' + str(stopmth) + '-' + hhdstart + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Numofdays%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22numofdays%22%2C%22value%22%3A%22" + str(dsdur) + "%22%2C%22inputType%22%3A%22Integer%22%7D%7D%5D"

                headers = {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
                response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
                print(clientname + 'Demand Stats will be emailed to: ' + usern)
            if response.status_code == 200:
                file = response.text
                df = pd.read_csv(StringIO(file), sep=',')
                df.to_excel('DS' + clientname + '.xlsx', index=False, header=True, sheet_name='DS' + clientname)
                cont = cont + 1
            else:
                print('An error occured when pulling the Demand Stats report for ' + clientname + '. This report will need to be processed manually.')
                cont = cont + 1
        except:
            print('An error occured when pulling the Demand Stats report for ' + clientname + '. This report will need to be processed manually.')
            cont = cont + 1

    # Setpoints
    print()
    print('Setpoints ' + clientname)
    time.sleep(5)
    cont = 0
    while cont == 0:
        try:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/78?outputformat=csv"

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%5D"
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }

            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print('Status Code:')
            print(response.status_code)
            if response.status_code == 400:
                url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/78?outputformat=csv&emailresult=" + usern

                payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%5D"
                headers = {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
                response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
                print(clientname + ' Setpoints will be emailed to ' + usern)
            if response.status_code == 504:
                url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/78?outputformat=csv&emailresult=" + usern

                payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%5D"
                headers = {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
                response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
                print(clientname + ' Setpoints will be emailed to ' + usern)

            if response.status_code == 200:
                file = response.text
                df = pd.read_csv(StringIO(file), sep=',')
                df.to_excel('SP' + clientname + '.xlsx', index=False, header=True, sheet_name='SP' + clientname)
                cont = cont + 1
            else:
                print('An error occured when pulling the Setpoints report for ' + clientname + '. This report will need to be processed manually.')
                cont = cont + 1
        except:
            print('An error occured when pulling the Setpoints report for ' + clientname + '. This report will need to be processed manually.')
            cont = cont + 1

# Security Integration
    print()
    print('Security Integration ' + clientname)
    time.sleep(5)
    cont = 0
    while cont == 0:
        try:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/266?outputformat=csv"

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%5D"
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print('Status Code:')
            print(response.status_code)
            if response.status_code == 400:
                url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/266?outputformat=csv&emailresult=" + usern

                payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%5D"
                headers = {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
                response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
                print(clientname + ' Security integration will be emailed to: ' + usern)
            if response.status_code == 504:
                url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/266?outputformat=csv&emailresult=" + usern

                payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%5D"
                headers = {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
                response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
                print(clientname + ' Security integration will be emailed to: ' + usern)

            if response.status_code == 200:
                file = response.text
                df = pd.read_csv(StringIO(file), sep=',')
                df.to_excel('SI' + clientname + '.xlsx', index=False, header=True, sheet_name='SI' + clientname)
                cont = cont + 1
            else:
                print('An error occurred when attempting to pull the Security Integration Report for ' + clientname +'. This report will need to be processed manually.')
                cont = cont + 1
        except:
            print('An error occurred when attempting to pull the Security Integration Report for ' + clientname +'. This report will need to be processed manually.')
            cont = cont + 1


print()
print('Report pulls completed')
comb = input("When all emailed reports are saved to the folder press enter to continue")
print()
print('Combining CSV sheets into one XLSX file.')
writer = ExcelWriter("combined_csv.xlsx")

for filename in glob.glob("*.csv"):
    csv_file = pd.read_csv(filename)
    (_, f_name) = os.path.split(filename)
    (f_short_name, _) = os.path.splitext(f_name)
    sheet_name = f_short_name
    df_excel = pd.read_csv(filename)
    df_excel.to_excel(writer, sheet_name, index=False)

writer.save()
print()
print('Combining all XLSX sheets into one file.')
writer = ExcelWriter("combined_all_unsorted.xlsx")

for filename in glob.glob("*.xlsx"):
    excel_file = pd.ExcelFile(filename)
    (_, f_name) = os.path.split(filename)
    (f_short_name, _) = os.path.splitext(f_name)
    for sheet_name in excel_file.sheet_names:
        df_excel = pd.read_excel(filename, sheet_name=sheet_name)
        df_excel.to_excel(writer, sheet_name, index=False)

writer.save()
print()
print('Sorting workbook tabs alphabetically.')
wb = load_workbook('combined_all_unsorted.xlsx')

wb._sheets.sort(key=lambda ws: ws.title)

wb.save('KPI Master.xlsx')
print()
print('KPI Master.xlsx saved.')

now = datetime.now()
stop_time = now.strftime("%H:%M:%S")
print("Start Time =", start_time)
print("Current Time = ", stop_time)
