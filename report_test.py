import requests
import csv
import getpass
from datetime import datetime

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

dsdur = hhdstop - int(hhdstart)

# Reports to run ONCE per KPI period

# Vertical Market KPI

url = "https://portal.site-controls.net/report-service/rest/v1/clients/56/runreport/109?outputformat=csv"

payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22After%22%2C%22rank%22%3A1%2C" \
          "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + \
          str(hhyrstart) + '-' + str(hhmstart) + '-' + str(hhdstart) + \
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
if response.status_code == 504:
    url = "https://portal.site-controls.net/report-service/rest/v1/clients/56/runreport/109?outputformat=csv&emailresult=" + usern

    payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22After%22%2C%22rank%22%3A1%2C" \
              "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22after%22%2C%22value%22%3A%22" + \
              str(hhyrstart) + '-' + str(hhmstart) + '-' + str(hhdstart) + \
              "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
              "%22Before%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
              "%22before%22%2C%22value%22%3A%22" + str(hhyrstop) + '-' + str(hhmstop) + '-' + str(hhdstop) + \
              "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    print('Vertical Market will be emailed to: ' + usern)
else:
    file = 'Vertical Market.csv'
    print('Saving ' + file + '.')
    with open(file, "w", newline='') as f:
        writer = csv.writer(f)
        for line in response.iter_lines():
            writer.writerow(line.decode('utf-8').split(','))


# Active Site Count KPI

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
    url = "https://portal.site-controls.net/report-service/rest/v1/clients/56/runreport/6?outputformat=csv&emailresult=" + usern

    payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Date%22%2C%22rank%22%3A1%2C" \
              "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22day%22%2C%22value%22%3A%22" + \
              str(stopyr) + '-' + str(stopday) + '-' + str(stopmth2) + "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    print('Active Site Count will be emailed to: ' + usern)
else:
    file = 'Active Site Count.csv'
    print('Saving ' + file + '.')
    with open(file, "w", newline='') as f:
        writer = csv.writer(f)
        for line in response.iter_lines():
            writer.writerow(line.decode('utf-8').split(','))


# Run for EACH Client

clientIDarr = ['3', '80', '75', '72', '88', '20', '52', '49', '94', '46', '55', '81']
cont = 0
clientname = 'quit'
clientID = '1'
while cont < len(clientIDarr):
    for clientID in clientIDarr:

        if clientID == '56':
            clientname = 'ALDI'
        elif clientID == '3':
            clientname = 'MSI'
        elif clientID == '80':
            clientname = 'HFT'
        elif clientID == '75':
            clientname = 'TM'
        elif clientID == '72':
            clientname = 'SCVL'
        elif clientID == '88':
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


        # Ongoing Exceptions Sensor No-Comm Report: NOT USED ON MONTHLY KPI
        #
        # url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + \
        #       "/runreport/109?outputformat=csv"
#
        # payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
        #           "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22" \
        #           "%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
        #           "%22Event+Category%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22" \
        #           "%3A%22eventcat%22%2C%22value%22%3A%22Data%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A" \
        #           "%22ReportInputField%22%2C%22label%22%3A%22All+Clients%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type" \
        #           "%22%3A%22InputField%22%2C%22name%22%3A%22allclients%22%2C%22value%22%3A%22false%22%2C%22inputType%22%3A" \
        #           "%22PickOne%22%7D%7D%5D "
        # headers = {
        #     'Content-Type': 'application/x-www-form-urlencoded'
        # }

        # response = requests.request('POST', url=url, headers=headers, data=payload, auth=(usern, pword))
        # print(response.status_code)
        # file = 'Ongoing Exceptions' + clientname + '.csv'
        # print('Saving ' + file + '.')
        # with open(file, "w") as f:
        #     writer = csv.writer(f)
        #     for line in response.iter_lines():
        #         writer.writerow(line.decode('utf-8').split(','))

        # Sleep for 10 seconds
        # time.sleep(10)

        # Monthly KPI for Enterprise

        url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/98?outputformat=csv"

        payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                  "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22" \
                  "%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                  "%22Month+Begin%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                  "%22month_begin%22%2C%22value%22%3A%22" + str(startmth) + \
                  "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                  "%22Year+Begin%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                  "%22year_begin%22%2C%22value%22%3A%22" + str(startyr) + \
                  "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                  "%22Month+End%22%2C%22rank%22%3A4%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                  "%22month_end%22%2C%22value%22%3A%22" + str(stopmth) + \
                  "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                  "%22Year+End%22%2C%22rank%22%3A5%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                  "%22year_end%22%2C%22value%22%3A%22" + str(stopyr) + "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%5D "
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
        print()
        print('Monthly KPI for Enteriprise')
        print('Status Code:')
        print(response.status_code)
        if response.status_code == 504:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/98?outputformat=csv&emailresult=" + usern

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                      "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22" \
                      "%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                      "%22Month+Begin%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                      "%22month_begin%22%2C%22value%22%3A%22" + str(startmth) + \
                      "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                      "%22Year+Begin%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                      "%22year_begin%22%2C%22value%22%3A%22" + str(startyr) + \
                      "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                      "%22Month+End%22%2C%22rank%22%3A4%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                      "%22month_end%22%2C%22value%22%3A%22" + str(stopmth) + \
                      "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                      "%22Year+End%22%2C%22rank%22%3A5%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                      "%22year_end%22%2C%22value%22%3A%22" + str(stopyr) + "%22%2C%22inputType%22%3A%22PickOne%22%7D%7D%5D "
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print(clientname + 'Monthly KPI for Enterprise will be emailed to: ' + usern)
        else:
            file = 'KPI' + clientname + '.csv'
            print('Saving ' + file + '.')
            with open(file, "w", newline='') as f:
                writer = csv.writer(f)
                for line in response.iter_lines():
                    writer.writerow(line.decode('utf-8').split(','))


        # HVAC Health current year

        url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + \
              "/runreport/187?outputformat=csv"

        payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                  "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22" \
                  "%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                  "%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22after" \
                  "%22%2C%22value%22%3A%22" + str(hhyrstart) + '-' + str(hhmstart) + '-' + str(hhdstart) + \
                  "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                  "%22Before%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                  "%22before%22%2C%22value%22%3A%22" + str(hhyrstop) + '-' + str(hhmstop) + '-' + str(hhdstop) + \
                  "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "

        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
        print()
        print('HVAC Health Current Year')
        print('Status Code:')
        print(response.status_code)
        if response.status_code == 504:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + \
                  "/runreport/187?outputformat=csv&emailresult=" + usern

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                      "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22" \
                      "%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                      "%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22after" \
                      "%22%2C%22value%22%3A%22" + str(hhyrstart) + '-' + str(hhmstart) + '-' + str(hhdstart) + \
                      "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                      "%22Before%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                      "%22before%22%2C%22value%22%3A%22" + str(hhyrstop) + '-' + str(hhmstop) + '-' + str(hhdstop) + \
                      "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "

            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print(clientname + 'HVAC Health current year will be emailed to: ' + usern)
        else:
            file = 'HH' + clientname + '20.csv'
            print('Saving ' + file + '.')
            with open(file, "w", newline='') as f:
                writer = csv.writer(f)
                for line in response.iter_lines():
                    writer.writerow(line.decode('utf-8').split(','))


        # HVAC Health previous year

        url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + \
              "/runreport/187?outputformat=csv"

        payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                  "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22" \
                  "%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                  "%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22after" \
                  "%22%2C%22value%22%3A%22" + str(hhyrstart2) + '-' + str(hhmstart) + '-' + str(hhdstart) + \
                  "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                  "%22Before%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                  "%22before%22%2C%22value%22%3A%22" + str(hhyrstop2) + '-' + str(hhmstop) + '-' + str(hhdstop) + \
                  "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "

        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
        print()
        print('HVAC Health Previous Year')
        print('Status Code:')
        print(response.status_code)
        if response.status_code == 504:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + \
                  "/runreport/187?outputformat=csv&emailresult=" + usern

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                      "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22%22" \
                      "%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                      "%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22after" \
                      "%22%2C%22value%22%3A%22" + str(hhyrstart2) + '-' + str(hhmstart) + '-' + str(hhdstart) + \
                      "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                      "%22Before%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                      "%22before%22%2C%22value%22%3A%22" + str(hhyrstop2) + '-' + str(hhmstop) + '-' + str(hhdstop) + \
                      "%22%2C%22inputType%22%3A%22Date%22%7D%7D%5D "

            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print('HVAC Health previous year will be emailed to: ' + usern)
        else:
            file = 'HH' + clientname + '19.csv'
            print(clientname + 'Saving ' + file + '.')
            with open(file, "w", newline='') as f:
                writer = csv.writer(f)
                for line in response.iter_lines():
                    writer.writerow(line.decode('utf-8').split(','))


        # Demand Stats
        url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/110?outputformat=csv"

        payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                  "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22" \
                  "%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                  "%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                  "%22after%22%2C%22value%22%3A%22" + str(hhyrstart) + '-' + str(hhmstart) + '-' + str(hhdstart) + \
                  "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                  "%22Numofdays%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                  "%22numofdays%22%2C%22value%22%3A%22" + str(dsdur) + "%22%2C%22inputType%22%3A%22Integer%22%7D%7D%5D "

        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
        print()
        print('Demand Stats')
        print('Status Code:')
        print(response.status_code)
        if response.status_code == 504:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/110?outputformat=csv&emailresult=" + usern

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                      "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22" \
                      "%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                      "%22After%22%2C%22rank%22%3A2%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                      "%22after%22%2C%22value%22%3A%22" + str(hhyrstart) + '-' + str(hhmstart) + '-' + str(hhdstart) + \
                      "%22%2C%22inputType%22%3A%22Date%22%7D%7D%2C%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A" \
                      "%22Numofdays%22%2C%22rank%22%3A3%2C%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A" \
                      "%22numofdays%22%2C%22value%22%3A%22" + str(dsdur) + "%22%2C%22inputType%22%3A%22Integer%22%7D%7D%5D "

            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print(clientname + 'Demand Stats will be emailed to: ' + usern)
        else:
            file = 'DS' + clientname + '.csv'
            print('Saving ' + file + '.')
            with open(file, "w", newline='') as f:
                writer = csv.writer(f)
                for line in response.iter_lines():
                    writer.writerow(line.decode('utf-8').split(','))


        # Setpoints
        url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/78?outputformat=csv"

        payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                  "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22" \
                  "%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%5D "
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
        print()
        print('Setpoints')
        print('Status Code:')
        print(response.status_code)
        if response.status_code == 504:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/78?outputformat=csv&emailresult=" + usern

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                      "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22" \
                      "%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%5D "
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print(clientname + ' Setpoints will be emailed to ' + usern)
        else:
            file = 'SP' + clientname + '.csv'
            print('Saving ' + file + '.')
            with open(file, "w", newline='') as f:
                writer = csv.writer(f)
                for line in response.iter_lines():
                    writer.writerow(line.decode('utf-8').split(','))


        # Security Integration
        url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/266?outputformat=csv"

        payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                  "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22" \
                  "%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%5D "
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
        print()
        print('Security Integration')
        print('Status Code:')
        print(response.status_code)
        if response.status_code == 504:
            url = "https://portal.site-controls.net/report-service/rest/v1/clients/" + clientID + "/runreport/266?outputformat=csv&emailresult=" + usern

            payload = "inputs=%5B%7B%22type%22%3A%22ReportInputField%22%2C%22label%22%3A%22Clientdb%22%2C%22rank%22%3A1%2C" \
                  "%22inputField%22%3A%7B%22type%22%3A%22InputField%22%2C%22name%22%3A%22clientdb%22%2C%22value%22%3A%22" \
                  "%22%2C%22inputType%22%3A%22ClientDB%22%7D%7D%5D "
            headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
            print(clientname + ' Security integration will be emailed to: ' + usern)
        else:
            file = 'SI' + clientname + '.csv'
            print('Saving ' + file + '.')
            with open(file, "w", newline='') as f:
                writer = csv.writer(f)
                for line in response.iter_lines():
                    writer.writerow(line.decode('utf-8').split(','))

    cont = cont + 1


now = datetime.now()
stop_time = now.strftime("%H:%M:%S")
print("Current Time = ", stop_time)
total_time = stop_time - start_time
print("Total runtime is: ", total_time)
