import requests
import json
import csv
import getpass
from array import *
import pandas as pd
from io import StringIO

# Login to Siemens RCS API
usern = input("Username: ")
pword = getpass.getpass("Password: ")

url = "https://portal.site-controls.net/authentication-service/rest/v1/login?j_username=" + usern + "&j_password=" + \
      pword
payload = {}
headers = {
    'Content-Type': 'application/x-www-form-urlencoded',
    'Cookie': "REMEMBER_USER=" + usern
}

response = requests.request("POST", url, headers=headers, data=payload)
jsonResponse = response.json()
json_data = json.loads(response.text)
# print()
# print(json.dumps(json_data, indent=4))
# print()
# print(json_data['data']['sessionToken'])
print(json_data['data']['message'])
token = (json_data['data']['sessionToken'])

cont = 1
while cont == 1:

    data = 0
    data_type = input("HVAC or METER?").lower()
    if data_type == "meter":
        data = 1
    if data_type == "hvac":
        data = 2

    # METER with array
    while data == 1:

        print('Meter data pull request')
        clientID = input("Client ID: ")
        dbIDarr = array('i', [])
        nDB = int(input('Enter number of sites: '))
        for i in range(nDB):
            xdb = int(input('Enter the next Site dbID: '))
            dbIDarr.append(xdb)

        meterarr = array('i', [])
        nmeter = int(input('Enter number of meters per site: '))
        for i in range(nmeter):
            xmeter = int(input('Enter the next Meter Key #. DEMAND(1, 2, 3, etc): '))
            meterarr.append(xmeter)

        # meterkey = input("Meter key (DEMAND1, DEMAND2, DEMAND3, METER, etc): ")
        count = 1
        while count == 1:
            start_date = input("Start date (mm-dd-yyyy): ")
            stop_date = input("Stop date (mm-dd-yyyy): ")
            dbarrnum = 0
            for i in range(nDB):
                meterarrnum = 0
                dbID = str(dbIDarr[dbarrnum])
                for i in range(nmeter):
                    meterkey = str(meterarr[meterarrnum])
                    # print(dbID)
                    url = 'https://portal.site-controls.net/device-service/rest/v1/clients/' + clientID + '/sites/' + dbID \
                          + '/emeterdata/DEMAND' + meterkey + '.csv?start=' + start_date + '&finish=' + stop_date + \
                          '&filename=' + meterkey + '.csv '

                    payload = {}
                    headers = {
                        'Cookie': "org.springframework.web.servlet.i18n.CookieLocaleResolver.LOCALE=en_US; "
                                  "PORTALAUTHCOOKIE=" + token
                    }
                    response = requests.request("GET", url, headers=headers, data=payload)
                    data = StringIO(response.text)
                    # print(response.text.encode('utf8'))

                    # Save file
                    # file = clientID + '_' + dbID + '_DEMAND' + meterkey + '_' + start_date + '_' + stop_date + '.csv'
                    df = pd.read_csv(data)
                    for col in df.columns:
                        print(col)
                    cols = array('i', [0, 1, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14])
                    df.drop(df.columns[cols], axis=1, inplace=True)
                    df.dropna(axis=0, how='any', thresh=None, subset=None, inplace=False)
                    df.to_csv(clientID + '_' + dbID + '_DEMAND' + meterkey + '_' + start_date + '_' + stop_date + '.csv')
                    # file1 = clientID + '_' + dbID + '_DEMAND' + meterkey + '_' + start_date + '_' + stop_date
                    # save = input("The file will be saved as: " + file1 + " Do you want to modify this? (y/n)").lower()

                    # if save == "y":
                    # file = (input('What do you want the file name to be? ') + '.csv')
                    # else:
                    # file = clientID + '_' + dbID + '_DEMAND' + meterkey + '_' + start_date + '_' + stop_date + '.csv'
                    # with open(file, "w") as f:
                    #    writer = csv.writer(f)
                    #    for line in response.iter_lines():
                    #        writer.writerow(line.decode('utf-8').split(','))
                    print(clientID + '_' + dbID + '_DEMAND' + meterkey + '_' + start_date + '_' + stop_date + '.csv' + \
                          " saved to folder.")
                    meterarrnum = meterarrnum + 1
                dbarrnum = dbarrnum + 1

            # Return or quit
            goback = input("Do you want to pull another report for this site and meter? (y/n) ").lower()

            if goback == "y":
                count = 1
                data = 1
            else:
                goback1 = input("Do you want to change sites or clients? (y/n)").lower()
                if goback1 == "n":
                    quit()
                else:
                    count = 0
                    data = 0
                    cont = 1

    # HVAC
    while data == 2:

        print('HVAC data pull request')
        clientID = input("Client ID: ")
        dbID = input("Site Database ID: ")
        hvackey = input("HVAC key (HVAC1, HVAC2, HVAC3, etc): ")

        count = 1
        while count == 1:
            start_date = input("Start date (mm-dd-yyyy): ")
            stop_date = input("Stop date (mm-dd-yyyy): ")
            url = 'https://portal.site-controls.net/device-service/rest/v1/clients/' + clientID + '/sites/' + dbID + \
                  '/hvacdata/' + hvackey + '.csv?start=' + start_date + '&finish=' + stop_date + '&filename=' + hvackey + \
                  '.csv '
            payload = {}
            headers = {
                'Cookie': "org.springframework.web.servlet.i18n.CookieLocaleResolver.LOCALE=en_US; PORTALAUTHCOOKIE=" +
                          token
            }

            response = requests.request("GET", url, headers=headers, data=payload)

            # print(response.text.encode('utf8'))
            file1 = clientID + '_' + dbID + '_' + hvackey + '_' + start_date + '_' + stop_date
            save = input("The file will be saved as: " + file1 + " Do you want to modify this? (y/n)").lower()

            if save == "y":
                file = (input('What do you want the file name to be? ') + '.csv')
            else:
                file = clientID + '_' + dbID + '_' + hvackey + '_' + start_date + '_' + stop_date + '.csv'
            with open(file, "w") as f:
                writer = csv.writer(f)
                for line in response.iter_lines():
                    writer.writerow(line.decode('utf-8').split(','))
            print(file + " saved to folder.")

            # Return or quit
            goback = input("Do you want to pull another report for this site and HVAC? (y/n) ").lower()

            if goback == "y":
                count = 1
                data = 2
            else:
                goback1 = input("Do you want to change sites or clients? (y/n)").lower()
                if goback1 == "n":
                    quit()
                else:
                    count = 0
                    data = 0
                    cont = 1