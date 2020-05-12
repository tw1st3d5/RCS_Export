import pandas as pd
from io import StringIO
import requests
import getpass
import csv
import os
import time

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
print(hhyrstart2)
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

print(payload)
response = requests.request("POST", url, headers=headers, data=payload, auth=(usern, pword))
print('Status Code:')
print(response.status_code)
file = 'Vertical Market.csv'
with open(file, "w", newline='') as f:
    writer = csv.writer(f)
    for line in response.iter_lines():
        writer.writerow(line.decode('utf-8').split(','))

time.sleep(5)

df = pd.read_csv('Vertical Market.csv')
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
df.to_excel('Vertical Market.xlsx', index=False, header=True)
os.remove('Vertical Market.csv')
