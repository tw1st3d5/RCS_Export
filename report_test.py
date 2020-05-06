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

clientID = input('ClientID:')

# print(dbID)
url = 'https://portal.site-controls.net/report-service/rest/v1/clients/' + clientID + '/runreport/98?outputformat=csv'
payload = {}
files = {
     {"type":"ReportInputField","label":"Clientdb","rank":1,"inputField":{"type":"InputField","name":"clientdb","value":"","inputType":"ClientDB"}},{"type":"ReportInputField","label":"Event Category","rank":2,"inputField":{"type":"InputField","name":"eventcat","value":"All","inputType":"PickOne"}},{"type":"ReportInputField","label":"All Clients","rank":3,"inputField":{"type":"InputField","name":"allclients","value":"false","inputType":"PickOne"}}
}
headers = {
    'Cookie': "org.springframework.web.servlet.i18n.CookieLocaleResolver.LOCALE=en_US; "
              "Content-Type: application/x-www-form-urlencoded"
              "PORTALAUTHCOOKIE=" + token
}

response = requests.request("GET", url, headers=headers, data=payload, files=files)
print(response.text)
data = StringIO(response.text)
df = pd.read_csv(data)
df.to_csv('Report98.csv')
