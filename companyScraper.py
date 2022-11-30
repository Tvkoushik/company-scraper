import requests
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials


SHEET_NAME = '2022-11-25 - LinkedIn Contacts'
SHEET_NUMBER = 1
API_KEY = '89f99489c4bd94c7b1152dc941875fd0'
API_SECRET = '8489e0bb76c4a17923b0d474297ed17a'
CELL_START_RANGE = 12596
CELL_END_RANGE = 12692
AUTH_JSON = 'koushik-349314-a25511eedd39.json'

# defining the scope of the application
scope_app = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

# credentials to the account
cred = ServiceAccountCredentials.from_json_keyfile_name(
    AUTH_JSON, scope_app)

# authorize the clientsheet
client = gspread.authorize(cred)

# get the sample of the Spreadsheet
sheet = client.open(SHEET_NAME)

# get the first sheet of the Spreadsheet
sheet_instance = sheet.get_worksheet(SHEET_NUMBER)


def get_access_token():
    params = {
        'grant_type': 'client_credentials',
        'client_id': API_KEY,
        'client_secret': API_SECRET
    }

    res = requests.post(
        'https://api.snov.io/v1/oauth/access_token', data=params)
    resText = res.text.encode('ascii', 'ignore')

    return json.loads(resText)['access_token']


def add_url_for_search(url):
    token = get_access_token()
    params = {'access_token': token,
              'url': url
              }

    res = requests.post(
        'https://api.snov.io/v1/add-url-for-search', data=params)

    return json.loads(res.text)


def get_emails_from_url(url):
    token = get_access_token()
    params = {'access_token': token,
              'url': url
              }

    res = requests.post(
        'https://api.snov.io/v1/get-emails-from-url', data=params)

    return json.loads(res.text)


def fill_sheet(CELL_START_RANGE, CELL_END_RANGE):
    for i in range(CELL_START_RANGE, CELL_END_RANGE):
        linked_url = sheet_instance.acell('I'+str(i)).value
        email = sheet_instance.acell('L'+str(i)).value
        if linked_url is None or email is not None:
            print('Skipping the Row')
            continue
        else:
            add_url_for_search(linked_url)
            new_email = ''
            try:
                full_data = get_emails_from_url(linked_url)['data']
                name = full_data['name']
                title = full_data['currentJob'][0]['position']
                email_list = full_data['emails']
                category = full_data['currentJob'][0]['industry']
                if not email_list:
                    new_email += 'Email Not Found,'
                else:
                    for prospect in email_list:
                        new_email += prospect['email']+','
                new_email = new_email[0:-1]

            except Exception as e:
                print('Could not find details')
                continue
            sheet_instance.update_acell('H'+str(i), name)
            sheet_instance.update_acell('G'+str(i), title)
            sheet_instance.update_acell('L'+str(i), new_email)
            sheet_instance.update_acell('B'+str(i), category)
            print(f'Updated Email for {linked_url} is {new_email}')
            print(f'Updated Name for {linked_url} is {name}')
            print(f'Updated Position for {linked_url} is {title}')
            print(f'Updated Industry for {linked_url} is {category}')


if __name__ == "__main__":
    fill_sheet(CELL_START_RANGE, CELL_END_RANGE)
