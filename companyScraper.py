import requests
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import yaml


with open("config.yaml", "r") as f:
    config = yaml.load(f, Loader=yaml.FullLoader)

SHEET_NAME = config['sheetName']
SHEET_NUMBER = config['sheetNumber']
API_KEY = config['apiKey']
API_SECRET = config['apiSecret']
CELL_START_RANGE = config['startRange']
CELL_END_RANGE = config['endRange']
AUTH_JSON = config['authJson']
LINKEDIN_COL = config['linkedinColumn']
EMAIL_COL = config['emailColumn']
NAME_COL = config['nameColumn']
TITLE_COL = config['titleColumn']
CATEGORY_COL = config['categoryColumn']

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
        linked_url = sheet_instance.acell(LINKEDIN_COL+str(i)).value
        email = sheet_instance.acell(EMAIL_COL+str(i)).value
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
            sheet_instance.update_acell(NAME_COL + str(i), name)
            sheet_instance.update_acell(TITLE_COL + str(i), title)
            sheet_instance.update_acell(EMAIL_COL + str(i), new_email)
            sheet_instance.update_acell(CATEGORY_COL + str(i), category)
            print(f'Updated Email for {linked_url} is {new_email}')
            print(f'Updated Name for {linked_url} is {name}')
            print(f'Updated Position for {linked_url} is {title}')
            print(f'Updated Industry for {linked_url} is {category}')


if __name__ == "__main__":
    fill_sheet(CELL_START_RANGE, CELL_END_RANGE)
