from flask import Flask
import requests

# Replace these with your Azure App details
CLIENT_ID = '1140a629-6ea1-41ec-9655-d5e1afab2408'
CLIENT_SECRET = 'wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ'
TENANT_ID = '6940843a-674d-4941-9ca2-dc5603f278df'

# Your file and Excel details
FILE_NAME = 'calculator.xlsx'
FILE_PATH = f'/Documents/{FILE_NAME}'  # Adjust path to where it's stored in OneDrive
SHEET_NAME = 'Sheet1'
CELL_ADDRESS = 'B2'  # Read this cell
UPDATE_CELL = 'A1'
UPDATE_VALUE = 'Hello from Flask!'

app = Flask(__name__)

def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'client_id': CLIENT_ID,
        'scope': 'https://graph.microsoft.com/.default',
        'client_secret': CLIENT_SECRET,
        'grant_type': 'client_credentials'
    }
    response = requests.post(url, headers=headers, data=data)
    return response.json()['access_token']

def update_cell(access_token):
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:{FILE_PATH}:/workbook/worksheets('{SHEET_NAME}')/range(address='{UPDATE_CELL}')"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    body = {
        "values": [[UPDATE_VALUE]]
    }
    response = requests.patch(url, headers=headers, json=body)
    return response.status_code == 200

def read_cell(access_token):
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:{FILE_PATH}:/workbook/worksheets('{SHEET_NAME}')/range(address='{CELL_ADDRESS}')"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()['values'][0][0]
    return "Error reading cell"

@app.route('/')
def index():
    token = get_access_token()
    update_cell(token)
    value = read_cell(token)
    return f"<h1>Cell {CELL_ADDRESS} says: {value}</h1>"

if __name__ == '__main__':
    app.run(debug=True)
