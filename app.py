from __future__ import print_function
from flask import Flask, request
from pymessenger.bot import Bot
from pymessenger import Button
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import difflib



app = Flask(__name__)
ACCESS_TOKEN = 'EAAK62dbFXnwBAEcKIHoDWytbfqYgH55WjUVgKE7ZBBafD8O8iZBOKVDG9gNx3WpbME4YVPGoNZBHxMAhRogWUF3uKXkZCFM2l7wOrEmBzgNQZCxjSoPQN1UlGigv8icZCKPxrYSZAbpQXrK5EZAx5wNAGOrS3t3oAxjh78ZCZBwKLCCrJRqcPyci2ohz4pQVixZB4cZD'
VERIFY_TOKEN = 'hkajsdbasldj'
bot = Bot(ACCESS_TOKEN)

@app.route("/", methods=['GET', 'POST'])
def receive_message():
    if request.method == 'GET':
        """Before allowing people to message your bot, Facebook has implemented a verify token
        that confirms all requests that your bot receives came from Facebook."""
        sent_verify_token = request.args.get("hub.verify_token")
        if sent_verify_token == VERIFY_TOKEN:
            return request.args.get("hub.challenge")
    else:
        # get whatever message a user sent the bot
        events = request.get_json()
        for event in events['entry']:
            messages = event['messaging']
            for message in messages:
                if message.get('message'):
                    recipient_id = message['sender']['id']
                    text = message['message']['text']
                    response = generate_message(text, recipient_id)
                    send_message(recipient_id, response)
                elif message.get('postback'):
                    recipient_id = message['sender']['id']
                    text = message['postback']['payload']
                    response = generate_message(text, recipient_id)
                    send_message(recipient_id, response)

    return "Message Processed"


def generate_message(text, recipient_id):
    data = get_data()
    books = data['books']
    names = data['names']

    if text in names:
        index = names.index(text)
        try:
            change_color(index+1, index+2, 2, 3, B=1)
            return "Ok"
        except:
            return "Error"
    elif text in books:
        index = books.index(text)
        try:
            change_color(index+1, index+2, 2, 3, G=1)
            return "Ok"
        except:
            return "Error"
    else:
        words = names+books
        matches = difflib.get_close_matches(text, words)
        buttons = [Button(type='postback', title=match, payload=match) for match in matches]
        res = bot.send_button_message(recipient_id, "did you mean:", buttons)




def change_color(row_start, row_end, col_start, col_end, R=0, G=0, B=0):
    body = {
        "requests": [
            {
                "updateCells": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": row_start,
                        "endRowIndex": row_end,
                        "startColumnIndex": col_start,
                        "endColumnIndex": col_end
                    },
                    "rows": [
                        {
                            "values": [
                                {
                                    "userEnteredFormat": {
                                        "backgroundColor": {
                                            "red": R,
                                            "green": G,
                                            "blue": B
                                        }
                                    }
                                }
                            ]
                        }
                    ],
                    "fields": "userEnteredFormat.backgroundColor"
                }
            }
        ]
    }
    res = service.spreadsheets().batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
    print(res)


#uses PyMessenger to send response to user
def send_message(recipient_id, response):
    #sends user the text message provided via input response parameter
    bot.send_text_message(recipient_id, response)
    return "success"


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1CcHUsSqQ7x3cXbC3TfAoIEpUexz9lNHwwSak9pNOn9M'
SAMPLE_RANGE_NAMES = 'Sheet1!A2:A148'
SAMPLE_RANGE_Books = 'Sheet1!C2:C148'

creds = None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()


def get_data():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAMES).execute()
    names = result.get('values', [])
    names = [name[0] for name in names]
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_Books).execute()
    books = result.get('values', [])
    books =[book[0]for book in books]

    return {'names': names, 'books': books}

if __name__ == "__main__":
    app.run()