from __future__ import print_function
import pandas as pd
import os
import os.path
import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive




def get_gmail_service():
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token_gmail.json'):
        creds = Credentials.from_authorized_user_file('token_gmail.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
                creds = flow.run_local_server(port=8080)
        # Save the credentials for the next run
        with open('token_gmail.json', 'w') as token:
            token.write(creds.to_json())

        try:
            # Call the Gmail API
            service = build('gmail', 'v1', credentials=creds)
            return service

        except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
            print(f'An error occurred: {error}')


def get_email_list():
    service = get_gmail_service()
    results = service.users().messages().list(userId='me',q='from:enter the email address subject: is:read').execute()
    return results.get('messages', [])[0].get('id', [])
    # return results.get('messages',[])

def get_email_content():
     x = get_email_list()
     service = get_gmail_service()
     message = service.users().messages().get(userId='me', id=x).execute()
     for part in message['payload']['parts']:
         if part['filename']:
             if 'data' in part['body']:
                 data = part['body']['data']
             else:
                 att_id = part['body']['attachmentId']
                 att = service.users().messages().attachments().get(userId='me', messageId=x,
                                                                    id=att_id).execute()

                 data = att['data']
                 file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                 path = part['filename']
                 # fullPath = path + ".msg"
                 with open(path, 'wb') as f:
                     f.write(file_data)
                     f.close()
                 df = pd.read_excel(path, skiprows=2)
                 pd.set_option('display.max_columns', None)
                 df.drop(df.columns[[0, 3, 4, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 27]],
                         axis=1, inplace=True)
                 df.drop(df.iloc[:, 12:28], inplace=True, axis=1)
                 # df1 = df.transpose()
                 df.to_excel(path, header = False, index_label=None,index=False)
                 return path

# df.to_csv(path + '.csv', header = False, index_label=None,index=False)
# print(get_email_content())

if __name__ == '__main__':
    get_email_content()

