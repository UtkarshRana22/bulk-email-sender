from __future__ import print_function

import os.path
from tkinter import messagebox

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://mail.google.com/']
global exists
exists=False

def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    happening=0
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        exists=True
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    else:
        exists=False
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                happening=1
                messagebox.showinfo('Gmail Authentication','User Token Refreshed!')
            except Exception as error:
                messagebox.showerror('Gmail Authentication Error!',error)
                os.remove('token.json')
                quit()
           
        
            
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            if exists!=True:
                
                messagebox.showinfo('Gmail Authentication','Authentication Successful!'+'\n'+'User Token Created Successfully!')
                happening=2

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        results = service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])
        if happening==0 or happening==3:
            messagebox.showinfo('Gmail Authentication','User Token Already Valid!')
        if not labels:
            print('No labels found.')
            return
        print('Labels:')
        for label in labels:
            print(label['name'])

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')




        
