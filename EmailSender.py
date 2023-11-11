import os,base64


from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials


from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

SCOPES= 'https://mail.google.com/'
# token checking
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

n=0
while os.path.exists('logfile'+str(n)+'.txt'):
    n+=1
logwrite=open('logfile'+str(n)+'.txt','x')


def sendemail(subject,body,emailto,kind):
    service = build('gmail', 'v1', credentials=creds)
    if kind=='html':
        message = MIMEText(body,'html')
    else:
        message=MIMEText(body)
    message['To'] = emailto
    message['From'] = 'ranautkarsh8709@gmail.com'
    message['Subject'] = subject

    # encoded message
    encoded_message = base64.urlsafe_b64encode(message.as_string().encode()).decode()
 #   print(encoded_message)
    create_message = {

        'raw': encoded_message
    }
    # sending message
    try:
        send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    except Exception as e:
        try:
            logwrite.write(e)
        except Exception as f:
            logwrite.write(f)

  
