import os.path

import base64
from email.message import EmailMessage

import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

'''

kvgmailsendsimple.py - send out simple emails from gmail using API methods

Main routine is:  gmail_send_simple()

This will create OATH json files to be used when executing.
And will require a one time authentication by the "email_from" user account
and approval to use this application

you will need to  install the following to make this work

pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib


@author:  Ken Venner
@contact: ken@vennerllc.com
@version:  1.04


Created:  2024-02-18;kv
Version:  2024-02-18;kv

'''


# If modifying these scopes, delete the file token.json.
SCOPES = [
  "https://www.googleapis.com/auth/gmail.readonly"
  ,"https://www.googleapis.com/auth/gmail.send"
]


# version number
AppVersion = '1.04'



def convert_email_to_filename( email_addr, file_ext='json' ):
  """ take an email address field in and convert characters to create a filename and an extension

    email_addr - input email address
    file_ext - the file extension to add to the filename
  """
  
  filename = email_addr.replace('@', '_')
  filename = filename.replace('.', '_')
  if file_ext[0] != '.':
    filename = filename + '.'
  return filename + file_ext


def google_creds_from_json(scopes=None, file_token_json=None, file_credentials_json=None):
  """ get and return creds from json.
      scopes - the scopes you are asking to be given permissions to - must be populated
      file_token_json - the tokens.json or email driven filename (eg. 210608thSt_gmail.json)
                        file created after we get permissions for this user
      file_credentials_json - the OATH file we are buildng tokens from

      the guide to getting this setup and executing:
      https://developers.google.com/gmail/api/quickstart/python

      add the email address to the list of users that can use the application
      create the token and save locally as "credential.json"
      run this program and it will generate the email based json file
      and force the local browser to authenticate

  """
  # if we don't have scopes - we error out
  if not scopes:
    scopes = SCOPES
  if not file_token_json:
    file_token_json = 'token.json'
  if not file_credentials_json:
    file_credentials_json = 'credentials.json'

  
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists(file_token_json):
    creds = Credentials.from_authorized_user_file(file_token_json, scopes)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          file_credentials_json, scopes
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open(file_token_json, "w") as token:
      token.write(creds.to_json())

  return creds

def gmail_refresh_token_take_no_action(email_from, scopes=None, file_token_json=None, file_credentials_json=None):
  '''
  When we don't send an email we should refresh the token to assure it stays current
  '''
  # determien the token.json file
  if not file_token_json:
    file_token_json = convert_email_to_filename(email_from)

   # set the credentials
  creds = google_creds_from_json(scopes, file_token_json, file_credentials_json)
  #creds, _ = google.auth.default()
  

def gmail_send_simple_message(email_from, email_to, email_subject, email_body, scopes=None, file_token_json=None, file_credentials_json=None):
  """Create and send an email message
  Print the returned  message id
  Returns: Message object, including message id

  Load pre-authorized user credentials from the environment.
  TODO(developer) - See https://developers.google.com/identity
  for guides on implementing OAuth2 for the application.

  email_from - the account that is sending out the email
  email_to - the address or set of addresses we are sending emails to
  email_subject - subject line of the email
  email_body - the text in the body being sent

  scope - the application scope we are giving out - if not sent we use the value set in SCOPES
  file_token_json - the filename holding the auth token for this email_from (not set we create it from the email_from address)
  file_credentials_json - the filename holding the OATH app approval credentials (default:  credentials.json)

  """
  # determien the token.json file
  if not file_token_json:
    file_token_json = convert_email_to_filename(email_from)

   # set the credentials
  creds = google_creds_from_json(scopes, file_token_json, file_credentials_json)
  #creds, _ = google.auth.default()

  try:
    service = build("gmail", "v1", credentials=creds)
    message = EmailMessage()

    message.set_content(email_body)

    message["To"] = email_to
    message["From"] = email_from
    message["Subject"] = email_subject

    # encoded message
    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    create_message = {"raw": encoded_message}
    # pylint: disable=E1101
    send_message = (
        service.users()
        .messages()
        .send(userId="me", body=create_message)
        .execute()
    )
    print(f'Message Id: {send_message["id"]}')
  except HttpError as error:
    print(f"An error occurred: {error}")
    send_message = None
  return send_message


if __name__ == "__main__":
  email_from = '210608thSt@gmail.com'
  email_to = 'ken@vennerllc.com'
  email_subject = 'kvgmailsendsimple.py - Test Message'
  email_body = 'This is a test run of this utility'
  scopes = None
  # file_token_json = '210608th.json'
  file_token_json = None
  file_credentials_json = None

  print('Refresh token - no email to send')
  gmail_refresh_token_take_no_action(email_from, scopes, file_token_json, file_credentials_json)

  
  print('Test email to filename conversion:  ', email_from)
  print(convert_email_to_filename(email_from))
  print('Test generation of email send through:  ', email_from)
  print('   Sent to...........................:  ', email_to)
  try:
    gmail_send_simple_message(email_from, email_to, email_subject, email_body, scopes, file_token_json, file_credentials_json)
  except Exception as err:
    print('Err: ', err)
    print('this failed - you must delete the input json and reauthenticate the application')
# eof
