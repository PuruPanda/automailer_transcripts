# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import numpy as np
import pandas as pd

import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage # Sheets API imports

from getpass import getpass

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None
    
    
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'C:/Users/ppandey/.credentials/client_secret.json'
APPLICATION_NAME = 'Python Sheets'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

credentials = get_credentials()
http = credentials.authorize(httplib2.Http())
discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                'version=v4')
service = discovery.build('sheets', 'v4', http=http,
                          discoveryServiceUrl=discoveryUrl)

spreadsheetId = '1GAYCS20L4Y_AwhL1kqHhQIJeW_G8YQY5Ee_YMrl5zdc'
rangeName = '2017-2018!A:Q'
result = service.spreadsheets().values().get(
    spreadsheetId=spreadsheetId, range=rangeName).execute()
values = result.get('values', [])

maildata = pd.DataFrame(values)
maildata.columns = maildata.loc[0]

maildata.fillna(value = np.nan, inplace = True)

maildata = maildata[(maildata['Ready'] != 'Ready') & (maildata['Ready'] != 'mailed')]

FROMADDR = 'ppandey@smuhsd.org'
msg = MIMEMultipart()
    
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(FROMADDR, getpass('Email password: '))

print('Emails will be sent to the following people:\n')
for recipient in maildata['Full Name']:
    print(recipient)
    
emailconfirm = input('Send email? (Y/N): ')

if emailconfirm == 'Y':
    for recipient in maildata.index:
        msg = MIMEMultipart()
        TOADDR = maildata['Email Address'][recipient]
        
        if maildata['Pick up or Mail'][recipient] == 'Pick up at Mills.':
            filltext = 'and is ready for pick up at the front desk of the Administration Office. Please see Tammy McGee to pick up your transcript and have your payment and proof of ID available.'
        else:
            filltext = 'and sent to the requested address. Please remit your payment within seven business days.'
            
        msg['From'] = FROMADDR
        msg['To'] = TOADDR
        msg['Subject'] = 'Transcript Request'
        
        body = '''Dear %s,\n
Your transcript request has been processed %s\n
Thank you,\n
Puru Pandey
Student Data Analyst
Mills High School
(650) 558-2519''' % (maildata['Full Name'][recipient], filltext)
    
        msg.attach(MIMEText(body, 'plain'))
        
        emailtext = msg.as_string()
        server.sendmail(FROMADDR, TOADDR, emailtext)

else:
    pass

server.quit()