from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


import base64
import email
from apiclient import errors
import pdb

from html.parser import *


from plus import *
from msg import *

import xlrd
 
from xlwt import Workbook, Formula
 
path = r"../contact.csv"



# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
            # print(creds.token)
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

    
    service = build('gmail', 'v1', credentials=creds)
    print(creds.token)
    #Create EXCEL
    classeur = Workbook()
    feuille = classeur.add_sheet("OCB", cell_overwrite_ok=True)
    feuille.write(0, 0, "Passenger")
    feuille.write(0, 1, "Phone")
    feuille.write(0, 2, "From")
    feuille.write(0, 3, "To")
    feuille.write(0, 4, "Date")



    l = 1
    # Call the Gmail API
    message = ListMessagesMatchingQuery(service,"ENTER_YOUR_EMAIL","subject:Ride assignment")
    for m in message:
        if(l!=2025):
            idMesssage = m['id']
            print(">_ "+idMesssage)
            ctx = GetMimeMessage(service,"ENTER_YOUR_EMAIL",idMesssage)
            getInfos(ctx,feuille,l)
            l = l +1
            classeur.save(path)
        else:
            print("ERREUR 2025")


    





if __name__ == '__main__':
    main()