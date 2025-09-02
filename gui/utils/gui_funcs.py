import tkinter as tk
import datetime
import pandas as pd

def CheckWhiteList(framelist, app):
    #CheckWhiteList takes arguments
    #framelist, which must be a list of mainEntryFrame.subEntryFrame objects,
    #and app, which is the app with all the stuff that needs to be used
    
    #Import the whitelist
    whitelist = app.Whitelist

    #initialize the output list
    outputlist = []

    #Iterate over the framelist
    for frameobj in framelist:
        #check if the entry window has a good name
        if not frameobj.sniped_ent.get() in whitelist and not frameobj.sniped_ent.get() == '':
            #if it's a bad name, set the background to error and output a false
            frameobj.sniped_ent.configure(style = 'entry_error.TEntry')
            outputlist.append(False)
        else:
            #otherwise, set the background to normal
            frameobj.sniped_ent.configure(style = 'entry_normal.TEntry')

        #repeat for the second entry. both are considered seperately so multiple bad values can be caught
        #in 1 check instead of one bad value per check
        if not frameobj.sniper_ent.get() in whitelist and not frameobj.sniper_ent.get() == '':
            frameobj.sniper_ent.configure(style = 'entry_error.TEntry')
            outputlist.append(False)
        else:
            frameobj.sniper_ent.configure(style = 'entry_normal.TEntry')

    #Return False if there are any false values in the outputlist
    return (all(outputlist))


def PrepareData(lst,yesterday):
    #Prepare Data takes a list of (sniper,sniped) and outputs a dataframe with columns (sniper, sniped, date)
    
    noblanks = []
    for pair1 in lst:
        if any(n == '' for n in pair1):
            pass
        else:
            noblanks.append(pair1)
    #check to see if this is a yesterday entry
    if yesterday:
        date = datetime.datetime.now() - datetime.timedelta(days = 1)
    else:
        date = datetime.datetime.now()

    #Extract the tuple into 2 points and add the date and a blank column for notes
    out = [(pair2[0],pair2[1],date, '') for pair2 in noblanks]

    return pd.DataFrame(out, dtype = 'string', columns = ['Sniper','Sniped','Date', 'Notes'])

def GenericError(err):
    tk.messagebox.showerror('Generic Exception',f'The script has encountered an error\n\n{err}')

def oauth_flow(base_path):
    import os
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    import sys

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
              "https://www.googleapis.com/auth/drive.metadata.readonly"
              ]
    creds = None

  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
    if os.path.exists(base_path + r"\gui\config\token.json"):
        creds = Credentials.from_authorized_user_file(base_path + r"\gui\config\token.json", SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        try:    
            if creds and creds.expired and creds.refresh_token:
                print('hit')
                try:
                    creds.refresh(Request())
                except:
                    #Run thke login flow if the refresh fails
                    flow = InstalledAppFlow.from_client_secrets_file(
                    base_path + r"\gui\config\client_secret.json", SCOPES
                    )
                    creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    base_path + r"\gui\config\client_secret.json", SCOPES
                )
                creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
            with open(base_path + r"\gui\config\token.json", "w") as token:
                token.write(creds.to_json())
        except Exception as e:
            tk.messagebox.showerror('Login Error',f'The script has encountered an error during login\n\n{e}')
            sys.exit()
    return creds