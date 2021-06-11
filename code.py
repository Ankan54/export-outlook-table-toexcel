'''
This code snippet reads outlook inbox mail and extract the tabular data
to export them in an excel file

Author: Ankan Bera
'''

import win32com.client as win32
import pandas as pd

def main_function():
  
    output_excel= "File path"  # replace with full file path having .xlsx extension
    try:
        outlook= win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except:
        print("Error accessing Outlook application")
        return False

    mails = outlook.GetDefaultFolder(6).Items
    mail= mails.GetLast() # latest mail

    mail_body= pd.read_html(mail.HTMLBody) #create dataframe

    column_header = mail_body.iloc[0]  # to make the first row of dataframe as header
    mail_body = mail_body[1:]
    mail_body.columns= column_header
    mail_body.to_excel(output_excel, index= False)
    print('Done!')
    return True
  
main_function()

