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
    for i in range(len(mail_body)):
        mail_body_df= mail_body[i]
        header= mail_body_df.iloc[0]
        mail_body_df= mail_body_df[1:]
        mail_body_df.columns= header
        mail_body_df.to_excel(output_excel, index= False)
        print('Done!')
    return True
  
main_function()

