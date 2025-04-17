#%%
import os
import time
import datetime as dt
import win32com.client

def email(message: str):

    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'helpdesk@ecology.co.uk'
    mail.Subject = 'BCP'
    mail.Body = f"""
    Hi,
        {message}
    """
    mail.Send()

def main():
    monthvals = {
        "Jan": 1
        , "Feb": 2
        , "Mar": 3
        , "Apr": 4
        , "May": 5
        , "Jun": 6
        , "Jul": 7
        , "Aug": 8
        , "Sep": 9
        , "Oct": 10
        , "Nov": 11
        , "Dec": 12
    }
    today = dt.date.today()
    filetime = os.path.getmtime("K:/IT/Restricted/Reporting/BCP/save_finl.xlsx")
    filetime = time.ctime(filetime)
    filemonth = filetime[4:7]
    filemonth = monthvals[filemonth]
    fileday = int(filetime[8:10])
    fileyear = int(filetime[-4:])
    filedate = dt.date(fileyear, filemonth, fileday)
    if filedate == today:
        email("The BCP File is fine, no action required. Close the ticket")
    else:
        email("BCP isn't fine, speak to Ian")

main() 


