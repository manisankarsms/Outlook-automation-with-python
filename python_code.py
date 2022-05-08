import pyautogui
import win32com.client as win32
import sys
import xlrd
import pandas as pd
import time

pyautogui.alert("Hey, Iam Triggered")

n = len(sys.argv)
mailBody = sys.argv[1]
for i in range(2, n):
    mailBody =mailBody +" "+ sys.argv[i]

df = pd.read_excel('Source.xlsx')
for i in range(len(df)):
  if df.loc[i, "JOB"]==mailBody:
      ownerMail = df.loc[i, "OWNER"]
      pyautogui.alert("Owner Found:"+ownerMail)
      break

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = ownerMail
mail.Subject = 'Python Mail Successful'
mail.Body = 'Arg...'+mailBody+'... end'
mail.HTMLBody = '<h1>'+mailBody+'</h1>' #this field is optional

# To attach a file to the email (optional):
#attachment  = "Path to the attachment"
#mail.Attachments.Add(attachment)
#pyautogui.alert(df.loc[i, "OWNER"])
mail.Send()
pyautogui.alert("Mail Sent")

