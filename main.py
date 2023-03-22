import zipfile

import pandas

import pandas as pd
import win32com.client
import os
import datetime
import win32com.client
import re

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
message = messages.GetFirst()
today_date = str(datetime.date.today())

for message in messages:
 try:
    current_sender = str(message.Sender).lower()
    current_subject = str(message.Subject).lower()
    message_date = str(message.senton.date())
    if re.search('coupa report: waw fsc',current_subject) != None and message_date == today_date:
      print(current_subject)
      print(current_sender)
      attachments = message.Attachments
      attachment = attachments.Item(1)
      attachment_name = str(attachment).lower()
      attachment.SaveASFile("C:\\Users\\skomuda\\Amadeus Workplace\\Testing" + '\\' + attachment_name)
    else:
        pass
    message = messages.GetNext()
 except:
    message = messages.GetNext()

dir_path = r'C:\\Users\\skomuda\\Amadeus Workplace\\Testing\\'

res = []
for file in os.listdir(dir_path):
    if file.endswith('.zip'):
        res.append(file)
print(res)

with zipfile.ZipFile("C:\\Users\\skomuda\\Amadeus Workplace\\Testing\\" + res[0], 'r') as zip_ref:
    zip_ref.extractall("C:\\Users\\skomuda\\Amadeus Workplace\\Testing")

excelFIles = []

for file in os.listdir(dir_path):
        if file.endswith('.xlsx'):
            excelFIles.append(file)
print(excelFIles)
df = pd.read_excel("C:\\Users\\skomuda\\Amadeus Workplace\\Testing\\" + excelFIles[0], 'sheet1')
df = pd.pivot_table(df, values=['Inbox Status'],
                                index=['Inbox Name'],
                                aggfunc='count',
                                fill_value=0)
writer = pd.ExcelWriter('C:\\Users\\skomuda\\Amadeus Workplace\\Testing\\First.xlsx')
df.to_excel(writer, sheet_name='PivotTable')

writer.save()