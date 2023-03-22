import datetime

import win32com.client
import re

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
message = messages.GetFirst()
today_date = str(datetime.date.today())
while True:
  try:
    current_sender = str(message.Sender).lower()
    current_subject = str(message.Subject).lower()
    messageDate = str(message.senton.date())
    print(messageDate)
    print('yy')
    print(today_date)
    # find the email from a specific sender with a specific subject
    # condition
    if re.search('x2',current_subject) != None and   \
            messageDate == today_date:
      print(current_subject) # verify the subject
      print(current_sender)  # verify the sender
      attachments = message.Attachments
      attachment = attachments.Item(1)
      attachment_name = str(attachment).lower()
      attachment.SaveASFile("C:\\Users\\skomuda\\Amadeus Workplace\\Testing" + "\\" + attachment_name)
    else:
      pass
    message = messages.GetNext()
  except:
    message = messages.GetNext()
exit

