import win32com.client
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
inbox = namespace.GetDefaultFolder(6)

cutoff = datetime.now() - timedelta(days=90)
messages = inbox.Items
date_filter = f"[ReceivedTime] >= '{cutoff.strftime('%m/%d/%Y %H:%M %p')}'"
filtered = messages.Restrict(date_filter)

for msg in filtered:
    try:
        sender = msg.SenderEmailAddress
        if sender and '@' not in str(sender):
            try:
                sender = msg.Sender.GetExchangeUser().PrimarySmtpAddress if msg.Sender else sender
            except:
                pass
        if sender and 'humanafterall.co.uk' in str(sender).lower():
            if 'large client contract' in str(msg.Subject).lower():
                print(f'Subject: {msg.Subject}')
                print(f'Date: {msg.ReceivedTime}')
                print(f'Attachments: {msg.Attachments.Count}')
                if msg.Attachments.Count > 0:
                    print('Attachment names:')
                    for att in msg.Attachments:
                        print(f'  - {att.FileName}')
                else:
                    print('No attachments')
    except Exception as e:
        continue
