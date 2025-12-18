import win32com.client
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
inbox = namespace.GetDefaultFolder(6)

cutoff = datetime.now() - timedelta(days=30)
messages = inbox.Items
date_filter = f"[ReceivedTime] >= '{cutoff.strftime('%m/%d/%Y %H:%M %p')}'"
filtered = messages.Restrict(date_filter)

emails = []
for msg in filtered:
    try:
        sender = msg.SenderEmailAddress
        if sender and '@' not in str(sender):
            try:
                sender = msg.Sender.GetExchangeUser().PrimarySmtpAddress if msg.Sender else sender
            except:
                pass
        if sender and 'lamplightdb.co.uk' in str(sender).lower():
            emails.append((msg.ReceivedTime, msg.Subject))
    except:
        continue

emails.sort(key=lambda x: x[0])
for date, subject in emails:
    print(f'{date.strftime("%Y-%m-%d %H:%M")} | {subject}')
