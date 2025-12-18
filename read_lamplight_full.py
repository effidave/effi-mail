import win32com.client
from datetime import datetime, timedelta

def get_lamplight_emails_full():
    """Get full emails from lamplightdb.co.uk"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        
        cutoff_date = datetime.now() - timedelta(days=30)
        cutoff_str = cutoff_date.strftime("%m/%d/%Y %H:%M %p")
        
        messages = inbox.Items
        date_filter = f"[ReceivedTime] >= '{cutoff_str}'"
        filtered_messages = messages.Restrict(date_filter)
        
        matching_emails = []
        
        for msg in filtered_messages:
            try:
                sender_email = msg.SenderEmailAddress
                
                if sender_email and '@' not in str(sender_email):
                    try:
                        sender_obj = msg.Sender
                        if sender_obj:
                            exchange_user = sender_obj.GetExchangeUser()
                            if exchange_user and exchange_user.PrimarySmtpAddress:
                                sender_email = exchange_user.PrimarySmtpAddress
                    except:
                        pass
                
                if sender_email and 'lamplightdb.co.uk' in str(sender_email).lower():
                    matching_emails.append({
                        'sender': msg.SenderName,
                        'email': sender_email,
                        'subject': msg.Subject,
                        'received': msg.ReceivedTime,
                        'body': msg.Body if msg.Body else "(no body)"
                    })
            except Exception as e:
                continue
        
        matching_emails.sort(key=lambda x: x['received'], reverse=True)
        
        for i, email in enumerate(matching_emails, 1):
            print(f"\n{'='*100}")
            print(f"EMAIL {i}")
            print(f"{'='*100}")
            print(f"From: {email['sender']} <{email['email']}>")
            print(f"Date: {email['received']}")
            print(f"Subject: {email['subject']}")
            print(f"\nFull Body:")
            print("-"*100)
            print(email['body'])
            print("-"*100)
        
        print(f"\n\nTotal: {len(matching_emails)} emails from lamplightdb.co.uk")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    get_lamplight_emails_full()
