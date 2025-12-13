import win32com.client
from datetime import datetime, timedelta
import sys

def get_emails_from_domain(domain, days=30):
    """
    Read all emails from a specific domain in the past month.
    """
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Get the Inbox folder
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        
        # Calculate date cutoff
        cutoff_date = datetime.now() - timedelta(days=days)
        cutoff_str = cutoff_date.strftime("%m/%d/%Y %H:%M %p")
        
        print(f"Searching for emails from domain: {domain}")
        print(f"Date range: {cutoff_date.strftime('%Y-%m-%d')} to now")
        print("="*100)
        
        # Get filtered messages
        messages = inbox.Items
        date_filter = f"[ReceivedTime] >= '{cutoff_str}'"
        filtered_messages = messages.Restrict(date_filter)
        
        # Collect matching emails
        matching_emails = []
        
        for msg in filtered_messages:
            try:
                sender_email = msg.SenderEmailAddress
                
                # Try to get SMTP address for Exchange users
                if sender_email and '@' not in str(sender_email):
                    try:
                        sender_obj = msg.Sender
                        if sender_obj:
                            exchange_user = sender_obj.GetExchangeUser()
                            if exchange_user and exchange_user.PrimarySmtpAddress:
                                sender_email = exchange_user.PrimarySmtpAddress
                    except:
                        pass
                
                # Check if email is from the specified domain
                if sender_email and domain.lower() in str(sender_email).lower():
                    matching_emails.append({
                        'sender': msg.SenderName,
                        'email': sender_email,
                        'subject': msg.Subject,
                        'received': msg.ReceivedTime,
                        'body': msg.Body[:2000] if msg.Body else "(no body)"  # Limit body length
                    })
            except Exception as e:
                continue
        
        # Sort by date (most recent first)
        matching_emails.sort(key=lambda x: x['received'], reverse=True)
        
        # Print results
        print(f"\nFound {len(matching_emails)} emails from {domain}\n")
        print("="*100)
        
        for i, email in enumerate(matching_emails, 1):
            print(f"\n--- EMAIL {i} ---")
            print(f"From: {email['sender']} <{email['email']}>")
            print(f"Date: {email['received']}")
            print(f"Subject: {email['subject']}")
            print(f"\nBody Preview:")
            print("-"*50)
            # Clean up body text
            body = email['body'].strip()
            # Remove excessive whitespace
            body = '\n'.join(line.strip() for line in body.split('\n') if line.strip())
            print(body[:1500] if len(body) > 1500 else body)
            print("-"*50)
        
        print(f"\n{'='*100}")
        print(f"Total: {len(matching_emails)} emails from {domain}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        domain = input("Enter domain to search (e.g., gmail.com): ").strip()
    else:
        domain = sys.argv[1]
    
    get_emails_from_domain(domain, days=30)
