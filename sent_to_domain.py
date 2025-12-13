import win32com.client
from datetime import datetime, timedelta
import sys

def get_sent_emails_to_domain(domain, days=30):
    """
    Read all emails sent TO a specific domain in the past month.
    """
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Get the Sent Items folder (5 = olFolderSentMail)
        sent_folder = namespace.GetDefaultFolder(5)
        
        # Calculate date cutoff
        cutoff_date = datetime.now() - timedelta(days=days)
        
        print(f"Searching for emails SENT TO domain: {domain}")
        print(f"Date range: {cutoff_date.strftime('%Y-%m-%d')} to now")
        print("="*100)
        
        # Get messages and sort by date (newest first)
        messages = sent_folder.Items
        messages.Sort("[SentOn]", True)
        
        # Collect matching emails
        matching_emails = []
        
        for msg in messages:
            try:
                # Check if message is within date range
                sent_time = msg.SentOn
                # Convert to naive datetime for comparison
                if hasattr(sent_time, 'replace'):
                    sent_naive = sent_time.replace(tzinfo=None)
                else:
                    sent_naive = sent_time
                
                if sent_naive < cutoff_date:
                    break  # Stop since messages are sorted by date
                
                # Collect all recipients (To, CC, BCC) that match the domain
                matched_recipients = []
                
                # Check all recipients - Recipients collection includes To, CC, and BCC
                for recipient in msg.Recipients:
                    recip_email = recipient.Address
                    recip_name = recipient.Name
                    
                    # Get recipient type: 1=To, 2=CC, 3=BCC
                    recip_type = recipient.Type
                    type_label = {1: "To", 2: "CC", 3: "BCC"}.get(recip_type, "To")
                    
                    # Try to get SMTP address for Exchange users
                    if recip_email and '@' not in str(recip_email):
                        try:
                            address_entry = recipient.AddressEntry
                            if address_entry:
                                exchange_user = address_entry.GetExchangeUser()
                                if exchange_user and exchange_user.PrimarySmtpAddress:
                                    recip_email = exchange_user.PrimarySmtpAddress
                        except:
                            pass
                    
                    # Check if recipient is from the specified domain
                    if recip_email and domain.lower() in str(recip_email).lower():
                        matched_recipients.append({
                            'name': recip_name,
                            'email': recip_email,
                            'type': type_label
                        })
                
                # If any recipients matched, add the email
                if matched_recipients:
                    matching_emails.append({
                        'recipients': matched_recipients,
                        'subject': msg.Subject,
                        'sent': msg.SentOn,
                        'body': msg.Body[:2000] if msg.Body else "(no body)"
                    })
            except Exception as e:
                continue
        
        # Sort by date (most recent first)
        matching_emails.sort(key=lambda x: x['sent'], reverse=True)
        
        # Print results
        print(f"\nFound {len(matching_emails)} emails sent to {domain}\n")
        print("="*100)
        
        for i, email in enumerate(matching_emails, 1):
            print(f"\n--- EMAIL {i} ---")
            for recip in email['recipients']:
                print(f"{recip['type']}: {recip['name']} <{recip['email']}>")
            print(f"Date: {email['sent']}")
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
        print(f"Total: {len(matching_emails)} emails sent to {domain}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        domain = input("Enter domain to search (e.g., gmail.com): ").strip()
    else:
        domain = sys.argv[1]
    
    get_sent_emails_to_domain(domain, days=30)
