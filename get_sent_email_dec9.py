import win32com.client
from datetime import datetime, timedelta

def get_sent_email_full(domain, subject_filter):
    """Get full sent email to a domain with subject filter"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        sent_folder = namespace.GetDefaultFolder(5)  # Sent Items
        
        cutoff_date = datetime.now() - timedelta(days=30)
        
        messages = sent_folder.Items
        messages.Sort("[SentOn]", True)
        
        for msg in messages:
            try:
                sent_time = msg.SentOn
                if hasattr(sent_time, 'replace'):
                    sent_naive = sent_time.replace(tzinfo=None)
                else:
                    sent_naive = sent_time
                
                if sent_naive < cutoff_date:
                    break
                
                # Check if any recipient matches domain
                has_domain = False
                for recipient in msg.Recipients:
                    recip_email = recipient.Address
                    if recip_email and domain.lower() in str(recip_email).lower():
                        has_domain = True
                        break
                
                if has_domain and subject_filter.lower() in str(msg.Subject).lower():
                    print(f"Date: {msg.SentOn}")
                    print(f"Subject: {msg.Subject}")
                    print(f"\nFull Body:")
                    print("="*100)
                    print(msg.Body if msg.Body else "(no body)")
                    print("="*100)
                    return
                    
            except Exception as e:
                continue
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    get_sent_email_full('lamplightdb.co.uk', 'T&C review')
