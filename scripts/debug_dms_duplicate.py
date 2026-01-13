"""Debug script for DMS duplicate detection."""
from outlook_client import OutlookClient

c = OutlookClient()
c._ensure_connection()

# Get the source email
email_id = None
for e in c.get_emails(days=1):
    email_id = e.id
    break

if not email_id:
    print("No emails found")
    exit()

message = c._namespace.GetItemFromID(email_id)
print("Source email:")
print(f"  Subject: {message.Subject}")
print(f"  ReceivedTime: {message.ReceivedTime} (type: {type(message.ReceivedTime)})")
print(f"  Sender: {message.SenderEmailAddress}")

# Get the filed email from DMS
path = f"{c.DMS_ROOT_FOLDER}\\Evotra Ltd\\7IM MSA Queries - Evotra Ltd (31857)\\{c.DMS_EMAILS_FOLDER}"
folder = c._get_folder_by_path(path)
print(f"\nDMS folder has {folder.Items.Count} items")

for item in folder.Items:
    print("\nDMS email:")
    print(f"  Subject: {item.Subject}")
    print(f"  ReceivedTime: {item.ReceivedTime} (type: {type(item.ReceivedTime)})")
    sender = getattr(item, "SenderEmailAddress", "")
    print(f"  Sender: {sender}")
    
    # Test comparison
    print("\nComparison:")
    print(f"  Subject match: {message.Subject == item.Subject}")
    print(f"  ReceivedTime match: {message.ReceivedTime == item.ReceivedTime}")
    if sender:
        print(f"  Sender match: {message.SenderEmailAddress.lower() == sender.lower()}")
    break
