"""
Script to delete all inbox emails from law360.com

This script connects to Outlook via COM and permanently deletes
all emails from the law360.com domain in the Inbox folder.
"""

import win32com.client


def delete_law360_emails():
    """Delete all inbox emails from law360.com domain."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # Get Inbox folder (olFolderInbox = 6)
    inbox = namespace.GetDefaultFolder(6)
    
    print(f"Scanning {inbox.Items.Count} emails in Inbox...")
    
    # Collect emails to delete (iterate backwards to avoid index issues)
    emails_to_delete = []
    
    for item in inbox.Items:
        try:
            sender_email = ""
            if hasattr(item, "SenderEmailAddress"):
                sender_email = item.SenderEmailAddress or ""
            
            # Check if from law360.com domain
            if sender_email.lower().endswith("@law360.com"):
                emails_to_delete.append({
                    "subject": item.Subject,
                    "sender": sender_email,
                    "received": str(item.ReceivedTime),
                    "entry_id": item.EntryID
                })
        except Exception as e:
            print(f"Error reading email: {e}")
            continue
    
    print(f"\nFound {len(emails_to_delete)} emails from law360.com")
    
    if not emails_to_delete:
        print("No emails to delete.")
        return
    
    # Show what will be deleted
    print("\nEmails to be deleted:")
    for i, email in enumerate(emails_to_delete, 1):
        print(f"  {i}. {email['subject'][:60]}... ({email['received'][:10]})")
    
    # Confirm deletion
    confirm = input(f"\nDelete {len(emails_to_delete)} emails permanently? (yes/no): ")
    
    if confirm.lower() != "yes":
        print("Cancelled.")
        return
    
    # Delete emails
    deleted_count = 0
    for email_info in emails_to_delete:
        try:
            item = namespace.GetItemFromID(email_info["entry_id"])
            item.Delete()  # Moves to Deleted Items
            deleted_count += 1
        except Exception as e:
            print(f"Failed to delete '{email_info['subject'][:40]}': {e}")
    
    print(f"\nDeleted {deleted_count} emails (moved to Deleted Items)")
    
    # Optionally empty deleted items
    empty_trash = input("Also empty from Deleted Items (permanent)? (yes/no): ")
    if empty_trash.lower() == "yes":
        deleted_folder = namespace.GetDefaultFolder(3)  # olFolderDeletedItems = 3
        permanent_count = 0
        for item in list(deleted_folder.Items):
            try:
                sender = getattr(item, "SenderEmailAddress", "") or ""
                if sender.lower().endswith("@law360.com"):
                    item.Delete()  # Permanent delete from Deleted Items
                    permanent_count += 1
            except:
                pass
        print(f"Permanently deleted {permanent_count} emails")


if __name__ == "__main__":
    delete_law360_emails()
