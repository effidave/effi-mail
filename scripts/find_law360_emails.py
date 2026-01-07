"""
Script to find all emails from law360.com across Outlook folders.
"""

import win32com.client


def find_law360_emails():
    """Find all emails from law360.com domain."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # Folders to search
    folders_to_search = [
        (6, "Inbox"),
    ]
    
    all_emails = []
    
    for folder_id, folder_name in folders_to_search:
        try:
            folder = namespace.GetDefaultFolder(folder_id)
            print(f"Scanning {folder_name} ({folder.Items.Count} items)...")
            
            for item in folder.Items:
                try:
                    sender_email = ""
                    if hasattr(item, "SenderEmailAddress"):
                        sender_email = item.SenderEmailAddress or ""
                    
                    if sender_email.lower().endswith("@law360.com"):
                        all_emails.append({
                            "folder": folder_name,
                            "subject": item.Subject,
                            "sender": sender_email,
                            "received": str(item.ReceivedTime),
                            "entry_id": item.EntryID
                        })
                except Exception as e:
                    continue
        except Exception as e:
            print(f"Could not access {folder_name}: {e}")
    
    print(f"\n{'='*70}")
    print(f"Found {len(all_emails)} emails from law360.com")
    print(f"{'='*70}\n")
    
    if not all_emails:
        print("No emails found from law360.com")
        return
    
    # Group by folder
    by_folder = {}
    for email in all_emails:
        folder = email["folder"]
        if folder not in by_folder:
            by_folder[folder] = []
        by_folder[folder].append(email)
    
    for folder, emails in by_folder.items():
        print(f"\n{folder} ({len(emails)} emails):")
        print("-" * 50)
        for i, email in enumerate(emails, 1):
            subject = email["subject"][:55] if email["subject"] else "(no subject)"
            date = email["received"][:10]
            print(f"  {i:3}. [{date}] {subject}")


if __name__ == "__main__":
    find_law360_emails()
