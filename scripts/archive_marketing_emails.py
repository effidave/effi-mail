"""
Script to move all Marketing emails from Inbox to Archive folder.

Finds all domains categorized as Marketing and moves their emails
from Inbox to Archive.
"""

import sys
import win32com.client
from pathlib import Path

# Add parent directory to path to import domain_categories
sys.path.insert(0, str(Path(__file__).parent.parent))
import domain_categories


def move_marketing_emails_to_archive():
    """Move all Marketing category emails from Inbox to Archive."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # Get Inbox folder (olFolderInbox = 6)
    inbox = namespace.GetDefaultFolder(6)
    
    # Get Archive folder - try to find it in root folders
    archive = None
    for folder in namespace.Folders:
        for subfolder in folder.Folders:
            if subfolder.Name == "Archive":
                archive = subfolder
                break
        if archive:
            break
    
    if not archive:
        print("Error: Could not find Archive folder in Outlook")
        return
    
    # Get all Marketing domains
    marketing_domains = domain_categories.get_domains_by_category("Marketing")
    
    if not marketing_domains:
        print("No domains categorized as Marketing")
        return
    
    print(f"Found {len(marketing_domains)} Marketing domains:")
    for domain in sorted(marketing_domains):
        print(f"  - {domain}")
    print()
    
    print(f"Scanning {inbox.Items.Count} emails in Inbox...")
    
    # Collect emails to move
    emails_to_move = []
    marketing_domains_lower = set(d.lower() for d in marketing_domains)
    
    for item in inbox.Items:
        try:
            sender_email = ""
            if hasattr(item, "SenderEmailAddress"):
                sender_email = item.SenderEmailAddress or ""
            
            if "@" in sender_email:
                domain = sender_email.split("@")[1].lower()
                if domain in marketing_domains_lower:
                    emails_to_move.append({
                        "subject": item.Subject or "(no subject)",
                        "sender": sender_email,
                        "domain": domain,
                        "received": str(item.ReceivedTime),
                        "entry_id": item.EntryID
                    })
        except Exception as e:
            continue
    
    print(f"\nFound {len(emails_to_move)} Marketing emails to move")
    
    if not emails_to_move:
        print("No Marketing emails to move.")
        return
    
    # Group by domain for display
    by_domain = {}
    for email in emails_to_move:
        domain = email["domain"]
        if domain not in by_domain:
            by_domain[domain] = 0
        by_domain[domain] += 1
    
    print("\nEmails by domain:")
    for domain, count in sorted(by_domain.items(), key=lambda x: x[1], reverse=True):
        print(f"  {domain}: {count} emails")
    
    # Confirm move
    confirm = input(f"\nMove {len(emails_to_move)} emails to Archive? (yes/no): ")
    
    if confirm.lower() != "yes":
        print("Cancelled.")
        return
    
    # Move emails
    moved_count = 0
    failed_count = 0
    
    print("\nMoving emails...")
    for email_info in emails_to_move:
        try:
            item = namespace.GetItemFromID(email_info["entry_id"])
            item.Move(archive)
            moved_count += 1
            if moved_count % 10 == 0:
                print(f"  Moved {moved_count}/{len(emails_to_move)}...")
        except Exception as e:
            failed_count += 1
            print(f"  Failed to move '{email_info['subject'][:40]}': {e}")
    
    print(f"\n{'='*70}")
    print(f"Complete!")
    print(f"  Moved: {moved_count}")
    print(f"  Failed: {failed_count}")
    print(f"{'='*70}")


if __name__ == "__main__":
    try:
        move_marketing_emails_to_archive()
    except KeyboardInterrupt:
        print("\n\nInterrupted by user")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
