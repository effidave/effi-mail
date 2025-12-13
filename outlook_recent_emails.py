import win32com.client
import json
from datetime import datetime, timedelta
from collections import defaultdict
import os

CATEGORIES_FILE = "domain_categories.json"
CATEGORY_OPTIONS = ["Client", "Internal", "Marketing", "Personal"]

def load_categories():
    """Load domain categories from JSON file."""
    if os.path.exists(CATEGORIES_FILE):
        with open(CATEGORIES_FILE, "r") as f:
            return json.load(f)
    return {}

def save_categories(categories):
    """Save domain categories to JSON file."""
    with open(CATEGORIES_FILE, "w") as f:
        json.dump(categories, f, indent=2)

def get_sender_email(message):
    """Extract sender email, handling Exchange addresses."""
    try:
        sender = message.Sender
        if sender is not None:
            # Try to get SMTP address from Exchange user
            if sender.AddressEntryUserType == 0:  # Exchange user
                exch_user = sender.GetExchangeUser()
                if exch_user:
                    return exch_user.PrimarySmtpAddress
            # Try PropertyAccessor for SMTP address
            try:
                return message.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                )
            except:
                pass
        # Fallback to SenderEmailAddress
        return message.SenderEmailAddress
    except:
        return message.SenderEmailAddress

def extract_domain(email):
    """Extract domain from email address."""
    if email and "@" in email:
        return email.split("@")[-1].lower()
    return "unknown"

def get_recent_emails(days=7):
    """Fetch emails from the past N days using Restrict filter."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    
    # Calculate date filter
    date_cutoff = datetime.now() - timedelta(days=days)
    date_str = date_cutoff.strftime("%m/%d/%Y %H:%M %p")
    filter_str = f"[ReceivedTime] >= '{date_str}'"
    
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    filtered = messages.Restrict(filter_str)
    
    emails = []
    for message in filtered:
        try:
            sender_email = get_sender_email(message)
            domain = extract_domain(sender_email)
            received = message.ReceivedTime
            # Remove timezone info for comparison
            if hasattr(received, 'replace'):
                received = received.replace(tzinfo=None)
            # Skip emails with "Unfocused" category
            categories_str = message.Categories if message.Categories else ""
            if "Unfocused" in categories_str:
                continue
            
            emails.append({
                "subject": message.Subject,
                "sender": message.SenderName,
                "sender_email": sender_email,
                "domain": domain,
                "received": received
            })
        except Exception as e:
            continue
    
    return emails

def group_by_domain(emails):
    """Group emails by sender domain."""
    grouped = defaultdict(list)
    for email in emails:
        grouped[email["domain"]].append(email)
    return grouped

def display_inbox_mode(emails, categories):
    """Display emails grouped by domain with categories."""
    grouped = group_by_domain(emails)
    
    # Sort domains by most recent email (newest first)
    def sort_key(domain):
        domain_emails = grouped[domain]
        most_recent = max(e["received"] for e in domain_emails if e["received"])
        return most_recent
    
    sorted_domains = sorted(grouped.keys(), key=sort_key, reverse=False)
    
    print(f"\n{'='*70}")
    print(f"INBOX - Emails from the past 7 days ({len(emails)} total)")
    print(f"{'='*70}")
    
    for domain in sorted_domains:
        domain_emails = grouped[domain]
        category = categories.get(domain, "Uncategorized")
        print(f"\n[{category}] {domain} ({len(domain_emails)} emails)")
        print("-" * 50)
        for email in domain_emails[:5]:  # Show max 5 per domain
            received_str = email["received"].strftime("%d/%m %H:%M") if email["received"] else ""
            subject = email["subject"][:50] + "..." if len(email["subject"]) > 50 else email["subject"]
            print(f"  {received_str} | {email['sender'][:20]:<20} | {subject}")
        if len(domain_emails) > 5:
            print(f"  ... and {len(domain_emails) - 5} more")

def categorization_mode(emails, categories):
    """Interactive mode to categorize domains."""
    grouped = group_by_domain(emails)
    uncategorized = [d for d in grouped.keys() if d not in categories]
    
    if not uncategorized:
        print("\nAll domains are already categorized!")
        return categories
    
    print(f"\n{'='*70}")
    print(f"CATEGORIZATION MODE - {len(uncategorized)} uncategorized domains")
    print(f"{'='*70}")
    print("\nCategories: " + ", ".join(f"{i+1}={c}" for i, c in enumerate(CATEGORY_OPTIONS)))
    print("Enter number to categorize, 's' to skip, 'q' to quit\n")
    
    for domain in uncategorized:
        domain_emails = grouped[domain]
        print(f"\nDomain: {domain} ({len(domain_emails)} emails)")
        print(f"  Sample senders: {', '.join(set(e['sender'] for e in domain_emails[:3]))}")
        print(f"  Sample subject: {domain_emails[0]['subject'][:60]}")
        
        while True:
            choice = input(f"  Category [1-{len(CATEGORY_OPTIONS)}/s/q]: ").strip().lower()
            if choice == 'q':
                save_categories(categories)
                return categories
            if choice == 's':
                break
            if choice.isdigit() and 1 <= int(choice) <= len(CATEGORY_OPTIONS):
                categories[domain] = CATEGORY_OPTIONS[int(choice) - 1]
                print(f"  -> Categorized as {categories[domain]}")
                save_categories(categories)
                break
            print("  Invalid choice. Try again.")
    
    return categories

def main_menu():
    """Main menu for the application."""
    categories = load_categories()
    emails = None
    
    while True:
        print(f"\n{'='*70}")
        print("OUTLOOK EMAIL MANAGER")
        print(f"{'='*70}")
        print("1. View Inbox (grouped by domain)")
        print("2. Categorize Domains")
        print("3. Refresh Emails")
        print("4. View Categories")
        print("5. Exit")
        
        choice = input("\nSelect option: ").strip()
        
        if choice == "1":
            if emails is None:
                print("\nFetching emails...")
                emails = get_recent_emails(days=7)
            display_inbox_mode(emails, categories)
        
        elif choice == "2":
            if emails is None:
                print("\nFetching emails...")
                emails = get_recent_emails(days=7)
            categories = categorization_mode(emails, categories)
        
        elif choice == "3":
            print("\nRefreshing emails...")
            emails = get_recent_emails(days=7)
            print(f"Loaded {len(emails)} emails from the past 7 days.")
        
        elif choice == "4":
            print("\nCurrent Domain Categories:")
            print("-" * 40)
            for cat in CATEGORY_OPTIONS:
                domains = [d for d, c in categories.items() if c == cat]
                if domains:
                    print(f"\n{cat}:")
                    for d in sorted(domains):
                        print(f"  - {d}")
        
        elif choice == "5":
            print("\nGoodbye!")
            break
        
        else:
            print("\nInvalid option. Please try again.")

if __name__ == "__main__":
    main_menu()
