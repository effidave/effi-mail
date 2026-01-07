"""
Script to report on all emails in Inbox grouped by domain.
Shows count per domain, ordered by largest number first.
"""

import win32com.client
from collections import defaultdict


def report_inbox_by_domain():
    """Report on all inbox emails grouped by domain."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # Get Inbox folder (olFolderInbox = 6)
    inbox = namespace.GetDefaultFolder(6)
    
    print(f"Scanning {inbox.Items.Count} emails in Inbox...\n")
    
    # Dictionary to count emails by domain
    domain_counts = defaultdict(int)
    no_domain_count = 0
    
    for item in inbox.Items:
        try:
            sender_email = ""
            if hasattr(item, "SenderEmailAddress"):
                sender_email = item.SenderEmailAddress or ""
            
            # Extract domain from email address
            if "@" in sender_email:
                domain = sender_email.split("@")[1].lower()
                domain_counts[domain] += 1
            else:
                no_domain_count += 1
        except Exception as e:
            no_domain_count += 1
            continue
    
    # Sort by count (descending)
    sorted_domains = sorted(domain_counts.items(), key=lambda x: x[1], reverse=True)
    
    print(f"{'='*70}")
    print(f"Inbox Email Report by Domain")
    print(f"{'='*70}\n")
    print(f"{'Domain':<40} {'Count':>10}")
    print(f"{'-'*40} {'-'*10}")
    
    total_emails = 0
    for domain, count in sorted_domains:
        print(f"{domain:<40} {count:>10}")
        total_emails += count
    
    if no_domain_count > 0:
        print(f"{'(no domain/unknown)':<40} {no_domain_count:>10}")
        total_emails += no_domain_count
    
    print(f"{'-'*40} {'-'*10}")
    print(f"{'TOTAL':<40} {total_emails:>10}")
    print(f"\nUnique domains: {len(sorted_domains)}")


if __name__ == "__main__":
    report_inbox_by_domain()
