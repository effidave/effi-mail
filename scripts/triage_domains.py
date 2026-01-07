"""
Interactive CLI tool to triage domain categories.

Shows uncategorized domains from Inbox with email counts and samples,
allows interactive categorization into Client/Internal/Marketing/Personal/Spam.
"""

import sys
import win32com.client
from collections import defaultdict
from pathlib import Path

# Add parent directory to path to import domain_categories
sys.path.insert(0, str(Path(__file__).parent.parent))
import domain_categories


def get_inbox_domains_with_samples():
    """Get all domains from Inbox with counts and sample emails."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    
    # Track domain info: count and sample emails
    domain_info = defaultdict(lambda: {"count": 0, "samples": []})
    
    print(f"Scanning {inbox.Items.Count} emails in Inbox...", end="", flush=True)
    
    for item in inbox.Items:
        try:
            sender_email = ""
            if hasattr(item, "SenderEmailAddress"):
                sender_email = item.SenderEmailAddress or ""
            
            if "@" in sender_email:
                domain = sender_email.split("@")[1].lower()
                domain_info[domain]["count"] += 1
                
                # Keep up to 3 sample subjects
                if len(domain_info[domain]["samples"]) < 3:
                    subject = item.Subject or "(no subject)"
                    domain_info[domain]["samples"].append(subject[:60])
        except:
            continue
    
    print(" done!\n")
    return domain_info


def triage_domains():
    """Interactive domain categorization."""
    print("="*70)
    print("Domain Triage Tool")
    print("="*70)
    print()
    
    # Get inbox domain info
    domain_info = get_inbox_domains_with_samples()
    
    # Get uncategorized domains
    all_domains = list(domain_info.keys())
    uncategorized = domain_categories.get_uncategorized_domains(all_domains)
    
    if not uncategorized:
        print("✓ All domains are already categorized!")
        print(f"\nTotal domains in inbox: {len(all_domains)}")
        return
    
    # Sort by email count (most emails first)
    uncategorized_sorted = sorted(
        uncategorized,
        key=lambda d: domain_info[d]["count"],
        reverse=True
    )
    
    print(f"Found {len(uncategorized_sorted)} uncategorized domains")
    print(f"(out of {len(all_domains)} total domains in inbox)\n")
    
    # Triage each domain
    categorized_count = 0
    skipped_count = 0
    
    for i, domain in enumerate(uncategorized_sorted, 1):
        info = domain_info[domain]
        
        print(f"\n{'='*70}")
        print(f"[{i}/{len(uncategorized_sorted)}] Domain: {domain}")
        print(f"{'='*70}")
        print(f"Email count: {info['count']}")
        print(f"\nSample subjects:")
        for j, subject in enumerate(info['samples'], 1):
            print(f"  {j}. {subject}")
        
        print(f"\nCategories:")
        print(f"  [c] Client")
        print(f"  [i] Internal")
        print(f"  [m] Marketing")
        print(f"  [p] Personal")
        print(f"  [s] Spam")
        print(f"  [skip] Skip this domain")
        print(f"  [q] Quit")
        
        while True:
            choice = input(f"\nCategory: ").strip().lower()
            
            if choice == "q":
                print(f"\nStopping. Categorized {categorized_count}, skipped {skipped_count}")
                return
            
            if choice == "skip":
                skipped_count += 1
                break
            
            category_map = {
                "c": "Client",
                "i": "Internal",
                "m": "Marketing",
                "p": "Personal",
                "s": "Spam"
            }
            
            if choice in category_map:
                category = category_map[choice]
                if domain_categories.set_domain_category(domain, category):
                    print(f"✓ Set {domain} → {category}")
                    categorized_count += 1
                    break
                else:
                    print(f"✗ Failed to set category")
            else:
                print("Invalid choice. Use c/i/m/p/s/skip/q")
    
    print(f"\n{'='*70}")
    print(f"Triage complete!")
    print(f"Categorized: {categorized_count}")
    print(f"Skipped: {skipped_count}")
    print(f"{'='*70}")


if __name__ == "__main__":
    try:
        triage_domains()
    except KeyboardInterrupt:
        print("\n\nInterrupted by user")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
