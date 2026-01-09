"""Test date formats for Outlook Restrict queries - debug version."""
import sys
sys.path.insert(0, '.')
from outlook_client import OutlookClient
from datetime import datetime, timedelta

client = OutlookClient()
client._ensure_connection()

# Navigate to a folder with known recent emails
inbox = client._namespace.GetDefaultFolder(6)
folder = None
for sub in inbox.Folders:
    if sub.Name == '~Zero':
        for sub2 in sub.Folders:
            if sub2.Name == 'PiP':
                folder = sub2
                break

if not folder:
    print("Folder not found")
    sys.exit(1)

print(f"Folder: {folder.Name} | Items: {folder.Items.Count}")

# Show actual dates in the folder
items = folder.Items
items.Sort("[ReceivedTime]", True)
print("\nActual items in folder:")
for i, item in enumerate(items):
    if i >= 5:
        break
    try:
        print(f"  {item.ReceivedTime} - {item.Subject[:40]}")
    except Exception as e:
        print(f"  Error reading item: {e}")

# Try the date query more carefully
date_from = datetime.now() - timedelta(days=100)
print(f"\nCutoff date: {date_from}")

# Try direct date comparison without Restrict
print("\nManual comparison (iterating):")
count_manual = 0
for item in folder.Items:
    try:
        recv = item.ReceivedTime
        # Convert to naive datetime for comparison
        if hasattr(recv, 'tzinfo') and recv.tzinfo:
            recv_naive = recv.replace(tzinfo=None)
        else:
            recv_naive = datetime(recv.year, recv.month, recv.day, recv.hour, recv.minute, recv.second)
        if recv_naive >= date_from:
            count_manual += 1
    except:
        pass
print(f"  Found {count_manual} items after {date_from}")

# Try Restrict with different approaches
print("\nTrying Restrict with different approaches:")

# Approach 1: Short date format (just date, no time)
date_str_short = date_from.strftime('%d/%m/%Y')
query1 = f"[ReceivedTime] >= '{date_str_short}'"
print(f"  Query 1: {query1}")
try:
    result1 = folder.Items.Restrict(query1)
    print(f"    Count: {sum(1 for _ in result1)}")
except Exception as e:
    print(f"    Error: {e}")

# Approach 2: US format short
date_str_us = date_from.strftime('%m/%d/%Y')
query2 = f"[ReceivedTime] >= '{date_str_us}'"
print(f"  Query 2 (US): {query2}")
try:
    result2 = folder.Items.Restrict(query2)
    print(f"    Count: {sum(1 for _ in result2)}")
except Exception as e:
    print(f"    Error: {e}")

# Check what dates look like in Restrict format
first_item = folder.Items.GetFirst()
if first_item:
    recv_time = first_item.ReceivedTime
    print(f"\n  First item ReceivedTime: {recv_time}")
    print(f"  Type: {type(recv_time)}")
    
# Check locale
import locale
print(f"\n  System locale: {locale.getdefaultlocale()}")
