import win32com.client
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('Outlook.Application')
ns = outlook.GetNamespace('MAPI')
inbox = ns.GetDefaultFolder(6)

# Test with days=30 US format - would this work?
date_from = datetime.now() - timedelta(days=30)
date_str = date_from.strftime('%m/%d/%Y %H:%M %p')  # US format
date_query = f"[ReceivedTime] >= '{date_str}'"

print(f'Query (days=30, US format): {date_query}')
print(f'This would be interpreted in UK locale as: July 12, 2025')

messages = inbox.Items
filtered = messages.Restrict(date_query)
print(f'Restrict returned: {filtered.Count} items')

# Compare with days=1
date_from = datetime.now() - timedelta(days=1)
date_str = date_from.strftime('%m/%d/%Y %H:%M %p')  # US format
date_query = f"[ReceivedTime] >= '{date_str}'"

print(f'\nQuery (days=1, US format): {date_query}')
print(f'This would be interpreted in UK locale as: May 1, 2026 (FUTURE!)')

messages2 = inbox.Items
filtered2 = messages2.Restrict(date_query)
print(f'Restrict returned: {filtered2.Count} items')

print(f"\nQuery: {date_query}")

messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

try:
    filtered = messages.Restrict(date_query)
    print(f"Filtered count: {filtered.Count}")
    
    print("\nFiltered emails:")
    count = 0
    for item in filtered:
        if count >= 10:
            break
        try:
            print(f"{count+1}. {item.ReceivedTime} - {item.Subject[:50]}")
            count += 1
        except Exception as e:
            print(f"Error: {e}")
except Exception as e:
    print(f"Restrict failed: {e}")
