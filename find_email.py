import win32com.client
import pythoncom

pythoncom.CoInitialize()
outlook = win32com.client.Dispatch('Outlook.Application')
ns = outlook.GetNamespace('MAPI')

email_id = '0000000072D31F9EDED3FD41970F5D5BAAF1D5B90700DC1DE0BEBA1FFC468F89E645E46AD1E0000007F7F43E850000'

# Try with each store's ID
for store in ns.Stores:
    try:
        store_id = store.StoreID
        message = ns.GetItemFromID(email_id, store_id)
        print(f"FOUND in store: {store.DisplayName}")
        print(f"Subject: {message.Subject}")
        print(f"From: {message.SenderEmailAddress}")
        print(f"Received: {message.ReceivedTime}")
        print(f"Attachments: {message.Attachments.Count}")
        for i in range(1, message.Attachments.Count + 1):
            att = message.Attachments.Item(i)
            print(f"  - {att.FileName} ({att.Size} bytes)")
        break
    except Exception as e:
        print(f"{store.DisplayName}: Not found - {e}")
