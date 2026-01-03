"""Count all emails in DMSforLegal matters."""
import win32com.client

ol = win32com.client.Dispatch('Outlook.Application')
ns = ol.GetNamespace('MAPI')
dms = ns.Folders['DMSforLegal']
my_matters = dms.Folders['_My Matters']

total_emails = 0
matter_count = 0

for client in my_matters.Folders:
    for matter in client.Folders:
        matter_count += 1
        for subfolder in matter.Folders:
            if subfolder.Name == 'Emails':
                count = subfolder.Items.Count
                if count > 0:
                    total_emails += count
                    print(f"{client.Name} / {matter.Name}: {count} emails")

print(f"\nTotal: {total_emails} emails across {matter_count} matters")
