"""Copy a test email to a DMS matter folder.

Creates a test email and moves it to the specified DMS Emails folder.
This is for testing DMS folder access.
"""

import win32com.client
import pythoncom
from datetime import datetime


def copy_test_to_dms(client_name: str, matter_name: str):
    """Create and copy a test email to DMS Emails folder."""
    
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # Find DMS store
    dms_store = None
    for store in namespace.Stores:
        if store.DisplayName == "DMSforLegal":
            dms_store = store
            break
    
    if not dms_store:
        print("ERROR: DMSforLegal store not found")
        return False
    
    # Navigate to the Emails folder
    # Note: DMS folders have "(read-only)" suffix
    path_parts = ["_My Matters", client_name, matter_name, "Emails (read-only)"]
    folder = dms_store.GetRootFolder()
    
    for part in path_parts:
        found = False
        for subfolder in folder.Folders:
            if subfolder.Name == part:
                folder = subfolder
                found = True
                break
        if not found:
            print(f"ERROR: Folder '{part}' not found")
            print(f"Available folders at this level: {[f.Name for f in folder.Folders]}")
            return False
    
    print(f"Found target folder: {folder.FolderPath}")
    
    # Create a test mail item
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.Subject = f"Test Email - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    mail.Body = "This is a test email created to verify DMS folder access."
    mail.Save()  # Save to Drafts first
    
    print(f"Created test email: {mail.Subject}")
    
    # Move to DMS folder
    try:
        filed = mail.Move(folder)
        print(f"SUCCESS: Email moved to DMS folder")
        print(f"Filed EntryID: {filed.EntryID}")
        return True
    except Exception as e:
        print(f"ERROR moving email: {e}")
        return False


if __name__ == "__main__":
    client = "Youtility Limited"
    matter = "Virgin Money Agreement - Data Protection Advice - Youtility Limited (21598)"
    
    print(f"Copying test email to DMS:")
    print(f"  Client: {client}")
    print(f"  Matter: {matter}")
    print()
    
    copy_test_to_dms(client, matter)
