"""DMS client for DMSforLegal operations."""

from datetime import datetime
from typing import List, Dict, Any, Optional

from outlook_client.base import BaseOutlookClient
from models import Email


class DMSClient(BaseOutlookClient):
    """Client for DMS (DMSforLegal) operations."""
    
    # DMS folder structure constants
    DMS_STORE_NAME = "DMSforLegal"
    DMS_ROOT_FOLDER = "_My Matters"
    DMS_EMAILS_FOLDER = "Emails"
    DMS_ADMIN_FOLDER = "Admin"
    
    def _get_dms_store(self):
        """Get the DMSforLegal Outlook store."""
        self._ensure_connection()
        try:
            for store in self._namespace.Stores:
                if store.DisplayName == self.DMS_STORE_NAME:
                    return store
        except Exception:
            pass
        return None
    
    def _get_folder_by_path(self, path: str):
        """Navigate to a folder by path within DMS store."""
        store = self._get_dms_store()
        if not store:
            return None
        
        try:
            folder = store.GetRootFolder()
            parts = path.split("\\")
            
            for part in parts:
                if not part:
                    continue
                found = False
                for subfolder in folder.Folders:
                    if subfolder.Name == part:
                        folder = subfolder
                        found = True
                        break
                if not found:
                    return None
            return folder
        except Exception:
            return None
    
    def list_dms_clients(self) -> List[str]:
        """List all client folders in DMS."""
        folder = self._get_folder_by_path(self.DMS_ROOT_FOLDER)
        if not folder:
            return []
        
        try:
            clients = [f.Name for f in folder.Folders]
            return sorted(clients)
        except Exception:
            return []
    
    def list_dms_matters(self, client: str) -> List[str]:
        """List all matter folders for a client in DMS."""
        path = f"{self.DMS_ROOT_FOLDER}\\{client}"
        folder = self._get_folder_by_path(path)
        if not folder:
            return []
        
        try:
            matters = [f.Name for f in folder.Folders]
            return sorted(matters)
        except Exception:
            return []
    
    def get_dms_emails(self, client: str, matter: str, limit: int = 50) -> List[Email]:
        """Get emails from a matter's Emails folder in DMS."""
        path = f"{self.DMS_ROOT_FOLDER}\\{client}\\{matter}\\{self.DMS_EMAILS_FOLDER}"
        folder = self._get_folder_by_path(path)
        if not folder:
            return []
        
        results = []
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            
            for message in items:
                if len(results) >= limit:
                    break
                try:
                    email = self._message_to_email(
                        message, 
                        folder_path=f"DMS/{client}/{matter}",
                        direction="filed"
                    )
                    if email:
                        results.append(email)
                except Exception:
                    continue
        except Exception:
            pass
        
        return results
    
    def get_dms_admin_emails(self, client: str, matter: str, limit: int = 50) -> List[Email]:
        """Get emails from a matter's Admin folder in DMS."""
        path = f"{self.DMS_ROOT_FOLDER}\\{client}\\{matter}\\{self.DMS_ADMIN_FOLDER}"
        folder = self._get_folder_by_path(path)
        if not folder:
            return []
        
        results = []
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            
            for message in items:
                if len(results) >= limit:
                    break
                try:
                    email = self._message_to_email(
                        message, 
                        folder_path=f"DMS/{client}/{matter}/Admin",
                        direction="filed"
                    )
                    if email:
                        results.append(email)
                except Exception:
                    continue
        except Exception:
            pass
        
        return results
    
    def search_dms_emails(
        self,
        client: str = None,
        matter: str = None,
        subject_contains: str = None,
        date_from: datetime = None,
        date_to: datetime = None,
        limit: int = 50,
    ) -> List[Email]:
        """Search emails across DMS with optional filters."""
        results = []
        
        if client:
            clients_to_search = [client]
        else:
            clients_to_search = self.list_dms_clients()
        
        for c in clients_to_search:
            if len(results) >= limit:
                break
            
            if matter and client:
                matters_to_search = [matter]
            else:
                matters_to_search = self.list_dms_matters(c)
            
            for m in matters_to_search:
                if len(results) >= limit:
                    break
                
                path = f"{self.DMS_ROOT_FOLDER}\\{c}\\{m}\\{self.DMS_EMAILS_FOLDER}"
                folder = self._get_folder_by_path(path)
                if not folder:
                    continue
                
                try:
                    items = folder.Items
                    items.Sort("[ReceivedTime]", True)
                    
                    if date_from:
                        date_str = date_from.strftime("%d/%m/%Y %H:%M")
                        items = items.Restrict(f"[ReceivedTime] >= '{date_str}'")
                    if date_to:
                        date_str = date_to.strftime("%d/%m/%Y %H:%M")
                        items = items.Restrict(f"[ReceivedTime] <= '{date_str}'")
                    
                    for message in items:
                        if len(results) >= limit:
                            break
                        
                        try:
                            if subject_contains:
                                subject = message.Subject or ""
                                if subject_contains.lower() not in subject.lower():
                                    continue
                            
                            email = self._message_to_email(
                                message,
                                folder_path=f"DMS/{c}/{m}",
                                direction="filed"
                            )
                            if email:
                                results.append(email)
                        except Exception:
                            continue
                except Exception:
                    continue
        
        return results
    
    def _check_dms_duplicate(self, message, target_folder) -> Optional[str]:
        """Check if an email already exists in the target DMS folder."""
        try:
            subject = message.Subject or ""
            received_time = message.ReceivedTime
            sender = getattr(message, "SenderEmailAddress", "") or ""
            
            for item in target_folder.Items:
                try:
                    item_subject = item.Subject or ""
                    item_received = item.ReceivedTime
                    item_sender = getattr(item, "SenderEmailAddress", "") or ""
                    
                    if (item_subject == subject and 
                        item_received == received_time and
                        item_sender.lower() == sender.lower()):
                        return item.EntryID
                except Exception:
                    continue
        except Exception:
            pass
        
        return None
    
    def file_email_to_dms(
        self,
        email_id: str,
        client_name: str,
        matter_name: str,
    ) -> Dict[str, Any]:
        """File an email to a DMS client/matter folder."""
        self._ensure_connection()
        
        dms_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}\\{self.DMS_EMAILS_FOLDER}"
        emails_folder = self._get_folder_by_path(dms_path)
        
        if not emails_folder:
            client_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}"
            client_folder = self._get_folder_by_path(client_path)
            
            if not client_folder:
                return {"success": False, "error": f"Client '{client_name}' not found in DMS"}
            
            matter_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}"
            matter_folder = self._get_folder_by_path(matter_path)
            
            if not matter_folder:
                return {"success": False, "error": f"Matter '{matter_name}' not found for client '{client_name}'"}
            
            return {"success": False, "error": f"Emails folder not found for matter '{matter_name}'. Please create the Emails subfolder in DMS."}
        
        try:
            message = self._namespace.GetItemFromID(email_id)
        except Exception as e:
            return {"success": False, "error": f"Email not found: {str(e)}"}
        
        duplicate_id = self._check_dms_duplicate(message, emails_folder)
        if duplicate_id:
            return {
                "success": False,
                "error": "Email already exists in destination folder",
                "duplicate_entry_id": duplicate_id,
                "subject": message.Subject,
                "received_time": message.ReceivedTime.isoformat() if hasattr(message.ReceivedTime, 'isoformat') else str(message.ReceivedTime)
            }
        
        try:
            copied = message.Copy()
            copied.Save()
            filed_message = copied.Move(emails_folder)
            filed_entry_id = filed_message.EntryID
        except Exception as e:
            return {"success": False, "error": f"Failed to copy email to DMS: {str(e)}"}
        
        try:
            existing_categories = message.Categories or ""
            categories = [c.strip() for c in existing_categories.split(",") if c.strip()]
            categories = [c for c in categories if not c.startswith(self.TRIAGE_CATEGORY_PREFIX)]
            
            if "Filed" not in categories:
                categories.append("Filed")
            categories.append(self.TRIAGE_CATEGORIES["processed"])
            
            message.Categories = ", ".join(categories)
            message.Save()
        except Exception as e:
            return {
                "success": True,
                "filed_entry_id": filed_entry_id,
                "subject": message.Subject,
                "received_time": message.ReceivedTime.isoformat() if hasattr(message.ReceivedTime, 'isoformat') else str(message.ReceivedTime),
                "filed_category": "Filed",
                "triage_status": "processed",
                "warning": f"Filed successfully but failed to update categories: {str(e)}"
            }
        
        return {
            "success": True,
            "filed_entry_id": filed_entry_id,
            "subject": message.Subject,
            "received_time": message.ReceivedTime.isoformat() if hasattr(message.ReceivedTime, 'isoformat') else str(message.ReceivedTime),
            "filed_category": "Filed",
            "triage_status": "processed"
        }
    
    def file_email_to_dms_admin(
        self,
        email_id: str,
        client_name: str,
        matter_name: str,
    ) -> Dict[str, Any]:
        """File an admin email to a DMS client/matter Admin folder."""
        self._ensure_connection()
        
        dms_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}\\{self.DMS_ADMIN_FOLDER}"
        admin_folder = self._get_folder_by_path(dms_path)
        
        if not admin_folder:
            client_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}"
            client_folder = self._get_folder_by_path(client_path)
            
            if not client_folder:
                return {"success": False, "error": f"Client '{client_name}' not found in DMS"}
            
            matter_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}"
            matter_folder = self._get_folder_by_path(matter_path)
            
            if not matter_folder:
                return {"success": False, "error": f"Matter '{matter_name}' not found for client '{client_name}'"}
            
            return {"success": False, "error": f"Admin folder not found for matter '{matter_name}'. Please create the Admin subfolder in DMS."}
        
        try:
            message = self._namespace.GetItemFromID(email_id)
        except Exception as e:
            return {"success": False, "error": f"Email not found: {str(e)}"}
        
        duplicate_id = self._check_dms_duplicate(message, admin_folder)
        if duplicate_id:
            return {
                "success": False,
                "error": "Email already exists in destination folder",
                "duplicate_entry_id": duplicate_id,
                "subject": message.Subject,
                "received_time": message.ReceivedTime.isoformat() if hasattr(message.ReceivedTime, 'isoformat') else str(message.ReceivedTime)
            }
        
        try:
            copied = message.Copy()
            copied.Save()
            filed_message = copied.Move(admin_folder)
            filed_entry_id = filed_message.EntryID
        except Exception as e:
            return {"success": False, "error": f"Failed to copy email to DMS Admin: {str(e)}"}
        
        try:
            existing_categories = message.Categories or ""
            categories = [c.strip() for c in existing_categories.split(",") if c.strip()]
            categories = [c for c in categories if not c.startswith(self.TRIAGE_CATEGORY_PREFIX)]
            
            if "Filed" not in categories:
                categories.append("Filed")
            categories.append(self.TRIAGE_CATEGORIES["processed"])
            
            message.Categories = ", ".join(categories)
            message.Save()
        except Exception as e:
            return {
                "success": True,
                "filed_entry_id": filed_entry_id,
                "subject": message.Subject,
                "received_time": message.ReceivedTime.isoformat() if hasattr(message.ReceivedTime, 'isoformat') else str(message.ReceivedTime),
                "filed_to": "Admin",
                "warning": f"Filed successfully but failed to update categories: {str(e)}"
            }
        
        return {
            "success": True,
            "filed_entry_id": filed_entry_id,
            "subject": message.Subject,
            "received_time": message.ReceivedTime.isoformat() if hasattr(message.ReceivedTime, 'isoformat') else str(message.ReceivedTime),
            "filed_to": "Admin",
            "triage_status": "processed"
        }
    
    def batch_file_emails_to_dms(
        self,
        email_ids: List[str],
        client_name: str,
        matter_name: str,
    ) -> Dict[str, Any]:
        """File multiple emails to a DMS client/matter folder."""
        self._ensure_connection()
        
        if not email_ids:
            return {
                "success": True,
                "filed_count": 0,
                "failed_count": 0,
                "filed_emails": [],
                "failed_emails": []
            }
        
        dms_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}\\{self.DMS_EMAILS_FOLDER}"
        emails_folder = self._get_folder_by_path(dms_path)
        
        if not emails_folder:
            client_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}"
            client_folder = self._get_folder_by_path(client_path)
            
            if not client_folder:
                return {"success": False, "error": f"Client '{client_name}' not found in DMS"}
            
            matter_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}"
            matter_folder = self._get_folder_by_path(matter_path)
            
            if not matter_folder:
                return {"success": False, "error": f"Matter '{matter_name}' not found for client '{client_name}'"}
            
            return {"success": False, "error": f"Emails folder not found for matter '{matter_name}'. Please create the Emails subfolder in DMS."}
        
        filed_emails = []
        failed_emails = []
        skipped_duplicates = []
        
        for email_id in email_ids:
            try:
                message = self._namespace.GetItemFromID(email_id)
                
                duplicate_id = self._check_dms_duplicate(message, emails_folder)
                if duplicate_id:
                    skipped_duplicates.append({
                        "email_id": email_id,
                        "duplicate_entry_id": duplicate_id,
                        "subject": message.Subject,
                        "received_time": message.ReceivedTime.isoformat() if hasattr(message.ReceivedTime, 'isoformat') else str(message.ReceivedTime),
                        "reason": "Already exists in destination folder"
                    })
                    continue
                
                copied = message.Copy()
                copied.Save()
                filed_message = copied.Move(emails_folder)
                
                existing_categories = message.Categories or ""
                categories = [c.strip() for c in existing_categories.split(",") if c.strip()]
                categories = [c for c in categories if not c.startswith(self.TRIAGE_CATEGORY_PREFIX)]
                
                if "Filed" not in categories:
                    categories.append("Filed")
                categories.append(self.TRIAGE_CATEGORIES["processed"])
                
                message.Categories = ", ".join(categories)
                message.Save()
                
                filed_emails.append({
                    "entry_id": filed_message.EntryID,
                    "subject": message.Subject,
                    "received_time": message.ReceivedTime.isoformat() if hasattr(message.ReceivedTime, 'isoformat') else str(message.ReceivedTime)
                })
                
            except Exception as e:
                failed_emails.append({
                    "email_id": email_id,
                    "error": str(e)
                })
        
        return {
            "success": True,
            "filed_count": len(filed_emails),
            "failed_count": len(failed_emails),
            "skipped_duplicates_count": len(skipped_duplicates),
            "filed_emails": filed_emails,
            "failed_emails": failed_emails,
            "skipped_duplicates": skipped_duplicates
        }
