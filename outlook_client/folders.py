"""Folders client for Outlook folder operations."""

from typing import List, Dict, Any

from outlook_client.base import BaseOutlookClient


class FoldersClient(BaseOutlookClient):
    """Client for Outlook folder operations."""
    
    def move_to_folder(self, email_id: str, folder_name: str) -> bool:
        """Move an email to a specified folder."""
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            inbox = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
            dest_folder = None
            
            for folder in inbox.Folders:
                if folder.Name.lower() == folder_name.lower():
                    dest_folder = folder
                    break
            
            if dest_folder:
                message.Move(dest_folder)
                return True
            return False
        except Exception:
            return False
    
    def move_to_archive(self, email_id: str, folder_path: str = "Archive", 
                        create_path: bool = False) -> Dict[str, Any]:
        r"""Move an email to a folder (default: Archive).
        
        Supports both simple folder names and full paths with subfolders.
        Examples:
          - "Archive" (root-level folder)
          - "Inbox\~Zero\Growth Engineering" (subfolder path)
          - "\\David.Sant@harperjames.co.uk\Inbox\~Zero" (full path - mailbox prefix stripped)
        """
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            
            inbox = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
            root = inbox.Parent
            
            clean_path = folder_path.strip("\\")
            first_part = clean_path.split("\\")[0] if "\\" in clean_path else ""
            if "@" in first_part:
                clean_path = "\\".join(clean_path.split("\\")[1:])
            
            path_parts = [p for p in clean_path.split("\\") if p]
            
            current_folder = root
            folders_created = []
            for part in path_parts:
                found = False
                for subfolder in current_folder.Folders:
                    if subfolder.Name.lower() == part.lower():
                        current_folder = subfolder
                        found = True
                        break
                if not found:
                    if create_path:
                        current_folder = current_folder.Folders.Add(part)
                        folders_created.append(part)
                    else:
                        return {"success": False, "error": f"Folder '{part}' not found in path '{folder_path}'"}
            target_folder = current_folder
            
            if target_folder == root:
                return {"success": False, "error": f"Invalid folder path: {folder_path}"}
            
            moved_message = message.Move(target_folder)
            
            result = {
                "success": True,
                "old_id": email_id,
                "new_id": moved_message.EntryID,
                "folder": target_folder.FolderPath
            }
            if folders_created:
                result["folders_created"] = folders_created
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def batch_move_to_archive(self, email_ids: List[str], folder_path: str = "Archive",
                              create_path: bool = False) -> Dict[str, Any]:
        """Move multiple emails to a folder (default: Archive)."""
        results = {"success": 0, "failed": 0, "moved": [], "errors": [], "folders_created": []}
        
        for email_id in email_ids:
            result = self.move_to_archive(email_id, folder_path=folder_path, create_path=create_path)
            if result.get("success"):
                results["success"] += 1
                results["moved"].append({"old_id": email_id, "new_id": result.get("new_id")})
                if result.get("folders_created"):
                    for folder in result["folders_created"]:
                        if folder not in results["folders_created"]:
                            results["folders_created"].append(folder)
            else:
                results["failed"] += 1
                results["errors"].append({"id": email_id, "error": result.get("error")})
        
        if not results["folders_created"]:
            del results["folders_created"]
        
        return results
    
    def list_subfolders(self, folder_path: str) -> List[str]:
        r"""List subfolders within a given folder path.
        
        Args:
            folder_path: Path to the folder, e.g. "Inbox\~Zero" or "Inbox"
        """
        self._ensure_connection()
        
        try:
            inbox = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
            root = inbox.Parent
            
            clean_path = folder_path.strip("\\")
            first_part = clean_path.split("\\")[0] if "\\" in clean_path else ""
            if "@" in first_part:
                clean_path = "\\".join(clean_path.split("\\")[1:])
            
            path_parts = [p for p in clean_path.split("\\") if p]
            
            current_folder = root
            for part in path_parts:
                found = False
                for subfolder in current_folder.Folders:
                    if subfolder.Name.lower() == part.lower():
                        current_folder = subfolder
                        found = True
                        break
                if not found:
                    return []
            
            subfolders = [f.Name for f in current_folder.Folders]
            return sorted(subfolders)
            
        except Exception:
            return []
    
    def set_category(self, email_id: str, category: str) -> bool:
        """Set Outlook category on an email."""
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            existing = message.Categories or ""
            if category not in existing:
                message.Categories = f"{existing}, {category}".strip(", ")
            message.Save()
            return True
        except:
            return False
