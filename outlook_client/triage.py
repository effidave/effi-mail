"""Triage client for email triage status via Outlook categories."""

from typing import Dict, Any, Optional, List

from outlook_client.base import BaseOutlookClient


class TriageClient(BaseOutlookClient):
    """Client for triage status operations using Outlook categories.
    
    Triage statuses are stored as Outlook categories with 'effi:' prefix
    to avoid conflicts with user categories.
    """
    
    def set_triage_status(self, email_id: str, status: str) -> bool:
        """Set triage status on an email using Outlook categories.
        
        Args:
            email_id: Outlook EntryID
            status: One of 'action', 'waiting', 'processed', 'archived'
            
        Returns:
            True if successful, False otherwise
        """
        if status not in self.TRIAGE_CATEGORIES:
            return False
            
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            existing = message.Categories or ""
            
            # Remove any existing effi: triage categories
            categories = [c.strip() for c in existing.split(",") if c.strip()]
            categories = [c for c in categories if not c.startswith(self.TRIAGE_CATEGORY_PREFIX)]
            
            # Add the new triage category
            categories.append(self.TRIAGE_CATEGORIES[status])
            
            message.Categories = ", ".join(categories)
            message.Save()
            return True
        except Exception as e:
            return False
    
    def get_triage_status(self, email_id: str) -> Optional[str]:
        """Get triage status from an email's Outlook categories.
        
        Args:
            email_id: Outlook EntryID
            
        Returns:
            Status string ('action', 'waiting', 'processed', 'archived') or None
        """
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            categories = message.Categories or ""
            
            for status, category in self.TRIAGE_CATEGORIES.items():
                if category in categories:
                    return status
            return None
        except:
            return None
    
    def clear_triage_status(self, email_id: str) -> bool:
        """Remove triage status from an email (reset to pending).
        
        Args:
            email_id: Outlook EntryID
            
        Returns:
            True if successful, False otherwise
        """
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            existing = message.Categories or ""
            
            # Remove all effi: categories
            categories = [c.strip() for c in existing.split(",") if c.strip()]
            categories = [c for c in categories if not c.startswith(self.TRIAGE_CATEGORY_PREFIX)]
            
            message.Categories = ", ".join(categories)
            message.Save()
            return True
        except:
            return False
    
    def batch_set_triage_status(self, email_ids: List[str], status: str) -> Dict[str, Any]:
        """Set triage status on multiple emails.
        
        Args:
            email_ids: List of Outlook EntryIDs
            status: One of 'action', 'waiting', 'processed', 'archived'
            
        Returns:
            Dict with success count, failure count, and failed IDs
        """
        results = {"success": 0, "failed": 0, "failed_ids": []}
        
        for email_id in email_ids:
            if self.set_triage_status(email_id, status):
                results["success"] += 1
            else:
                results["failed"] += 1
                results["failed_ids"].append(email_id)
        
        return results
    
    def get_pending_emails_from_domain(
        self,
        domain: str,
        days: int = 30,
        limit: int = 100,
    ) -> list:
        """Get pending (un-triaged) emails from a specific domain.
        
        Args:
            domain: Sender domain to filter
            days: Days to look back
            limit: Maximum results
            
        Returns:
            List of pending Email objects from that domain
        """
        from datetime import datetime, timedelta
        
        self._ensure_connection()
        
        folder = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
        
        date_from = datetime.now() - timedelta(days=days)
        date_str = date_from.strftime("%d/%m/%Y %H:%M")
        
        dasl_query = f"@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%@{domain}'"
        date_query = f"[ReceivedTime] >= '{date_str}'"
        
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)
        
        try:
            filtered_by_date = messages.Restrict(date_query)
            filtered = filtered_by_date.Restrict(dasl_query)
        except:
            filtered = messages.Restrict(date_query)
        
        pending_emails = []
        
        for message in filtered:
            if len(pending_emails) >= limit:
                break
            
            try:
                categories = message.Categories or ""
                has_triage = any(cat.strip().startswith(self.TRIAGE_CATEGORY_PREFIX) 
                                for cat in categories.split(",") if cat.strip())
                
                if not has_triage:
                    email = self._message_to_email(message, folder.Name, "inbound")
                    if email:
                        pending_emails.append(email)
            except:
                continue
        
        return pending_emails
