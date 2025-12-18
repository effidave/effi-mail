"""Outlook COM interface for email access."""

import win32com.client
from datetime import datetime, timedelta
from typing import List, Optional, Generator
import pythoncom

from models import Email, Domain, EmailCategory
from database import Database


class OutlookClient:
    """Windows COM client for Outlook email access."""
    
    # Outlook folder constants
    FOLDER_INBOX = 6
    FOLDER_SENT = 5
    FOLDER_DRAFTS = 16
    FOLDER_DELETED = 3
    
    def __init__(self, db: Optional[Database] = None):
        self.db = db or Database()
        self._outlook = None
        self._namespace = None
    
    def _ensure_connection(self):
        """Ensure COM connection is established."""
        if self._outlook is None:
            pythoncom.CoInitialize()
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
    
    def _get_sender_email(self, message) -> str:
        """Extract sender email, handling Exchange addresses."""
        try:
            sender = message.Sender
            if sender is not None:
                # Try to get SMTP address from Exchange user
                if sender.AddressEntryUserType == 0:  # Exchange user
                    exch_user = sender.GetExchangeUser()
                    if exch_user:
                        return exch_user.PrimarySmtpAddress
                # Try PropertyAccessor for SMTP address
                try:
                    return message.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                    )
                except:
                    pass
            return message.SenderEmailAddress
        except:
            return message.SenderEmailAddress or ""
    
    def _extract_domain(self, email: str) -> str:
        """Extract domain from email address."""
        if email and "@" in email:
            return email.split("@")[-1].lower()
        return "(no domain)"
    
    def _message_to_email(self, message, folder_path: str = "Inbox") -> Optional[Email]:
        """Convert Outlook message to Email object."""
        try:
            sender_email = self._get_sender_email(message)
            received = message.ReceivedTime
            if hasattr(received, 'replace'):
                received = received.replace(tzinfo=None)
            
            # Get attachments
            attachments = []
            has_attachments = message.Attachments.Count > 0
            if has_attachments:
                for i in range(1, min(message.Attachments.Count + 1, 11)):  # Max 10 attachment names
                    try:
                        attachments.append(message.Attachments.Item(i).FileName)
                    except:
                        pass
            
            # Get body preview (first 500 chars)
            body_preview = ""
            try:
                body = message.Body or ""
                body_preview = body[:500].replace("\r\n", " ").strip()
            except:
                pass
            
            return Email(
                id=message.EntryID,
                subject=message.Subject or "(No Subject)",
                sender_name=message.SenderName or "",
                sender_email=sender_email,
                domain=self._extract_domain(sender_email),
                received_time=received,
                body_preview=body_preview,
                has_attachments=has_attachments,
                attachment_names=attachments,
                categories=message.Categories or "",
                conversation_id=getattr(message, 'ConversationID', None),
                folder_path=folder_path
            )
        except Exception as e:
            return None
    
    def get_emails(self, days: int = 7, folder_id: int = None, 
                   exclude_categories: List[str] = None) -> Generator[Email, None, None]:
        """
        Fetch emails from Outlook.
        
        Args:
            days: Number of days to look back
            folder_id: Outlook folder constant (default: Inbox)
            exclude_categories: Categories to exclude (e.g., ["Unfocused"])
        
        Yields:
            Email objects
        """
        self._ensure_connection()
        
        if folder_id is None:
            folder_id = self.FOLDER_INBOX
        
        folder = self._namespace.GetDefaultFolder(folder_id)
        folder_path = folder.Name
        
        # Calculate date filter
        date_cutoff = datetime.now() - timedelta(days=days)
        date_str = date_cutoff.strftime("%m/%d/%Y %H:%M %p")
        filter_str = f"[ReceivedTime] >= '{date_str}'"
        
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)  # Descending
        filtered = messages.Restrict(filter_str)
        
        exclude_categories = exclude_categories or []
        
        for message in filtered:
            try:
                # Check for excluded categories
                msg_categories = message.Categories or ""
                if any(cat in msg_categories for cat in exclude_categories):
                    continue
                
                email = self._message_to_email(message, folder_path)
                if email:
                    yield email
            except:
                continue
    
    def get_email_body(self, email_id: str, max_length: int = 10000) -> str:
        """
        Get full email body by EntryID.
        
        Args:
            email_id: Outlook EntryID
            max_length: Maximum body length to return
        
        Returns:
            Email body text, truncated if necessary
        """
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            body = message.Body or ""
            if len(body) > max_length:
                return body[:max_length] + f"\n\n[Truncated - {len(body) - max_length} more characters]"
            return body
        except Exception as e:
            return f"Error retrieving email body: {e}"
    
    def get_email_html(self, email_id: str) -> str:
        """Get email HTML body by EntryID."""
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            return message.HTMLBody or ""
        except Exception as e:
            return f"Error retrieving email HTML: {e}"
    
    def sync_emails_to_db(self, days: int = 7, 
                          exclude_categories: List[str] = None) -> dict:
        """
        Sync emails from Outlook to database.
        
        Returns:
            Statistics about the sync operation
        """
        stats = {"new": 0, "updated": 0, "domains_updated": 0}
        domain_data = {}  # Track domain statistics
        
        for email in self.get_emails(days=days, exclude_categories=exclude_categories):
            # Check if email exists
            existing = self.db.get_email(email.id)
            if existing:
                # Preserve triage data from existing email
                email.triage_status = existing.triage_status
                email.client_id = existing.client_id
                email.matter_id = existing.matter_id
                email.notes = existing.notes
                email.processed_at = existing.processed_at
                stats["updated"] += 1
            else:
                stats["new"] += 1
            
            self.db.upsert_email(email)
            
            # Track domain data
            if email.domain not in domain_data:
                domain_data[email.domain] = {
                    "count": 0,
                    "last_seen": email.received_time,
                    "senders": set()
                }
            domain_data[email.domain]["count"] += 1
            domain_data[email.domain]["senders"].add(email.sender_name)
            if email.received_time > domain_data[email.domain]["last_seen"]:
                domain_data[email.domain]["last_seen"] = email.received_time
        
        # Update domain records
        for domain_name, data in domain_data.items():
            existing_domain = self.db.get_domain(domain_name)
            category = existing_domain.category if existing_domain else EmailCategory.UNCATEGORIZED
            
            domain = Domain(
                name=domain_name,
                category=category,
                email_count=data["count"],
                last_seen=data["last_seen"],
                sample_senders=list(data["senders"])[:5]
            )
            self.db.upsert_domain(domain)
            stats["domains_updated"] += 1
        
        return stats
    
    def move_to_folder(self, email_id: str, folder_name: str) -> bool:
        """Move an email to a specified folder."""
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            # Find destination folder
            inbox = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
            dest_folder = None
            
            # Search in inbox subfolders
            for folder in inbox.Folders:
                if folder.Name.lower() == folder_name.lower():
                    dest_folder = folder
                    break
            
            if dest_folder:
                message.Move(dest_folder)
                return True
            return False
        except Exception as e:
            return False
    
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
