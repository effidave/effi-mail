"""Base Outlook client with connection management and shared utilities."""

import win32com.client
from datetime import datetime, timedelta
from typing import List, Optional, Dict, Any
import pythoncom
import os
import mimetypes

from models import Email, TriageStatus


class BaseOutlookClient:
    """Base class for Outlook COM operations.
    
    Provides connection management and shared utility methods.
    All specialized clients inherit from this.
    """
    
    # Outlook folder constants
    FOLDER_INBOX = 6
    FOLDER_SENT = 5
    FOLDER_DRAFTS = 16
    FOLDER_DELETED = 3
    
    # MAPI property for PR_INTERNET_MESSAGE_ID
    PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
    
    # Triage category constants
    TRIAGE_CATEGORY_PREFIX = "effi:"
    TRIAGE_CATEGORIES = {
        "action": "effi:action",
        "waiting": "effi:waiting",
        "processed": "effi:processed",
        "archived": "effi:archived",
    }
    
    # DMS constants
    DMS_STORE_NAME = "DMSforLegal"
    DMS_ROOT_FOLDER = "_My Matters"
    DMS_EMAILS_FOLDER = "Emails"
    DMS_ADMIN_FOLDER = "Admin"
    
    # DASL property for custom RecipientDomain field
    RECIPIENT_DOMAIN_PROP = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/RecipientDomain"
    
    def __init__(self):
        self._outlook = None
        self._namespace = None
    
    def _reset_connection(self):
        """Reset COM connection (call when Outlook restarts)."""
        self._outlook = None
        self._namespace = None
    
    def _ensure_connection(self):
        """Ensure COM connection is established. Auto-reconnects if stale."""
        if self._outlook is None:
            pythoncom.CoInitialize()
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
        else:
            try:
                _ = self._namespace.CurrentUser
            except Exception:
                self._reset_connection()
                pythoncom.CoInitialize()
                self._outlook = win32com.client.Dispatch("Outlook.Application")
                self._namespace = self._outlook.GetNamespace("MAPI")
    
    # =========================================================================
    # Email Address Extraction
    # =========================================================================
    
    def _get_recipient_email(self, recipient) -> Optional[str]:
        """Extract SMTP email from a recipient (To, CC, or BCC)."""
        try:
            if recipient.AddressEntry:
                if recipient.AddressEntry.AddressEntryUserType == 0:
                    exch_user = recipient.AddressEntry.GetExchangeUser()
                    if exch_user:
                        return exch_user.PrimarySmtpAddress
                try:
                    return recipient.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                    )
                except:
                    return recipient.Address
        except:
            pass
        return None
    
    def _get_sender_email(self, message) -> str:
        """Extract sender email, handling Exchange addresses."""
        try:
            sender = message.Sender
            if sender is not None:
                if sender.AddressEntryUserType == 0:
                    exch_user = sender.GetExchangeUser()
                    if exch_user:
                        return exch_user.PrimarySmtpAddress
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
    
    def _get_internet_message_id(self, message) -> Optional[str]:
        """Extract the permanent Internet Message-ID (RFC2822 header)."""
        try:
            return message.PropertyAccessor.GetProperty(self.PR_INTERNET_MESSAGE_ID)
        except:
            return None
    
    def _compute_recipient_domains(self, recipients_to: List[str], recipients_cc: List[str]) -> str:
        """Compute unique recipient domains from To and CC lists."""
        domains = set()
        for email in recipients_to + recipients_cc:
            if email and "@" in email:
                domain = email.split("@")[-1].lower()
                if domain:
                    domains.add(domain)
        return ",".join(sorted(domains))
    
    def _extract_recipients(self, message, recipient_type: str) -> List[str]:
        """Extract recipient email addresses from a message."""
        recipients = []
        try:
            recipient_collection = message.Recipients
            type_map = {"To": 1, "CC": 2, "BCC": 3}
            target_type = type_map.get(recipient_type, 1)
            
            for i in range(1, recipient_collection.Count + 1):
                recipient = recipient_collection.Item(i)
                if recipient.Type == target_type:
                    try:
                        smtp_address = recipient.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                        )
                        recipients.append(smtp_address.lower())
                    except:
                        addr = recipient.Address
                        if addr and "@" in addr:
                            recipients.append(addr.lower())
        except:
            pass
        return recipients
    
    def _get_primary_recipient_domain(self, message) -> str:
        """Extract domain from primary recipient of a sent message."""
        try:
            recipients = message.Recipients
            if recipients.Count > 0:
                recipient = recipients.Item(1)
                try:
                    smtp_address = recipient.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                    )
                    return self._extract_domain(smtp_address)
                except:
                    return self._extract_domain(recipient.Address)
        except:
            pass
        return "(no domain)"
    
    # =========================================================================
    # Message Conversion
    # =========================================================================
    
    def _message_to_email(self, message, folder_path: str = "Inbox", direction: str = "inbound", 
                           recipient_domain: str = None) -> Optional[Email]:
        """Convert Outlook message to Email object."""
        try:
            sender_email = self._get_sender_email(message)
            received = message.ReceivedTime
            if hasattr(received, 'replace'):
                received = received.replace(tzinfo=None)
            
            # Get attachments - filter out inline images
            attachments = []
            has_attachments = message.Attachments.Count > 0
            if has_attachments:
                for i in range(1, message.Attachments.Count + 1):
                    try:
                        att = message.Attachments.Item(i)
                        filename = att.FileName
                        lower_name = filename.lower()
                        
                        doc_extensions = ('.docx', '.doc', '.pdf', '.xlsx', '.xls', 
                                         '.pptx', '.ppt', '.zip', '.rar', '.csv', '.txt')
                        is_document = lower_name.endswith(doc_extensions)
                        
                        if is_document:
                            attachments.append(filename)
                        else:
                            is_image = lower_name.endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))
                            if is_image:
                                try:
                                    content_id = att.PropertyAccessor.GetProperty(
                                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                                    )
                                    has_content_id = bool(content_id)
                                except:
                                    has_content_id = False
                                is_inline = has_content_id or lower_name.startswith('image')
                                if is_inline:
                                    continue
                            attachments.append(filename)
                    except:
                        pass
            
            attachments = attachments[:20]
            
            body_preview = ""
            try:
                body = message.Body or ""
                body_preview = body[:500].replace("\r\n", " ").strip()
            except:
                pass
            
            domain = recipient_domain if recipient_domain else self._extract_domain(sender_email)
            triage_status = TriageStatus.PROCESSED if direction == "outbound" else TriageStatus.PENDING
            
            recipients_to = self._extract_recipients(message, "To")
            recipients_cc = self._extract_recipients(message, "CC")
            recipient_domains = self._compute_recipient_domains(recipients_to, recipients_cc)
            internet_message_id = self._get_internet_message_id(message)
            
            return Email(
                id=message.EntryID,
                subject=message.Subject or "(No Subject)",
                sender_name=message.SenderName or "",
                sender_email=sender_email,
                domain=domain,
                received_time=received,
                body_preview=body_preview,
                has_attachments=has_attachments,
                attachment_names=attachments,
                categories=message.Categories or "",
                conversation_id=getattr(message, 'ConversationID', None),
                folder_path=folder_path,
                direction=direction,
                recipients_to=recipients_to,
                recipients_cc=recipients_cc,
                recipient_domains=recipient_domains,
                internet_message_id=internet_message_id,
                triage_status=triage_status,
            )
        except Exception as e:
            return None
    
    # =========================================================================
    # Category Operations
    # =========================================================================
    
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
