"""Outlook COM interface for email access."""

import win32com.client
from datetime import datetime, timedelta
from typing import List, Optional, Generator, Dict, Any
import pythoncom
import os
import mimetypes

from models import Email, Domain, EmailCategory, TriageStatus


class OutlookClient:
    """Windows COM client for Outlook email access.
    
    No database dependency - triage status is stored in Outlook categories.
    """
    
    # Outlook folder constants
    FOLDER_INBOX = 6
    FOLDER_SENT = 5
    FOLDER_DRAFTS = 16
    FOLDER_DELETED = 3
    
    # MAPI property for PR_INTERNET_MESSAGE_ID
    PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
    
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
            # Test if connection is still valid
            try:
                _ = self._namespace.CurrentUser
            except Exception:
                # Connection stale, reconnect
                self._reset_connection()
                pythoncom.CoInitialize()
                self._outlook = win32com.client.Dispatch("Outlook.Application")
                self._namespace = self._outlook.GetNamespace("MAPI")
    
    def _set_recipient_domains(self, limit: int = 200) -> dict:
        """Set RecipientDomain custom property on recent Sent Items.
        
        Populates a searchable field with recipient domains from To, CC, and BCC,
        enabling DASL queries to filter sent emails by recipient domain.
        
        Args:
            limit: Maximum items to process (default 200)
            
        Returns:
            Dict with counts of processed/updated items
        """
        self._ensure_connection()
        
        sent_folder = self._namespace.GetDefaultFolder(self.FOLDER_SENT)
        messages = sent_folder.Items
        messages.Sort("[SentOn]", True)  # Most recent first
        
        prop_path = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/RecipientDomain"
        processed = 0
        updated = 0
        
        for message in messages:
            if processed >= limit:
                break
            processed += 1
            
            try:
                # Check if already has RecipientDomain set
                try:
                    existing = message.PropertyAccessor.GetProperty(prop_path)
                    if existing:
                        continue  # Already set, skip
                except:
                    pass  # Property doesn't exist yet
                
                # Extract recipient domains from To, CC, and BCC
                domains = set()
                for recipient in message.Recipients:
                    try:
                        # Recipients collection includes To, CC, and BCC
                        email = self._get_recipient_email(recipient)
                        if email and "@" in email:
                            domain = email.split("@")[-1].lower()
                            domains.add(domain)
                    except:
                        continue
                
                if domains:
                    domain_str = ";".join(sorted(domains))
                    message.PropertyAccessor.SetProperty(prop_path, domain_str)
                    message.Save()
                    updated += 1
                    
            except:
                continue  # Skip problematic messages
        
        return {"processed": processed, "updated": updated}
    
    def _get_recipient_email(self, recipient) -> Optional[str]:
        """Extract SMTP email from a recipient (To, CC, or BCC)."""
        try:
            if recipient.AddressEntry:
                # Exchange user
                if recipient.AddressEntry.AddressEntryUserType == 0:
                    exch_user = recipient.AddressEntry.GetExchangeUser()
                    if exch_user:
                        return exch_user.PrimarySmtpAddress
                # Try SMTP property
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
    
    def _get_internet_message_id(self, message) -> Optional[str]:
        """Extract the permanent Internet Message-ID (RFC2822 header) from a message.
        
        This ID persists across folder moves and mailbox migrations, unlike EntryID.
        """
        try:
            return message.PropertyAccessor.GetProperty(self.PR_INTERNET_MESSAGE_ID)
        except:
            return None
    
    def _compute_recipient_domains(self, recipients_to: List[str], recipients_cc: List[str]) -> str:
        """Compute unique recipient domains from To and CC lists.
        
        Args:
            recipients_to: List of To recipient email addresses
            recipients_cc: List of CC recipient email addresses
            
        Returns:
            Comma-separated string of unique domains
        """
        domains = set()
        for email in recipients_to + recipients_cc:
            if email and "@" in email:
                domain = email.split("@")[-1].lower()
                if domain:
                    domains.add(domain)
        return ",".join(sorted(domains))
    
    def _message_to_email(self, message, folder_path: str = "Inbox", direction: str = "inbound", 
                           recipient_domain: str = None) -> Optional[Email]:
        """Convert Outlook message to Email object."""
        try:
            sender_email = self._get_sender_email(message)
            received = message.ReceivedTime
            if hasattr(received, 'replace'):
                received = received.replace(tzinfo=None)
            
            # Get attachments - filter out inline images (email signatures)
            attachments = []
            has_attachments = message.Attachments.Count > 0
            if has_attachments:
                for i in range(1, message.Attachments.Count + 1):
                    try:
                        att = message.Attachments.Item(i)
                        filename = att.FileName
                        lower_name = filename.lower()
                        
                        # Document file extensions are ALWAYS real attachments
                        doc_extensions = ('.docx', '.doc', '.pdf', '.xlsx', '.xls', 
                                         '.pptx', '.ppt', '.zip', '.rar', '.csv', '.txt')
                        is_document = lower_name.endswith(doc_extensions)
                        
                        if is_document:
                            attachments.append(filename)
                        else:
                            # For images, check if they're inline (signature images)
                            is_image = lower_name.endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))
                            
                            if is_image:
                                # Check for ContentID (inline images have this)
                                try:
                                    content_id = att.PropertyAccessor.GetProperty(
                                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                                    )
                                    has_content_id = bool(content_id)
                                except:
                                    has_content_id = False
                                
                                # Skip if inline OR matches signature pattern (image001.png etc)
                                is_inline = has_content_id or lower_name.startswith('image')
                                if is_inline:
                                    continue  # Skip this attachment
                            
                            # Keep non-inline attachments
                            attachments.append(filename)
                    except:
                        pass
            
            # Limit to 20 attachments max
            attachments = attachments[:20]
            
            # Get body preview (first 500 chars)
            body_preview = ""
            try:
                body = message.Body or ""
                body_preview = body[:500].replace("\r\n", " ").strip()
            except:
                pass
            
            # For outbound emails, use recipient domain instead of sender domain
            domain = recipient_domain if recipient_domain else self._extract_domain(sender_email)
            
            # Set triage status based on direction - sent emails are auto-processed
            triage_status = TriageStatus.PROCESSED if direction == "outbound" else TriageStatus.PENDING
            
            # Extract recipients (To and CC)
            recipients_to = self._extract_recipients(message, "To")
            recipients_cc = self._extract_recipients(message, "CC")
            
            # Compute recipient domains at sync time
            recipient_domains = self._compute_recipient_domains(recipients_to, recipients_cc)
            
            # Extract permanent message ID
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
    
    def _extract_recipients(self, message, recipient_type: str) -> List[str]:
        """Extract recipient email addresses from a message.
        
        Args:
            message: Outlook message object
            recipient_type: 'To' or 'CC'
        
        Returns:
            List of email addresses
        """
        recipients = []
        try:
            recipient_collection = message.Recipients
            # Outlook recipient types: 1=To, 2=CC, 3=BCC
            type_map = {"To": 1, "CC": 2, "BCC": 3}
            target_type = type_map.get(recipient_type, 1)
            
            for i in range(1, recipient_collection.Count + 1):
                recipient = recipient_collection.Item(i)
                if recipient.Type == target_type:
                    # Try to get SMTP address
                    try:
                        smtp_address = recipient.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                        )
                        recipients.append(smtp_address.lower())
                    except:
                        # Fall back to Address property
                        addr = recipient.Address
                        if addr and "@" in addr:
                            recipients.append(addr.lower())
        except:
            pass
        return recipients
    
    def get_emails(self, days: int = 7, folder_id: int = None, 
                   exclude_categories: List[str] = None, direction: str = "inbound",
                   since_time: datetime = None) -> Generator[Email, None, None]:
        """
        Fetch emails from Outlook.
        
        Args:
            days: Number of days to look back (used if since_time not provided)
            folder_id: Outlook folder constant (default: Inbox)
            exclude_categories: Categories to exclude (e.g., ["Unfocused"])
            direction: 'inbound' or 'outbound' - affects how domain is extracted
            since_time: Fetch emails after this timestamp (overrides days if provided)
        
        Yields:
            Email objects
        """
        self._ensure_connection()
        
        if folder_id is None:
            folder_id = self.FOLDER_INBOX
        
        folder = self._namespace.GetDefaultFolder(folder_id)
        folder_path = folder.Name
        
        # Calculate date filter - use since_time if provided, otherwise use days
        if since_time:
            date_cutoff = since_time
        else:
            date_cutoff = datetime.now() - timedelta(days=days)
        date_str = date_cutoff.strftime("%d/%m/%Y %H:%M")
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
                
                # For sent items, extract recipient domain
                recipient_domain = None
                if direction == "outbound":
                    recipient_domain = self._get_primary_recipient_domain(message)
                
                email = self._message_to_email(message, folder_path, direction, recipient_domain)
                if email:
                    yield email
            except:
                continue
    
    def _get_primary_recipient_domain(self, message) -> str:
        """Extract domain from primary recipient of a sent message."""
        try:
            recipients = message.Recipients
            if recipients.Count > 0:
                # Get first recipient
                recipient = recipients.Item(1)
                # Try to get SMTP address
                try:
                    smtp_address = recipient.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                    )
                    return self._extract_domain(smtp_address)
                except:
                    # Fall back to Address property
                    return self._extract_domain(recipient.Address)
        except:
            pass
        return "(no domain)"
    
    def get_emails_by_conversation_id(
        self,
        conversation_id: str,
        include_sent: bool = True,
        include_dms: bool = False,
        limit: int = 50,
        conversation_topic: str = None
    ) -> List[Email]:
        """Get all emails matching a ConversationID across folders.
        
        Note: Outlook's Restrict() doesn't support ConversationID filtering directly.
        We use ConversationTopic to find candidates, then verify ConversationID matches.
        
        Args:
            conversation_id: Exchange ConversationID to search for
            include_sent: Include Sent Items folder (default: True)
            include_dms: Include DMS store folders (default: False)
            limit: Maximum emails to return (default: 50)
            conversation_topic: ConversationTopic for filtering (optional but recommended)
            
        Returns:
            List of Email objects matching the ConversationID
        """
        self._ensure_connection()
        
        results = []
        
        # If no topic provided, we can't filter efficiently
        # ConversationID is not a filterable property in Outlook Restrict()
        if not conversation_topic:
            return results
        
        # Escape single quotes for Jet filter
        escaped_topic = conversation_topic.replace("'", "''")
        filter_str = f"[ConversationTopic] = '{escaped_topic}'"
        
        def search_folder(folder, direction: str, folder_path: str = None):
            """Search a folder and add matching emails to results."""
            nonlocal results
            if folder_path is None:
                folder_path = folder.Name
            try:
                messages = folder.Items.Restrict(filter_str)
                for message in messages:
                    if len(results) >= limit:
                        return
                    # Verify ConversationID matches exactly
                    msg_conv_id = getattr(message, 'ConversationID', None)
                    if msg_conv_id != conversation_id:
                        continue
                    if direction == "outbound":
                        recipient_domain = self._get_primary_recipient_domain(message)
                        email = self._message_to_email(message, folder_path, direction, recipient_domain)
                    else:
                        email = self._message_to_email(message, folder_path, direction)
                    if email:
                        results.append(email)
            except Exception:
                pass
        
        # Search Inbox
        inbox = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
        search_folder(inbox, "inbound")
        
        # Search Sent Items if requested
        if include_sent and len(results) < limit:
            sent = self._namespace.GetDefaultFolder(self.FOLDER_SENT)
            search_folder(sent, "outbound")
        
        # Search DMS if requested (limited to top-level DMS folders)
        if include_dms and len(results) < limit:
            try:
                dms_store = self._get_dms_store()
                if dms_store:
                    for folder in dms_store.Folders:
                        if len(results) >= limit:
                            break
                        folder_path = f"DMS\\{folder.Name}"
                        search_folder(folder, "inbound", folder_path)
            except Exception:
                pass
        
        return results
    
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
    
    def download_attachment(self, email_id: str, attachment_name: str, 
                           save_path: Optional[str] = None) -> Dict[str, Any]:
        """
        Download an attachment from an email and save it to disk.
        
        Args:
            email_id: The Outlook EntryID of the email
            attachment_name: The name of the attachment to download
            save_path: Where to save the file. If not provided, saves to 
                      ./attachments/{domain}/{date}/{filename}
        
        Returns:
            Dict with success status, file_path, file_size, and content_type
        """
        self._ensure_connection()
        
        # Get the email
        try:
            message = self._namespace.GetItemFromID(email_id)
        except Exception as e:
            return {"success": False, "error": f"Email not found: {e}"}
        
        # Find the attachment
        attachment = None
        for i in range(1, message.Attachments.Count + 1):
            att = message.Attachments.Item(i)
            if att.FileName == attachment_name:
                attachment = att
                break
        
        if not attachment:
            # List available attachments for debugging
            available = [message.Attachments.Item(i).FileName 
                        for i in range(1, message.Attachments.Count + 1)]
            return {
                "success": False, 
                "error": f"Attachment '{attachment_name}' not found",
                "available_attachments": available
            }
        
        # Determine save path
        if not save_path:
            # Get email metadata for default path
            try:
                # For sent emails, use recipient domain; for received, use sender domain
                sender = message.SenderEmailAddress or ""
                is_sent = sender.lower().endswith("harperjames.co.uk") or "@harperjames" in sender.lower()
                
                if is_sent and message.Recipients.Count > 0:
                    # Use first recipient's domain for sent emails
                    try:
                        recip = message.Recipients.Item(1)
                        recip_addr = recip.Address
                        # Handle Exchange addresses
                        if "@" not in recip_addr:
                            try:
                                recip_addr = recip.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                            except:
                                recip_addr = recip.PropertyAccessor.GetProperty(
                                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                                )
                        domain = self._extract_domain(recip_addr) or "unknown"
                    except:
                        domain = "sent"
                else:
                    domain = self._extract_domain(sender) or "unknown"
                
                received = message.ReceivedTime
                if hasattr(received, 'strftime'):
                    date_str = received.strftime('%Y-%m-%d')
                else:
                    date_str = str(received)[:10]
            except:
                domain = "unknown"
                date_str = datetime.now().strftime('%Y-%m-%d')
            
            # Sanitize domain for folder name
            domain = "".join(c if c.isalnum() or c in '.-_' else '_' for c in domain)
            
            # Build default path
            base_dir = os.path.join(os.path.dirname(__file__), "attachments")
            save_path = os.path.join(base_dir, domain, date_str, attachment_name)
        
        # Ensure absolute path (required by Outlook's SaveAsFile)
        save_path = os.path.abspath(save_path)
        
        # Create directory if needed
        try:
            os.makedirs(os.path.dirname(save_path), exist_ok=True)
        except Exception as e:
            return {"success": False, "error": f"Failed to create directory: {e}"}
        
        # Save attachment
        try:
            attachment.SaveAsFile(save_path)
            file_size = os.path.getsize(save_path)
            
            # Determine content type
            content_type, _ = mimetypes.guess_type(attachment_name)
            if not content_type:
                content_type = "application/octet-stream"
            
            return {
                "success": True,
                "file_path": os.path.abspath(save_path),
                "file_size": file_size,
                "content_type": content_type
            }
        except Exception as e:
            return {"success": False, "error": f"Failed to save attachment: {e}"}
    
    def list_attachments(self, email_id: str) -> Dict[str, Any]:
        """
        List all attachments for an email with details.
        
        Args:
            email_id: The Outlook EntryID of the email
            
        Returns:
            Dict with list of attachment details (name, size, type)
        """
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
        except Exception as e:
            return {"success": False, "error": f"Email not found: {e}"}
        
        attachments = []
        for i in range(1, message.Attachments.Count + 1):
            att = message.Attachments.Item(i)
            filename = att.FileName
            lower_name = filename.lower()
            
            # Determine if it's a real document (not inline image)
            doc_extensions = ('.docx', '.doc', '.pdf', '.xlsx', '.xls', 
                             '.pptx', '.ppt', '.zip', '.rar', '.csv', '.txt')
            is_document = lower_name.endswith(doc_extensions)
            
            is_image = lower_name.endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))
            is_inline = is_image and lower_name.startswith('image')
            
            content_type, _ = mimetypes.guess_type(filename)
            
            attachments.append({
                "name": filename,
                "size": att.Size,
                "content_type": content_type or "application/octet-stream",
                "is_document": is_document,
                "is_inline_image": is_inline
            })
        
        return {
            "success": True,
            "count": len(attachments),
            "attachments": attachments,
            "documents": [a for a in attachments if a["is_document"]],
            "inline_images": [a for a in attachments if a["is_inline_image"]]
        }
    
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
    
    def move_to_archive(self, email_id: str, folder_path: str = "Archive", create_path: bool = False) -> Dict[str, Any]:
        r"""Move an email to a folder (default: Archive).
        
        Supports both simple folder names and full paths with subfolders.
        Examples:
          - "Archive" (root-level folder)
          - "Inbox\~Zero\Growth Engineering" (subfolder path)
          - "\\David.Sant@harperjames.co.uk\Inbox\~Zero" (full path - mailbox prefix stripped)
        
        Args:
            email_id: Outlook EntryID of the email to archive
            folder_path: Folder name or path (default: "Archive")
            create_path: If True, create missing folders in the path (default: False)
            
        Returns:
            Dict with success status, new_id (EntryID changes after move),
            and folders_created list if any were created
        """
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            
            # Get root folder (the mailbox itself)
            inbox = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
            root = inbox.Parent  # Parent of Inbox = mailbox root
            
            # Parse folder path - handle full paths and relative paths
            # Strip leading backslashes and mailbox name if present
            clean_path = folder_path.strip("\\")
            first_part = clean_path.split("\\")[0] if "\\" in clean_path else ""
            if "@" in first_part:
                # Full path like "David.Sant@harperjames.co.uk\Inbox\~Zero" - strip mailbox
                clean_path = "\\".join(clean_path.split("\\")[1:])
            
            # Split into folder components
            path_parts = [p for p in clean_path.split("\\") if p]
            
            # Navigate to target folder, optionally creating missing folders
            target_folder = None
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
                        # Create the missing folder
                        current_folder = current_folder.Folders.Add(part)
                        folders_created.append(part)
                    else:
                        return {"success": False, "error": f"Folder '{part}' not found in path '{folder_path}'"}
            target_folder = current_folder
            
            if target_folder == root:
                return {"success": False, "error": f"Invalid folder path: {folder_path}"}
            
            # Move returns the new MailItem (with new EntryID)
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
    
    def batch_move_to_archive(self, email_ids: List[str], folder_path: str = "Archive", create_path: bool = False) -> Dict[str, Any]:
        r"""Move multiple emails to a folder (default: Archive).
        
        Supports folder paths with subfolders (see move_to_archive).
        
        Args:
            email_ids: List of Outlook EntryIDs
            folder_path: Folder name or path (default: "Archive")
            create_path: If True, create missing folders in the path (default: False)
            
        Returns:
            Dict with success/failed counts, details, and folders_created if any
        """
        results = {"success": 0, "failed": 0, "moved": [], "errors": [], "folders_created": []}
        
        for email_id in email_ids:
            result = self.move_to_archive(email_id, folder_path=folder_path, create_path=create_path)
            if result.get("success"):
                results["success"] += 1
                results["moved"].append({"old_id": email_id, "new_id": result.get("new_id")})
                # Track folders created (only first email will create them)
                if result.get("folders_created"):
                    for folder in result["folders_created"]:
                        if folder not in results["folders_created"]:
                            results["folders_created"].append(folder)
            else:
                results["failed"] += 1
                results["errors"].append({"id": email_id, "error": result.get("error")})
        
        # Remove empty folders_created list for cleaner output
        if not results["folders_created"]:
            del results["folders_created"]
        
        return results
    
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

    # Triage Status Methods (using Outlook Categories)
    # Categories are prefixed with "effi:" to avoid conflicts with user categories
    TRIAGE_CATEGORY_PREFIX = "effi:"
    TRIAGE_CATEGORIES = {
        "action": "effi:action",        # I need to do something
        "waiting": "effi:waiting",      # Ball in someone else's court
        "processed": "effi:processed",  # Dealt with, linked to matter
        "archived": "effi:archived",    # Reference only, no action needed
    }
    
    def set_triage_status(self, email_id: str, status: str) -> bool:
        """
        Set triage status on an email using Outlook categories.
        
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
        """
        Get triage status from an email's Outlook categories.
        
        Args:
            email_id: Outlook EntryID
            
        Returns:
            Status string ('action', 'waiting', 'processed', 'archived') or None if pending/not triaged
        """
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            categories = message.Categories or ""
            
            for status, category in self.TRIAGE_CATEGORIES.items():
                if category in categories:
                    return status
            return None  # No triage category = pending
        except:
            return None
    
    def clear_triage_status(self, email_id: str) -> bool:
        """
        Remove triage status from an email (reset to pending).
        
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
        """
        Set triage status on multiple emails.
        
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

    def get_pending_emails(
        self,
        days: int = 30,
        date_from: datetime = None,
        limit: int = 200,
        group_by_domain: bool = True,
    ) -> Dict[str, Any]:
        """
        Get inbound emails that haven't been triaged (no effi: category).
        
        Args:
            days: Days to look back (default 30)
            date_from: Start date (overrides days)
            limit: Maximum results
            group_by_domain: If True, group results by sender domain
            
        Returns:
            Dict with emails grouped by domain or flat list
        """
        self._ensure_connection()
        
        folder = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
        
        # Set date range
        if not date_from:
            date_from = datetime.now() - timedelta(days=days)
        
        date_str = date_from.strftime("%d/%m/%Y %H:%M")
        date_query = f"[ReceivedTime] >= '{date_str}'"
        
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)
        
        try:
            filtered = messages.Restrict(date_query)
        except:
            filtered = messages
        
        pending_emails = []
        
        for message in filtered:
            if len(pending_emails) >= limit:
                break
            
            try:
                # Check if message has any effi: triage categories
                categories = message.Categories or ""
                has_triage = any(cat.strip().startswith(self.TRIAGE_CATEGORY_PREFIX) 
                                for cat in categories.split(",") if cat.strip())
                
                if not has_triage:
                    email = self._message_to_email(message, folder.Name, "inbound")
                    if email:
                        pending_emails.append(email)
            except:
                continue
        
        if not group_by_domain:
            return {"emails": pending_emails, "total": len(pending_emails)}
        
        # Group by domain
        by_domain = {}
        for email in pending_emails:
            domain = email.domain or "(no domain)"
            if domain not in by_domain:
                by_domain[domain] = []
            by_domain[domain].append(email)
        
        # Sort domains by email count (highest first)
        sorted_domains = sorted(by_domain.items(), key=lambda x: len(x[1]), reverse=True)
        
        return {
            "domains": [
                {
                    "domain": domain,
                    "count": len(emails),
                    "emails": emails
                }
                for domain, emails in sorted_domains
            ],
            "total": len(pending_emails)
        }

    def get_domain_counts(
        self,
        days: int = 30,
        limit: Optional[int] = None,
        pending_only: bool = True,
    ) -> Dict[str, Any]:
        """
        Fast method to get domain counts from emails.
        
        Only extracts sender domain and subject - no full email conversion.
        Much faster than get_pending_emails for domain discovery.
        
        Args:
            days: Days to look back (default 30)
            limit: Maximum messages to scan (None = no limit)
            pending_only: If True, only count pending emails. If False, count all emails.
            
        Returns:
            Dict with domains, counts, and sample subjects
        """
        self._ensure_connection()
        
        folder = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
        
        date_from = datetime.now() - timedelta(days=days)
        date_str = date_from.strftime("%d/%m/%Y %H:%M")
        date_query = f"[ReceivedTime] >= '{date_str}'"
        
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)
        
        try:
            filtered = messages.Restrict(date_query)
        except:
            filtered = messages
        
        # Track domains with counts and sample subjects
        domain_data: Dict[str, Dict] = {}
        scanned = 0
        
        for message in filtered:
            if limit is not None and scanned >= limit:
                break
            scanned += 1
            
            try:
                # Check if message has any effi: triage categories
                if pending_only:
                    categories = message.Categories or ""
                    has_triage = any(cat.strip().startswith(self.TRIAGE_CATEGORY_PREFIX) 
                                    for cat in categories.split(",") if cat.strip())
                    if has_triage:
                        continue
                
                # Fast extraction - just sender and subject
                sender_email = self._get_sender_email(message)
                domain = self._extract_domain(sender_email)
                subject = message.Subject or "(No Subject)"
                received_time = message.ReceivedTime
                
                if domain not in domain_data:
                    domain_data[domain] = {"count": 0, "subjects": [], "latest": received_time}
                domain_data[domain]["count"] += 1
                # Track most recent email for this domain
                if received_time > domain_data[domain]["latest"]:
                    domain_data[domain]["latest"] = received_time
                if len(domain_data[domain]["subjects"]) < 3:
                    domain_data[domain]["subjects"].append(subject)
            except:
                continue
        
        # Sort by most recent email first
        sorted_domains = sorted(domain_data.items(), key=lambda x: x[1]["latest"], reverse=True)
        
        return {
            "domains": [
                {
                    "domain": domain,
                    "count": data["count"],
                    "sample_subjects": data["subjects"]
                }
                for domain, data in sorted_domains
            ],
            "total_scanned": scanned,
            "total_pending": sum(d["count"] for d in domain_data.values())
        }

    def get_pending_emails_from_domain(
        self,
        domain: str,
        days: int = 30,
        limit: int = 100,
    ) -> List[Email]:
        """
        Get pending (un-triaged) emails from a specific domain.
        
        Args:
            domain: Sender domain to filter
            days: Days to look back
            limit: Maximum results
            
        Returns:
            List of pending Email objects from that domain
        """
        self._ensure_connection()
        
        folder = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
        
        date_from = datetime.now() - timedelta(days=days)
        date_str = date_from.strftime("%d/%m/%Y %H:%M")
        
        # Use DASL for domain filter
        dasl_query = f"@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%@{domain}'"
        date_query = f"[ReceivedTime] >= '{date_str}'"
        
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)
        
        try:
            # Apply date filter first
            filtered_by_date = messages.Restrict(date_query)
            # Then DASL filter for domain
            filtered = filtered_by_date.Restrict(dasl_query)
        except:
            # Fallback
            filtered = messages.Restrict(date_query)
        
        pending_emails = []
        
        for message in filtered:
            if len(pending_emails) >= limit:
                break
            
            try:
                # Only include if no effi: category
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

    # DASL Query Support for Direct Outlook Searches
    # DASL property path for custom RecipientDomain field (PS_PUBLIC_STRINGS namespace)
    RECIPIENT_DOMAIN_PROP = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/RecipientDomain"

    def _build_query(
        self,
        sender_domain: str = None,
        sender_email: str = None,
        recipient_domain: str = None,
        recipient_email: str = None,
        subject_contains: str = None,
        body_contains: str = None,
        date_from: datetime = None,
        date_to: datetime = None,
    ) -> tuple:
        """Build query strings for Outlook Items.Restrict().
        
        Returns separate Jet and DASL queries. DASL cannot be applied to an
        already-restricted collection, so if we have DASL conditions, we include
        dates in the DASL query to make it a single query.
        
        Strategy:
        - If only dates: use Jet (indexed, fast)
        - If any DASL conditions: include dates in DASL query (single Restrict)
        - RecipientDomain uses DASL with full MAPI property path (supports LIKE)
        - Sender filters use DASL (urn:schemas:httpmail:fromemail) for Exchange compatibility
        
        Args:
            sender_domain: Filter by sender's domain (e.g., 'client.com')
            sender_email: Filter by exact sender email
            recipient_domain: Filter by recipient domain (uses custom RecipientDomain field)
            recipient_email: Filter by exact recipient email
            subject_contains: Subject line contains this text
            body_contains: Body contains this text
            date_from: Start date
            date_to: End date
            
        Returns:
            Tuple of (jet_query, dasl_query) - if dasl_query is set, jet_query will be None
        """
        dasl_conditions = []
        
        # Recipient domain uses custom RecipientDomain field (DASL with MAPI path)
        # This field is populated by VBA macro with semicolon-separated domains
        # DASL syntax requires single quotes for values
        if recipient_domain:
            dasl_conditions.append(f'"{self.RECIPIENT_DOMAIN_PROP}" LIKE \'%{recipient_domain}%\'')
        
        # Sender filters use DASL for Exchange compatibility
        # [SenderEmailAddress] in Jet may return Exchange DN instead of SMTP address
        if sender_email:
            dasl_conditions.append(f"\"urn:schemas:httpmail:fromemail\" LIKE '%{sender_email}%'")
        elif sender_domain:
            dasl_conditions.append(f"\"urn:schemas:httpmail:fromemail\" LIKE '%@{sender_domain}'")
        
        # Recipient email uses DASL (displayto has display names, but works for exact emails)
        if recipient_email:
            dasl_conditions.append(f"\"urn:schemas:httpmail:displayto\" LIKE '%{recipient_email}%'")
        
        # Subject filter (DASL)
        if subject_contains:
            dasl_conditions.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{subject_contains}%'")
        
        # Body filter (DASL)
        if body_contains:
            dasl_conditions.append(f"\"urn:schemas:httpmail:textdescription\" LIKE '%{body_contains}%'")
        
        # If we have DASL conditions, include dates in DASL (can't chain Restrict calls)
        if dasl_conditions:
            # Add dates to DASL query using urn:schemas:httpmail:datereceived
            if date_from:
                date_str = date_from.strftime("%d/%m/%Y %H:%M")
                dasl_conditions.append(f"\"urn:schemas:httpmail:datereceived\" >= '{date_str}'")
            if date_to:
                date_str = date_to.strftime("%d/%m/%Y %H:%M")
                dasl_conditions.append(f"\"urn:schemas:httpmail:datereceived\" <= '{date_str}'")
            
            dasl_query = "@SQL=" + " AND ".join(dasl_conditions)
            return (None, dasl_query)
        
        # No DASL conditions - use Jet for dates only (indexed, fast)
        jet_conditions = []
        if date_from:
            date_str = date_from.strftime("%d/%m/%Y %H:%M")
            jet_conditions.append(f"[ReceivedTime] >= '{date_str}'")
        if date_to:
            date_str = date_to.strftime("%d/%m/%Y %H:%M")
            jet_conditions.append(f"[ReceivedTime] <= '{date_str}'")
        
        jet_query = " AND ".join(jet_conditions) if jet_conditions else None
        return (jet_query, None)

    def search_outlook(
        self,
        sender_domain: str = None,
        sender_email: str = None,
        recipient_domain: str = None,
        recipient_email: str = None,
        subject_contains: str = None,
        body_contains: str = None,
        date_from: datetime = None,
        date_to: datetime = None,
        days: int = 30,
        folder: str = "Inbox",
        limit: int = 50,
    ) -> List[Email]:
        """Search Outlook directly with flexible filters.
        
        This searches Outlook directly without syncing to the database.
        Use for historical lookups or ad-hoc searches.
        
        Query strategy:
        - Jet query first (dates + RecipientDomain) for indexed performance
        - DASL query second (sender filters) for Exchange SMTP address compatibility
        
        Args:
            sender_domain: Filter by sender's domain
            sender_email: Filter by exact sender email
            recipient_domain: Filter by recipient domain
            recipient_email: Filter by exact recipient email
            subject_contains: Subject line contains text
            body_contains: Body contains text
            date_from: Start date (overrides days)
            date_to: End date
            days: Days to look back (default 30)
            folder: Outlook folder name ('Inbox', 'Sent Items', etc.)
            limit: Maximum results
            
        Returns:
            List of Email objects matching the filters
        """
        self._ensure_connection()
        
        # Determine folder - support paths like "Inbox\~Zero" or simple names
        direction = "inbound"
        folder_obj = None
        
        if folder.lower() in ["sent", "sent items"]:
            folder_obj = self._namespace.GetDefaultFolder(self.FOLDER_SENT)
            direction = "outbound"
        elif "\\" in folder or "/" in folder:
            # Subfolder path - navigate to it
            # Normalize to backslash
            folder_path = folder.replace("/", "\\")
            path_parts = [p for p in folder_path.split("\\") if p]
            
            if path_parts:
                # Start from the root folder (first part)
                first_part = path_parts[0].lower()
                if first_part in ["sent", "sent items"]:
                    folder_obj = self._namespace.GetDefaultFolder(self.FOLDER_SENT)
                    direction = "outbound"
                else:
                    folder_obj = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
                
                # Navigate to subfolders
                for part in path_parts[1:]:
                    found = False
                    for subfolder in folder_obj.Folders:
                        if subfolder.Name.lower() == part.lower():
                            folder_obj = subfolder
                            found = True
                            break
                    if not found:
                        # Subfolder not found - return empty
                        return []
        else:
            # Simple folder name
            folder_obj = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
        
        # Set date range
        if not date_from:
            date_from = datetime.now() - timedelta(days=days)
        
        # Build separate Jet and DASL queries (can't be mixed)
        jet_query, dasl_query = self._build_query(
            sender_domain=sender_domain,
            sender_email=sender_email,
            recipient_domain=recipient_domain,
            recipient_email=recipient_email,
            subject_contains=subject_contains,
            body_contains=body_contains,
            date_from=date_from,
            date_to=date_to,
        )
        
        results = []
        messages = folder_obj.Items
        messages.Sort("[ReceivedTime]", True)
        
        try:
            # Apply query - either Jet (dates only) or DASL (all conditions including dates)
            # Only one will be set based on _build_query logic
            if dasl_query:
                # DASL query includes dates - single Restrict call
                filtered = messages.Restrict(dasl_query)
            elif jet_query:
                # Jet query for dates only
                filtered = messages.Restrict(jet_query)
            else:
                # No filters
                filtered = messages
        except Exception as e:
            # Fallback to simple date filter
            date_str = date_from.strftime("%d/%m/%Y %H:%M")
            filtered = messages.Restrict(f"[ReceivedTime] >= '{date_str}'")
        
        for message in filtered:
            if len(results) >= limit:
                break
            
            try:
                email = self._message_to_email(message, folder_obj.Name, direction)
                if email:
                    results.append(email)
            except:
                continue
        
        return results

    def search_outlook_by_identifiers(
        self,
        domains: List[str],
        contact_emails: List[str] = None,
        days: int = 30,
        date_from: datetime = None,
        date_to: datetime = None,
        limit: int = 100,
    ) -> List[Email]:
        """Search Outlook for emails matching client domains/contact emails.
        
        Searches both Inbox (inbound) and Sent Items (outbound).
        
        Args:
            domains: List of domains to search for
            contact_emails: List of specific email addresses to search for
            days: Days to look back
            date_from: Start date (overrides days)
            date_to: End date
            limit: Maximum results
            
        Returns:
            List of matching Email objects
        """
        results = []
        contact_emails = contact_emails or []
        
        if not date_from:
            date_from = datetime.now() - timedelta(days=days)
        
        # Search Inbox for inbound emails
        for domain in domains:
            inbox_results = self.search_outlook(
                sender_domain=domain,
                date_from=date_from,
                date_to=date_to,
                folder="Inbox",
                limit=limit,
            )
            results.extend(inbox_results)
        
        # Search for contact emails (personal addresses)
        for email in contact_emails:
            inbox_results = self.search_outlook(
                sender_email=email,
                date_from=date_from,
                date_to=date_to,
                folder="Inbox",
                limit=limit,
            )
            results.extend(inbox_results)
        
        # Search Sent Items for outbound emails
        for domain in domains:
            sent_results = self.search_outlook(
                recipient_domain=domain,
                date_from=date_from,
                date_to=date_to,
                folder="Sent Items",
                limit=limit,
            )
            results.extend(sent_results)
        
        # Deduplicate by email ID and limit
        seen_ids = set()
        unique_results = []
        for email in results:
            if email.id not in seen_ids:
                seen_ids.add(email.id)
                unique_results.append(email)
                if len(unique_results) >= limit:
                    break
        
        return unique_results

    def get_email_full(self, email_id: str) -> Dict[str, Any]:
        """Get full email details by EntryID including body and attachments.
        
        Args:
            email_id: Outlook EntryID
            
        Returns:
            Dict with full email details including body and attachment metadata
        """
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            
            # Get body
            body = message.Body or ""
            html_body = message.HTMLBody or ""
            
            # Get attachments
            attachments = []
            for i in range(1, message.Attachments.Count + 1):
                att = message.Attachments.Item(i)
                attachments.append({
                    "name": att.FileName,
                    "size": att.Size,
                })
            
            # Get recipients
            recipients_to = self._extract_recipients(message, "To")
            recipients_cc = self._extract_recipients(message, "CC")
            
            return {
                "id": email_id,
                "subject": message.Subject or "(No Subject)",
                "sender_name": message.SenderName or "",
                "sender_email": self._get_sender_email(message),
                "received_time": message.ReceivedTime.isoformat() if hasattr(message.ReceivedTime, 'isoformat') else str(message.ReceivedTime),
                "body": body,
                "html_body": html_body,
                "recipients_to": recipients_to,
                "recipients_cc": recipients_cc,
                "attachments": attachments,
                "internet_message_id": self._get_internet_message_id(message),
                "conversation_id": getattr(message, 'ConversationID', None),
                "conversation_topic": getattr(message, 'ConversationTopic', None),
            }
        except Exception as e:
            raise Exception(f"Error retrieving email: {e}")

    def get_email_for_sync(self, email_id: str) -> Optional[Email]:
        """Get an email from Outlook ready for syncing to database.
        
        Args:
            email_id: Outlook EntryID
            
        Returns:
            Email object ready for database insertion
        """
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            
            # Determine direction based on folder
            folder_path = message.Parent.Name if hasattr(message, 'Parent') else "Inbox"
            direction = "outbound" if "Sent" in folder_path else "inbound"
            
            return self._message_to_email(message, folder_path, direction)
        except Exception as e:
            return None

    # =========================================================================
    # DMS (DMSforLegal) Methods - Read-only access to filed emails
    # Structure: \\DMSforLegal\_My Matters\{Client}\{Matter}\Emails
    # =========================================================================
    
    DMS_STORE_NAME = "DMSforLegal"
    DMS_ROOT_FOLDER = "_My Matters"
    DMS_EMAILS_FOLDER = "Emails"
    DMS_ADMIN_FOLDER = "Admin"
    
    def _get_dms_store(self):
        """Get the DMSforLegal Outlook store.
        
        Returns:
            Store object or None if not found
        """
        self._ensure_connection()
        
        try:
            for store in self._namespace.Stores:
                if store.DisplayName == self.DMS_STORE_NAME:
                    return store
        except Exception:
            pass
        return None
    
    def _get_folder_by_path(self, path: str):
        """Navigate to a folder by path within DMS store.
        
        Args:
            path: Backslash-separated path like "_My Matters\\Client\\Matter\\Emails"
            
        Returns:
            Folder object or None if not found
        """
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
        """List all client folders in DMS.
        
        Returns:
            Sorted list of client folder names
        """
        folder = self._get_folder_by_path(self.DMS_ROOT_FOLDER)
        if not folder:
            return []
        
        try:
            clients = [f.Name for f in folder.Folders]
            return sorted(clients)
        except Exception:
            return []
    
    def list_dms_matters(self, client: str) -> List[str]:
        """List all matter folders for a client in DMS.
        
        Args:
            client: Client folder name (exact match)
            
        Returns:
            Sorted list of matter folder names
        """
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
        """Get emails from a matter's Emails folder in DMS.
        
        Args:
            client: Client folder name
            matter: Matter folder name
            limit: Maximum emails to return (default 50)
            
        Returns:
            List of Email objects
        """
        path = f"{self.DMS_ROOT_FOLDER}\\{client}\\{matter}\\{self.DMS_EMAILS_FOLDER}"
        folder = self._get_folder_by_path(path)
        if not folder:
            return []
        
        results = []
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # Most recent first
            
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
        """Get emails from a matter's Admin folder in DMS.
        
        Args:
            client: Client folder name
            matter: Matter folder name
            limit: Maximum emails to return (default 50)
            
        Returns:
            List of Email objects
        """
        path = f"{self.DMS_ROOT_FOLDER}\\{client}\\{matter}\\{self.DMS_ADMIN_FOLDER}"
        folder = self._get_folder_by_path(path)
        if not folder:
            return []
        
        results = []
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # Most recent first
            
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
        """Search emails across DMS with optional filters.
        
        Args:
            client: Filter by client name (optional - searches all if not provided)
            matter: Filter by matter name (requires client)
            subject_contains: Filter by subject text
            date_from: Filter by start date
            date_to: Filter by end date
            limit: Maximum results (default 50)
            
        Returns:
            List of Email objects matching filters
        """
        results = []
        
        # Determine which clients to search
        if client:
            clients_to_search = [client]
        else:
            clients_to_search = self.list_dms_clients()
        
        for c in clients_to_search:
            if len(results) >= limit:
                break
            
            # Determine which matters to search
            if matter and client:
                matters_to_search = [matter]
            else:
                matters_to_search = self.list_dms_matters(c)
            
            for m in matters_to_search:
                if len(results) >= limit:
                    break
                
                # Get emails from this matter
                path = f"{self.DMS_ROOT_FOLDER}\\{c}\\{m}\\{self.DMS_EMAILS_FOLDER}"
                folder = self._get_folder_by_path(path)
                if not folder:
                    continue
                
                try:
                    items = folder.Items
                    items.Sort("[ReceivedTime]", True)
                    
                    # Apply date filter if provided
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
                            # Apply subject filter
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

    def file_email_to_dms(
        self,
        email_id: str,
        client_name: str,
        matter_name: str,
    ) -> Dict[str, Any]:
        """File an email to a DMS client/matter folder.
        
        Copies the email to the matter's Emails folder, adds "Filed" category
        to the original, and marks it as effi:processed.
        
        Args:
            email_id: EntryID of the email to file
            client_name: Client folder name in DMS
            matter_name: Matter folder name under the client
            
        Returns:
            Dict with:
            - success: bool
            - filed_entry_id: EntryID of the filed copy (if successful)
            - subject: Email subject
            - received_time: Email received time (ISO format)
            - filed_category: "Filed"
            - triage_status: "processed"
            - error: Error message (if failed)
        """
        self._ensure_connection()
        
        # Validate DMS folder exists
        dms_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}\\{self.DMS_EMAILS_FOLDER}"
        emails_folder = self._get_folder_by_path(dms_path)
        
        if not emails_folder:
            # Determine which part is missing for better error message
            client_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}"
            client_folder = self._get_folder_by_path(client_path)
            
            if not client_folder:
                return {
                    "success": False,
                    "error": f"Client '{client_name}' not found in DMS"
                }
            
            matter_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}"
            matter_folder = self._get_folder_by_path(matter_path)
            
            if not matter_folder:
                return {
                    "success": False,
                    "error": f"Matter '{matter_name}' not found for client '{client_name}'"
                }
            
            return {
                "success": False,
                "error": f"Emails folder not found for matter '{matter_name}'. Please create the Emails subfolder in DMS."
            }
        
        # Get the email to file
        try:
            message = self._namespace.GetItemFromID(email_id)
        except Exception as e:
            return {
                "success": False,
                "error": f"Email not found: {str(e)}"
            }
        
        # Copy email to DMS
        # Note: Copy().Move() across different store types can create empty shells
        # Saving the copy first forces content sync before the cross-store move
        try:
            copied = message.Copy()
            copied.Save()  # Force content materialization before cross-store move
            filed_message = copied.Move(emails_folder)
            filed_entry_id = filed_message.EntryID
        except Exception as e:
            return {
                "success": False,
                "error": f"Failed to copy email to DMS: {str(e)}"
            }
        
        # Add "Filed" category and effi:processed to original
        try:
            existing_categories = message.Categories or ""
            categories = [c.strip() for c in existing_categories.split(",") if c.strip()]
            
            # Remove any existing effi: triage categories
            categories = [c for c in categories if not c.startswith(self.TRIAGE_CATEGORY_PREFIX)]
            
            # Add Filed and effi:processed
            if "Filed" not in categories:
                categories.append("Filed")
            categories.append(self.TRIAGE_CATEGORIES["processed"])
            
            message.Categories = ", ".join(categories)
            message.Save()
        except Exception as e:
            # Email was filed but category update failed - still report success
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
        """File an admin email to a DMS client/matter Admin folder.
        
        Same as file_email_to_dms but files to Admin subfolder instead of Emails.
        Used for internal/system emails related to a matter (e.g. new project notifications).
        
        Args:
            email_id: EntryID of the email to file
            client_name: Client folder name in DMS
            matter_name: Matter folder name under the client
            
        Returns:
            Dict with success status, filed email details, or error message.
        """
        self._ensure_connection()
        
        # Build path to Admin folder
        dms_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}\\{self.DMS_ADMIN_FOLDER}"
        admin_folder = self._get_folder_by_path(dms_path)
        
        if not admin_folder:
            # Determine which part is missing for better error message
            client_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}"
            client_folder = self._get_folder_by_path(client_path)
            
            if not client_folder:
                return {
                    "success": False,
                    "error": f"Client '{client_name}' not found in DMS"
                }
            
            matter_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}"
            matter_folder = self._get_folder_by_path(matter_path)
            
            if not matter_folder:
                return {
                    "success": False,
                    "error": f"Matter '{matter_name}' not found for client '{client_name}'"
                }
            
            return {
                "success": False,
                "error": f"Admin folder not found for matter '{matter_name}'. Please create the Admin subfolder in DMS."
            }
        
        # Get the email to file
        try:
            message = self._namespace.GetItemFromID(email_id)
        except Exception as e:
            return {
                "success": False,
                "error": f"Email not found: {str(e)}"
            }
        
        # Copy email to DMS Admin folder
        # Note: Copy().Move() across different store types can create empty shells
        # Saving the copy first forces content sync before the cross-store move
        try:
            copied = message.Copy()
            copied.Save()  # Force content materialization before cross-store move
            filed_message = copied.Move(admin_folder)
            filed_entry_id = filed_message.EntryID
        except Exception as e:
            return {
                "success": False,
                "error": f"Failed to copy email to DMS Admin: {str(e)}"
            }
        
        # Add "Filed" category and effi:processed to original
        try:
            existing_categories = message.Categories or ""
            categories = [c.strip() for c in existing_categories.split(",") if c.strip()]
            
            # Remove any existing effi: triage categories
            categories = [c for c in categories if not c.startswith(self.TRIAGE_CATEGORY_PREFIX)]
            
            # Add Filed and effi:processed
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
        """File multiple emails to a DMS client/matter folder.
        
        Validates the DMS folder once upfront, then files each email.
        Continues processing if individual emails fail.
        
        Args:
            email_ids: List of EntryIDs to file
            client_name: Client folder name in DMS
            matter_name: Matter folder name under the client
            
        Returns:
            Dict with:
            - success: bool (True if folder validation passed)
            - filed_count: Number of emails successfully filed
            - failed_count: Number of emails that failed
            - filed_emails: List of {entry_id, subject, received_time}
            - failed_emails: List of {email_id, error}
            - error: Error message (if folder validation failed)
        """
        self._ensure_connection()
        
        # Handle empty list
        if not email_ids:
            return {
                "success": True,
                "filed_count": 0,
                "failed_count": 0,
                "filed_emails": [],
                "failed_emails": []
            }
        
        # Validate DMS folder exists upfront
        dms_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}\\{self.DMS_EMAILS_FOLDER}"
        emails_folder = self._get_folder_by_path(dms_path)
        
        if not emails_folder:
            # Determine which part is missing for better error message
            client_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}"
            client_folder = self._get_folder_by_path(client_path)
            
            if not client_folder:
                return {
                    "success": False,
                    "error": f"Client '{client_name}' not found in DMS"
                }
            
            matter_path = f"{self.DMS_ROOT_FOLDER}\\{client_name}\\{matter_name}"
            matter_folder = self._get_folder_by_path(matter_path)
            
            if not matter_folder:
                return {
                    "success": False,
                    "error": f"Matter '{matter_name}' not found for client '{client_name}'"
                }
            
            return {
                "success": False,
                "error": f"Emails folder not found for matter '{matter_name}'. Please create the Emails subfolder in DMS."
            }
        
        filed_emails = []
        failed_emails = []
        
        for email_id in email_ids:
            try:
                message = self._namespace.GetItemFromID(email_id)
                
                # Copy to DMS - Save before Move to prevent empty shell issue
                copied = message.Copy()
                copied.Save()  # Force content materialization before cross-store move
                filed_message = copied.Move(emails_folder)
                
                # Add categories to original
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
            "filed_emails": filed_emails,
            "failed_emails": failed_emails
        }
