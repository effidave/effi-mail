"""Retrieval client for fetching emails from Outlook."""

from datetime import datetime, timedelta
from typing import List, Optional, Generator, Dict, Any
import os
import mimetypes

from outlook_client.base import BaseOutlookClient
from models import Email


class RetrievalClient(BaseOutlookClient):
    """Client for email retrieval operations."""
    
    def _set_recipient_domains(self, limit: int = 200) -> dict:
        """Set RecipientDomain custom property on recent Sent Items."""
        self._ensure_connection()
        
        sent_folder = self._namespace.GetDefaultFolder(self.FOLDER_SENT)
        messages = sent_folder.Items
        messages.Sort("[SentOn]", True)
        
        prop_path = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/RecipientDomain"
        processed = 0
        updated = 0
        
        for message in messages:
            if processed >= limit:
                break
            processed += 1
            
            try:
                try:
                    existing = message.PropertyAccessor.GetProperty(prop_path)
                    if existing:
                        continue
                except:
                    pass
                
                domains = set()
                for recipient in message.Recipients:
                    try:
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
                continue
        
        return {"processed": processed, "updated": updated}
    
    def get_emails(self, days: int = 7, folder_id: int = None, 
                   exclude_categories: List[str] = None, direction: str = "inbound",
                   since_time: datetime = None) -> Generator[Email, None, None]:
        """Fetch emails from Outlook.
        
        Args:
            days: Number of days to look back
            folder_id: Outlook folder constant (default: Inbox)
            exclude_categories: Categories to exclude
            direction: 'inbound' or 'outbound'
            since_time: Fetch emails after this timestamp
        
        Yields:
            Email objects
        """
        self._ensure_connection()
        
        if folder_id is None:
            folder_id = self.FOLDER_INBOX
        
        folder = self._namespace.GetDefaultFolder(folder_id)
        folder_path = folder.Name
        
        if since_time:
            date_cutoff = since_time
        else:
            date_cutoff = datetime.now() - timedelta(days=days)
        date_str = date_cutoff.strftime("%d/%m/%Y %H:%M")
        filter_str = f"[ReceivedTime] >= '{date_str}'"
        
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)
        filtered = messages.Restrict(filter_str)
        
        exclude_categories = exclude_categories or []
        
        for message in filtered:
            try:
                msg_categories = message.Categories or ""
                if any(cat in msg_categories for cat in exclude_categories):
                    continue
                
                recipient_domain = None
                if direction == "outbound":
                    recipient_domain = self._get_primary_recipient_domain(message)
                
                email = self._message_to_email(message, folder_path, direction, recipient_domain)
                if email:
                    yield email
            except:
                continue
    
    def get_emails_by_conversation_id(
        self,
        conversation_id: str,
        include_sent: bool = True,
        include_dms: bool = False,
        limit: int = 50,
        conversation_topic: str = None
    ) -> List[Email]:
        """Get all emails matching a ConversationID across folders."""
        self._ensure_connection()
        
        results = []
        
        if not conversation_topic:
            return results
        
        escaped_topic = conversation_topic.replace("'", "''")
        filter_str = f"[ConversationTopic] = '{escaped_topic}'"
        
        def search_folder(folder, direction: str, folder_path: str = None):
            nonlocal results
            if folder_path is None:
                folder_path = folder.Name
            try:
                messages = folder.Items.Restrict(filter_str)
                for message in messages:
                    if len(results) >= limit:
                        return
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
        
        inbox = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
        search_folder(inbox, "inbound")
        
        if include_sent and len(results) < limit:
            sent = self._namespace.GetDefaultFolder(self.FOLDER_SENT)
            search_folder(sent, "outbound")
        
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
    
    def get_email_body(self, email_id: str, max_length: int = 10000) -> str:
        """Get full email body by EntryID."""
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
    
    def get_email_full(self, email_id: str) -> Dict[str, Any]:
        """Get full email details by EntryID including body and attachments."""
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            
            body = message.Body or ""
            
            try:
                html_body = message.HTMLBody or ""
            except Exception:
                html_body = ""
            
            attachments = []
            try:
                for i in range(1, message.Attachments.Count + 1):
                    att = message.Attachments.Item(i)
                    attachments.append({
                        "name": att.FileName,
                        "size": att.Size,
                    })
            except Exception:
                pass
            
            recipients_to = self._extract_recipients(message, "To")
            recipients_cc = self._extract_recipients(message, "CC")
            
            message_class = getattr(message, 'MessageClass', 'IPM.Note')
            
            result = {
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
                "message_class": message_class,
            }
            
            if message_class.startswith('IPM.Schedule.Meeting'):
                result["is_meeting_request"] = True
                try:
                    result["start_time"] = message.Start.isoformat() if hasattr(message, 'Start') and message.Start else None
                    result["end_time"] = message.End.isoformat() if hasattr(message, 'End') and message.End else None
                    result["location"] = getattr(message, 'Location', None)
                except Exception:
                    pass
            
            return result
        except Exception as e:
            raise Exception(f"Error retrieving email: {e}")
    
    def get_email_for_sync(self, email_id: str) -> Optional[Email]:
        """Get an email from Outlook ready for syncing."""
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
            folder_path = message.Parent.Name if hasattr(message, 'Parent') else "Inbox"
            direction = "outbound" if "Sent" in folder_path else "inbound"
            return self._message_to_email(message, folder_path, direction)
        except Exception as e:
            return None
    
    def get_pending_emails(
        self,
        days: int = 30,
        date_from: datetime = None,
        limit: int = 200,
        group_by_domain: bool = True,
    ) -> Dict[str, Any]:
        """Get inbound emails that haven't been triaged (no effi: category)."""
        self._ensure_connection()
        
        folder = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
        
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
        
        by_domain = {}
        for email in pending_emails:
            domain = email.domain or "(no domain)"
            if domain not in by_domain:
                by_domain[domain] = []
            by_domain[domain].append(email)
        
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
        """Fast method to get domain counts from emails."""
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
        
        domain_data: Dict[str, Dict] = {}
        scanned = 0
        
        for message in filtered:
            if limit is not None and scanned >= limit:
                break
            scanned += 1
            
            try:
                if pending_only:
                    categories = message.Categories or ""
                    has_triage = any(cat.strip().startswith(self.TRIAGE_CATEGORY_PREFIX) 
                                    for cat in categories.split(",") if cat.strip())
                    if has_triage:
                        continue
                
                sender_email = self._get_sender_email(message)
                domain = self._extract_domain(sender_email)
                subject = message.Subject or "(No Subject)"
                received_time = message.ReceivedTime
                
                if domain not in domain_data:
                    domain_data[domain] = {"count": 0, "subjects": [], "latest": received_time}
                domain_data[domain]["count"] += 1
                if received_time > domain_data[domain]["latest"]:
                    domain_data[domain]["latest"] = received_time
                if len(domain_data[domain]["subjects"]) < 3:
                    domain_data[domain]["subjects"].append(subject)
            except:
                continue
        
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
    
    def download_attachment(self, email_id: str, attachment_name: str, 
                           save_path: Optional[str] = None) -> Dict[str, Any]:
        """Download an attachment from an email and save it to disk."""
        self._ensure_connection()
        
        try:
            message = self._namespace.GetItemFromID(email_id)
        except Exception as e:
            return {"success": False, "error": f"Email not found: {e}"}
        
        attachment = None
        for i in range(1, message.Attachments.Count + 1):
            att = message.Attachments.Item(i)
            if att.FileName == attachment_name:
                attachment = att
                break
        
        if not attachment:
            available = [message.Attachments.Item(i).FileName 
                        for i in range(1, message.Attachments.Count + 1)]
            return {
                "success": False, 
                "error": f"Attachment '{attachment_name}' not found",
                "available_attachments": available
            }
        
        if not save_path:
            try:
                sender = message.SenderEmailAddress or ""
                is_sent = sender.lower().endswith("harperjames.co.uk") or "@harperjames" in sender.lower()
                
                if is_sent and message.Recipients.Count > 0:
                    try:
                        recip = message.Recipients.Item(1)
                        recip_addr = recip.Address
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
            
            domain = "".join(c if c.isalnum() or c in '.-_' else '_' for c in domain)
            base_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "attachments")
            save_path = os.path.join(base_dir, domain, date_str, attachment_name)
        
        save_path = os.path.abspath(save_path)
        
        try:
            os.makedirs(os.path.dirname(save_path), exist_ok=True)
        except Exception as e:
            return {"success": False, "error": f"Failed to create directory: {e}"}
        
        try:
            attachment.SaveAsFile(save_path)
            file_size = os.path.getsize(save_path)
            
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
        """List all attachments for an email with details."""
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
