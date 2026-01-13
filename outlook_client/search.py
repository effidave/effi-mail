"""Search client for Outlook DASL queries and searches."""

from datetime import datetime, timedelta
from typing import List, Optional, Tuple

from outlook_client.base import BaseOutlookClient
from models import Email


class SearchClient(BaseOutlookClient):
    """Client for Outlook search operations using DASL queries."""
    
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
    ) -> Tuple[Optional[str], Optional[str]]:
        """Build query strings for Outlook Items.Restrict().
        
        Returns separate Jet and DASL queries. DASL cannot be applied to an
        already-restricted collection, so if we have DASL conditions, we include
        dates in the DASL query to make it a single query.
        """
        dasl_conditions = []
        
        if recipient_domain:
            dasl_conditions.append(f'"{self.RECIPIENT_DOMAIN_PROP}" LIKE \'%{recipient_domain}%\'')
        
        if sender_email:
            dasl_conditions.append(f"\"urn:schemas:httpmail:fromemail\" LIKE '%{sender_email}%'")
        elif sender_domain:
            dasl_conditions.append(f"\"urn:schemas:httpmail:fromemail\" LIKE '%@{sender_domain}'")
        
        if recipient_email:
            dasl_conditions.append(f"\"urn:schemas:httpmail:displayto\" LIKE '%{recipient_email}%'")
        
        if subject_contains:
            dasl_conditions.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{subject_contains}%'")
        
        if body_contains:
            dasl_conditions.append(f"\"urn:schemas:httpmail:textdescription\" LIKE '%{body_contains}%'")
        
        if dasl_conditions:
            if date_from:
                date_str = date_from.strftime("%d/%m/%Y %H:%M")
                dasl_conditions.append(f"\"urn:schemas:httpmail:datereceived\" >= '{date_str}'")
            if date_to:
                date_str = date_to.strftime("%d/%m/%Y %H:%M")
                dasl_conditions.append(f"\"urn:schemas:httpmail:datereceived\" <= '{date_str}'")
            
            dasl_query = "@SQL=" + " AND ".join(dasl_conditions)
            return (None, dasl_query)
        
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
        """Search Outlook directly with flexible filters."""
        self._ensure_connection()
        
        direction = "inbound"
        folder_obj = None
        
        if folder.lower() in ["sent", "sent items"]:
            folder_obj = self._namespace.GetDefaultFolder(self.FOLDER_SENT)
            direction = "outbound"
        elif "\\" in folder or "/" in folder:
            folder_path = folder.replace("/", "\\")
            path_parts = [p for p in folder_path.split("\\") if p]
            
            if path_parts:
                first_part = path_parts[0].lower()
                if first_part in ["sent", "sent items"]:
                    folder_obj = self._namespace.GetDefaultFolder(self.FOLDER_SENT)
                    direction = "outbound"
                else:
                    folder_obj = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
                
                for part in path_parts[1:]:
                    found = False
                    for subfolder in folder_obj.Folders:
                        if subfolder.Name.lower() == part.lower():
                            folder_obj = subfolder
                            found = True
                            break
                    if not found:
                        return []
        else:
            folder_obj = self._namespace.GetDefaultFolder(self.FOLDER_INBOX)
        
        if not date_from:
            date_from = datetime.now() - timedelta(days=days)
        
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
            if dasl_query:
                filtered = messages.Restrict(dasl_query)
            elif jet_query:
                filtered = messages.Restrict(jet_query)
            else:
                filtered = messages
        except Exception:
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
        """Search Outlook for emails matching client domains/contact emails."""
        results = []
        contact_emails = contact_emails or []
        
        if not date_from:
            date_from = datetime.now() - timedelta(days=days)
        
        for domain in domains:
            inbox_results = self.search_outlook(
                sender_domain=domain,
                date_from=date_from,
                date_to=date_to,
                folder="Inbox",
                limit=limit,
            )
            results.extend(inbox_results)
        
        for email in contact_emails:
            inbox_results = self.search_outlook(
                sender_email=email,
                date_from=date_from,
                date_to=date_to,
                folder="Inbox",
                limit=limit,
            )
            results.extend(inbox_results)
        
        for domain in domains:
            sent_results = self.search_outlook(
                recipient_domain=domain,
                date_from=date_from,
                date_to=date_to,
                folder="Sent Items",
                limit=limit,
            )
            results.extend(sent_results)
        
        seen_ids = set()
        unique_results = []
        for email in results:
            if email.id not in seen_ids:
                seen_ids.add(email.id)
                unique_results.append(email)
                if len(unique_results) >= limit:
                    break
        
        return unique_results
