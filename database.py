"""SQLite database for email storage and management."""

import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, List, Dict
from contextlib import contextmanager

from models import Email, Domain, Client, Matter, Counterparty, EmailCategory, TriageStatus, GENERIC_DOMAINS


class Database:
    """SQLite database manager for effi-mail."""
    
    def __init__(self, db_path: str = "effi_mail.db"):
        self.db_path = Path(db_path)
        self._init_database()
    
    @contextmanager
    def connection(self):
        """Context manager for database connections."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
    
    def _init_database(self):
        """Initialize database schema."""
        with self.connection() as conn:
            cursor = conn.cursor()
            
            # Emails table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS emails (
                    id TEXT PRIMARY KEY,
                    subject TEXT NOT NULL,
                    sender_name TEXT NOT NULL,
                    sender_email TEXT NOT NULL,
                    domain TEXT NOT NULL,
                    received_time TEXT NOT NULL,
                    body_preview TEXT DEFAULT '',
                    has_attachments INTEGER DEFAULT 0,
                    attachment_names TEXT DEFAULT '',
                    categories TEXT DEFAULT '',
                    conversation_id TEXT,
                    folder_path TEXT DEFAULT 'Inbox',
                    direction TEXT DEFAULT 'inbound',
                    recipients_to TEXT DEFAULT '[]',
                    recipients_cc TEXT DEFAULT '[]',
                    recipient_domains TEXT DEFAULT '',
                    internet_message_id TEXT,
                    triage_status TEXT DEFAULT 'pending',
                    client_id TEXT,
                    matter_id TEXT,
                    processed_at TEXT,
                    notes TEXT DEFAULT '',
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Migration: Add direction column if it doesn't exist
            cursor.execute("PRAGMA table_info(emails)")
            columns = [col[1] for col in cursor.fetchall()]
            if 'direction' not in columns:
                cursor.execute("ALTER TABLE emails ADD COLUMN direction TEXT DEFAULT 'inbound'")
            
            # Migration: Add recipients columns if they don't exist
            if 'recipients_to' not in columns:
                cursor.execute("ALTER TABLE emails ADD COLUMN recipients_to TEXT DEFAULT '[]'")
            if 'recipients_cc' not in columns:
                cursor.execute("ALTER TABLE emails ADD COLUMN recipients_cc TEXT DEFAULT '[]'")
            
            # Migration: Add recipient_domains column if it doesn't exist
            if 'recipient_domains' not in columns:
                cursor.execute("ALTER TABLE emails ADD COLUMN recipient_domains TEXT DEFAULT ''")
            
            # Migration: Add internet_message_id column if it doesn't exist
            if 'internet_message_id' not in columns:
                cursor.execute("ALTER TABLE emails ADD COLUMN internet_message_id TEXT")
            
            # Domains table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS domains (
                    name TEXT PRIMARY KEY,
                    category TEXT DEFAULT 'Uncategorized',
                    email_count INTEGER DEFAULT 0,
                    last_seen TEXT,
                    sample_senders TEXT DEFAULT ''
                )
            """)
            
            # Clients table (domains now come from effi-clients, not stored locally)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS clients (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    folder_path TEXT
                )
            """)
            
            # Matters table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS matters (
                    id TEXT PRIMARY KEY,
                    client_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    description TEXT DEFAULT '',
                    folder_path TEXT,
                    active INTEGER DEFAULT 1,
                    FOREIGN KEY (client_id) REFERENCES clients(id)
                )
            """)
            
            # Counterparties table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS counterparties (
                    id TEXT PRIMARY KEY,
                    matter_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    contact_name TEXT,
                    contact_email TEXT,
                    domains TEXT DEFAULT '',
                    notes TEXT DEFAULT '',
                    FOREIGN KEY (matter_id) REFERENCES matters(id)
                )
            """)
            
            # Contact emails table (for mapping individual emails to clients, esp. from generic domains)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS contact_emails (
                    email TEXT PRIMARY KEY,
                    client_id TEXT NOT NULL,
                    contact_name TEXT DEFAULT '',
                    FOREIGN KEY (client_id) REFERENCES clients(id)
                )
            """)
            
            # Sync metadata table (for tracking last sync time)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS sync_metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL,
                    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Create indexes
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_emails_domain ON emails(domain)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_emails_triage ON emails(triage_status)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_emails_received ON emails(received_time)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_emails_client ON emails(client_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_emails_sender_email ON emails(sender_email)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_emails_internet_message_id ON emails(internet_message_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_matters_client ON matters(client_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_contact_emails_client ON contact_emails(client_id)")
    
    # Sync metadata operations
    def get_last_sync_time(self) -> Optional[datetime]:
        """Get the timestamp of the last successful sync."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT value FROM sync_metadata WHERE key = 'last_sync_time'")
            row = cursor.fetchone()
            if row:
                return datetime.fromisoformat(row[0])
            return None
    
    def set_last_sync_time(self, sync_time: datetime) -> None:
        """Set the timestamp of the last successful sync."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO sync_metadata (key, value, updated_at)
                VALUES ('last_sync_time', ?, ?)
            """, (sync_time.isoformat(), datetime.now().isoformat()))
    
    # Email operations
    def upsert_email(self, email: Email) -> None:
        """Insert or update an email."""
        import json
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO emails (
                    id, subject, sender_name, sender_email, domain,
                    received_time, body_preview, has_attachments, attachment_names,
                    categories, conversation_id, folder_path, direction,
                    recipients_to, recipients_cc, recipient_domains, internet_message_id,
                    triage_status, client_id, matter_id, processed_at, notes, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                email.id, email.subject, email.sender_name, email.sender_email,
                email.domain, email.received_time.isoformat(),
                email.body_preview, int(email.has_attachments),
                ",".join(email.attachment_names), email.categories,
                email.conversation_id, email.folder_path, email.direction,
                json.dumps(email.recipients_to), json.dumps(email.recipients_cc),
                email.recipient_domains, email.internet_message_id,
                email.triage_status.value, email.client_id, email.matter_id,
                email.processed_at.isoformat() if email.processed_at else None,
                email.notes, datetime.now().isoformat()
            ))
    
    def get_email(self, email_id: str) -> Optional[Email]:
        """Get an email by ID."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM emails WHERE id = ?", (email_id,))
            row = cursor.fetchone()
            return self._row_to_email(row) if row else None
    
    def get_emails_by_status(self, status: TriageStatus, limit: int = 50, direction: str = None) -> List[Email]:
        """Get emails by triage status, optionally filtered by direction."""
        with self.connection() as conn:
            cursor = conn.cursor()
            if direction:
                cursor.execute(
                    "SELECT * FROM emails WHERE triage_status = ? AND direction = ? ORDER BY received_time DESC LIMIT ?",
                    (status.value, direction, limit)
                )
            else:
                cursor.execute(
                    "SELECT * FROM emails WHERE triage_status = ? ORDER BY received_time DESC LIMIT ?",
                    (status.value, limit)
                )
            return [self._row_to_email(row) for row in cursor.fetchall()]
    
    def get_emails_by_domain(self, domain: str, limit: int = 50) -> List[Email]:
        """Get emails from a specific domain."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM emails WHERE domain = ? ORDER BY received_time DESC LIMIT ?",
                (domain, limit)
            )
            return [self._row_to_email(row) for row in cursor.fetchall()]
    
    def get_sent_emails(self, client_id: str = None, days: int = None, limit: int = 50) -> List[Email]:
        """Get sent (outbound) emails, optionally filtered by client or date range."""
        with self.connection() as conn:
            cursor = conn.cursor()
            query = "SELECT * FROM emails WHERE direction = 'outbound'"
            params = []
            
            if client_id:
                query += " AND client_id = ?"
                params.append(client_id)
            
            if days:
                cutoff = (datetime.now() - timedelta(days=days)).isoformat()
                query += " AND received_time >= ?"
                params.append(cutoff)
            
            query += " ORDER BY received_time DESC LIMIT ?"
            params.append(limit)
            
            cursor.execute(query, params)
            return [self._row_to_email(row) for row in cursor.fetchall()]
    
    def get_conversation_thread(self, subject: str = None, client_id: str = None, 
                                 conversation_id: str = None, participant_email: str = None,
                                 limit: int = 50) -> List[Email]:
        """Get both inbound and outbound emails for a conversation.
        
        Args:
            subject: Subject line to match (partial)
            client_id: Client ID to scope search
            conversation_id: Outlook conversation ID
            participant_email: Email address to find in From, To, or CC
            limit: Maximum emails to return
        """
        with self.connection() as conn:
            cursor = conn.cursor()
            
            if conversation_id:
                cursor.execute(
                    "SELECT * FROM emails WHERE conversation_id = ? ORDER BY received_time ASC LIMIT ?",
                    (conversation_id, limit)
                )
            elif participant_email:
                # Find all emails where this person is involved (From, To, or CC)
                email_lower = participant_email.lower()
                cursor.execute("""
                    SELECT * FROM emails 
                    WHERE LOWER(sender_email) = ?
                       OR LOWER(recipients_to) LIKE ?
                       OR LOWER(recipients_cc) LIKE ?
                    ORDER BY received_time ASC LIMIT ?
                """, (email_lower, f'%"{email_lower}"%', f'%"{email_lower}"%', limit))
            elif subject and client_id:
                # Match by normalized subject and client
                cursor.execute(
                    """SELECT * FROM emails 
                       WHERE client_id = ? AND subject LIKE ? 
                       ORDER BY received_time ASC LIMIT ?""",
                    (client_id, f"%{subject}%", limit)
                )
            elif subject:
                # Match by subject pattern (handles RE:, FW: prefixes)
                base_subject = subject.lstrip("RE: ").lstrip("FW: ").lstrip("Re: ").lstrip("Fw: ")
                cursor.execute(
                    "SELECT * FROM emails WHERE subject LIKE ? ORDER BY received_time ASC LIMIT ?",
                    (f"%{base_subject}%", limit)
                )
            else:
                return []
            
            return [self._row_to_email(row) for row in cursor.fetchall()]
    
    def get_emails_with_recipient(self, email_address: str, days: int = 30, limit: int = 100) -> List[Email]:
        """Get emails where the given address appears in From, To, or CC.
        
        Args:
            email_address: Email address to search for
            days: Number of days to look back
            limit: Maximum emails to return
        
        Returns:
            List of emails involving this recipient
        """
        with self.connection() as conn:
            cursor = conn.cursor()
            cutoff = (datetime.now() - timedelta(days=days)).isoformat()
            email_lower = email_address.lower()
            
            # Search in sender_email, recipients_to (JSON), and recipients_cc (JSON)
            # Use LIKE for JSON arrays since SQLite's JSON functions may not be available
            cursor.execute("""
                SELECT * FROM emails 
                WHERE received_time >= ?
                AND (
                    LOWER(sender_email) = ?
                    OR LOWER(recipients_to) LIKE ?
                    OR LOWER(recipients_cc) LIKE ?
                )
                ORDER BY received_time DESC
                LIMIT ?
            """, (cutoff, email_lower, f'%"{email_lower}"%', f'%"{email_lower}"%', limit))
            
            return [self._row_to_email(row) for row in cursor.fetchall()]
    
    def get_last_contact_by_client(self) -> List[dict]:
        """Get last inbound and outbound email dates per client."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT client_id, direction, MAX(received_time) as last_time
                FROM emails
                WHERE client_id IS NOT NULL
                GROUP BY client_id, direction
            """)
            
            result = {}
            for row in cursor.fetchall():
                client_id = row["client_id"]
                if client_id not in result:
                    result[client_id] = {"client_id": client_id, "last_inbound": None, "last_outbound": None}
                
                if row["direction"] == "inbound":
                    result[client_id]["last_inbound"] = row["last_time"]
                else:
                    result[client_id]["last_outbound"] = row["last_time"]
            
            return list(result.values())
    
    def update_email_triage(self, email_id: str, status: TriageStatus,
                            client_id: Optional[str] = None,
                            matter_id: Optional[str] = None,
                            notes: str = "") -> None:
        """Update email triage status."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE emails SET
                    triage_status = ?,
                    client_id = COALESCE(?, client_id),
                    matter_id = COALESCE(?, matter_id),
                    notes = CASE WHEN ? != '' THEN ? ELSE notes END,
                    processed_at = ?,
                    updated_at = ?
                WHERE id = ?
            """, (
                status.value, client_id, matter_id, notes, notes,
                datetime.now().isoformat(), datetime.now().isoformat(), email_id
            ))
    
    def _row_to_email(self, row: sqlite3.Row) -> Email:
        """Convert database row to Email object."""
        import json
        
        # Parse recipients JSON, with fallback for older rows
        recipients_to = []
        recipients_cc = []
        if "recipients_to" in row.keys() and row["recipients_to"]:
            try:
                recipients_to = json.loads(row["recipients_to"])
            except:
                recipients_to = []
        if "recipients_cc" in row.keys() and row["recipients_cc"]:
            try:
                recipients_cc = json.loads(row["recipients_cc"])
            except:
                recipients_cc = []
        
        return Email(
            id=row["id"],
            subject=row["subject"],
            sender_name=row["sender_name"],
            sender_email=row["sender_email"],
            domain=row["domain"],
            received_time=datetime.fromisoformat(row["received_time"]),
            body_preview=row["body_preview"],
            has_attachments=bool(row["has_attachments"]),
            attachment_names=row["attachment_names"].split(",") if row["attachment_names"] else [],
            categories=row["categories"],
            conversation_id=row["conversation_id"],
            folder_path=row["folder_path"],
            direction=row["direction"] if "direction" in row.keys() else "inbound",
            recipients_to=recipients_to,
            recipients_cc=recipients_cc,
            recipient_domains=row["recipient_domains"] if "recipient_domains" in row.keys() else "",
            internet_message_id=row["internet_message_id"] if "internet_message_id" in row.keys() else None,
            triage_status=TriageStatus(row["triage_status"]),
            client_id=row["client_id"],
            matter_id=row["matter_id"],
            processed_at=datetime.fromisoformat(row["processed_at"]) if row["processed_at"] else None,
            notes=row["notes"]
        )
    
    # Domain operations
    def upsert_domain(self, domain: Domain) -> None:
        """Insert or update a domain."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO domains (name, category, email_count, last_seen, sample_senders)
                VALUES (?, ?, ?, ?, ?)
            """, (
                domain.name, domain.category.value, domain.email_count,
                domain.last_seen.isoformat() if domain.last_seen else None,
                ",".join(domain.sample_senders[:5])
            ))
    
    def get_domain(self, name: str) -> Optional[Domain]:
        """Get a domain by name."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM domains WHERE name = ?", (name,))
            row = cursor.fetchone()
            return self._row_to_domain(row) if row else None
    
    def get_domains_by_category(self, category: EmailCategory) -> List[Domain]:
        """Get domains by category."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM domains WHERE category = ? ORDER BY email_count DESC",
                (category.value,)
            )
            return [self._row_to_domain(row) for row in cursor.fetchall()]
    
    def get_uncategorized_domains(self) -> List[Domain]:
        """Get domains that haven't been categorized."""
        return self.get_domains_by_category(EmailCategory.UNCATEGORIZED)
    
    def update_domain_category(self, name: str, category: EmailCategory) -> None:
        """Update domain category. Generic domains (gmail, etc) are protected."""
        # Don't categorize generic domains - they require contact_email matching
        if name.lower() in GENERIC_DOMAINS:
            return
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE domains SET category = ? WHERE name = ?",
                (category.value, name)
            )
    
    def _row_to_domain(self, row: sqlite3.Row) -> Domain:
        """Convert database row to Domain object."""
        return Domain(
            name=row["name"],
            category=EmailCategory(row["category"]),
            email_count=row["email_count"],
            last_seen=datetime.fromisoformat(row["last_seen"]) if row["last_seen"] else None,
            sample_senders=row["sample_senders"].split(",") if row["sample_senders"] else []
        )
    
    # Client operations
    def upsert_client(self, client: Client) -> None:
        """Insert or update a client."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO clients (id, name, folder_path)
                VALUES (?, ?, ?)
            """, (client.id, client.name, client.folder_path))
    
    def get_client(self, client_id: str) -> Optional[Client]:
        """Get a client by ID."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM clients WHERE id = ?", (client_id,))
            row = cursor.fetchone()
            return self._row_to_client(row) if row else None
    
    def get_all_clients(self) -> List[Client]:
        """Get all clients."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM clients ORDER BY name")
            return [self._row_to_client(row) for row in cursor.fetchall()]
    
    def _row_to_client(self, row: sqlite3.Row) -> Client:
        """Convert database row to Client object."""
        return Client(
            id=row["id"],
            name=row["name"],
            folder_path=row["folder_path"]
        )
    
    # Matter operations
    def upsert_matter(self, matter: Matter) -> None:
        """Insert or update a matter."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO matters (id, client_id, name, description, folder_path, active)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (matter.id, matter.client_id, matter.name, matter.description,
                  matter.folder_path, int(matter.active)))
    
    def get_matter(self, matter_id: str) -> Optional[Matter]:
        """Get a matter by ID."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM matters WHERE id = ?", (matter_id,))
            row = cursor.fetchone()
            return self._row_to_matter(row) if row else None
    
    def get_matters_for_client(self, client_id: str, active_only: bool = True) -> List[Matter]:
        """Get matters for a client."""
        with self.connection() as conn:
            cursor = conn.cursor()
            if active_only:
                cursor.execute(
                    "SELECT * FROM matters WHERE client_id = ? AND active = 1 ORDER BY name",
                    (client_id,)
                )
            else:
                cursor.execute(
                    "SELECT * FROM matters WHERE client_id = ? ORDER BY name",
                    (client_id,)
                )
            return [self._row_to_matter(row) for row in cursor.fetchall()]
    
    def _row_to_matter(self, row: sqlite3.Row) -> Matter:
        """Convert database row to Matter object."""
        return Matter(
            id=row["id"],
            client_id=row["client_id"],
            name=row["name"],
            description=row["description"],
            folder_path=row["folder_path"],
            active=bool(row["active"])
        )
    
    # Counterparty operations
    def upsert_counterparty(self, counterparty: Counterparty) -> None:
        """Insert or update a counterparty."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO counterparties (id, matter_id, name, contact_name, contact_email, domains, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (counterparty.id, counterparty.matter_id, counterparty.name,
                  counterparty.contact_name, counterparty.contact_email,
                  ",".join(counterparty.domains), counterparty.notes))
    
    def get_counterparty(self, counterparty_id: str) -> Optional[Counterparty]:
        """Get a counterparty by ID."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM counterparties WHERE id = ?", (counterparty_id,))
            row = cursor.fetchone()
            return self._row_to_counterparty(row) if row else None
    
    def get_counterparties_for_matter(self, matter_id: str) -> List[Counterparty]:
        """Get counterparties for a matter."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM counterparties WHERE matter_id = ? ORDER BY name",
                (matter_id,)
            )
            return [self._row_to_counterparty(row) for row in cursor.fetchall()]
    
    def get_counterparty_by_domain(self, domain: str) -> Optional[Counterparty]:
        """Find a counterparty by one of their email domains."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM counterparties WHERE domains LIKE ?",
                (f"%{domain}%",)
            )
            row = cursor.fetchone()
            return self._row_to_counterparty(row) if row else None
    
    def delete_counterparty(self, counterparty_id: str) -> None:
        """Delete a counterparty."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM counterparties WHERE id = ?", (counterparty_id,))
    
    def _row_to_counterparty(self, row: sqlite3.Row) -> Counterparty:
        """Convert database row to Counterparty object."""
        return Counterparty(
            id=row["id"],
            matter_id=row["matter_id"],
            name=row["name"],
            contact_name=row["contact_name"],
            contact_email=row["contact_email"],
            domains=row["domains"].split(",") if row["domains"] else [],
            notes=row["notes"] or ""
        )
    
    # Statistics
    def get_triage_stats(self) -> Dict[str, int]:
        """Get email counts by triage status."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT triage_status, COUNT(*) as count
                FROM emails GROUP BY triage_status
            """)
            return {row["triage_status"]: row["count"] for row in cursor.fetchall()}
    
    def get_domain_stats(self) -> Dict[str, int]:
        """Get email counts by domain category."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT d.category, COUNT(e.id) as count
                FROM domains d
                LEFT JOIN emails e ON e.domain = d.name
                GROUP BY d.category
            """)
            return {row["category"]: row["count"] for row in cursor.fetchall()}


    # Contact email operations
    def upsert_contact_email(self, email: str, client_id: str, contact_name: str = "") -> None:
        """Add or update a contact email to client mapping."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO contact_emails (email, client_id, contact_name)
                VALUES (?, ?, ?)
            """, (email.lower(), client_id, contact_name))

    def get_client_by_contact_email(self, email: str) -> Optional[Client]:
        """Find a client by a specific contact email address."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT client_id FROM contact_emails WHERE email = ?",
                (email.lower(),)
            )
            row = cursor.fetchone()
            if row:
                return self.get_client(row["client_id"])
            return None

    def get_contact_emails_for_client(self, client_id: str) -> List[str]:
        """Get all contact emails for a client."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT email FROM contact_emails WHERE client_id = ?",
                (client_id,)
            )
            return [row["email"] for row in cursor.fetchall()]

    def link_emails_to_clients(self, client_domains_map: Dict[str, List[str]] = None) -> Dict[str, int]:
        """
        Link emails to clients based on domain matching.
        
        Args:
            client_domains_map: Dict mapping client_id to list of domains.
                               If not provided, only contact_emails linking will be done.
        
        For each client's domains, update emails with matching domain.
        For generic domains (gmail, outlook, etc.), check contact_emails table.
        
        Returns:
            Dict with counts of emails linked by domain and by contact_email
        """
        stats = {"by_domain": 0, "by_contact_email": 0}
        
        with self.connection() as conn:
            cursor = conn.cursor()
            
            # Link by domain if client_domains_map provided
            if client_domains_map:
                for client_id, domains in client_domains_map.items():
                    for domain in domains:
                        if domain and domain not in GENERIC_DOMAINS:
                            # Update emails with matching domain that don't have a client_id
                            cursor.execute("""
                                UPDATE emails 
                                SET client_id = ?, updated_at = ?
                                WHERE domain = ? 
                                AND client_id IS NULL
                            """, (client_id, datetime.now().isoformat(), domain))
                            stats["by_domain"] += cursor.rowcount
            
            # For generic domains, check contact_emails table
            # Get emails from generic domains without client_id
            generic_domains_tuple = tuple(GENERIC_DOMAINS)
            placeholders = ",".join(["?"] * len(generic_domains_tuple))
            
            cursor.execute(f"""
                SELECT id, sender_email FROM emails 
                WHERE domain IN ({placeholders})
                AND client_id IS NULL
            """, generic_domains_tuple)
            
            generic_emails = cursor.fetchall()
            
            for row in generic_emails:
                email_id = row["id"]
                sender_email = row["sender_email"].lower()
                
                # Check if sender_email is in contact_emails
                cursor.execute(
                    "SELECT client_id FROM contact_emails WHERE email = ?",
                    (sender_email,)
                )
                contact_row = cursor.fetchone()
                
                if contact_row:
                    cursor.execute("""
                        UPDATE emails 
                        SET client_id = ?, updated_at = ?
                        WHERE id = ?
                    """, (contact_row["client_id"], datetime.now().isoformat(), email_id))
                    stats["by_contact_email"] += 1
            
            conn.commit()
        
        return stats

    # Client-centric search operations
    def get_email_by_internet_message_id(self, internet_message_id: str) -> Optional[Email]:
        """Get an email by its permanent internet_message_id (RFC2822 Message-ID)."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM emails WHERE internet_message_id = ?", (internet_message_id,))
            row = cursor.fetchone()
            return self._row_to_email(row) if row else None

    # NOTE: get_client_identifiers() removed - use effi_work_client.get_client_identifiers_from_effi_work() instead
    # Client domains now come from effi-work as the single source of truth

    def search_emails_by_recipient_domain(self, domain: str, limit: int = 100) -> List[Email]:
        """Search for emails sent to a specific recipient domain.
        
        Args:
            domain: The recipient domain to search for
            limit: Maximum emails to return
            
        Returns:
            List of emails where recipient_domains contains the domain
        """
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM emails 
                WHERE recipient_domains LIKE ?
                ORDER BY received_time DESC
                LIMIT ?
            """, (f"%{domain}%", limit))
            return [self._row_to_email(row) for row in cursor.fetchall()]

    def search_emails_by_client(
        self,
        domains: List[str],
        contact_emails: List[str] = None,
        include_inbound: bool = True,
        include_outbound: bool = True,
        include_cc: bool = True,
        days: int = 30,
        date_from: str = None,
        date_to: str = None,
        limit: int = 100,
    ) -> List[Email]:
        """Search for all emails related to a client by their domains and contact emails.
        
        Matches emails where:
        - Sender domain matches any client domain (inbound)
        - Sender email matches any contact email (inbound from personal addresses)
        - Any To/CC recipient domain matches (outbound)
        - Any CC recipient matches client domain (cc'd)
        
        Args:
            domains: List of client email domains
            contact_emails: List of personal/contact emails associated with client
            include_inbound: Include emails FROM client domains/contacts
            include_outbound: Include emails TO client domains/contacts
            include_cc: Include emails where client appears in CC
            days: Days to look back (default 30)
            date_from: Start date (YYYY-MM-DD), overrides days if provided
            date_to: End date (YYYY-MM-DD)
            limit: Maximum results
            
        Returns:
            List of matching emails
        """
        contact_emails = contact_emails or []
        
        if not domains and not contact_emails:
            return []
        
        # Build date filter - use full datetime for proper comparison
        if date_from:
            cutoff_start = date_from if " " in date_from else f"{date_from} 00:00:00"
        else:
            cutoff_start = (datetime.now() - timedelta(days=days)).isoformat()
        
        if date_to:
            cutoff_end = date_to if " " in date_to else f"{date_to} 23:59:59"
        else:
            cutoff_end = datetime.now().isoformat()
        
        # Build the query conditions
        conditions = []
        params = []
        
        # Inbound: sender domain matches
        if include_inbound and domains:
            domain_placeholders = ",".join(["?"] * len(domains))
            conditions.append(f"(direction = 'inbound' AND domain IN ({domain_placeholders}))")
            params.extend(domains)
        
        # Inbound from contact emails (personal addresses)
        if include_inbound and contact_emails:
            email_placeholders = ",".join(["?"] * len(contact_emails))
            conditions.append(f"(direction = 'inbound' AND LOWER(sender_email) IN ({email_placeholders}))")
            params.extend([e.lower() for e in contact_emails])
        
        # Outbound: recipient domain matches
        if include_outbound and domains:
            outbound_conditions = []
            for domain in domains:
                outbound_conditions.append("recipient_domains LIKE ?")
                params.append(f"%{domain}%")
            conditions.append(f"(direction = 'outbound' AND ({' OR '.join(outbound_conditions)}))")
        
        # CC: client domain appears in CC
        if include_cc and domains:
            cc_conditions = []
            for domain in domains:
                cc_conditions.append("LOWER(recipients_cc) LIKE ?")
                params.append(f"%{domain}%")
            conditions.append(f"({' OR '.join(cc_conditions)})")
        
        if not conditions:
            return []
        
        query = f"""
            SELECT * FROM emails 
            WHERE ({' OR '.join(conditions)})
            AND received_time >= ?
            AND received_time <= ?
            ORDER BY received_time DESC
            LIMIT ?
        """
        params.extend([cutoff_start, cutoff_end, limit])
        
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [self._row_to_email(row) for row in cursor.fetchall()]
