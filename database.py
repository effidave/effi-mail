"""SQLite database for email storage and management."""

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict
from contextlib import contextmanager

from models import Email, Domain, Client, Matter, EmailCategory, TriageStatus


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
                    triage_status TEXT DEFAULT 'pending',
                    client_id TEXT,
                    matter_id TEXT,
                    processed_at TEXT,
                    notes TEXT DEFAULT '',
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
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
            
            # Clients table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS clients (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    domains TEXT DEFAULT '',
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
            
            # Create indexes
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_emails_domain ON emails(domain)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_emails_triage ON emails(triage_status)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_emails_received ON emails(received_time)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_emails_client ON emails(client_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_matters_client ON matters(client_id)")
    
    # Email operations
    def upsert_email(self, email: Email) -> None:
        """Insert or update an email."""
        with self.connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO emails (
                    id, subject, sender_name, sender_email, domain,
                    received_time, body_preview, has_attachments, attachment_names,
                    categories, conversation_id, folder_path, triage_status,
                    client_id, matter_id, processed_at, notes, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                email.id, email.subject, email.sender_name, email.sender_email,
                email.domain, email.received_time.isoformat(),
                email.body_preview, int(email.has_attachments),
                ",".join(email.attachment_names), email.categories,
                email.conversation_id, email.folder_path,
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
    
    def get_emails_by_status(self, status: TriageStatus, limit: int = 50) -> List[Email]:
        """Get emails by triage status."""
        with self.connection() as conn:
            cursor = conn.cursor()
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
        """Update domain category."""
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
                INSERT OR REPLACE INTO clients (id, name, domains, folder_path)
                VALUES (?, ?, ?, ?)
            """, (client.id, client.name, ",".join(client.domains), client.folder_path))
    
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
            domains=row["domains"].split(",") if row["domains"] else [],
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
