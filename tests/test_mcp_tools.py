"""Comprehensive tests for MCP server tools."""

import pytest
import json
import asyncio
from datetime import datetime, timedelta
from unittest.mock import Mock, MagicMock, patch, AsyncMock
import tempfile
import os

# Import the modules under test
from models import Email, Domain, Client, Matter, Counterparty, EmailCategory, TriageStatus
from database import Database


# ============================================================================
# Fixtures
# ============================================================================

@pytest.fixture
def temp_db():
    """Create a temporary database for testing."""
    with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as f:
        db_path = f.name
    
    db = Database(db_path)
    yield db
    
    # Cleanup
    try:
        os.unlink(db_path)
    except:
        pass


@pytest.fixture
def sample_email():
    """Create a sample email for testing."""
    return Email(
        id="test-email-001",
        subject="Test Subject",
        sender_name="John Doe",
        sender_email="john@example.com",
        domain="example.com",
        received_time=datetime.now() - timedelta(hours=1),
        body_preview="This is a test email preview.",
        has_attachments=True,
        attachment_names=["doc.pdf", "image.png"],
        categories="",
        conversation_id="conv-001",
        folder_path="Inbox",
        triage_status=TriageStatus.PENDING,
    )


@pytest.fixture
def sample_emails(sample_email):
    """Create multiple sample emails."""
    emails = [sample_email]
    
    # Add more emails with different domains
    for i in range(2, 6):
        emails.append(Email(
            id=f"test-email-{i:03d}",
            subject=f"Test Subject {i}",
            sender_name=f"Sender {i}",
            sender_email=f"sender{i}@domain{i}.com",
            domain=f"domain{i}.com",
            received_time=datetime.now() - timedelta(hours=i),
            body_preview=f"Preview for email {i}",
            has_attachments=i % 2 == 0,
            attachment_names=[],
            categories="",
            triage_status=TriageStatus.PENDING,
        ))
    
    return emails


@pytest.fixture
def sample_domain():
    """Create a sample domain for testing."""
    return Domain(
        name="example.com",
        category=EmailCategory.CLIENT,
        email_count=5,
        last_seen=datetime.now(),
        sample_senders=["John Doe", "Jane Smith"],
    )


@pytest.fixture
def sample_client():
    """Create a sample client for testing."""
    return Client(
        id="acme-corp",
        name="Acme Corporation",
        domains=["acme.com", "acme.co.uk"],
        folder_path="/clients/acme",
    )


@pytest.fixture
def sample_matter():
    """Create a sample matter for testing."""
    return Matter(
        id="acme-contract-2024",
        client_id="acme-corp",
        name="Contract Review 2024",
        description="Annual contract review",
        folder_path="/clients/acme/contracts/2024",
        active=True,
    )


@pytest.fixture
def sample_counterparty():
    """Create a sample counterparty for testing."""
    return Counterparty(
        id="bigco-cp",
        matter_id="acme-contract-2024",
        name="Big Corporation Ltd",
        contact_name="Jane Smith",
        contact_email="jane.smith@bigcorp.com",
        domains=["bigcorp.com", "bigcorp.co.uk"],
        notes="Primary contact for contract negotiations",
    )


# ============================================================================
# Database Tests
# ============================================================================

class TestDatabase:
    """Tests for the Database class."""
    
    def test_database_initialization(self, temp_db):
        """Test database creates tables on init."""
        # Check tables exist by trying operations
        emails = temp_db.get_emails_by_status(TriageStatus.PENDING)
        assert emails == []
    
    def test_upsert_and_get_email(self, temp_db, sample_email):
        """Test inserting and retrieving an email."""
        temp_db.upsert_email(sample_email)
        
        retrieved = temp_db.get_email(sample_email.id)
        assert retrieved is not None
        assert retrieved.id == sample_email.id
        assert retrieved.subject == sample_email.subject
        assert retrieved.sender_email == sample_email.sender_email
        assert retrieved.domain == sample_email.domain
        assert retrieved.triage_status == TriageStatus.PENDING
    
    def test_get_emails_by_status(self, temp_db, sample_emails):
        """Test filtering emails by triage status."""
        # Insert emails
        for email in sample_emails:
            temp_db.upsert_email(email)
        
        # All should be pending
        pending = temp_db.get_emails_by_status(TriageStatus.PENDING)
        assert len(pending) == len(sample_emails)
        
        # Update one to processed
        temp_db.update_email_triage(sample_emails[0].id, TriageStatus.PROCESSED)
        
        pending = temp_db.get_emails_by_status(TriageStatus.PENDING)
        assert len(pending) == len(sample_emails) - 1
        
        processed = temp_db.get_emails_by_status(TriageStatus.PROCESSED)
        assert len(processed) == 1
    
    def test_get_emails_by_domain(self, temp_db, sample_emails):
        """Test filtering emails by domain."""
        for email in sample_emails:
            temp_db.upsert_email(email)
        
        emails = temp_db.get_emails_by_domain("example.com")
        assert len(emails) == 1
        assert emails[0].domain == "example.com"
    
    def test_update_email_triage(self, temp_db, sample_email):
        """Test updating email triage status."""
        temp_db.upsert_email(sample_email)
        
        temp_db.update_email_triage(
            sample_email.id,
            TriageStatus.PROCESSED,
            client_id="client-001",
            matter_id="matter-001",
            notes="Reviewed and completed"
        )
        
        updated = temp_db.get_email(sample_email.id)
        assert updated.triage_status == TriageStatus.PROCESSED
        assert updated.client_id == "client-001"
        assert updated.matter_id == "matter-001"
        assert updated.notes == "Reviewed and completed"
        assert updated.processed_at is not None
    
    def test_upsert_and_get_domain(self, temp_db, sample_domain):
        """Test inserting and retrieving a domain."""
        temp_db.upsert_domain(sample_domain)
        
        retrieved = temp_db.get_domain(sample_domain.name)
        assert retrieved is not None
        assert retrieved.name == sample_domain.name
        assert retrieved.category == EmailCategory.CLIENT
        assert retrieved.email_count == 5
    
    def test_get_domains_by_category(self, temp_db):
        """Test filtering domains by category."""
        domains = [
            Domain(name="client1.com", category=EmailCategory.CLIENT),
            Domain(name="client2.com", category=EmailCategory.CLIENT),
            Domain(name="internal.com", category=EmailCategory.INTERNAL),
            Domain(name="marketing.com", category=EmailCategory.MARKETING),
        ]
        
        for domain in domains:
            temp_db.upsert_domain(domain)
        
        client_domains = temp_db.get_domains_by_category(EmailCategory.CLIENT)
        assert len(client_domains) == 2
        
        internal_domains = temp_db.get_domains_by_category(EmailCategory.INTERNAL)
        assert len(internal_domains) == 1
    
    def test_get_uncategorized_domains(self, temp_db):
        """Test getting uncategorized domains."""
        domains = [
            Domain(name="known.com", category=EmailCategory.CLIENT),
            Domain(name="unknown1.com", category=EmailCategory.UNCATEGORIZED),
            Domain(name="unknown2.com", category=EmailCategory.UNCATEGORIZED),
        ]
        
        for domain in domains:
            temp_db.upsert_domain(domain)
        
        uncategorized = temp_db.get_uncategorized_domains()
        assert len(uncategorized) == 2
    
    def test_update_domain_category(self, temp_db, sample_domain):
        """Test updating domain category."""
        sample_domain.category = EmailCategory.UNCATEGORIZED
        temp_db.upsert_domain(sample_domain)
        
        temp_db.update_domain_category(sample_domain.name, EmailCategory.MARKETING)
        
        updated = temp_db.get_domain(sample_domain.name)
        assert updated.category == EmailCategory.MARKETING
    
    def test_upsert_and_get_client(self, temp_db, sample_client):
        """Test inserting and retrieving a client."""
        temp_db.upsert_client(sample_client)
        
        retrieved = temp_db.get_client(sample_client.id)
        assert retrieved is not None
        assert retrieved.id == sample_client.id
        assert retrieved.name == sample_client.name
        assert "acme.com" in retrieved.domains
    
    def test_get_all_clients(self, temp_db):
        """Test getting all clients."""
        clients = [
            Client(id="client1", name="Client One"),
            Client(id="client2", name="Client Two"),
            Client(id="client3", name="Client Three"),
        ]
        
        for client in clients:
            temp_db.upsert_client(client)
        
        all_clients = temp_db.get_all_clients()
        assert len(all_clients) == 3
    
    def test_upsert_and_get_matter(self, temp_db, sample_client, sample_matter):
        """Test inserting and retrieving a matter."""
        temp_db.upsert_client(sample_client)
        temp_db.upsert_matter(sample_matter)
        
        retrieved = temp_db.get_matter(sample_matter.id)
        assert retrieved is not None
        assert retrieved.id == sample_matter.id
        assert retrieved.client_id == sample_matter.client_id
        assert retrieved.name == sample_matter.name
    
    def test_get_matters_for_client(self, temp_db, sample_client):
        """Test getting matters for a specific client."""
        temp_db.upsert_client(sample_client)
        
        matters = [
            Matter(id="matter1", client_id=sample_client.id, name="Matter 1", active=True),
            Matter(id="matter2", client_id=sample_client.id, name="Matter 2", active=True),
            Matter(id="matter3", client_id=sample_client.id, name="Matter 3", active=False),
        ]
        
        for matter in matters:
            temp_db.upsert_matter(matter)
        
        # Active only (default)
        active_matters = temp_db.get_matters_for_client(sample_client.id)
        assert len(active_matters) == 2
        
        # All matters
        all_matters = temp_db.get_matters_for_client(sample_client.id, active_only=False)
        assert len(all_matters) == 3
    
    def test_get_triage_stats(self, temp_db, sample_emails):
        """Test getting triage statistics."""
        for email in sample_emails:
            temp_db.upsert_email(email)
        
        # Update some statuses
        temp_db.update_email_triage(sample_emails[0].id, TriageStatus.PROCESSED)
        temp_db.update_email_triage(sample_emails[1].id, TriageStatus.ARCHIVED)
        
        stats = temp_db.get_triage_stats()
        assert stats.get("processed", 0) == 1
        assert stats.get("archived", 0) == 1
        assert stats.get("pending", 0) == 3
    
    def test_get_domain_stats(self, temp_db, sample_emails):
        """Test getting domain statistics."""
        # Add emails and domains
        for email in sample_emails:
            temp_db.upsert_email(email)
            temp_db.upsert_domain(Domain(
                name=email.domain,
                category=EmailCategory.CLIENT if email.domain == "example.com" else EmailCategory.UNCATEGORIZED
            ))
        
        stats = temp_db.get_domain_stats()
        assert "Client" in stats or "Uncategorized" in stats
    
    def test_upsert_and_get_counterparty(self, temp_db, sample_client, sample_matter, sample_counterparty):
        """Test inserting and retrieving a counterparty."""
        temp_db.upsert_client(sample_client)
        temp_db.upsert_matter(sample_matter)
        temp_db.upsert_counterparty(sample_counterparty)
        
        retrieved = temp_db.get_counterparty(sample_counterparty.id)
        assert retrieved is not None
        assert retrieved.id == sample_counterparty.id
        assert retrieved.name == sample_counterparty.name
        assert retrieved.matter_id == sample_counterparty.matter_id
        assert "bigcorp.com" in retrieved.domains
    
    def test_get_counterparties_for_matter(self, temp_db, sample_client, sample_matter):
        """Test getting counterparties for a specific matter."""
        temp_db.upsert_client(sample_client)
        temp_db.upsert_matter(sample_matter)
        
        counterparties = [
            Counterparty(id="cp1", matter_id=sample_matter.id, name="Counterparty 1"),
            Counterparty(id="cp2", matter_id=sample_matter.id, name="Counterparty 2"),
            Counterparty(id="cp3", matter_id="other-matter", name="Counterparty 3"),
        ]
        
        for cp in counterparties:
            temp_db.upsert_counterparty(cp)
        
        matter_cps = temp_db.get_counterparties_for_matter(sample_matter.id)
        assert len(matter_cps) == 2
    
    def test_get_counterparty_by_domain(self, temp_db, sample_client, sample_matter, sample_counterparty):
        """Test finding counterparty by email domain."""
        temp_db.upsert_client(sample_client)
        temp_db.upsert_matter(sample_matter)
        temp_db.upsert_counterparty(sample_counterparty)
        
        found = temp_db.get_counterparty_by_domain("bigcorp.com")
        assert found is not None
        assert found.id == sample_counterparty.id
        
        not_found = temp_db.get_counterparty_by_domain("unknown.com")
        assert not_found is None
    
    def test_delete_counterparty(self, temp_db, sample_client, sample_matter, sample_counterparty):
        """Test deleting a counterparty."""
        temp_db.upsert_client(sample_client)
        temp_db.upsert_matter(sample_matter)
        temp_db.upsert_counterparty(sample_counterparty)
        
        # Verify it exists
        assert temp_db.get_counterparty(sample_counterparty.id) is not None
        
        # Delete it
        temp_db.delete_counterparty(sample_counterparty.id)
        
        # Verify it's gone
        assert temp_db.get_counterparty(sample_counterparty.id) is None
    
    def test_sync_metadata_get_set(self, temp_db):
        """Test getting and setting last sync time."""
        # Initially no sync time
        assert temp_db.get_last_sync_time() is None
        
        # Set a sync time
        sync_time = datetime.now()
        temp_db.set_last_sync_time(sync_time)
        
        # Retrieve it
        retrieved = temp_db.get_last_sync_time()
        assert retrieved is not None
        # Compare within a second (datetime serialization may lose microseconds)
        assert abs((retrieved - sync_time).total_seconds()) < 1
        
        # Update it
        new_sync_time = datetime.now() + timedelta(hours=1)
        temp_db.set_last_sync_time(new_sync_time)
        
        retrieved2 = temp_db.get_last_sync_time()
        assert abs((retrieved2 - new_sync_time).total_seconds()) < 1


# ============================================================================
# MCP Tool Handler Tests
# ============================================================================

class TestMCPToolHandlers:
    """Tests for MCP tool handler functions."""
    
    @pytest.fixture
    def mock_db(self):
        """Create a mock database."""
        return Mock(spec=Database)
    
    @pytest.fixture
    def mock_outlook(self):
        """Create a mock Outlook client."""
        mock = Mock()
        mock.sync_emails_to_db = Mock(return_value={"new": 10, "updated": 5, "domains_updated": 8})
        mock.get_email_body = Mock(return_value="Full email body content here...")
        return mock
    
    @pytest.mark.asyncio
    async def test_sync_emails_tool(self, mock_db, mock_outlook, sample_emails):
        """Test sync_emails tool handler."""
        with patch('mcp_server.db', mock_db), patch('mcp_server.outlook', mock_outlook):
            from mcp_server import call_tool
            
            result = await call_tool("sync_emails", {"days": 7})
            
            assert len(result) == 1
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert "stats" in data
            mock_outlook.sync_emails_to_db.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_sync_emails_with_custom_days(self, mock_db, mock_outlook):
        """Test sync_emails with custom days parameter."""
        with patch('mcp_server.db', mock_db), patch('mcp_server.outlook', mock_outlook):
            from mcp_server import call_tool
            
            result = await call_tool("sync_emails", {"days": 14, "exclude_unfocused": False})
            
            mock_outlook.sync_emails_to_db.assert_called_with(days=14, exclude_categories=[])
    
    @pytest.mark.asyncio
    async def test_get_pending_emails_tool(self, mock_db, sample_emails):
        """Test get_pending_emails tool handler."""
        mock_db.get_emails_by_status.return_value = sample_emails
        mock_db.get_domain.return_value = Domain(name="example.com", category=EmailCategory.CLIENT)
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("get_pending_emails", {"limit": 50})
            
            assert len(result) == 1
            data = json.loads(result[0].text)
            assert "total_pending" in data
            assert "domains" in data
    
    @pytest.mark.asyncio
    async def test_get_pending_emails_with_category_filter(self, mock_db, sample_emails):
        """Test get_pending_emails with category filter."""
        mock_db.get_emails_by_status.return_value = sample_emails
        mock_db.get_domain.return_value = Domain(name="example.com", category=EmailCategory.CLIENT)
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("get_pending_emails", {
                "limit": 50,
                "category_filter": "Client"
            })
            
            data = json.loads(result[0].text)
            assert "domains" in data
    
    @pytest.mark.asyncio
    async def test_get_emails_by_domain_tool(self, mock_db, sample_emails):
        """Test get_emails_by_domain tool handler."""
        mock_db.get_emails_by_domain.return_value = [sample_emails[0]]
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("get_emails_by_domain", {
                "domain": "example.com",
                "limit": 20
            })
            
            data = json.loads(result[0].text)
            assert data["domain"] == "example.com"
            assert data["count"] == 1
            assert len(data["emails"]) == 1
    
    @pytest.mark.asyncio
    async def test_get_email_content_tool(self, mock_db, mock_outlook, sample_email):
        """Test get_email_content tool handler."""
        mock_db.get_email.return_value = sample_email
        mock_outlook.get_email_body.return_value = "Full email body text..."
        
        with patch('mcp_server.db', mock_db), patch('mcp_server.outlook', mock_outlook):
            from mcp_server import call_tool
            
            result = await call_tool("get_email_content", {
                "email_id": sample_email.id,
                "max_length": 5000
            })
            
            data = json.loads(result[0].text)
            assert data["subject"] == sample_email.subject
            assert "body" in data
    
    @pytest.mark.asyncio
    async def test_get_email_content_not_found(self, mock_db, mock_outlook):
        """Test get_email_content when email not found."""
        mock_db.get_email.return_value = None
        
        with patch('mcp_server.db', mock_db), patch('mcp_server.outlook', mock_outlook):
            from mcp_server import call_tool
            
            result = await call_tool("get_email_content", {
                "email_id": "nonexistent"
            })
            
            data = json.loads(result[0].text)
            assert "error" in data
    
    @pytest.mark.asyncio
    async def test_triage_email_tool(self, mock_db):
        """Test triage_email tool handler."""
        mock_db.update_email_triage = Mock()
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("triage_email", {
                "email_id": "test-001",
                "status": "processed",
                "client_id": "client-001",
                "notes": "Reviewed"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["status"] == "processed"
            mock_db.update_email_triage.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_batch_triage_tool(self, mock_db):
        """Test batch_triage tool handler."""
        mock_db.update_email_triage = Mock()
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            email_ids = ["email-001", "email-002", "email-003"]
            result = await call_tool("batch_triage", {
                "email_ids": email_ids,
                "status": "archived"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["count"] == 3
            assert mock_db.update_email_triage.call_count == 3
    
    @pytest.mark.asyncio
    async def test_batch_archive_domain_tool(self, mock_db, sample_emails):
        """Test batch_archive_domain tool handler."""
        # All emails are pending
        mock_db.get_emails_by_domain.return_value = sample_emails
        mock_db.update_email_triage = Mock()
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("batch_archive_domain", {
                "domain": "example.com"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["domain"] == "example.com"
            assert data["archived_count"] == len(sample_emails)
    
    @pytest.mark.asyncio
    async def test_get_uncategorized_domains_tool(self, mock_db):
        """Test get_uncategorized_domains tool handler."""
        uncategorized = [
            Domain(name="unknown1.com", category=EmailCategory.UNCATEGORIZED, email_count=5),
            Domain(name="unknown2.com", category=EmailCategory.UNCATEGORIZED, email_count=3),
        ]
        mock_db.get_uncategorized_domains.return_value = uncategorized
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("get_uncategorized_domains", {"limit": 20})
            
            data = json.loads(result[0].text)
            assert data["count"] == 2
            assert len(data["domains"]) == 2
    
    @pytest.mark.asyncio
    async def test_categorize_domain_tool(self, mock_db):
        """Test categorize_domain tool handler."""
        mock_db.update_domain_category = Mock()
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("categorize_domain", {
                "domain": "newclient.com",
                "category": "Client"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["domain"] == "newclient.com"
            assert data["category"] == "Client"
    
    @pytest.mark.asyncio
    async def test_get_domain_summary_tool(self, mock_db):
        """Test get_domain_summary tool handler."""
        mock_db.get_domains_by_category.return_value = [
            Domain(name="client.com", category=EmailCategory.CLIENT, email_count=10),
        ]
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("get_domain_summary", {})
            
            data = json.loads(result[0].text)
            # Should have all category keys
            assert "Client" in data or "Internal" in data
    
    @pytest.mark.asyncio
    async def test_create_client_tool(self, mock_db):
        """Test create_client tool handler."""
        mock_db.upsert_client = Mock()
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("create_client", {
                "id": "new-client",
                "name": "New Client Inc",
                "domains": ["newclient.com"],
                "folder_path": "/clients/new-client"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["client_id"] == "new-client"
            mock_db.upsert_client.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_create_matter_tool(self, mock_db):
        """Test create_matter tool handler."""
        mock_db.upsert_matter = Mock()
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("create_matter", {
                "id": "matter-001",
                "client_id": "client-001",
                "name": "New Matter",
                "description": "Description here"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["matter_id"] == "matter-001"
            mock_db.upsert_matter.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_list_clients_tool(self, mock_db, sample_client, sample_matter):
        """Test list_clients tool handler."""
        mock_db.get_all_clients.return_value = [sample_client]
        mock_db.get_matters_for_client.return_value = [sample_matter]
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("list_clients", {})
            
            data = json.loads(result[0].text)
            assert "clients" in data
            assert len(data["clients"]) == 1
            assert data["clients"][0]["id"] == sample_client.id
    
    @pytest.mark.asyncio
    async def test_get_triage_stats_tool(self, mock_db):
        """Test get_triage_stats tool handler."""
        mock_db.get_triage_stats.return_value = {
            "pending": 10,
            "processed": 5,
            "archived": 3
        }
        mock_db.get_domain_stats.return_value = {
            "Client": 8,
            "Marketing": 7,
            "Internal": 3
        }
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("get_triage_stats", {})
            
            data = json.loads(result[0].text)
            assert "by_triage_status" in data
            assert "by_domain_category" in data
    
    @pytest.mark.asyncio
    async def test_create_counterparty_tool(self, mock_db):
        """Test create_counterparty tool handler."""
        mock_db.upsert_counterparty = Mock()
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("create_counterparty", {
                "id": "cp-001",
                "matter_id": "matter-001",
                "name": "Opposing Corp",
                "contact_name": "John Smith",
                "contact_email": "john@opposing.com",
                "domains": ["opposing.com"],
                "notes": "Main counterparty"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["counterparty_id"] == "cp-001"
            assert data["matter_id"] == "matter-001"
            mock_db.upsert_counterparty.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_get_counterparties_tool(self, mock_db, sample_counterparty):
        """Test get_counterparties tool handler."""
        mock_db.get_counterparties_for_matter.return_value = [sample_counterparty]
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("get_counterparties", {
                "matter_id": "acme-contract-2024"
            })
            
            data = json.loads(result[0].text)
            assert data["matter_id"] == "acme-contract-2024"
            assert data["count"] == 1
            assert len(data["counterparties"]) == 1
            assert data["counterparties"][0]["name"] == sample_counterparty.name
    
    @pytest.mark.asyncio
    async def test_find_counterparty_by_domain_found(self, mock_db, sample_counterparty):
        """Test find_counterparty_by_domain when found."""
        mock_db.get_counterparty_by_domain.return_value = sample_counterparty
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("find_counterparty_by_domain", {
                "domain": "bigcorp.com"
            })
            
            data = json.loads(result[0].text)
            assert data["found"] is True
            assert data["counterparty"]["name"] == sample_counterparty.name
    
    @pytest.mark.asyncio
    async def test_find_counterparty_by_domain_not_found(self, mock_db):
        """Test find_counterparty_by_domain when not found."""
        mock_db.get_counterparty_by_domain.return_value = None
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("find_counterparty_by_domain", {
                "domain": "unknown.com"
            })
            
            data = json.loads(result[0].text)
            assert data["found"] is False
    
    @pytest.mark.asyncio
    async def test_delete_counterparty_tool(self, mock_db):
        """Test delete_counterparty tool handler."""
        mock_db.delete_counterparty = Mock()
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("delete_counterparty", {
                "counterparty_id": "cp-001"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["deleted_id"] == "cp-001"
            mock_db.delete_counterparty.assert_called_once_with("cp-001")
    
    @pytest.mark.asyncio
    async def test_unknown_tool_error(self, mock_db):
        """Test handling of unknown tool name."""
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("nonexistent_tool", {})
            
            data = json.loads(result[0].text)
            assert "error" in data
            assert "Unknown tool" in data["error"]
    
    @pytest.mark.asyncio
    async def test_tool_exception_handling(self, mock_db):
        """Test that exceptions are handled gracefully."""
        mock_db.get_emails_by_status.side_effect = Exception("Database error")
        
        with patch('mcp_server.db', mock_db):
            from mcp_server import call_tool
            
            result = await call_tool("get_pending_emails", {})
            
            data = json.loads(result[0].text)
            assert "error" in data
            assert "Database error" in data["error"]


# ============================================================================
# Helper Function Tests
# ============================================================================

class TestHelperFunctions:
    """Tests for helper functions in mcp_server."""
    
    def test_truncate_text_short(self):
        """Test truncate_text with short text."""
        from mcp_server import truncate_text
        
        result = truncate_text("Short text", 500)
        assert result == "Short text"
    
    def test_truncate_text_long(self):
        """Test truncate_text with long text."""
        from mcp_server import truncate_text
        
        long_text = "A" * 1000
        result = truncate_text(long_text, 500)
        
        assert len(result) < len(long_text)
        assert "500 more chars" in result
    
    def test_format_email_summary_basic(self, sample_email):
        """Test format_email_summary without preview."""
        from mcp_server import format_email_summary
        
        result = format_email_summary(sample_email, include_preview=False)
        
        assert result["id"] == sample_email.id
        assert result["subject"] == sample_email.subject
        assert "preview" not in result
    
    def test_format_email_summary_with_preview(self, sample_email):
        """Test format_email_summary with preview."""
        from mcp_server import format_email_summary
        
        result = format_email_summary(sample_email, include_preview=True)
        
        assert "preview" in result
    
    def test_format_email_summary_with_client_matter(self, sample_email):
        """Test format_email_summary includes client/matter when present."""
        from mcp_server import format_email_summary
        
        sample_email.client_id = "client-001"
        sample_email.matter_id = "matter-001"
        
        result = format_email_summary(sample_email)
        
        assert result["client_id"] == "client-001"
        assert result["matter_id"] == "matter-001"


# ============================================================================
# Integration Tests
# ============================================================================

class TestIntegration:
    """Integration tests using real database."""
    
    @pytest.mark.asyncio
    async def test_full_triage_workflow(self, temp_db, sample_emails):
        """Test complete email triage workflow."""
        # Setup: Add emails to database
        for email in sample_emails:
            temp_db.upsert_email(email)
        
        # Add domains
        for email in sample_emails:
            temp_db.upsert_domain(Domain(
                name=email.domain,
                category=EmailCategory.UNCATEGORIZED
            ))
        
        # Verify pending emails
        pending = temp_db.get_emails_by_status(TriageStatus.PENDING)
        assert len(pending) == len(sample_emails)
        
        # Categorize a domain
        temp_db.update_domain_category("example.com", EmailCategory.CLIENT)
        domain = temp_db.get_domain("example.com")
        assert domain.category == EmailCategory.CLIENT
        
        # Triage an email
        temp_db.update_email_triage(
            sample_emails[0].id,
            TriageStatus.PROCESSED,
            client_id="client-001"
        )
        
        # Verify stats
        stats = temp_db.get_triage_stats()
        assert stats.get("processed", 0) == 1
        assert stats.get("pending", 0) == len(sample_emails) - 1
    
    @pytest.mark.asyncio
    async def test_client_matter_workflow(self, temp_db):
        """Test client and matter creation workflow."""
        # Create client
        client = Client(
            id="test-client",
            name="Test Client",
            domains=["testclient.com"]
        )
        temp_db.upsert_client(client)
        
        # Create matters
        matter1 = Matter(id="m1", client_id="test-client", name="Matter 1")
        matter2 = Matter(id="m2", client_id="test-client", name="Matter 2")
        temp_db.upsert_matter(matter1)
        temp_db.upsert_matter(matter2)
        
        # Verify
        all_clients = temp_db.get_all_clients()
        assert len(all_clients) == 1
        
        matters = temp_db.get_matters_for_client("test-client")
        assert len(matters) == 2
    
    @pytest.mark.asyncio
    async def test_batch_archive_workflow(self, temp_db, sample_emails):
        """Test batch archiving of marketing emails."""
        # Setup: Add emails all from one marketing domain
        for i, email in enumerate(sample_emails):
            email.domain = "marketing.example.com"
            temp_db.upsert_email(email)
        
        # Add marketing domain
        temp_db.upsert_domain(Domain(
            name="marketing.example.com",
            category=EmailCategory.MARKETING
        ))
        
        # Archive all emails from this domain
        emails = temp_db.get_emails_by_domain("marketing.example.com")
        for email in emails:
            if email.triage_status == TriageStatus.PENDING:
                temp_db.update_email_triage(email.id, TriageStatus.ARCHIVED)
        
        # Verify all archived
        archived = temp_db.get_emails_by_status(TriageStatus.ARCHIVED)
        assert len(archived) == len(sample_emails)
        
        pending = temp_db.get_emails_by_status(TriageStatus.PENDING)
        assert len(pending) == 0


# ============================================================================
# List Tools Tests
# ============================================================================

class TestListTools:
    """Tests for the list_tools handler."""
    
    @pytest.mark.asyncio
    async def test_list_tools_returns_all_tools(self):
        """Test that list_tools returns all expected tools."""
        from mcp_server import list_tools
        
        tools = await list_tools()
        
        tool_names = [tool.name for tool in tools]
        
        expected_tools = [
            "sync_emails",
            "get_pending_emails",
            "get_emails_by_domain",
            "get_email_content",
            "triage_email",
            "batch_triage",
            "batch_archive_domain",
            "get_uncategorized_domains",
            "categorize_domain",
            "get_domain_summary",
            "create_client",
            "create_matter",
            "list_clients",
            "get_triage_stats",
        ]
        
        for expected in expected_tools:
            assert expected in tool_names, f"Missing tool: {expected}"
    
    @pytest.mark.asyncio
    async def test_tools_have_valid_schemas(self):
        """Test that all tools have valid input schemas."""
        from mcp_server import list_tools
        
        tools = await list_tools()
        
        for tool in tools:
            assert tool.inputSchema is not None
            assert "type" in tool.inputSchema
            assert tool.inputSchema["type"] == "object"
            assert "properties" in tool.inputSchema


# ============================================================================
# Model Tests
# ============================================================================

class TestModels:
    """Tests for data models."""
    
    def test_email_default_values(self):
        """Test Email dataclass default values."""
        email = Email(
            id="test-id",
            subject="Test",
            sender_name="Sender",
            sender_email="sender@test.com",
            domain="test.com",
            received_time=datetime.now()
        )
        
        assert email.triage_status == TriageStatus.PENDING
        assert email.has_attachments is False
        assert email.attachment_names == []
        assert email.body_preview == ""
    
    def test_domain_default_values(self):
        """Test Domain dataclass default values."""
        domain = Domain(name="test.com")
        
        assert domain.category == EmailCategory.UNCATEGORIZED
        assert domain.email_count == 0
        assert domain.sample_senders == []
    
    def test_email_category_enum_values(self):
        """Test EmailCategory enum values."""
        assert EmailCategory.CLIENT.value == "Client"
        assert EmailCategory.INTERNAL.value == "Internal"
        assert EmailCategory.MARKETING.value == "Marketing"
        assert EmailCategory.PERSONAL.value == "Personal"
        assert EmailCategory.UNCATEGORIZED.value == "Uncategorized"
    
    def test_triage_status_enum_values(self):
        """Test TriageStatus enum values."""
        assert TriageStatus.PENDING.value == "pending"
        assert TriageStatus.PROCESSED.value == "processed"
        assert TriageStatus.DEFERRED.value == "deferred"
        assert TriageStatus.ARCHIVED.value == "archived"
    
    def test_counterparty_default_values(self):
        """Test Counterparty dataclass default values."""
        counterparty = Counterparty(
            id="cp-id",
            matter_id="matter-id",
            name="Test Counterparty"
        )
        
        assert counterparty.contact_name is None
        assert counterparty.contact_email is None
        assert counterparty.domains == []
        assert counterparty.notes == ""


# ============================================================================
# Sent Email / Direction Tests
# ============================================================================

class TestEmailDirection:
    """Tests for email direction and sent email functionality."""
    
    def test_email_direction_default_inbound(self):
        """Test that email direction defaults to inbound."""
        email = Email(
            id="test-id",
            subject="Test",
            sender_name="Sender",
            sender_email="sender@test.com",
            domain="test.com",
            received_time=datetime.now()
        )
        assert email.direction == "inbound"
    
    def test_email_direction_outbound(self):
        """Test that outbound direction can be set."""
        email = Email(
            id="test-id",
            subject="Test",
            sender_name="Sender",
            sender_email="sender@test.com",
            domain="test.com",
            received_time=datetime.now(),
            direction="outbound"
        )
        assert email.direction == "outbound"
    
    def test_database_stores_direction(self, temp_db, sample_email):
        """Test that database correctly stores and retrieves direction."""
        sample_email.direction = "outbound"
        temp_db.upsert_email(sample_email)
        
        retrieved = temp_db.get_email(sample_email.id)
        assert retrieved.direction == "outbound"
    
    def test_get_sent_emails_empty(self, temp_db):
        """Test get_sent_emails returns empty list when no sent emails."""
        emails = temp_db.get_sent_emails()
        assert emails == []
    
    def test_get_sent_emails_filters_outbound(self, temp_db, sample_emails):
        """Test get_sent_emails only returns outbound emails."""
        # Insert both inbound and outbound emails
        for i, email in enumerate(sample_emails):
            email.direction = "outbound" if i % 2 == 0 else "inbound"
            temp_db.upsert_email(email)
        
        sent_emails = temp_db.get_sent_emails()
        assert all(e.direction == "outbound" for e in sent_emails)
        # Should have 3 outbound emails (indices 0, 2, 4)
        assert len(sent_emails) == 3
    
    def test_get_sent_emails_filters_by_client(self, temp_db, sample_emails):
        """Test get_sent_emails can filter by client_id."""
        for i, email in enumerate(sample_emails):
            email.direction = "outbound"
            email.client_id = "client-a" if i < 3 else "client-b"
            temp_db.upsert_email(email)
        
        client_a_emails = temp_db.get_sent_emails(client_id="client-a")
        assert len(client_a_emails) == 3
        assert all(e.client_id == "client-a" for e in client_a_emails)
    
    def test_get_emails_by_status_filters_direction(self, temp_db, sample_emails):
        """Test get_emails_by_status respects direction filter."""
        for i, email in enumerate(sample_emails):
            email.direction = "outbound" if i % 2 == 0 else "inbound"
            temp_db.upsert_email(email)
        
        inbound_pending = temp_db.get_emails_by_status(TriageStatus.PENDING, direction="inbound")
        outbound_pending = temp_db.get_emails_by_status(TriageStatus.PENDING, direction="outbound")
        
        assert all(e.direction == "inbound" for e in inbound_pending)
        assert all(e.direction == "outbound" for e in outbound_pending)
    
    def test_get_conversation_thread_empty(self, temp_db):
        """Test get_conversation_thread returns empty list when no matches."""
        emails = temp_db.get_conversation_thread(subject="nonexistent")
        assert emails == []
    
    def test_get_conversation_thread_by_subject(self, temp_db, sample_emails):
        """Test get_conversation_thread finds emails by subject."""
        # Set a common subject for some emails
        sample_emails[0].subject = "Re: Project Update"
        sample_emails[0].direction = "inbound"
        sample_emails[1].subject = "Re: Project Update"
        sample_emails[1].direction = "outbound"
        sample_emails[2].subject = "Different Subject"
        
        for email in sample_emails[:3]:
            temp_db.upsert_email(email)
        
        thread = temp_db.get_conversation_thread(subject="Project Update")
        assert len(thread) == 2
    
    def test_get_conversation_thread_by_conversation_id(self, temp_db, sample_emails):
        """Test get_conversation_thread finds emails by conversation_id."""
        sample_emails[0].conversation_id = "conv-123"
        sample_emails[0].direction = "inbound"
        sample_emails[1].conversation_id = "conv-123"
        sample_emails[1].direction = "outbound"
        sample_emails[2].conversation_id = "conv-456"
        
        for email in sample_emails[:3]:
            temp_db.upsert_email(email)
        
        thread = temp_db.get_conversation_thread(conversation_id="conv-123")
        assert len(thread) == 2
    
    def test_get_last_contact_by_client_empty(self, temp_db):
        """Test get_last_contact_by_client returns empty list when no clients."""
        contacts = temp_db.get_last_contact_by_client()
        assert contacts == []
    
    def test_get_last_contact_by_client(self, temp_db, sample_emails):
        """Test get_last_contact_by_client returns most recent contact per client."""
        # Set up emails for two clients
        for i, email in enumerate(sample_emails[:4]):
            email.client_id = "client-a" if i < 2 else "client-b"
            email.direction = "outbound" if i % 2 == 0 else "inbound"
            temp_db.upsert_email(email)
        
        contacts = temp_db.get_last_contact_by_client()
        assert len(contacts) == 2
        assert any(c["client_id"] == "client-a" for c in contacts)
        assert any(c["client_id"] == "client-b" for c in contacts)


class TestSentEmailMCPTools:
    """Tests for sent email MCP tool handlers."""
    
    @pytest.fixture
    def mock_outlook(self):
        """Create a mock OutlookClient."""
        mock = MagicMock()
        mock.get_email_body.return_value = "Email body text"
        mock.sync_emails_to_db.return_value = {
            "new": 5, "updated": 2, "domains_updated": 3,
            "inbound": 4, "outbound": 3
        }
        return mock
    
    @pytest.mark.asyncio
    async def test_get_sent_emails_tool(self, temp_db, sample_emails, mock_outlook):
        """Test get_sent_emails MCP tool handler."""
        from mcp_server import call_tool
        import mcp_server
        
        # Setup: insert outbound emails
        original_db = mcp_server.db
        mcp_server.db = temp_db
        
        try:
            for email in sample_emails[:3]:
                email.direction = "outbound"
                email.triage_status = TriageStatus.PROCESSED
                temp_db.upsert_email(email)
            
            result = await call_tool("get_sent_emails", {"limit": 10})
            
            assert len(result) == 1
            data = json.loads(result[0].text)
            assert data["count"] == 3
            assert "emails" in data
        finally:
            mcp_server.db = original_db
    
    @pytest.mark.asyncio
    async def test_get_conversation_thread_tool(self, temp_db, sample_emails, mock_outlook):
        """Test get_conversation_thread MCP tool handler."""
        from mcp_server import call_tool
        import mcp_server
        
        original_db = mcp_server.db
        mcp_server.db = temp_db
        
        try:
            # Setup: insert emails with matching conversation
            sample_emails[0].subject = "Re: Important Matter"
            sample_emails[0].direction = "inbound"
            sample_emails[1].subject = "Re: Important Matter"
            sample_emails[1].direction = "outbound"
            
            for email in sample_emails[:2]:
                temp_db.upsert_email(email)
            
            result = await call_tool("get_conversation_thread", {"subject": "Important Matter"})
            
            assert len(result) == 1
            data = json.loads(result[0].text)
            assert data["count"] == 2
            assert "thread" in data
        finally:
            mcp_server.db = original_db
    
    @pytest.mark.asyncio
    async def test_get_last_contact_dates_tool(self, temp_db, sample_emails, mock_outlook):
        """Test get_last_contact_dates MCP tool handler."""
        from mcp_server import call_tool
        import mcp_server
        
        original_db = mcp_server.db
        mcp_server.db = temp_db
        
        try:
            # Setup: insert emails with client assignments
            for i, email in enumerate(sample_emails[:2]):
                email.client_id = f"client-{i}"
                temp_db.upsert_email(email)
            
            result = await call_tool("get_last_contact_dates", {})
            
            assert len(result) == 1
            data = json.loads(result[0].text)
            assert data["client_count"] == 2
            assert "clients" in data
        finally:
            mcp_server.db = original_db
    
    @pytest.mark.asyncio
    async def test_format_email_summary_includes_direction(self, sample_email):
        """Test that format_email_summary includes direction field."""
        from mcp_server import format_email_summary
        
        sample_email.direction = "outbound"
        summary = format_email_summary(sample_email)
        
        assert "direction" in summary
        assert summary["direction"] == "outbound"


# ============================================================================
# Client Linking Tests
# ============================================================================

class TestClientLinking:
    """Tests for linking emails to clients based on domain and contact email."""
    
    def test_link_emails_to_clients_by_domain(self, temp_db, sample_emails, sample_client):
        """Test that emails are linked to clients by matching domain."""
        # Insert client
        temp_db.upsert_client(sample_client)
        
        # Insert emails with matching domain (no client_id set)
        sample_emails[0].domain = "example.com"
        sample_emails[0].client_id = None
        temp_db.upsert_email(sample_emails[0])
        
        # Run linking with domains map (domains now come from effi-clients)
        client_domains_map = {sample_client.id: ["example.com"]}
        stats = temp_db.link_emails_to_clients(client_domains_map)
        
        # Verify email was linked
        assert stats["by_domain"] >= 1
        
        linked_email = temp_db.get_email(sample_emails[0].id)
        assert linked_email.client_id == sample_client.id
    
    def test_link_emails_skips_generic_domains(self, temp_db, sample_emails, sample_client):
        """Test that generic domains (gmail, outlook) are not auto-linked by domain."""
        # Insert client
        temp_db.upsert_client(sample_client)
        
        # Insert email from gmail (no client_id)
        sample_emails[0].domain = "gmail.com"
        sample_emails[0].client_id = None
        temp_db.upsert_email(sample_emails[0])
        
        # Run linking with generic domain in map (shouldn't link)
        client_domains_map = {sample_client.id: ["gmail.com"]}
        stats = temp_db.link_emails_to_clients(client_domains_map)
        
        # Email should NOT be linked (generic domain)
        linked_email = temp_db.get_email(sample_emails[0].id)
        assert linked_email.client_id is None
    
    def test_link_emails_by_contact_email(self, temp_db, sample_emails, sample_client):
        """Test that emails from generic domains are linked via contact_emails table."""
        # Insert client
        temp_db.upsert_client(sample_client)
        
        # Add contact email mapping for a gmail address
        temp_db.upsert_contact_email("john.smith@gmail.com", sample_client.id, "John Smith")
        
        # Insert email from that gmail address
        sample_emails[0].domain = "gmail.com"
        sample_emails[0].sender_email = "john.smith@gmail.com"
        sample_emails[0].client_id = None
        temp_db.upsert_email(sample_emails[0])
        
        # Run linking (no domains map needed for contact_email linking)
        stats = temp_db.link_emails_to_clients()
        
        # Email should be linked via contact_email
        assert stats["by_contact_email"] >= 1
        
        linked_email = temp_db.get_email(sample_emails[0].id)
        assert linked_email.client_id == sample_client.id
    
    def test_link_emails_preserves_existing_client_id(self, temp_db, sample_emails, sample_client):
        """Test that emails already linked to a client are not overwritten."""
        # Insert two clients
        temp_db.upsert_client(sample_client)
        
        from models import Client
        other_client = Client(id="other-client", name="Other Client")
        temp_db.upsert_client(other_client)
        
        # Insert email already linked to other_client
        sample_emails[0].domain = "example.com"
        sample_emails[0].client_id = "other-client"
        temp_db.upsert_email(sample_emails[0])
        
        # Run linking with domains map (should skip since client_id is already set)
        client_domains_map = {sample_client.id: ["example.com"]}
        stats = temp_db.link_emails_to_clients(client_domains_map)
        
        # Email should still be linked to other_client
        linked_email = temp_db.get_email(sample_emails[0].id)
        assert linked_email.client_id == "other-client"
    
    def test_contact_email_crud(self, temp_db, sample_client):
        """Test contact email CRUD operations."""
        temp_db.upsert_client(sample_client)
        
        # Add contact email
        temp_db.upsert_contact_email("test@gmail.com", sample_client.id, "Test User")
        
        # Get contact emails for client
        emails = temp_db.get_contact_emails_for_client(sample_client.id)
        assert "test@gmail.com" in emails
        
        # Get client by contact email
        client = temp_db.get_client_by_contact_email("test@gmail.com")
        assert client is not None
        assert client.id == sample_client.id
    
    @pytest.mark.asyncio
    async def test_add_contact_email_tool(self, temp_db, sample_client):
        """Test add_contact_email MCP tool."""
        from mcp_server import call_tool
        import mcp_server
        
        original_db = mcp_server.db
        mcp_server.db = temp_db
        
        try:
            temp_db.upsert_client(sample_client)
            
            result = await call_tool("add_contact_email", {
                "email": "contact@hotmail.com",
                "client_id": sample_client.id,
                "contact_name": "Contact Person"
            })
            
            assert len(result) == 1
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["email"] == "contact@hotmail.com"
            
            # Verify stored
            emails = temp_db.get_contact_emails_for_client(sample_client.id)
            assert "contact@hotmail.com" in emails
        finally:
            mcp_server.db = original_db
    
    @pytest.mark.asyncio
    async def test_get_contact_emails_tool(self, temp_db, sample_client):
        """Test get_contact_emails MCP tool."""
        from mcp_server import call_tool
        import mcp_server
        
        original_db = mcp_server.db
        mcp_server.db = temp_db
        
        try:
            temp_db.upsert_client(sample_client)
            temp_db.upsert_contact_email("email1@gmail.com", sample_client.id)
            temp_db.upsert_contact_email("email2@outlook.com", sample_client.id)
            
            result = await call_tool("get_contact_emails", {"client_id": sample_client.id})
            
            assert len(result) == 1
            data = json.loads(result[0].text)
            assert data["count"] == 2
            assert "email1@gmail.com" in data["contact_emails"]
            assert "email2@outlook.com" in data["contact_emails"]
        finally:
            mcp_server.db = original_db


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
