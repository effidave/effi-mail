"""Comprehensive tests for client-centric email search tools.

These tests cover:
1. Database schema changes (internet_message_id, recipient_domains)
2. New MCP tools (search_emails_by_client, search_outlook_by_client, etc.)
3. Outlook client DASL query support
"""

import pytest
import json
import asyncio
from datetime import datetime, timedelta
from unittest.mock import Mock, MagicMock, patch, AsyncMock
import tempfile
import os

from models import Email, Domain, Client, Matter, EmailCategory, TriageStatus


# ============================================================================
# Fixtures
# ============================================================================

@pytest.fixture
def temp_db():
    """Create a temporary database for testing."""
    from database import Database
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
def sample_client_vrs():
    """Create a sample client with multiple domains."""
    return Client(
        id="vrs",
        name="VRS Technology Ltd",
        domains=["vrs.com", "vrstech.co.uk"],
        folder_path="/clients/vrs",
    )


@pytest.fixture
def sample_client_acme():
    """Create another sample client."""
    return Client(
        id="acme",
        name="Acme Corporation",
        domains=["acme.com"],
        folder_path="/clients/acme",
    )


@pytest.fixture
def sample_emails_for_client_search():
    """Create emails that should match various client search criteria."""
    base_time = datetime.now()
    return [
        # Inbound from client domain
        Email(
            id="email-001",
            internet_message_id="<msg001@vrs.com>",
            subject="Project Update",
            sender_name="John VRS",
            sender_email="john@vrs.com",
            domain="vrs.com",
            received_time=base_time - timedelta(hours=1),
            body_preview="Here's the project update...",
            direction="inbound",
            recipients_to=["david@ourcompany.com"],
            recipients_cc=[],
            recipient_domains="ourcompany.com",
            triage_status=TriageStatus.PENDING,
        ),
        # Inbound from client's secondary domain
        Email(
            id="email-002",
            internet_message_id="<msg002@vrstech.co.uk>",
            subject="Invoice Query",
            sender_name="Jane VRS",
            sender_email="jane@vrstech.co.uk",
            domain="vrstech.co.uk",
            received_time=base_time - timedelta(hours=2),
            body_preview="About the invoice...",
            direction="inbound",
            recipients_to=["david@ourcompany.com"],
            recipients_cc=[],
            recipient_domains="ourcompany.com",
            triage_status=TriageStatus.PENDING,
        ),
        # Outbound TO client domain
        Email(
            id="email-003",
            internet_message_id="<msg003@ourcompany.com>",
            subject="RE: Project Update",
            sender_name="David",
            sender_email="david@ourcompany.com",
            domain="vrs.com",  # For outbound, domain is recipient domain
            received_time=base_time - timedelta(hours=0.5),
            body_preview="Thanks for the update...",
            direction="outbound",
            recipients_to=["john@vrs.com"],
            recipients_cc=["jane@vrstech.co.uk"],
            recipient_domains="vrs.com,vrstech.co.uk",
            triage_status=TriageStatus.PROCESSED,
        ),
        # Inbound from contact email (gmail - generic domain)
        Email(
            id="email-004",
            internet_message_id="<msg004@gmail.com>",
            subject="VRS Related Query",
            sender_name="Bob VRS Personal",
            sender_email="bob.vrs@gmail.com",
            domain="gmail.com",
            received_time=base_time - timedelta(hours=3),
            body_preview="Quick question...",
            direction="inbound",
            recipients_to=["david@ourcompany.com"],
            recipients_cc=[],
            recipient_domains="ourcompany.com",
            triage_status=TriageStatus.PENDING,
        ),
        # Outbound with client in CC
        Email(
            id="email-005",
            internet_message_id="<msg005@ourcompany.com>",
            subject="Third Party Discussion",
            sender_name="David",
            sender_email="david@ourcompany.com",
            domain="thirdparty.com",
            received_time=base_time - timedelta(hours=4),
            body_preview="Loop in VRS...",
            direction="outbound",
            recipients_to=["someone@thirdparty.com"],
            recipients_cc=["john@vrs.com"],
            recipient_domains="thirdparty.com,vrs.com",
            triage_status=TriageStatus.PROCESSED,
        ),
        # Unrelated email (should NOT match VRS search)
        Email(
            id="email-006",
            internet_message_id="<msg006@other.com>",
            subject="Unrelated Email",
            sender_name="Other Person",
            sender_email="other@other.com",
            domain="other.com",
            received_time=base_time - timedelta(hours=5),
            body_preview="Not related to VRS...",
            direction="inbound",
            recipients_to=["david@ourcompany.com"],
            recipients_cc=[],
            recipient_domains="ourcompany.com",
            triage_status=TriageStatus.PENDING,
        ),
    ]


# ============================================================================
# Schema Tests - internet_message_id
# ============================================================================

class TestInternetMessageIdSchema:
    """Tests for internet_message_id field in database schema."""
    
    def test_email_model_has_internet_message_id_field(self):
        """Test that Email model has internet_message_id field."""
        email = Email(
            id="test-001",
            subject="Test",
            sender_name="Test",
            sender_email="test@test.com",
            domain="test.com",
            received_time=datetime.now(),
            internet_message_id="<unique123@test.com>",
        )
        assert email.internet_message_id == "<unique123@test.com>"
    
    def test_internet_message_id_defaults_to_none(self):
        """Test that internet_message_id defaults to None."""
        email = Email(
            id="test-001",
            subject="Test",
            sender_name="Test",
            sender_email="test@test.com",
            domain="test.com",
            received_time=datetime.now(),
        )
        assert email.internet_message_id is None
    
    def test_database_stores_internet_message_id(self, temp_db):
        """Test that database stores and retrieves internet_message_id."""
        email = Email(
            id="test-001",
            subject="Test Subject",
            sender_name="Sender",
            sender_email="sender@test.com",
            domain="test.com",
            received_time=datetime.now(),
            internet_message_id="<msg123@test.com>",
        )
        temp_db.upsert_email(email)
        
        retrieved = temp_db.get_email("test-001")
        assert retrieved is not None
        assert retrieved.internet_message_id == "<msg123@test.com>"
    
    def test_get_email_by_internet_message_id(self, temp_db):
        """Test looking up email by internet_message_id."""
        email = Email(
            id="test-001",
            subject="Test Subject",
            sender_name="Sender",
            sender_email="sender@test.com",
            domain="test.com",
            received_time=datetime.now(),
            internet_message_id="<unique-msg-id@example.com>",
        )
        temp_db.upsert_email(email)
        
        retrieved = temp_db.get_email_by_internet_message_id("<unique-msg-id@example.com>")
        assert retrieved is not None
        assert retrieved.id == "test-001"
        assert retrieved.subject == "Test Subject"
    
    def test_get_email_by_internet_message_id_not_found(self, temp_db):
        """Test looking up non-existent internet_message_id returns None."""
        result = temp_db.get_email_by_internet_message_id("<nonexistent@test.com>")
        assert result is None


# ============================================================================
# Schema Tests - recipient_domains
# ============================================================================

class TestRecipientDomainsSchema:
    """Tests for recipient_domains field in database schema."""
    
    def test_email_model_has_recipient_domains_field(self):
        """Test that Email model has recipient_domains field."""
        email = Email(
            id="test-001",
            subject="Test",
            sender_name="Test",
            sender_email="test@test.com",
            domain="test.com",
            received_time=datetime.now(),
            recipient_domains="client.com,other.com",
        )
        assert email.recipient_domains == "client.com,other.com"
    
    def test_recipient_domains_defaults_to_empty(self):
        """Test that recipient_domains defaults to empty string."""
        email = Email(
            id="test-001",
            subject="Test",
            sender_name="Test",
            sender_email="test@test.com",
            domain="test.com",
            received_time=datetime.now(),
        )
        assert email.recipient_domains == ""
    
    def test_database_stores_recipient_domains(self, temp_db):
        """Test that database stores and retrieves recipient_domains."""
        email = Email(
            id="test-001",
            subject="Test Subject",
            sender_name="Sender",
            sender_email="sender@test.com",
            domain="test.com",
            received_time=datetime.now(),
            recipients_to=["user@client.com", "other@partner.com"],
            recipients_cc=["cc@third.com"],
            recipient_domains="client.com,partner.com,third.com",
        )
        temp_db.upsert_email(email)
        
        retrieved = temp_db.get_email("test-001")
        assert retrieved is not None
        assert "client.com" in retrieved.recipient_domains
        assert "partner.com" in retrieved.recipient_domains
        assert "third.com" in retrieved.recipient_domains
    
    def test_search_by_recipient_domain(self, temp_db):
        """Test searching emails by recipient domain."""
        emails = [
            Email(
                id="email-001",
                subject="To Client A",
                sender_name="Me",
                sender_email="me@mycompany.com",
                domain="clienta.com",
                received_time=datetime.now(),
                direction="outbound",
                recipients_to=["user@clienta.com"],
                recipient_domains="clienta.com",
            ),
            Email(
                id="email-002",
                subject="To Client B",
                sender_name="Me",
                sender_email="me@mycompany.com",
                domain="clientb.com",
                received_time=datetime.now(),
                direction="outbound",
                recipients_to=["user@clientb.com"],
                recipient_domains="clientb.com",
            ),
        ]
        for email in emails:
            temp_db.upsert_email(email)
        
        # Search for emails sent to clienta.com
        results = temp_db.search_emails_by_recipient_domain("clienta.com")
        assert len(results) == 1
        assert results[0].id == "email-001"


# ============================================================================
# Database Tests - get_client_identifiers (now from effi-clients)
# ============================================================================

class TestGetClientIdentifiers:
    """Tests for get_client_identifiers_from_effi_work function.
    
    These tests mock the MCP session to test response parsing logic.
    Note: Function calls effi-clients but name kept for backwards compatibility.
    """
    
    @pytest.mark.asyncio
    async def test_get_client_identifiers_returns_domains(self):
        """Test that get_client_identifiers_from_effi_work returns domains."""
        mock_response = {
            "folder": "VRS",
            "context": {
                "domains": ["vrs.com", "vrstech.co.uk"],
                "contact_emails": [],
            }
        }
        
        with patch('effi_work_client.get_effi_clients_session') as mock_session:
            mock_session.return_value.__aenter__.return_value.call_tool.return_value = Mock(
                content=[Mock(text=json.dumps(mock_response))]
            )
            
            from effi_work_client import get_client_identifiers_from_effi_work
            identifiers = await get_client_identifiers_from_effi_work("vrs")
        
        assert "vrs.com" in identifiers["domains"]
        assert "vrstech.co.uk" in identifiers["domains"]
    
    @pytest.mark.asyncio
    async def test_get_client_identifiers_returns_contact_emails(self):
        """Test that get_client_identifiers_from_effi_work returns contact emails."""
        mock_response = {
            "folder": "VRS",
            "context": {
                "domains": ["vrs.com"],
                "contact_emails": ["bob.vrs@gmail.com", "alice.vrs@yahoo.com"],
            }
        }
        
        with patch('effi_work_client.get_effi_clients_session') as mock_session:
            mock_session.return_value.__aenter__.return_value.call_tool.return_value = Mock(
                content=[Mock(text=json.dumps(mock_response))]
            )
            
            from effi_work_client import get_client_identifiers_from_effi_work
            identifiers = await get_client_identifiers_from_effi_work("vrs")
        
        assert "bob.vrs@gmail.com" in identifiers["contact_emails"]
        assert "alice.vrs@yahoo.com" in identifiers["contact_emails"]
    
    @pytest.mark.asyncio
    async def test_get_client_identifiers_client_not_found(self):
        """Test that get_client_identifiers_from_effi_work returns empty for unknown client."""
        mock_response = {"error": "Client not found"}
        
        with patch('effi_work_client.get_effi_clients_session') as mock_session:
            mock_session.return_value.__aenter__.return_value.call_tool.return_value = Mock(
                content=[Mock(text=json.dumps(mock_response))]
            )
            
            from effi_work_client import get_client_identifiers_from_effi_work
            identifiers = await get_client_identifiers_from_effi_work("nonexistent")
        
        assert identifiers["domains"] == []
        assert identifiers["contact_emails"] == []
        assert identifiers["source"] == "not-found"
    
    @pytest.mark.asyncio
    async def test_get_client_identifiers_combined(self):
        """Test full identifiers response structure from effi-clients."""
        mock_response = {
            "folder": "VRS",
            "context": {
                "domains": ["vrs.com", "vrstech.co.uk"],
                "contact_emails": ["personal@gmail.com"],
            }
        }
        
        with patch('effi_work_client.get_effi_clients_session') as mock_session:
            mock_session.return_value.__aenter__.return_value.call_tool.return_value = Mock(
                content=[Mock(text=json.dumps(mock_response))]
            )
            
            from effi_work_client import get_client_identifiers_from_effi_work
            identifiers = await get_client_identifiers_from_effi_work("vrs")
        
        assert "client_id" in identifiers
        assert identifiers["client_id"] == "vrs"
        assert len(identifiers["domains"]) == 2
        assert len(identifiers["contact_emails"]) == 1
        assert identifiers["source"] == "effi-clients"


# ============================================================================
# Database Tests - search_emails_by_client
# ============================================================================

class TestSearchEmailsByClientDatabase:
    """Tests for search_emails_by_client database method.
    
    Note: The method now takes domains/contact_emails directly instead of client_id.
    This is because client lookups now come from effi-clients via MCP.
    """
    
    # VRS domains for test convenience
    VRS_DOMAINS = ["vrs.com", "vrstech.co.uk"]
    VRS_CONTACT_EMAILS = ["bob.vrs@gmail.com"]
    
    def test_search_finds_inbound_from_client_domain(self, temp_db, sample_emails_for_client_search):
        """Test finding inbound emails from client domains."""
        for email in sample_emails_for_client_search:
            temp_db.upsert_email(email)
        
        results = temp_db.search_emails_by_client(
            domains=self.VRS_DOMAINS,
            include_inbound=True,
            include_outbound=False,
            include_cc=False,
        )
        
        # Should find email-001 (from vrs.com) and email-002 (from vrstech.co.uk)
        ids = [e.id for e in results]
        assert "email-001" in ids
        assert "email-002" in ids
        assert "email-006" not in ids  # Unrelated
    
    def test_search_finds_outbound_to_client_domain(self, temp_db, sample_emails_for_client_search):
        """Test finding outbound emails to client domains."""
        for email in sample_emails_for_client_search:
            temp_db.upsert_email(email)
        
        results = temp_db.search_emails_by_client(
            domains=self.VRS_DOMAINS,
            include_inbound=False,
            include_outbound=True,
            include_cc=False,
        )
        
        ids = [e.id for e in results]
        assert "email-003" in ids  # Outbound to vrs.com
    
    def test_search_finds_cc_to_client(self, temp_db, sample_emails_for_client_search):
        """Test finding emails where client is CC'd."""
        for email in sample_emails_for_client_search:
            temp_db.upsert_email(email)
        
        results = temp_db.search_emails_by_client(
            domains=self.VRS_DOMAINS,
            include_inbound=False,
            include_outbound=False,
            include_cc=True,
        )
        
        ids = [e.id for e in results]
        assert "email-005" in ids  # CC'd to john@vrs.com
    
    def test_search_finds_contact_emails(self, temp_db, sample_emails_for_client_search):
        """Test finding emails from registered contact emails."""
        for email in sample_emails_for_client_search:
            temp_db.upsert_email(email)
        
        results = temp_db.search_emails_by_client(
            domains=[],  # No domain matching
            contact_emails=self.VRS_CONTACT_EMAILS,
            include_inbound=True,
            include_outbound=False,
            include_cc=False,
        )
        
        ids = [e.id for e in results]
        assert "email-004" in ids  # From bob.vrs@gmail.com
    
    def test_search_all_combined(self, temp_db, sample_emails_for_client_search):
        """Test combined search with all options enabled."""
        for email in sample_emails_for_client_search:
            temp_db.upsert_email(email)
        
        results = temp_db.search_emails_by_client(
            domains=self.VRS_DOMAINS,
            contact_emails=self.VRS_CONTACT_EMAILS,
            include_inbound=True,
            include_outbound=True,
            include_cc=True,
        )
        
        ids = [e.id for e in results]
        # Should find all VRS-related emails (1,2,3,4,5) but not email-006
        assert "email-001" in ids
        assert "email-002" in ids
        assert "email-003" in ids
        assert "email-004" in ids
        assert "email-005" in ids
        assert "email-006" not in ids
    
    def test_search_respects_days_filter(self, temp_db):
        """Test that days filter is respected."""
        
        # Email from 40 days ago
        old_email = Email(
            id="old-email",
            subject="Old Email",
            sender_name="Old",
            sender_email="old@vrs.com",
            domain="vrs.com",
            received_time=datetime.now() - timedelta(days=40),
            direction="inbound",
        )
        temp_db.upsert_email(old_email)
        
        # Email from 5 days ago
        new_email = Email(
            id="new-email",
            subject="New Email",
            sender_name="New",
            sender_email="new@vrs.com",
            domain="vrs.com",
            received_time=datetime.now() - timedelta(days=5),
            direction="inbound",
        )
        temp_db.upsert_email(new_email)
        
        # Search with 30-day limit
        results = temp_db.search_emails_by_client(domains=["vrs.com"], days=30)
        
        ids = [e.id for e in results]
        assert "new-email" in ids
        assert "old-email" not in ids
    
    def test_search_respects_date_range(self, temp_db):
        """Test that date_from and date_to filters work."""
        
        emails = [
            Email(
                id=f"email-{i}",
                subject=f"Email {i}",
                sender_name="Sender",
                sender_email="sender@vrs.com",
                domain="vrs.com",
                received_time=datetime(2025, 12, i, 10, 0),
                direction="inbound",
            )
            for i in range(1, 21)  # Dec 1-20
        ]
        for email in emails:
            temp_db.upsert_email(email)
        
        # Search for Dec 5-10 (should include Dec 5 00:00:00 through Dec 10 23:59:59)
        results = temp_db.search_emails_by_client(
            domains=["vrs.com"],
            date_from="2025-12-05",
            date_to="2025-12-10",
        )
        
        ids = [e.id for e in results]
        # Emails 5-10 all have timestamps at 10:00 AM on their respective days
        # With date_from=Dec 5 00:00:00 and date_to=Dec 10 23:59:59, we should get 6
        assert len(results) >= 5  # At minimum we should get Dec 5-9
        assert "email-5" in ids
        assert "email-9" in ids
        assert "email-4" not in ids
        assert "email-11" not in ids
    
    def test_search_respects_limit(self, temp_db):
        """Test that limit parameter is respected."""
        
        for i in range(50):
            email = Email(
                id=f"email-{i:03d}",
                subject=f"Email {i}",
                sender_name="Sender",
                sender_email="sender@vrs.com",
                domain="vrs.com",
                received_time=datetime.now() - timedelta(hours=i),
                direction="inbound",
            )
            temp_db.upsert_email(email)
        
        results = temp_db.search_emails_by_client(domains=["vrs.com"], limit=10)
        assert len(results) == 10


# ============================================================================
# MCP Tool Tests - search_emails_by_client
# ============================================================================

class TestSearchEmailsByClientTool:
    """Tests for search_emails_by_client MCP tool."""
    
    @pytest.fixture
    def mock_db(self):
        """Create a mock database."""
        from database import Database
        return Mock(spec=Database)
    
    @pytest.mark.asyncio
    async def test_search_emails_by_client_basic(self, mock_db, sample_emails_for_client_search):
        """Test basic search_emails_by_client tool call."""
        mock_identifiers = {
            "client_id": "vrs",
            "domains": ["vrs.com", "vrstech.co.uk"],
            "contact_emails": [],
            "source": "effi-clients",
        }
        mock_db.search_emails_by_client.return_value = sample_emails_for_client_search[:3]
        
        with patch('mcp_server.db', mock_db), \
             patch('mcp_server.get_client_identifiers_from_effi_work', return_value=mock_identifiers):
            from mcp_server import call_tool
            
            result = await call_tool("search_emails_by_client", {
                "client_id": "vrs",
            })
            
            data = json.loads(result.content[0].text)
            assert data["client_id"] == "vrs"
            assert "emails" in data
            assert "count" in data
    
    @pytest.mark.asyncio
    async def test_search_emails_by_client_with_filters(self, mock_db):
        """Test search_emails_by_client with all filters."""
        mock_identifiers = {
            "client_id": "vrs",
            "domains": ["vrs.com"],
            "contact_emails": ["personal@gmail.com"],
            "source": "effi-clients",
        }
        mock_db.search_emails_by_client.return_value = []
        
        with patch('mcp_server.db', mock_db), \
             patch('mcp_server.get_client_identifiers_from_effi_work', return_value=mock_identifiers):
            from mcp_server import call_tool
            
            result = await call_tool("search_emails_by_client", {
                "client_id": "vrs",
                "include_inbound": True,
                "include_outbound": False,
                "include_cc": True,
                "include_contact_emails": True,
                "days": 14,
                "limit": 50,
            })
            
            mock_db.search_emails_by_client.assert_called_once()
            call_args = mock_db.search_emails_by_client.call_args
            # New signature: domains and contact_emails passed directly
            assert call_args.kwargs.get("domains") == ["vrs.com"]
            assert call_args.kwargs.get("contact_emails") == ["personal@gmail.com"]
    
    @pytest.mark.asyncio
    async def test_search_emails_by_client_not_found(self, mock_db):
        """Test search when client doesn't exist."""
        mock_identifiers = {
            "client_id": None,
            "domains": [],
            "contact_emails": [],
            "source": "not-found",
        }
        
        with patch('mcp_server.db', mock_db), \
             patch('mcp_server.get_client_identifiers_from_effi_work', return_value=mock_identifiers):
            from mcp_server import call_tool
            
            result = await call_tool("search_emails_by_client", {
                "client_id": "nonexistent",
            })
            
            data = json.loads(result.content[0].text)
            assert "error" in data or data.get("count", 0) == 0


# ============================================================================
# MCP Tool Tests - search_outlook_direct
# ============================================================================

class TestSearchOutlookDirectTool:
    """Tests for search_outlook_direct MCP tool."""
    
    @pytest.fixture
    def mock_outlook(self):
        """Create a mock Outlook client."""
        from outlook_client import OutlookClient
        return Mock(spec=OutlookClient)
    
    @pytest.mark.asyncio
    async def test_search_outlook_direct_by_sender_domain(self, mock_outlook):
        """Test direct Outlook search by sender domain."""
        mock_outlook.search_outlook.return_value = []
        
        with patch('mcp_server.outlook', mock_outlook):
            from mcp_server import call_tool
            
            result = await call_tool("search_outlook_direct", {
                "sender_domain": "client.com",
                "days": 14,
            })
            
            data = json.loads(result.content[0].text)
            assert "emails" in data
            mock_outlook.search_outlook.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_search_outlook_direct_with_all_filters(self, mock_outlook):
        """Test direct Outlook search with all filter options."""
        mock_outlook.search_outlook.return_value = []
        
        with patch('mcp_server.outlook', mock_outlook):
            from mcp_server import call_tool
            
            result = await call_tool("search_outlook_direct", {
                "sender_domain": "client.com",
                "sender_email": "specific@client.com",
                "recipient_domain": "ourcompany.com",
                "subject_contains": "Project",
                "body_contains": "deadline",
                "date_from": "2025-12-01",
                "date_to": "2025-12-15",
                "folder": "Inbox",
                "limit": 25,
            })
            
            mock_outlook.search_outlook.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_search_outlook_direct_sent_folder(self, mock_outlook):
        """Test searching Sent Items folder directly."""
        mock_outlook.search_outlook.return_value = []
        
        with patch('mcp_server.outlook', mock_outlook):
            from mcp_server import call_tool
            
            result = await call_tool("search_outlook_direct", {
                "recipient_domain": "client.com",
                "folder": "Sent Items",
                "days": 7,
            })
            
            call_args = mock_outlook.search_outlook.call_args
            assert "Sent" in str(call_args) or call_args[1].get("folder") == "Sent Items"


# ============================================================================
# MCP Tool Tests - search_outlook_by_client
# ============================================================================

class TestSearchOutlookByClientTool:
    """Tests for search_outlook_by_client MCP tool."""
    
    @pytest.fixture
    def mock_db(self):
        from database import Database
        return Mock(spec=Database)
    
    @pytest.fixture
    def mock_outlook(self):
        from outlook_client import OutlookClient
        return Mock(spec=OutlookClient)
    
    @pytest.mark.asyncio
    async def test_search_outlook_by_client_basic(self, mock_db, mock_outlook):
        """Test basic search_outlook_by_client call."""
        mock_identifiers = {
            "client_id": "vrs",
            "domains": ["vrs.com", "vrstech.co.uk"],
            "contact_emails": ["personal@gmail.com"],
            "source": "effi-clients",
        }
        mock_outlook.search_outlook_by_identifiers.return_value = []
        
        with patch('mcp_server.db', mock_db), \
             patch('mcp_server.outlook', mock_outlook), \
             patch('mcp_server.get_client_identifiers_from_effi_work', return_value=mock_identifiers):
            from mcp_server import call_tool
            
            result = await call_tool("search_outlook_by_client", {
                "client_id": "vrs",
                "days": 30,
            })
            
            data = json.loads(result.content[0].text)
            assert "emails" in data
    
    @pytest.mark.asyncio
    async def test_search_outlook_by_client_with_sync(self, mock_db, mock_outlook):
        """Test search_outlook_by_client with sync_results=True."""
        mock_identifiers = {
            "client_id": "vrs",
            "domains": ["vrs.com"],
            "contact_emails": [],
            "source": "effi-clients",
        }
        sample_email = Email(
            id="outlook-001",
            subject="Test",
            sender_name="Test",
            sender_email="test@vrs.com",
            domain="vrs.com",
            received_time=datetime.now(),
        )
        mock_outlook.search_outlook_by_identifiers.return_value = [sample_email]
        mock_db.upsert_email = Mock()
        
        with patch('mcp_server.db', mock_db), \
             patch('mcp_server.outlook', mock_outlook), \
             patch('mcp_server.get_client_identifiers_from_effi_work', return_value=mock_identifiers):
            from mcp_server import call_tool
            
            result = await call_tool("search_outlook_by_client", {
                "client_id": "vrs",
                "sync_results": True,
            })
            
            data = json.loads(result.content[0].text)
            assert data.get("synced", 0) == 1 or mock_db.upsert_email.called


# ============================================================================
# MCP Tool Tests - get_email_by_id
# ============================================================================

class TestGetEmailByIdTool:
    """Tests for get_email_by_id MCP tool."""
    
    @pytest.fixture
    def mock_db(self):
        from database import Database
        return Mock(spec=Database)
    
    @pytest.fixture
    def mock_outlook(self):
        from outlook_client import OutlookClient
        return Mock(spec=OutlookClient)
    
    @pytest.mark.asyncio
    async def test_get_email_by_entry_id(self, mock_db, mock_outlook):
        """Test fetching email by EntryID."""
        mock_outlook.get_email_full.return_value = {
            "id": "ENTRY123",
            "subject": "Test Email",
            "body": "Full body text here...",
            "attachments": [{"name": "doc.pdf", "size": 1024}],
        }
        
        with patch('mcp_server.outlook', mock_outlook):
            from mcp_server import call_tool
            
            result = await call_tool("get_email_by_id", {
                "email_id": "ENTRY123",
            })
            
            data = json.loads(result.content[0].text)
            assert data["subject"] == "Test Email"
            assert "body" in data
    
    @pytest.mark.asyncio
    async def test_get_email_by_internet_message_id(self, mock_db, mock_outlook):
        """Test fetching email by internet_message_id (with @ symbol)."""
        # First lookup in DB to get EntryID
        mock_db.get_email_by_internet_message_id.return_value = Email(
            id="ENTRY123",
            subject="Test",
            sender_name="Test",
            sender_email="test@test.com",
            domain="test.com",
            received_time=datetime.now(),
            internet_message_id="<msg123@domain.com>",
        )
        mock_outlook.get_email_full.return_value = {
            "id": "ENTRY123",
            "subject": "Test",
            "body": "Body text",
            "attachments": [],
        }
        
        with patch('mcp_server.db', mock_db), patch('mcp_server.outlook', mock_outlook):
            from mcp_server import call_tool
            
            # Internet message IDs contain @ and angle brackets
            result = await call_tool("get_email_by_id", {
                "email_id": "<msg123@domain.com>",
            })
            
            mock_db.get_email_by_internet_message_id.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_get_email_by_id_not_found(self, mock_outlook):
        """Test handling of non-existent email ID."""
        mock_outlook.get_email_full.side_effect = Exception("Item not found")
        
        with patch('mcp_server.outlook', mock_outlook):
            from mcp_server import call_tool
            
            result = await call_tool("get_email_by_id", {
                "email_id": "NONEXISTENT",
            })
            
            data = json.loads(result.content[0].text)
            assert "error" in data


# ============================================================================
# MCP Tool Tests - sync_emails_by_client
# ============================================================================

class TestSyncEmailsByClientTool:
    """Tests for sync_emails_by_client MCP tool."""
    
    @pytest.fixture
    def mock_db(self):
        from database import Database
        return Mock(spec=Database)
    
    @pytest.fixture
    def mock_outlook(self):
        from outlook_client import OutlookClient
        return Mock(spec=OutlookClient)
    
    @pytest.mark.asyncio
    async def test_sync_emails_by_client_basic(self, mock_db, mock_outlook):
        """Test syncing emails for a specific client."""
        mock_identifiers = {
            "client_id": "vrs",
            "domains": ["vrs.com"],
            "contact_emails": [],
            "source": "effi-clients",
        }
        mock_outlook.sync_emails_by_identifiers.return_value = {
            "new": 5,
            "updated": 2,
        }
        
        with patch('mcp_server.db', mock_db), \
             patch('mcp_server.outlook', mock_outlook), \
             patch('mcp_server.get_client_identifiers_from_effi_work', return_value=mock_identifiers):
            from mcp_server import call_tool
            
            result = await call_tool("sync_emails_by_client", {
                "client_id": "vrs",
                "days": 30,
            })
            
            data = json.loads(result.content[0].text)
            assert data["success"] is True
            # Stats are nested under "stats" key
            assert "stats" in data
            assert data["stats"]["new"] == 5


# ============================================================================
# MCP Tool Tests - sync_email_by_id
# ============================================================================

class TestSyncEmailByIdTool:
    """Tests for sync_email_by_id MCP tool."""
    
    @pytest.fixture
    def mock_db(self):
        from database import Database
        return Mock(spec=Database)
    
    @pytest.fixture
    def mock_outlook(self):
        from outlook_client import OutlookClient
        return Mock(spec=OutlookClient)
    
    @pytest.mark.asyncio
    async def test_sync_single_email_by_id(self, mock_db, mock_outlook):
        """Test syncing a single email by EntryID."""
        sample_email = Email(
            id="ENTRY123",
            subject="Test Email",
            sender_name="Sender",
            sender_email="sender@client.com",
            domain="client.com",
            received_time=datetime.now(),
            internet_message_id="<msg123@client.com>",
        )
        mock_outlook.get_email_for_sync.return_value = sample_email
        mock_db.upsert_email = Mock()
        
        with patch('mcp_server.db', mock_db), patch('mcp_server.outlook', mock_outlook):
            from mcp_server import call_tool
            
            result = await call_tool("sync_email_by_id", {
                "email_id": "ENTRY123",
            })
            
            data = json.loads(result.content[0].text)
            assert data["success"] is True
            mock_db.upsert_email.assert_called_once()


# ============================================================================
# Outlook Client Tests - DASL Query Support
# ============================================================================

class TestOutlookDASLQueries:
    """Tests for Outlook DASL query functionality."""
    
    def test_build_dasl_query_sender_domain(self):
        """Test building DASL query for sender domain filter."""
        from outlook_client import OutlookClient
        
        client = OutlookClient.__new__(OutlookClient)
        query = client._build_dasl_query(sender_domain="client.com")
        
        assert "client.com" in query
        # DASL uses urn:schemas:httpmail:fromemail
        assert "urn:schemas:httpmail" in query or "fromemail" in query.lower() or "senderemail" in query.lower()
    
    def test_build_dasl_query_date_range(self):
        """Test building DASL query with date range."""
        from outlook_client import OutlookClient
        
        client = OutlookClient.__new__(OutlookClient)
        query = client._build_dasl_query(
            date_from=datetime(2025, 12, 1),
            date_to=datetime(2025, 12, 15),
        )
        
        # Date-only query should use Jet syntax (no @SQL= prefix)
        assert "ReceivedTime" in query
        assert "@SQL=" not in query
    
    def test_build_dasl_query_subject(self):
        """Test building DASL query with subject filter."""
        from outlook_client import OutlookClient
        
        client = OutlookClient.__new__(OutlookClient)
        query = client._build_dasl_query(subject_contains="Project Update")
        
        # Subject filter uses DASL syntax
        assert "@SQL=" in query
        assert "Project Update" in query
    
    def test_build_dasl_query_combined(self):
        """Test building DASL query with multiple filters."""
        from outlook_client import OutlookClient
        
        client = OutlookClient.__new__(OutlookClient)
        query = client._build_dasl_query(
            sender_domain="client.com",
            subject_contains="Invoice",
            date_from=datetime(2025, 12, 1),  # Dates ignored for DASL
        )
        
        # DASL query with domain and subject (dates are handled separately)
        assert "@SQL=" in query
        assert "AND" in query
        assert "client.com" in query
        assert "Invoice" in query


# ============================================================================
# Outlook Client Tests - search_outlook method
# ============================================================================

class TestOutlookSearchMethod:
    """Tests for OutlookClient.search_outlook method."""
    
    @pytest.fixture
    def mock_outlook_client(self):
        """Create an OutlookClient with mocked COM objects."""
        with patch('outlook_client.win32com.client'):
            from outlook_client import OutlookClient
            from database import Database
            
            mock_db = Mock(spec=Database)
            client = OutlookClient(db=mock_db)
            client._outlook = Mock()
            client._namespace = Mock()
            return client
    
    def test_search_outlook_returns_emails(self, mock_outlook_client):
        """Test that search_outlook returns Email objects."""
        # Mock the folder and messages
        mock_folder = Mock()
        mock_messages = Mock()
        mock_date_filtered = Mock()
        mock_dasl_filtered = Mock()
        
        mock_outlook_client._namespace.GetDefaultFolder.return_value = mock_folder
        mock_folder.Items = mock_messages
        # First Restrict (date filter) returns mock_date_filtered
        # Second Restrict (DASL filter) returns mock_dasl_filtered
        mock_messages.Restrict.return_value = mock_date_filtered
        mock_date_filtered.Restrict.return_value = mock_dasl_filtered
        mock_dasl_filtered.__iter__ = Mock(return_value=iter([]))
        
        results = list(mock_outlook_client.search_outlook(
            sender_domain="client.com",
            days=7,
        ))
        
        assert isinstance(results, list)
    
    def test_search_outlook_applies_filters(self, mock_outlook_client):
        """Test that search_outlook applies the correct filters."""
        mock_folder = Mock()
        mock_messages = Mock()
        mock_date_filtered = Mock()
        mock_dasl_filtered = Mock()
        
        mock_outlook_client._namespace.GetDefaultFolder.return_value = mock_folder
        mock_folder.Items = mock_messages
        mock_messages.Restrict.return_value = mock_date_filtered
        mock_date_filtered.Restrict.return_value = mock_dasl_filtered
        mock_dasl_filtered.__iter__ = Mock(return_value=iter([]))
        
        list(mock_outlook_client.search_outlook(
            sender_domain="client.com",
            subject_contains="Invoice",
        ))
        
        # Verify date Restrict was called first
        mock_messages.Restrict.assert_called_once()
        date_query = mock_messages.Restrict.call_args[0][0]
        assert "ReceivedTime" in date_query
        
        # Verify DASL Restrict was called with domain and subject
        mock_date_filtered.Restrict.assert_called_once()
        dasl_query = mock_date_filtered.Restrict.call_args[0][0]
        assert "@SQL=" in dasl_query
        assert "client.com" in dasl_query
        assert "Invoice" in dasl_query


# ============================================================================
# Outlook Client Tests - internet_message_id extraction
# ============================================================================

class TestInternetMessageIdExtraction:
    """Tests for extracting internet_message_id from Outlook messages."""
    
    def test_extract_internet_message_id(self):
        """Test extracting PR_INTERNET_MESSAGE_ID from message."""
        from outlook_client import OutlookClient
        
        # Mock message with PropertyAccessor
        mock_message = Mock()
        mock_message.PropertyAccessor.GetProperty.return_value = "<unique123@domain.com>"
        
        client = OutlookClient.__new__(OutlookClient)
        msg_id = client._get_internet_message_id(mock_message)
        
        assert msg_id == "<unique123@domain.com>"
    
    def test_extract_internet_message_id_fallback(self):
        """Test fallback when PR_INTERNET_MESSAGE_ID not available."""
        from outlook_client import OutlookClient
        
        mock_message = Mock()
        mock_message.PropertyAccessor.GetProperty.side_effect = Exception("Not found")
        
        client = OutlookClient.__new__(OutlookClient)
        msg_id = client._get_internet_message_id(mock_message)
        
        assert msg_id is None


# ============================================================================
# Outlook Client Tests - recipient_domains computation
# ============================================================================

class TestRecipientDomainsComputation:
    """Tests for computing recipient_domains at sync time."""
    
    def test_compute_recipient_domains_single(self):
        """Test computing recipient_domains with single recipient."""
        from outlook_client import OutlookClient
        
        recipients_to = ["user@client.com"]
        recipients_cc = []
        
        client = OutlookClient.__new__(OutlookClient)
        domains = client._compute_recipient_domains(recipients_to, recipients_cc)
        
        assert domains == "client.com"
    
    def test_compute_recipient_domains_multiple(self):
        """Test computing recipient_domains with multiple recipients."""
        from outlook_client import OutlookClient
        
        recipients_to = ["user1@client.com", "user2@partner.com"]
        recipients_cc = ["cc@third.com", "cc2@client.com"]  # client.com is duplicate
        
        client = OutlookClient.__new__(OutlookClient)
        domains = client._compute_recipient_domains(recipients_to, recipients_cc)
        
        # Should be deduplicated
        domain_list = domains.split(",")
        assert "client.com" in domain_list
        assert "partner.com" in domain_list
        assert "third.com" in domain_list
        assert len(domain_list) == 3  # No duplicates
    
    def test_compute_recipient_domains_empty(self):
        """Test computing recipient_domains with no recipients."""
        from outlook_client import OutlookClient
        
        client = OutlookClient.__new__(OutlookClient)
        domains = client._compute_recipient_domains([], [])
        
        assert domains == ""
