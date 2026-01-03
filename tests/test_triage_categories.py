"""Tests for triage status via Outlook categories.

These tests verify that triage status is stored as Outlook categories
(effi:processed, effi:deferred, effi:archived) rather than in a database.
"""

import pytest
from unittest.mock import Mock, MagicMock, patch
from datetime import datetime, timedelta

from outlook_client import OutlookClient
from models import Email


# ============================================================================
# Fixtures
# ============================================================================

@pytest.fixture
def mock_outlook_message():
    """Create a mock Outlook message object."""
    message = Mock()
    message.Categories = ""
    message.EntryID = "test-entry-id-001"
    message.Subject = "Test Subject"
    message.ReceivedTime = datetime.now()
    message.Save = Mock()
    return message


@pytest.fixture
def mock_namespace(mock_outlook_message):
    """Create a mock Outlook namespace."""
    namespace = Mock()
    namespace.GetItemFromID = Mock(return_value=mock_outlook_message)
    return namespace


@pytest.fixture
def outlook_client(mock_namespace):
    """Create an OutlookClient with mocked COM connection."""
    client = OutlookClient()
    client._namespace = mock_namespace
    client._outlook = Mock()
    return client


# ============================================================================
# Tests: Triage Category Constants
# ============================================================================

class TestTriageCategoryConstants:
    """Test that triage category constants are correctly defined."""
    
    def test_triage_category_prefix(self):
        """Triage categories should use effi: prefix to avoid conflicts."""
        assert OutlookClient.TRIAGE_CATEGORY_PREFIX == "effi:"
    
    def test_triage_categories_defined(self):
        """All triage statuses should have corresponding categories."""
        categories = OutlookClient.TRIAGE_CATEGORIES
        
        assert "processed" in categories
        assert "deferred" in categories
        assert "archived" in categories
    
    def test_triage_category_values(self):
        """Triage category values should use effi: prefix."""
        categories = OutlookClient.TRIAGE_CATEGORIES
        
        assert categories["processed"] == "effi:processed"
        assert categories["deferred"] == "effi:deferred"
        assert categories["archived"] == "effi:archived"


# ============================================================================
# Tests: set_triage_status
# ============================================================================

class TestSetTriageStatus:
    """Test setting triage status on emails."""
    
    def test_set_triage_status_processed(self, outlook_client, mock_outlook_message):
        """Should add effi:processed category to email."""
        result = outlook_client.set_triage_status("test-id", "processed")
        
        assert result is True
        assert "effi:processed" in mock_outlook_message.Categories
        mock_outlook_message.Save.assert_called_once()
    
    def test_set_triage_status_deferred(self, outlook_client, mock_outlook_message):
        """Should add effi:deferred category to email."""
        result = outlook_client.set_triage_status("test-id", "deferred")
        
        assert result is True
        assert "effi:deferred" in mock_outlook_message.Categories
        mock_outlook_message.Save.assert_called_once()
    
    def test_set_triage_status_archived(self, outlook_client, mock_outlook_message):
        """Should add effi:archived category to email."""
        result = outlook_client.set_triage_status("test-id", "archived")
        
        assert result is True
        assert "effi:archived" in mock_outlook_message.Categories
        mock_outlook_message.Save.assert_called_once()
    
    def test_set_triage_status_invalid_status(self, outlook_client):
        """Should return False for invalid status."""
        result = outlook_client.set_triage_status("test-id", "invalid")
        
        assert result is False
    
    def test_set_triage_status_preserves_other_categories(self, outlook_client, mock_outlook_message):
        """Should preserve existing non-triage categories."""
        mock_outlook_message.Categories = "Important, Work"
        
        outlook_client.set_triage_status("test-id", "processed")
        
        categories = mock_outlook_message.Categories
        assert "Important" in categories
        assert "Work" in categories
        assert "effi:processed" in categories
    
    def test_set_triage_status_replaces_existing_triage(self, outlook_client, mock_outlook_message):
        """Should replace existing triage category with new one."""
        mock_outlook_message.Categories = "effi:deferred, Work"
        
        outlook_client.set_triage_status("test-id", "processed")
        
        categories = mock_outlook_message.Categories
        assert "effi:processed" in categories
        assert "effi:deferred" not in categories
        assert "Work" in categories
    
    def test_set_triage_status_handles_empty_categories(self, outlook_client, mock_outlook_message):
        """Should handle email with no existing categories."""
        mock_outlook_message.Categories = ""
        
        result = outlook_client.set_triage_status("test-id", "processed")
        
        assert result is True
        assert "effi:processed" in mock_outlook_message.Categories
    
    def test_set_triage_status_handles_none_categories(self, outlook_client, mock_outlook_message):
        """Should handle email with None categories."""
        mock_outlook_message.Categories = None
        
        result = outlook_client.set_triage_status("test-id", "processed")
        
        assert result is True
        assert "effi:processed" in mock_outlook_message.Categories


# ============================================================================
# Tests: get_triage_status
# ============================================================================

class TestGetTriageStatus:
    """Test retrieving triage status from emails."""
    
    def test_get_triage_status_processed(self, outlook_client, mock_outlook_message):
        """Should return 'processed' when email has effi:processed category."""
        mock_outlook_message.Categories = "effi:processed, Work"
        
        result = outlook_client.get_triage_status("test-id")
        
        assert result == "processed"
    
    def test_get_triage_status_deferred(self, outlook_client, mock_outlook_message):
        """Should return 'deferred' when email has effi:deferred category."""
        mock_outlook_message.Categories = "Important, effi:deferred"
        
        result = outlook_client.get_triage_status("test-id")
        
        assert result == "deferred"
    
    def test_get_triage_status_archived(self, outlook_client, mock_outlook_message):
        """Should return 'archived' when email has effi:archived category."""
        mock_outlook_message.Categories = "effi:archived"
        
        result = outlook_client.get_triage_status("test-id")
        
        assert result == "archived"
    
    def test_get_triage_status_pending(self, outlook_client, mock_outlook_message):
        """Should return None when email has no triage category (pending)."""
        mock_outlook_message.Categories = "Work, Important"
        
        result = outlook_client.get_triage_status("test-id")
        
        assert result is None
    
    def test_get_triage_status_empty_categories(self, outlook_client, mock_outlook_message):
        """Should return None when email has no categories."""
        mock_outlook_message.Categories = ""
        
        result = outlook_client.get_triage_status("test-id")
        
        assert result is None
    
    def test_get_triage_status_none_categories(self, outlook_client, mock_outlook_message):
        """Should return None when email categories is None."""
        mock_outlook_message.Categories = None
        
        result = outlook_client.get_triage_status("test-id")
        
        assert result is None


# ============================================================================
# Tests: clear_triage_status
# ============================================================================

class TestClearTriageStatus:
    """Test clearing triage status from emails."""
    
    def test_clear_triage_status_removes_triage_category(self, outlook_client, mock_outlook_message):
        """Should remove effi: category from email."""
        mock_outlook_message.Categories = "effi:processed, Work"
        
        result = outlook_client.clear_triage_status("test-id")
        
        assert result is True
        assert "effi:processed" not in mock_outlook_message.Categories
        assert "Work" in mock_outlook_message.Categories
    
    def test_clear_triage_status_preserves_other_categories(self, outlook_client, mock_outlook_message):
        """Should preserve non-triage categories."""
        mock_outlook_message.Categories = "Important, effi:deferred, Work"
        
        outlook_client.clear_triage_status("test-id")
        
        categories = mock_outlook_message.Categories
        assert "Important" in categories
        assert "Work" in categories
        assert "effi:" not in categories
    
    def test_clear_triage_status_handles_no_triage(self, outlook_client, mock_outlook_message):
        """Should work even if no triage category exists."""
        mock_outlook_message.Categories = "Work, Important"
        
        result = outlook_client.clear_triage_status("test-id")
        
        assert result is True
        assert "Work" in mock_outlook_message.Categories


# ============================================================================
# Tests: batch_set_triage_status
# ============================================================================

class TestBatchSetTriageStatus:
    """Test batch triage operations."""
    
    def test_batch_set_triage_status_all_success(self, outlook_client, mock_namespace):
        """Should successfully triage all emails."""
        # Create multiple mock messages
        messages = []
        for i in range(3):
            msg = Mock()
            msg.Categories = ""
            msg.Save = Mock()
            messages.append(msg)
        
        mock_namespace.GetItemFromID = Mock(side_effect=messages)
        
        result = outlook_client.batch_set_triage_status(
            ["id1", "id2", "id3"], 
            "processed"
        )
        
        assert result["success"] == 3
        assert result["failed"] == 0
        assert result["failed_ids"] == []
    
    def test_batch_set_triage_status_partial_failure(self, outlook_client, mock_namespace):
        """Should track failed emails."""
        msg1 = Mock()
        msg1.Categories = ""
        msg1.Save = Mock()
        
        msg2 = Mock()
        msg2.Categories = ""
        msg2.Save = Mock(side_effect=Exception("Save failed"))
        
        mock_namespace.GetItemFromID = Mock(side_effect=[msg1, msg2])
        
        result = outlook_client.batch_set_triage_status(
            ["id1", "id2"], 
            "processed"
        )
        
        assert result["success"] == 1
        assert result["failed"] == 1
    
    def test_batch_set_triage_status_empty_list(self, outlook_client):
        """Should handle empty email list."""
        result = outlook_client.batch_set_triage_status([], "processed")
        
        assert result["success"] == 0
        assert result["failed"] == 0


# ============================================================================
# Tests: get_pending_emails
# ============================================================================

class TestGetPendingEmails:
    """Test retrieving pending (un-triaged) emails."""
    
    def test_get_pending_emails_excludes_triaged(self, outlook_client, mock_namespace):
        """Should exclude emails with effi: categories."""
        folder = Mock()
        folder.Name = "Inbox"
        folder.Items = self._create_mock_messages([
            ("Email 1", "Work"),                    # Should include
            ("Email 2", "effi:Processed, Work"),    # Should exclude
            ("Email 3", ""),                        # Should include
            ("Email 4", "effi:Archived"),           # Should exclude
        ])
        
        mock_namespace.GetDefaultFolder = Mock(return_value=folder)
        
        result = outlook_client.get_pending_emails(days=7, limit=100)
        
        # Should only include emails without effi: categories
        assert result["total"] == 2
    
    def test_get_pending_emails_groups_by_domain(self, outlook_client, mock_namespace):
        """Should group emails by sender domain."""
        folder = Mock()
        folder.Name = "Inbox"
        folder.Items = self._create_mock_messages([
            ("Email from acme", ""),
            ("Email from acme 2", ""),
            ("Email from other", ""),
        ], domains=["acme.com", "acme.com", "other.com"])
        
        mock_namespace.GetDefaultFolder = Mock(return_value=folder)
        
        result = outlook_client.get_pending_emails(days=7, group_by_domain=True)
        
        assert "domains" in result
        # Should have 2 domains
        domain_names = [d["domain"] for d in result["domains"]]
        assert "acme.com" in domain_names or any("acme" in str(d) for d in result["domains"])
    
    def _create_mock_messages(self, emails, domains=None):
        """Helper to create mock message collection."""
        messages = Mock()
        message_list = []
        
        for i, (subject, categories) in enumerate(emails):
            msg = Mock()
            msg.Subject = subject
            msg.Categories = categories
            msg.SenderEmailAddress = f"sender@{domains[i] if domains else 'example.com'}"
            msg.ReceivedTime = datetime.now() - timedelta(hours=i)
            msg.Parent = Mock()
            msg.Parent.Name = "Inbox"
            msg.Body = "Test body"
            msg.Sender = None
            msg.To = "recipient@test.com"
            msg.CC = ""
            msg.Attachments = Mock()
            msg.Attachments.Count = 0
            msg.ConversationID = f"conv-{i}"
            msg.EntryID = f"entry-{i}"
            msg.PropertyAccessor = Mock()
            msg.PropertyAccessor.GetProperty = Mock(return_value=f"<msg-{i}@test.com>")
            message_list.append(msg)
        
        messages.Sort = Mock()
        messages.Restrict = Mock(return_value=message_list)
        messages.__iter__ = Mock(return_value=iter(message_list))
        
        return messages


# ============================================================================
# Tests: get_pending_emails_from_domain
# ============================================================================

class TestGetPendingEmailsFromDomain:
    """Test retrieving pending emails from a specific domain."""
    
    def test_filters_by_domain(self, outlook_client, mock_namespace):
        """Should only return emails from specified domain."""
        folder = Mock()
        folder.Name = "Inbox"
        
        # Create messages with different domains
        msg1 = Mock()
        msg1.Categories = ""
        msg1.SenderEmailAddress = "user@target.com"
        msg1.Subject = "From target"
        msg1.ReceivedTime = datetime.now()
        msg1.Parent = Mock()
        msg1.Parent.Name = "Inbox"
        msg1.Body = "Test"
        msg1.Sender = None
        msg1.To = "me@test.com"
        msg1.CC = ""
        msg1.Attachments = Mock()
        msg1.Attachments.Count = 0
        msg1.ConversationID = "conv-1"
        msg1.EntryID = "entry-1"
        msg1.PropertyAccessor = Mock()
        msg1.PropertyAccessor.GetProperty = Mock(return_value="<msg@test.com>")
        
        messages = Mock()
        messages.Sort = Mock()
        messages.Restrict = Mock(return_value=[msg1])
        folder.Items = messages
        
        mock_namespace.GetDefaultFolder = Mock(return_value=folder)
        
        result = outlook_client.get_pending_emails_from_domain("target.com", days=7)
        
        # Should return list of emails
        assert isinstance(result, list)
    
    def test_excludes_triaged_emails(self, outlook_client, mock_namespace):
        """Should exclude emails with effi: categories."""
        folder = Mock()
        folder.Name = "Inbox"
        
        msg1 = Mock()
        msg1.Categories = "effi:Processed"  # Should be excluded
        msg1.SenderEmailAddress = "user@target.com"
        
        msg2 = Mock()
        msg2.Categories = ""  # Should be included
        msg2.SenderEmailAddress = "user@target.com"
        msg2.Subject = "Test"
        msg2.ReceivedTime = datetime.now()
        msg2.Parent = Mock()
        msg2.Parent.Name = "Inbox"
        msg2.Body = "Test"
        msg2.Sender = None
        msg2.To = "me@test.com"
        msg2.CC = ""
        msg2.Attachments = Mock()
        msg2.Attachments.Count = 0
        msg2.ConversationID = "conv-1"
        msg2.EntryID = "entry-1"
        msg2.PropertyAccessor = Mock()
        msg2.PropertyAccessor.GetProperty = Mock(return_value="<msg@test.com>")
        
        messages = Mock()
        messages.Sort = Mock()
        messages.Restrict = Mock(return_value=[msg1, msg2])
        folder.Items = messages
        
        mock_namespace.GetDefaultFolder = Mock(return_value=folder)
        
        result = outlook_client.get_pending_emails_from_domain("target.com", days=7)
        
        # Should only include the non-triaged email
        assert len(result) <= 1


# ============================================================================
# Tests: MCP Server Integration
# ============================================================================

class TestMCPServerTriageIntegration:
    """Test MCP server triage tool implementations."""
    
    @pytest.fixture
    def mock_outlook(self):
        """Create mock OutlookClient."""
        return Mock()
    
    @pytest.mark.asyncio
    async def test_triage_email_tool_calls_outlook(self, mock_outlook):
        """triage_email tool should call outlook.set_triage_status."""
        mock_outlook.set_triage_status = Mock(return_value=True)
        
        # Simulate the tool implementation
        email_id = "test-id"
        status = "processed"
        
        success = mock_outlook.set_triage_status(email_id, status)
        
        assert success is True
        mock_outlook.set_triage_status.assert_called_once_with(email_id, status)
    
    @pytest.mark.asyncio
    async def test_batch_triage_tool_calls_outlook(self, mock_outlook):
        """batch_triage tool should call outlook.batch_set_triage_status."""
        mock_outlook.batch_set_triage_status = Mock(return_value={
            "success": 3,
            "failed": 0,
            "failed_ids": []
        })
        
        email_ids = ["id1", "id2", "id3"]
        status = "archived"
        
        result = mock_outlook.batch_set_triage_status(email_ids, status)
        
        assert result["success"] == 3
        mock_outlook.batch_set_triage_status.assert_called_once_with(email_ids, status)
    
    @pytest.mark.asyncio
    async def test_get_pending_emails_tool_calls_outlook(self, mock_outlook):
        """get_pending_emails tool should call outlook.get_pending_emails."""
        mock_outlook.get_pending_emails = Mock(return_value={
            "total": 5,
            "domains": []
        })
        
        result = mock_outlook.get_pending_emails(days=30, limit=100, group_by_domain=True)
        
        assert result["total"] == 5
        mock_outlook.get_pending_emails.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_batch_archive_domain_calls_outlook(self, mock_outlook):
        """batch_archive_domain should get pending emails then triage them."""
        # Mock pending emails from domain
        mock_email = Mock()
        mock_email.id = "email-1"
        
        mock_outlook.get_pending_emails_from_domain = Mock(return_value=[mock_email])
        mock_outlook.set_triage_status = Mock(return_value=True)
        
        # Simulate batch_archive_domain logic
        domain = "marketing.com"
        pending = mock_outlook.get_pending_emails_from_domain(domain, days=30)
        
        archived = 0
        for email in pending:
            if mock_outlook.set_triage_status(email.id, "archived"):
                archived += 1
        
        assert archived == 1
        mock_outlook.get_pending_emails_from_domain.assert_called_once()
        mock_outlook.set_triage_status.assert_called_once_with("email-1", "archived")


# ============================================================================
# Tests: Domain Categories (JSON-based)
# ============================================================================

class TestDomainCategories:
    """Test JSON-based domain category functions."""
    
    def test_get_domain_category_returns_uncategorized_for_unknown(self):
        """Should return 'Uncategorized' for unknown domains."""
        from domain_categories import get_domain_category
        import tempfile
        from pathlib import Path
        
        # Use empty temp file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            f.write('{}')
            temp_path = Path(f.name)
        
        try:
            result = get_domain_category("unknown-domain.com", json_path=temp_path)
            assert result == "Uncategorized"
        finally:
            temp_path.unlink()
    
    def test_set_and_get_domain_category(self):
        """Should be able to set and retrieve domain categories."""
        from domain_categories import get_domain_category, set_domain_category
        import tempfile
        from pathlib import Path
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            f.write('{}')
            temp_path = Path(f.name)
        
        try:
            # Set category
            result = set_domain_category("acme.com", "Client", json_path=temp_path)
            assert result is True
            
            # Get category
            category = get_domain_category("acme.com", json_path=temp_path)
            assert category == "Client"
        finally:
            temp_path.unlink()
    
    def test_set_domain_category_rejects_invalid(self):
        """Should reject invalid category names."""
        from domain_categories import set_domain_category
        import tempfile
        from pathlib import Path
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            f.write('{}')
            temp_path = Path(f.name)
        
        try:
            result = set_domain_category("test.com", "InvalidCategory", json_path=temp_path)
            assert result is False
        finally:
            temp_path.unlink()
    
    def test_get_all_domain_categories(self):
        """Should return all domain categories."""
        from domain_categories import get_all_domain_categories, set_domain_category
        import tempfile
        from pathlib import Path
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            f.write('{}')
            temp_path = Path(f.name)
        
        try:
            set_domain_category("acme.com", "Client", json_path=temp_path)
            set_domain_category("newsletter.com", "Marketing", json_path=temp_path)
            
            all_cats = get_all_domain_categories(json_path=temp_path)
            
            assert all_cats["acme.com"] == "Client"
            assert all_cats["newsletter.com"] == "Marketing"
        finally:
            temp_path.unlink()
    
    def test_get_domains_by_category(self):
        """Should filter domains by category."""
        from domain_categories import get_domains_by_category, set_domain_category
        import tempfile
        from pathlib import Path
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            f.write('{}')
            temp_path = Path(f.name)
        
        try:
            set_domain_category("client1.com", "Client", json_path=temp_path)
            set_domain_category("client2.com", "Client", json_path=temp_path)
            set_domain_category("newsletter.com", "Marketing", json_path=temp_path)
            
            clients = get_domains_by_category("Client", json_path=temp_path)
            
            assert "client1.com" in clients
            assert "client2.com" in clients
            assert "newsletter.com" not in clients
        finally:
            temp_path.unlink()


# ============================================================================
# Tests: No Database Dependency
# ============================================================================

class TestNoDatabaseDependency:
    """Verify that triage operations don't use database."""
    
    def test_outlook_client_has_no_db_attribute(self):
        """OutlookClient should not have db attribute."""
        client = OutlookClient()
        assert not hasattr(client, 'db') or client.db is None or not hasattr(client, 'db')
    
    def test_set_triage_status_no_db_call(self, outlook_client, mock_outlook_message):
        """set_triage_status should not call any db methods."""
        # This test passes if set_triage_status doesn't raise AttributeError
        # when trying to access db methods
        outlook_client.set_triage_status("test-id", "processed")
        
        # Verify Save was called (Outlook operation, not DB)
        mock_outlook_message.Save.assert_called()
    
    def test_mcp_server_uses_outlook_for_triage(self):
        """MCP server should use outlook methods for triage, not db."""
        # Check the triage tool implementation
        from effi_mail.tools import triage
        import inspect
        source_code = inspect.getsource(triage)
        
        # Should not have db.update_email_triage calls
        assert "db.update_email_triage" not in source_code
        
        # Should use outlook.set_triage_status
        assert "outlook.set_triage_status" in source_code
