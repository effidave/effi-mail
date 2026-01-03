"""Comprehensive tests for MCP server tools (database-free architecture).

This test file tests the effi-mail tools with the new architecture:
- Triage status stored as Outlook categories
- Domain categories stored in domain_categories.json
- Client data retrieved from effi-core MCP server
- All email queries go directly to Outlook COM
"""

import pytest
import json
import asyncio
from datetime import datetime, timedelta
from unittest.mock import Mock, MagicMock, patch, AsyncMock
import tempfile
import os

from models import Email
from effi_mail import list_tools, call_tool
from effi_mail.helpers import truncate_text, format_email_summary


# ============================================================================
# Fixtures
# ============================================================================

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
    )


@pytest.fixture
def sample_emails():
    """Create multiple sample emails."""
    emails = []
    for i in range(5):
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
        ))
    return emails


@pytest.fixture
def mock_outlook():
    """Create a mock OutlookClient."""
    mock = Mock()
    
    # Triage methods
    mock.set_triage_status = Mock(return_value=True)
    mock.get_triage_status = Mock(return_value=None)
    mock.clear_triage_status = Mock(return_value=True)
    mock.batch_set_triage_status = Mock(return_value={"success": 3, "failed": 0, "failed_ids": []})
    
    # Email retrieval methods
    mock.get_pending_emails = Mock(return_value={
        "total": 5,
        "domains": [
            {"domain": "example.com", "count": 3, "emails": []},
            {"domain": "other.com", "count": 2, "emails": []}
        ]
    })
    mock.get_pending_emails_from_domain = Mock(return_value=[])
    mock.search_outlook = Mock(return_value=[])
    mock.search_outlook_by_identifiers = Mock(return_value=[])
    mock.get_email_full = Mock(return_value={
        "subject": "Test",
        "sender": "test@example.com",
        "received": datetime.now().isoformat(),
        "body": "Email body",
        "attachments": []
    })
    
    # DMS methods
    mock.list_dms_clients = Mock(return_value=[])
    mock.list_dms_matters = Mock(return_value=[])
    mock.get_dms_emails = Mock(return_value=[])
    mock.search_dms_emails = Mock(return_value=[])
    mock.get_domain_counts = Mock(return_value={"domains": []})
    
    return mock


from contextlib import contextmanager

@contextmanager
def patch_outlook(mock_outlook):
    """Patch outlook in all tool modules."""
    with patch('effi_mail.helpers.outlook', mock_outlook), \
         patch('effi_mail.tools.triage.outlook', mock_outlook), \
         patch('effi_mail.tools.email_retrieval.outlook', mock_outlook), \
         patch('effi_mail.tools.domain_categories.outlook', mock_outlook), \
         patch('effi_mail.tools.client_search.outlook', mock_outlook), \
         patch('effi_mail.tools.dms.outlook', mock_outlook):
        yield


# ============================================================================
# Test: Tool Definitions
# ============================================================================

class TestToolDefinitions:
    """Test that correct tools are defined."""
    
    @pytest.mark.asyncio
    async def test_list_tools_includes_triage_tools(self):
        """Should include triage tools."""
        tools = await list_tools()
        tool_names = [t.name for t in tools]
        
        assert "triage_email" in tool_names
        assert "batch_triage" in tool_names
        assert "batch_archive_domain" in tool_names
    
    @pytest.mark.asyncio
    async def test_list_tools_includes_domain_tools(self):
        """Should include domain categorization tools."""
        tools = await list_tools()
        tool_names = [t.name for t in tools]
        
        assert "get_uncategorized_domains" in tool_names
        assert "categorize_domain" in tool_names
        assert "get_domain_summary" in tool_names
    
    @pytest.mark.asyncio
    async def test_list_tools_includes_email_retrieval_tools(self):
        """Should include email retrieval tools."""
        tools = await list_tools()
        tool_names = [t.name for t in tools]
        
        assert "get_pending_emails" in tool_names
        assert "get_inbox_emails_by_domain" in tool_names
        assert "get_email_by_id" in tool_names
    
    @pytest.mark.asyncio
    async def test_list_tools_includes_search_tools(self):
        """Should include client search tools."""
        tools = await list_tools()
        tool_names = [t.name for t in tools]
        
        assert "get_emails_by_client" in tool_names
        assert "search_outlook_direct" in tool_names
    
    @pytest.mark.asyncio
    async def test_list_tools_excludes_removed_tools(self):
        """Should NOT include removed database-based tools."""
        tools = await list_tools()
        tool_names = [t.name for t in tools]
        
        # These tools should have been removed
        assert "sync_emails" not in tool_names
        assert "sync_emails_by_client" not in tool_names
        assert "sync_email_by_id" not in tool_names
        assert "create_client" not in tool_names
        assert "create_matter" not in tool_names
        assert "list_clients" not in tool_names
        assert "get_triage_stats" not in tool_names


# ============================================================================
# Test: Triage Tool Handlers
# ============================================================================

class TestTriageToolHandlers:
    """Test triage tool handlers."""
    
    @pytest.mark.asyncio
    async def test_triage_email_calls_outlook(self, mock_outlook):
        """triage_email should call outlook.set_triage_status."""
        with patch_outlook(mock_outlook):
            result = await call_tool("triage_email", {
                "email_id": "test-id",
                "status": "processed"
            })
            
            mock_outlook.set_triage_status.assert_called_once_with("test-id", "processed")
            
            data = json.loads(result[0].text)
            assert data["success"] is True
    
    @pytest.mark.asyncio
    async def test_triage_email_handles_failure(self, mock_outlook):
        """triage_email should handle Outlook failures."""
        mock_outlook.set_triage_status = Mock(return_value=False)
        
        with patch_outlook(mock_outlook):
            result = await call_tool("triage_email", {
                "email_id": "test-id",
                "status": "processed"
            })
            
            data = json.loads(result[0].text)
            assert "error" in data
    
    @pytest.mark.asyncio
    async def test_batch_triage_calls_outlook(self, mock_outlook):
        """batch_triage should call outlook.batch_set_triage_status."""
        with patch_outlook(mock_outlook):
            result = await call_tool("batch_triage", {
                "email_ids": ["id1", "id2", "id3"],
                "status": "archived"
            })
            
            mock_outlook.batch_set_triage_status.assert_called_once_with(
                ["id1", "id2", "id3"], 
                "archived"
            )
            
            data = json.loads(result[0].text)
            assert data["triaged"] == 3
    
    @pytest.mark.asyncio
    async def test_batch_archive_domain(self, mock_outlook, sample_emails):
        """batch_archive_domain should archive all pending emails from domain."""
        mock_outlook.get_pending_emails_from_domain = Mock(return_value=sample_emails[:2])
        
        with patch_outlook(mock_outlook):
            result = await call_tool("batch_archive_domain", {
                "domain": "marketing.com",
                "days": 30
            })
            
            mock_outlook.get_pending_emails_from_domain.assert_called_once_with(
                "marketing.com", days=30
            )
            
            # Should have called set_triage_status for each email
            assert mock_outlook.set_triage_status.call_count == 2
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["archived_count"] == 2


# ============================================================================
# Test: Email Retrieval Tool Handlers
# ============================================================================

class TestEmailRetrievalHandlers:
    """Test email retrieval tool handlers."""
    
    @pytest.mark.asyncio
    async def test_get_pending_emails_calls_outlook(self, mock_outlook):
        """get_pending_emails should call outlook.get_pending_emails."""
        with patch_outlook(mock_outlook):
            result = await call_tool("get_pending_emails", {
                "days": 14,
                "limit": 50
            })
            
            mock_outlook.get_pending_emails.assert_called_once()
            
            # Check call arguments
            call_args = mock_outlook.get_pending_emails.call_args
            assert call_args.kwargs["days"] == 14
            assert call_args.kwargs["limit"] == 50
    
    @pytest.mark.asyncio
    async def test_get_inbox_emails_by_domain_calls_outlook(self, mock_outlook, sample_emails):
        """get_inbox_emails_by_domain should call outlook.search_outlook."""
        mock_outlook.search_outlook = Mock(return_value=sample_emails[:2])
        
        with patch_outlook(mock_outlook):
            result = await call_tool("get_inbox_emails_by_domain", {
                "domain": "example.com",
                "limit": 20
            })
            
            mock_outlook.search_outlook.assert_called_once()
            call_args = mock_outlook.search_outlook.call_args
            assert call_args.kwargs["sender_domain"] == "example.com"
    
    @pytest.mark.asyncio
    async def test_get_email_by_id_calls_outlook(self, mock_outlook):
        """get_email_by_id should call outlook.get_email_full."""
        with patch_outlook(mock_outlook):
            result = await call_tool("get_email_by_id", {
                "email_id": "test-id",
                "include_body": True,
                "include_attachments": True
            })
            
            mock_outlook.get_email_full.assert_called_once_with("test-id")
    
    @pytest.mark.asyncio
    async def test_get_email_by_id_with_max_body_length(self, mock_outlook):
        """get_email_by_id should truncate body when max_body_length is set."""
        mock_outlook.get_email_full = Mock(return_value={
            "subject": "Test",
            "body": "A" * 10000,
            "attachments": []
        })
        
        with patch_outlook(mock_outlook):
            result = await call_tool("get_email_by_id", {
                "email_id": "test-id",
                "max_body_length": 100
            })
            
            data = json.loads(result[0].text)
            # Body should be truncated - format is "AAA... [X more chars]"
            assert len(data["body"]) < 10000
            assert "more chars" in data["body"]


# ============================================================================
# Test: Domain Categorization Tool Handlers
# ============================================================================

class TestDomainToolHandlers:
    """Test domain categorization tool handlers."""
    
    @pytest.mark.asyncio
    async def test_categorize_domain_saves_to_json(self, mock_outlook):
        """categorize_domain should save to domain_categories.json."""
        import tempfile
        from pathlib import Path
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            f.write('{}')
            temp_path = Path(f.name)
        
        try:
            with patch_outlook(mock_outlook), \
                 patch('effi_mail.tools.domain_categories.set_domain_category') as mock_set:
                mock_set.return_value = True
                
                result = await call_tool("categorize_domain", {
                    "domain": "acme.com",
                    "category": "Client"
                })
                
                mock_set.assert_called_once_with("acme.com", "Client")
                
                data = json.loads(result[0].text)
                assert data["success"] is True
                assert data["category"] == "Client"
        finally:
            temp_path.unlink()
    
    @pytest.mark.asyncio
    async def test_get_domain_summary_reads_from_json(self, mock_outlook):
        """get_domain_summary should read from domain_categories.json."""
        with patch_outlook(mock_outlook), \
             patch('effi_mail.tools.domain_categories.get_all_domain_categories') as mock_get:
            mock_get.return_value = {
                "acme.com": "Client",
                "internal.com": "Internal",
                "newsletter.com": "Marketing"
            }
            
            result = await call_tool("get_domain_summary", {})
            
            mock_get.assert_called_once()
            
            data = json.loads(result[0].text)
            assert "Client" in data or "Marketing" in data


# ============================================================================
# Test: Client Search Tool Handlers
# ============================================================================

class TestClientSearchHandlers:
    """Test client search tool handlers."""
    
    @pytest.fixture
    def mock_effi_work_client(self):
        """Mock the effi_work_client module."""
        async def mock_get_identifiers(client_id):
            return {
                "source": "effi-core",
                "client_id": client_id,
                "domains": ["acme.com", "acme.co.uk"],
                "contact_emails": ["ceo@personal.com"]
            }
        return mock_get_identifiers
    
    @pytest.mark.asyncio
    async def test_get_emails_by_client_calls_outlook(self, mock_outlook, mock_effi_work_client, sample_emails):
        """get_emails_by_client should search Outlook with client domains."""
        mock_outlook.search_outlook_by_identifiers = Mock(return_value=sample_emails[:2])
        
        with patch_outlook(mock_outlook), \
             patch('effi_mail.tools.client_search.get_client_identifiers_from_effi_work', mock_effi_work_client):
            result = await call_tool("get_emails_by_client", {
                "client_id": "acme-corp",
                "days": 30
            })
            
            mock_outlook.search_outlook_by_identifiers.assert_called_once()
            
            data = json.loads(result[0].text)
            assert data["client_id"] == "acme-corp"
            assert "identifiers" in data
    
    @pytest.mark.asyncio
    async def test_search_outlook_direct(self, mock_outlook, sample_emails):
        """search_outlook_direct should search with flexible filters."""
        mock_outlook.search_outlook = Mock(return_value=sample_emails[:3])
        
        with patch_outlook(mock_outlook):
            result = await call_tool("search_outlook_direct", {
                "sender_domain": "example.com",
                "subject_contains": "invoice",
                "days": 30,
                "folder": "Inbox"
            })
            
            mock_outlook.search_outlook.assert_called_once()
            
            data = json.loads(result[0].text)
            assert data["count"] == 3


# ============================================================================
# Test: Error Handling
# ============================================================================

class TestErrorHandling:
    """Test error handling in tool handlers."""
    
    @pytest.mark.asyncio
    async def test_unknown_tool_returns_error(self, mock_outlook):
        """Unknown tool should return error."""
        with patch_outlook(mock_outlook):
            result = await call_tool("nonexistent_tool", {})
            
            data = json.loads(result[0].text)
            assert "error" in data
            assert "Unknown tool" in data["error"]
    
    @pytest.mark.asyncio
    async def test_exception_returns_error(self, mock_outlook):
        """Exceptions should be caught and returned as errors."""
        mock_outlook.get_pending_emails = Mock(side_effect=Exception("COM error"))
        
        with patch_outlook(mock_outlook):
            result = await call_tool("get_pending_emails", {"days": 7})
            
            data = json.loads(result[0].text)
            assert "error" in data


# ============================================================================
# Test: Helper Functions
# ============================================================================

class TestHelperFunctions:
    """Test helper functions."""
    
    def test_truncate_text_short(self):
        """Short text should not be truncated."""
        result = truncate_text("Short text", 100)
        assert result == "Short text"
    
    def test_truncate_text_long(self):
        """Long text should be truncated with indicator."""
        long_text = "A" * 200
        result = truncate_text(long_text, 50)
        
        assert len(result) < 200
        assert "more chars" in result
    
    def test_format_email_summary(self, sample_email, mock_outlook):
        """format_email_summary should format email correctly."""
        mock_outlook.get_triage_status = Mock(return_value=None)
        
        with patch_outlook(mock_outlook):
            result = format_email_summary(sample_email)
            
            assert result["id"] == sample_email.id
            assert result["subject"] == sample_email.subject
            assert sample_email.sender_email in result["sender"]
            assert result["domain"] == sample_email.domain
    
    def test_format_email_summary_with_triage(self, sample_email, mock_outlook):
        """format_email_summary should include triage status from Outlook."""
        mock_outlook.get_triage_status = Mock(return_value="processed")
        
        with patch_outlook(mock_outlook):
            result = format_email_summary(sample_email)
            
            assert result.get("triage_status") == "processed"


# ============================================================================
# Integration Tests
# ============================================================================

class TestIntegration:
    """Integration tests for the tools."""
    
    @pytest.mark.asyncio
    async def test_triage_workflow(self, mock_outlook, sample_emails):
        """Test complete triage workflow."""
        # Setup: emails are pending (no triage category)
        mock_outlook.get_pending_emails = Mock(return_value={
            "total": 2,
            "domains": [{"domain": "example.com", "count": 2, "emails": sample_emails[:2]}]
        })
        # Override batch_set_triage_status to return correct count for this test
        mock_outlook.batch_set_triage_status = Mock(return_value={"success": 1, "failed": 0, "failed_ids": []})
        
        with patch_outlook(mock_outlook):
            # 1. Get pending emails
            result = await call_tool("get_pending_emails", {"days": 7})
            data = json.loads(result[0].text)
            assert data["total_pending"] == 2
            
            # 2. Triage one email as processed
            result = await call_tool("triage_email", {
                "email_id": sample_emails[0].id,
                "status": "processed"
            })
            data = json.loads(result[0].text)
            assert data["success"] is True
            
            # 3. Batch triage remaining
            result = await call_tool("batch_triage", {
                "email_ids": [sample_emails[1].id],
                "status": "archived"
            })
            data = json.loads(result[0].text)
            assert data["triaged"] == 1
