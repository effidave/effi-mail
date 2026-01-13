"""Comprehensive tests for scan_for_commitments and mark_scanned tools.

These tools support the EmailAssistant's commitment tracking workflow:
- scan_for_commitments: Fetches unscanned sent emails with full body for commitment detection
- mark_scanned: Marks emails as scanned with effi:scanned category
"""

import pytest
import json
from datetime import datetime, timedelta
from unittest.mock import Mock, patch

from models import Email


# ============================================================================
# Fixtures
# ============================================================================

@pytest.fixture
def sent_emails_with_commitments():
    """Create sample sent emails with various commitment patterns."""
    base_time = datetime.now()
    return [
        Email(
            id="sent-001",
            subject="Re: LA Service Agreement",
            sender_name="David Sant",
            sender_email="david.sant@harperjames.co.uk",
            domain="harperjames.co.uk",
            received_time=base_time - timedelta(hours=12),
            body_preview="I'll take a look through your LA service agreement tomorrow",
            has_attachments=False,
            attachment_names=[],
            categories="",
            folder_path="Sent Items",
        ),
        Email(
            id="sent-002",
            subject="Re: Contract review",
            sender_name="David Sant",
            sender_email="david.sant@harperjames.co.uk",
            domain="harperjames.co.uk",
            received_time=base_time - timedelta(hours=24),
            body_preview="I will send you the draft by Monday",
            has_attachments=False,
            attachment_names=[],
            categories="",
            folder_path="Sent Items",
        ),
        Email(
            id="sent-003",
            subject="Re: NDA query",
            sender_name="David Sant",
            sender_email="david.sant@harperjames.co.uk",
            domain="harperjames.co.uk",
            received_time=base_time - timedelta(hours=36),
            body_preview="Thanks for your email. No further action needed.",
            has_attachments=False,
            attachment_names=[],
            categories="",  # No commitment
            folder_path="Sent Items",
        ),
        Email(
            id="sent-004",
            subject="Re: Inline comments",
            sender_name="David Sant",
            sender_email="david.sant@harperjames.co.uk",
            domain="harperjames.co.uk",
            received_time=base_time - timedelta(hours=48),
            body_preview="I have added my comments inline below",
            has_attachments=False,
            attachment_names=[],
            categories="",  # Inline indicator - needs manual review
            folder_path="Sent Items",
        ),
    ]


@pytest.fixture
def already_scanned_email():
    """Email that has already been scanned."""
    return Email(
        id="sent-005",
        subject="Re: Already processed",
        sender_name="David Sant",
        sender_email="david.sant@harperjames.co.uk",
        domain="harperjames.co.uk",
        received_time=datetime.now() - timedelta(hours=6),
        body_preview="I'll follow up on this tomorrow",
        has_attachments=False,
        attachment_names=[],
        categories="effi:scanned",  # Already scanned
        folder_path="Sent Items",
    )


@pytest.fixture
def mock_outlook():
    """Create a mock OutlookClient."""
    mock = Mock()
    mock.search_outlook = Mock(return_value=[])
    mock.get_email_full = Mock(return_value={})
    mock.set_category = Mock(return_value=True)
    return mock


# ============================================================================
# Test: Tool Registration
# ============================================================================

class TestToolRegistration:
    """Test that commitment tools are properly registered."""
    
    @pytest.mark.asyncio
    async def test_scan_for_commitments_is_registered(self):
        """scan_for_commitments should be in the list of available tools."""
        from effi_mail import list_tools
        tools = await list_tools()
        tool_names = [t.name for t in tools]
        
        assert "scan_for_commitments" in tool_names
    
    @pytest.mark.asyncio
    async def test_mark_scanned_is_registered(self):
        """mark_scanned should be in the list of available tools."""
        from effi_mail import list_tools
        tools = await list_tools()
        tool_names = [t.name for t in tools]
        
        assert "mark_scanned" in tool_names


# ============================================================================
# Test: scan_for_commitments
# ============================================================================

class TestScanForCommitments:
    """Tests for scan_for_commitments tool."""
    
    @pytest.mark.asyncio
    async def test_returns_sent_emails_from_sent_items_folder(
        self, mock_outlook, sent_emails_with_commitments
    ):
        """Should query Sent Items folder."""
        mock_outlook.search_outlook = Mock(return_value=sent_emails_with_commitments)
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            result = await call_tool("scan_for_commitments", {"days": 7})
            
            # Verify Sent Items folder was queried
            mock_outlook.search_outlook.assert_called_once()
            call_args = mock_outlook.search_outlook.call_args
            assert call_args.kwargs.get('folder') == "Sent Items"
    
    @pytest.mark.asyncio
    async def test_excludes_already_scanned_emails(
        self, mock_outlook, sent_emails_with_commitments, already_scanned_email
    ):
        """Should exclude emails with effi:scanned category."""
        # Include an already scanned email in the results
        all_emails = sent_emails_with_commitments + [already_scanned_email]
        mock_outlook.search_outlook = Mock(return_value=all_emails)
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            result = await call_tool("scan_for_commitments", {"days": 7})
            
            data = json.loads(result[0].text)
            email_ids = [e["id"] for e in data["emails"]]
            
            # Already scanned email should be excluded
            assert "sent-005" not in email_ids
            # Unscanned emails should be included
            assert "sent-001" in email_ids
    
    @pytest.mark.asyncio
    async def test_returns_full_body_content(
        self, mock_outlook, sent_emails_with_commitments
    ):
        """Should return full email body, not just preview."""
        mock_outlook.search_outlook = Mock(return_value=sent_emails_with_commitments[:1])
        mock_outlook.get_email_full = Mock(return_value={
            "id": "sent-001",
            "subject": "Re: LA Service Agreement",
            "sender_email": "david.sant@harperjames.co.uk",
            "received_time": datetime.now().isoformat(),
            "body": "Hi Deven,\n\nI'll take a look through your LA service agreement tomorrow and get back to you.\n\nBest regards,\nDavid",
            "recipients_to": [{"name": "Deven", "email": "deven@policyinpractice.co.uk"}],
            "attachments": [],
        })
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            result = await call_tool("scan_for_commitments", {"days": 7})
            
            data = json.loads(result[0].text)
            # Should have full body, not just preview
            assert "body" in data["emails"][0]
            assert len(data["emails"][0]["body"]) > 50  # Full body is longer than preview
    
    @pytest.mark.asyncio
    async def test_returns_recipient_information(
        self, mock_outlook, sent_emails_with_commitments
    ):
        """Should include recipient info for commitment context."""
        mock_outlook.search_outlook = Mock(return_value=sent_emails_with_commitments[:1])
        mock_outlook.get_email_full = Mock(return_value={
            "id": "sent-001",
            "subject": "Re: LA Service Agreement",
            "sender_email": "david.sant@harperjames.co.uk",
            "received_time": datetime.now().isoformat(),
            "body": "I'll take a look tomorrow",
            "recipients_to": [{"name": "Deven", "email": "deven@policyinpractice.co.uk"}],
            "recipients_cc": [],
            "attachments": [],
        })
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            result = await call_tool("scan_for_commitments", {"days": 7})
            
            data = json.loads(result[0].text)
            assert "recipients_to" in data["emails"][0]
            assert data["emails"][0]["recipients_to"][0]["email"] == "deven@policyinpractice.co.uk"
    
    @pytest.mark.asyncio
    async def test_respects_days_parameter(self, mock_outlook):
        """Should use days parameter for date filtering."""
        mock_outlook.search_outlook = Mock(return_value=[])
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            await call_tool("scan_for_commitments", {"days": 14})
            
            call_args = mock_outlook.search_outlook.call_args
            assert call_args.kwargs.get('days') == 14
    
    @pytest.mark.asyncio
    async def test_respects_limit_parameter(self, mock_outlook):
        """Should use limit parameter."""
        mock_outlook.search_outlook = Mock(return_value=[])
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            await call_tool("scan_for_commitments", {"days": 7, "limit": 50})
            
            call_args = mock_outlook.search_outlook.call_args
            # Implementation fetches limit*2 to account for filtered emails
            assert call_args.kwargs.get('limit') == 100
    
    @pytest.mark.asyncio
    async def test_default_parameters(self, mock_outlook):
        """Should use sensible defaults (14 days, 100 limit)."""
        mock_outlook.search_outlook = Mock(return_value=[])
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            await call_tool("scan_for_commitments", {})
            
            call_args = mock_outlook.search_outlook.call_args
            assert call_args.kwargs.get('days') == 14
            # Implementation fetches limit*2 to account for filtered emails
            assert call_args.kwargs.get('limit') == 200
    
    @pytest.mark.asyncio
    async def test_returns_count_of_emails(
        self, mock_outlook, sent_emails_with_commitments
    ):
        """Should return count of emails found."""
        mock_outlook.search_outlook = Mock(return_value=sent_emails_with_commitments)
        mock_outlook.get_email_full = Mock(return_value={
            "id": "test",
            "subject": "Test",
            "sender_email": "test@test.com",
            "received_time": datetime.now().isoformat(),
            "body": "Test body",
            "recipients_to": [],
            "attachments": [],
        })
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            result = await call_tool("scan_for_commitments", {"days": 7})
            
            data = json.loads(result[0].text)
            assert "count" in data
            assert data["count"] == 4  # 4 unscanned emails


# ============================================================================
# Test: mark_scanned
# ============================================================================

class TestMarkScanned:
    """Tests for mark_scanned tool."""
    
    @pytest.mark.asyncio
    async def test_adds_scanned_category_to_email(self, mock_outlook):
        """Should add effi:scanned category to email."""
        mock_outlook.set_category = Mock(return_value=True)
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            result = await call_tool("mark_scanned", {"email_id": "sent-001"})
            
            mock_outlook.set_category.assert_called_once_with("sent-001", "effi:scanned")
    
    @pytest.mark.asyncio
    async def test_returns_success_on_valid_email(self, mock_outlook):
        """Should return success response."""
        mock_outlook.set_category = Mock(return_value=True)
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            result = await call_tool("mark_scanned", {"email_id": "sent-001"})
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["email_id"] == "sent-001"
    
    @pytest.mark.asyncio
    async def test_returns_error_on_failure(self, mock_outlook):
        """Should return error if category cannot be set."""
        mock_outlook.set_category = Mock(return_value=False)
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            result = await call_tool("mark_scanned", {"email_id": "invalid-id"})
            
            data = json.loads(result[0].text)
            assert data["success"] is False
            assert "error" in data
    
    @pytest.mark.asyncio
    async def test_batch_mark_scanned(self, mock_outlook):
        """Should support marking multiple emails as scanned."""
        mock_outlook.set_category = Mock(return_value=True)
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            result = await call_tool("batch_mark_scanned", {
                "email_ids": ["sent-001", "sent-002", "sent-003"]
            })
            
            data = json.loads(result[0].text)
            assert data["marked_count"] == 3
            assert data["failed_count"] == 0
            
            # Should have called set_category for each email
            assert mock_outlook.set_category.call_count == 3
    
    @pytest.mark.asyncio
    async def test_batch_mark_scanned_partial_failure(self, mock_outlook):
        """Should handle partial failures in batch mode."""
        # First two succeed, third fails
        mock_outlook.set_category = Mock(side_effect=[True, True, False])
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            result = await call_tool("batch_mark_scanned", {
                "email_ids": ["sent-001", "sent-002", "sent-003"]
            })
            
            data = json.loads(result[0].text)
            assert data["marked_count"] == 2
            assert data["failed_count"] == 1
            assert "sent-003" in data["failed_ids"]


# ============================================================================
# Test: batch_mark_scanned registration
# ============================================================================

class TestBatchMarkScannedRegistration:
    """Test that batch_mark_scanned is properly registered."""
    
    @pytest.mark.asyncio
    async def test_batch_mark_scanned_is_registered(self):
        """batch_mark_scanned should be in the list of available tools."""
        from effi_mail import list_tools
        tools = await list_tools()
        tool_names = [t.name for t in tools]
        
        assert "batch_mark_scanned" in tool_names


# ============================================================================
# Test: Integration scenarios
# ============================================================================

class TestIntegrationScenarios:
    """Test realistic usage scenarios."""
    
    @pytest.mark.asyncio
    async def test_full_workflow_scan_then_mark(
        self, mock_outlook, sent_emails_with_commitments
    ):
        """Test the full workflow: scan, process, mark as scanned."""
        # First scan returns unscanned emails
        mock_outlook.search_outlook = Mock(return_value=sent_emails_with_commitments)
        mock_outlook.get_email_full = Mock(return_value={
            "id": "sent-001",
            "subject": "Test",
            "sender_email": "test@test.com",
            "received_time": datetime.now().isoformat(),
            "body": "I'll send the draft tomorrow",
            "recipients_to": [{"name": "Test", "email": "test@example.com"}],
            "attachments": [],
        })
        mock_outlook.set_category = Mock(return_value=True)
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            
            # Step 1: Scan for commitments
            scan_result = await call_tool("scan_for_commitments", {"days": 7})
            scan_data = json.loads(scan_result[0].text)
            
            # Step 2: Mark scanned emails
            email_ids = [e["id"] for e in scan_data["emails"]]
            mark_result = await call_tool("batch_mark_scanned", {"email_ids": email_ids})
            mark_data = json.loads(mark_result[0].text)
            
            assert mark_data["marked_count"] == len(email_ids)
    
    @pytest.mark.asyncio
    async def test_rescan_excludes_previously_scanned(
        self, mock_outlook, sent_emails_with_commitments, already_scanned_email
    ):
        """Running scan again should exclude previously scanned emails."""
        # Simulate second scan where some emails now have effi:scanned
        all_emails = sent_emails_with_commitments + [already_scanned_email]
        mock_outlook.search_outlook = Mock(return_value=all_emails)
        
        with patch('effi_mail.tools.client_search.search', mock_outlook), \
             patch('effi_mail.tools.client_search.retrieval', mock_outlook), \
             patch('effi_mail.tools.client_search.folders', mock_outlook):
            from effi_mail import call_tool
            
            result = await call_tool("scan_for_commitments", {"days": 7})
            data = json.loads(result[0].text)
            
            # Already scanned email should not be in results
            email_ids = [e["id"] for e in data["emails"]]
            assert already_scanned_email.id not in email_ids
