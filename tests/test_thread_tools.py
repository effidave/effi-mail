"""Comprehensive tests for email thread tracking tools.

Tests the thread retrieval functionality that uses Exchange ConversationID
to deterministically find related emails across multiple Outlook folders.
"""

import pytest
import json
from datetime import datetime, timedelta
from unittest.mock import Mock, patch, MagicMock

from models import Email


# ============================================================================
# Fixtures
# ============================================================================

@pytest.fixture
def thread_email_1():
    """First email in a conversation thread."""
    return Email(
        id="entry-001",
        subject="Project Alpha - Initial Discussion",
        sender_name="Alice Client",
        sender_email="alice@client.com",
        domain="client.com",
        received_time=datetime.now() - timedelta(hours=3),
        body_preview="Hi David, I wanted to discuss the project scope...",
        has_attachments=False,
        attachment_names=[],
        categories="",
        conversation_id="CONV-ABC123",
        folder_path="Inbox",
        direction="inbound",
        recipients_to=["david.sant@harperjames.co.uk"],
        recipients_cc=[],
        internet_message_id="<msg001@client.com>",
    )


@pytest.fixture
def thread_email_2():
    """Second email in same conversation - outbound reply."""
    return Email(
        id="entry-002",
        subject="RE: Project Alpha - Initial Discussion",
        sender_name="David Sant",
        sender_email="david.sant@harperjames.co.uk",
        domain="client.com",
        received_time=datetime.now() - timedelta(hours=2),
        body_preview="Thanks Alice, I've reviewed the scope and...",
        has_attachments=True,
        attachment_names=["scope-review.pdf"],
        categories="",
        conversation_id="CONV-ABC123",
        folder_path="Sent Items",
        direction="outbound",
        recipients_to=["alice@client.com"],
        recipients_cc=[],
        internet_message_id="<msg002@harperjames.co.uk>",
    )


@pytest.fixture
def thread_email_3():
    """Third email in same conversation - inbound reply."""
    return Email(
        id="entry-003",
        subject="RE: Project Alpha - Initial Discussion",
        sender_name="Alice Client",
        sender_email="alice@client.com",
        domain="client.com",
        received_time=datetime.now() - timedelta(hours=1),
        body_preview="Great, that looks good. Can we schedule a call?",
        has_attachments=False,
        attachment_names=[],
        categories="",
        conversation_id="CONV-ABC123",
        folder_path="Inbox",
        direction="inbound",
        recipients_to=["david.sant@harperjames.co.uk"],
        recipients_cc=["bob@client.com"],
        internet_message_id="<msg003@client.com>",
    )


@pytest.fixture
def dms_filed_email():
    """Email from same thread, filed to DMS."""
    return Email(
        id="entry-004",
        subject="RE: Project Alpha - Initial Discussion",
        sender_name="David Sant",
        sender_email="david.sant@harperjames.co.uk",
        domain="client.com",
        received_time=datetime.now() - timedelta(minutes=30),
        body_preview="Call confirmed for tomorrow at 2pm.",
        has_attachments=False,
        attachment_names=[],
        categories="",
        conversation_id="CONV-ABC123",
        folder_path="DMS\\Client\\Alpha Project",
        direction="outbound",
        recipients_to=["alice@client.com"],
        recipients_cc=[],
        internet_message_id="<msg004@harperjames.co.uk>",
    )


@pytest.fixture
def unrelated_email():
    """Email with different ConversationID."""
    return Email(
        id="entry-999",
        subject="Different Topic",
        sender_name="Other Person",
        sender_email="other@different.com",
        domain="different.com",
        received_time=datetime.now() - timedelta(hours=1),
        body_preview="Unrelated content",
        has_attachments=False,
        attachment_names=[],
        categories="",
        conversation_id="CONV-XYZ999",
        folder_path="Inbox",
        direction="inbound",
        recipients_to=["david.sant@harperjames.co.uk"],
        recipients_cc=[],
        internet_message_id="<msg999@different.com>",
    )


@pytest.fixture
def mock_outlook():
    """Create a mock OutlookClient with thread-related methods."""
    mock = Mock()
    mock.get_email_full = Mock()
    mock.search_outlook = Mock(return_value=[])
    mock.get_emails_by_conversation_id = Mock(return_value=[])
    mock.get_conversation_id_for_email = Mock(return_value=None)
    return mock


# ============================================================================
# Tests for get_email_thread tool
# ============================================================================

class TestGetEmailThread:
    """Tests for the get_email_thread tool."""

    def test_get_thread_returns_all_messages_from_inbox_and_sent(
        self, mock_outlook, thread_email_1, thread_email_2, thread_email_3
    ):
        """Should return all emails sharing the same ConversationID."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_1.id,
            "conversation_id": "CONV-ABC123",
            "subject": thread_email_1.subject,
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [
            thread_email_1, thread_email_2, thread_email_3
        ]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            result = get_email_thread(email_id="entry-001")
            data = json.loads(result)
            
            assert data["message_count"] == 3
            assert len(data["messages"]) == 3
            mock_outlook.get_emails_by_conversation_id.assert_called_once()

    def test_get_thread_returns_messages_sorted_chronologically(
        self, mock_outlook, thread_email_1, thread_email_2, thread_email_3
    ):
        """Messages should be sorted by received_time, oldest first."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_2.id,
            "conversation_id": "CONV-ABC123",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [
            thread_email_3, thread_email_1, thread_email_2
        ]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            result = get_email_thread(email_id="entry-002")
            data = json.loads(result)
            
            messages = data["messages"]
            assert messages[0]["id"] == "entry-001"
            assert messages[1]["id"] == "entry-002"
            assert messages[2]["id"] == "entry-003"

    def test_get_thread_includes_folder_location_for_each_message(
        self, mock_outlook, thread_email_1, thread_email_2
    ):
        """Each message should include its folder path."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_1.id,
            "conversation_id": "CONV-ABC123",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [
            thread_email_1, thread_email_2
        ]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            result = get_email_thread(email_id="entry-001")
            data = json.loads(result)
            
            assert data["messages"][0]["folder"] == "Inbox"
            assert data["messages"][1]["folder"] == "Sent Items"

    def test_get_thread_includes_dms_when_requested(
        self, mock_outlook, thread_email_1, thread_email_2, dms_filed_email
    ):
        """Should include DMS folders when include_dms=True."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_1.id,
            "conversation_id": "CONV-ABC123",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [
            thread_email_1, thread_email_2, dms_filed_email
        ]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            result = get_email_thread(email_id="entry-001", include_dms=True)
            data = json.loads(result)
            
            assert data["message_count"] == 3
            folders = [m["folder"] for m in data["messages"]]
            assert "DMS\\Client\\Alpha Project" in folders

    def test_get_thread_excludes_dms_by_default(
        self, mock_outlook, thread_email_1, thread_email_2
    ):
        """DMS folders should not be searched by default."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_1.id,
            "conversation_id": "CONV-ABC123",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [
            thread_email_1, thread_email_2
        ]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            result = get_email_thread(email_id="entry-001")
            
            call_args = mock_outlook.get_emails_by_conversation_id.call_args
            assert call_args.kwargs.get("include_dms") is False

    def test_get_thread_excludes_sent_when_requested(
        self, mock_outlook, thread_email_1, thread_email_3
    ):
        """Should exclude Sent Items when include_sent=False."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_1.id,
            "conversation_id": "CONV-ABC123",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [
            thread_email_1, thread_email_3
        ]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            result = get_email_thread(email_id="entry-001", include_sent=False)
            
            call_args = mock_outlook.get_emails_by_conversation_id.call_args
            assert call_args.kwargs.get("include_sent") is False

    def test_get_thread_respects_limit(self, mock_outlook, thread_email_1):
        """Should respect the limit parameter."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_1.id,
            "conversation_id": "CONV-ABC123",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [thread_email_1]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            get_email_thread(email_id="entry-001", limit=10)
            
            call_args = mock_outlook.get_emails_by_conversation_id.call_args
            assert call_args.kwargs.get("limit") == 10

    def test_get_thread_returns_error_for_invalid_email_id(self, mock_outlook):
        """Should return error JSON if email not found."""
        mock_outlook.get_email_full.return_value = None
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            result = get_email_thread(email_id="invalid-id")
            data = json.loads(result)
            
            assert "error" in data

    def test_get_thread_returns_error_for_email_without_conversation_id(
        self, mock_outlook
    ):
        """Should return error if source email has no ConversationID."""
        mock_outlook.get_email_full.return_value = {
            "id": "entry-001",
            "conversation_id": None,
        }
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            result = get_email_thread(email_id="entry-001")
            data = json.loads(result)
            
            assert "error" in data

    def test_get_thread_includes_thread_metadata(
        self, mock_outlook, thread_email_1, thread_email_2, thread_email_3
    ):
        """Should include thread-level metadata."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_1.id,
            "conversation_id": "CONV-ABC123",
            "subject": "Project Alpha - Initial Discussion",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [
            thread_email_1, thread_email_2, thread_email_3
        ]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            result = get_email_thread(email_id="entry-001")
            data = json.loads(result)
            
            assert "conversation_id" in data
            assert "message_count" in data
            assert "participants" in data
            assert "date_range" in data
            assert data["conversation_id"] == "CONV-ABC123"

    def test_get_thread_extracts_unique_participants(
        self, mock_outlook, thread_email_1, thread_email_2, thread_email_3
    ):
        """Should list all unique participants in the thread."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_1.id,
            "conversation_id": "CONV-ABC123",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [
            thread_email_1, thread_email_2, thread_email_3
        ]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            result = get_email_thread(email_id="entry-001")
            data = json.loads(result)
            
            participants = data["participants"]
            assert "alice@client.com" in participants
            assert "david.sant@harperjames.co.uk" in participants
            assert "bob@client.com" in participants
            assert len(participants) == 3


# ============================================================================
# Tests for get_thread_locations tool
# ============================================================================

class TestGetThreadLocations:
    """Tests for the get_thread_locations tool - lightweight version."""

    def test_returns_ids_and_folders_only(
        self, mock_outlook, thread_email_1, thread_email_2, thread_email_3
    ):
        """Should return minimal data: just IDs and folder paths."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_1.id,
            "conversation_id": "CONV-ABC123",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [
            thread_email_1, thread_email_2, thread_email_3
        ]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_thread_locations
            
            result = get_thread_locations(email_id="entry-001")
            data = json.loads(result)
            
            assert "locations" in data
            assert len(data["locations"]) == 3
            
            loc = data["locations"][0]
            assert "id" in loc
            assert "folder" in loc
            assert "preview" not in loc
            assert "body" not in loc

    def test_includes_direction_for_each_message(
        self, mock_outlook, thread_email_1, thread_email_2
    ):
        """Each location should include direction (inbound/outbound)."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_1.id,
            "conversation_id": "CONV-ABC123",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [
            thread_email_1, thread_email_2
        ]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_thread_locations
            
            result = get_thread_locations(email_id="entry-001")
            data = json.loads(result)
            
            assert data["locations"][0]["direction"] == "inbound"
            assert data["locations"][1]["direction"] == "outbound"

    def test_includes_received_time_for_each_message(
        self, mock_outlook, thread_email_1, thread_email_2
    ):
        """Each location should include received timestamp."""
        mock_outlook.get_email_full.return_value = {
            "id": thread_email_1.id,
            "conversation_id": "CONV-ABC123",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [
            thread_email_1, thread_email_2
        ]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_thread_locations
            
            result = get_thread_locations(email_id="entry-001")
            data = json.loads(result)
            
            assert "received" in data["locations"][0]
            assert "received" in data["locations"][1]

    def test_returns_error_for_invalid_email(self, mock_outlook):
        """Should return error if email not found."""
        mock_outlook.get_email_full.return_value = None
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_thread_locations
            
            result = get_thread_locations(email_id="invalid-id")
            data = json.loads(result)
            
            assert "error" in data


# ============================================================================
# Tests for OutlookClient.get_emails_by_conversation_id
# ============================================================================

class TestOutlookClientGetEmailsByConversationId:
    """Tests for the OutlookClient method."""

    def test_queries_inbox_by_default(self):
        """Should query Inbox folder by default."""
        with patch('outlook_client.win32com.client') as mock_win32:
            mock_namespace = MagicMock()
            mock_outlook = MagicMock()
            mock_outlook.GetNamespace.return_value = mock_namespace
            mock_win32.Dispatch.return_value = mock_outlook
            
            mock_inbox = MagicMock()
            mock_inbox.Name = "Inbox"
            mock_inbox.Items.Restrict.return_value = []
            mock_namespace.GetDefaultFolder.return_value = mock_inbox
            
            from outlook_client import OutlookClient
            client = OutlookClient()
            
            client.get_emails_by_conversation_id("CONV-123")
            
            mock_namespace.GetDefaultFolder.assert_called()

    def test_queries_sent_items_when_include_sent_true(self):
        """Should query Sent Items when include_sent=True."""
        with patch('outlook_client.win32com.client') as mock_win32:
            mock_namespace = MagicMock()
            mock_outlook = MagicMock()
            mock_outlook.GetNamespace.return_value = mock_namespace
            mock_win32.Dispatch.return_value = mock_outlook
            
            mock_folder = MagicMock()
            mock_folder.Name = "Inbox"
            mock_folder.Items.Restrict.return_value = []
            mock_namespace.GetDefaultFolder.return_value = mock_folder
            
            from outlook_client import OutlookClient
            client = OutlookClient()
            
            client.get_emails_by_conversation_id("CONV-123", include_sent=True)
            
            calls = mock_namespace.GetDefaultFolder.call_args_list
            folder_ids = [call[0][0] for call in calls]
            assert 6 in folder_ids  # FOLDER_INBOX
            assert 5 in folder_ids  # FOLDER_SENT

    def test_uses_restrict_with_conversation_id_filter(self):
        """Should use Outlook Restrict with ConversationID filter."""
        with patch('outlook_client.win32com.client') as mock_win32:
            mock_namespace = MagicMock()
            mock_outlook = MagicMock()
            mock_outlook.GetNamespace.return_value = mock_namespace
            mock_win32.Dispatch.return_value = mock_outlook
            
            mock_items = MagicMock()
            mock_items.Restrict.return_value = []
            mock_folder = MagicMock()
            mock_folder.Name = "Inbox"
            mock_folder.Items = mock_items
            mock_namespace.GetDefaultFolder.return_value = mock_folder
            
            from outlook_client import OutlookClient
            client = OutlookClient()
            
            client.get_emails_by_conversation_id("CONV-ABC123")
            
            mock_items.Restrict.assert_called()
            filter_arg = mock_items.Restrict.call_args[0][0]
            assert "CONV-ABC123" in filter_arg


# ============================================================================
# Tests for DASL query helper
# ============================================================================

class TestBuildConversationFilter:
    """Tests for the DASL query builder helper."""

    def test_builds_valid_restrict_filter(self):
        """Should build a valid Outlook Restrict filter string."""
        from effi_mail.helpers import build_conversation_filter
        
        filter_str = build_conversation_filter("CONV-ABC123")
        
        assert "ConversationID" in filter_str
        assert "CONV-ABC123" in filter_str

    def test_escapes_special_characters(self):
        """Should escape quotes in ConversationID."""
        from effi_mail.helpers import build_conversation_filter
        
        filter_str = build_conversation_filter("CONV-ABC'123")
        
        assert "CONV-ABC''123" in filter_str

    def test_handles_empty_conversation_id(self):
        """Should raise ValueError for empty ConversationID."""
        from effi_mail.helpers import build_conversation_filter
        
        with pytest.raises(ValueError):
            build_conversation_filter("")

    def test_handles_none_conversation_id(self):
        """Should raise ValueError for None ConversationID."""
        from effi_mail.helpers import build_conversation_filter
        
        with pytest.raises(ValueError):
            build_conversation_filter(None)


# ============================================================================
# Integration tests
# ============================================================================

class TestThreadToolsIntegration:
    """Integration tests that verify the full flow."""

    def test_get_email_thread_calls_outlook_with_correct_parameters(
        self, mock_outlook, thread_email_1
    ):
        """Verify the tool correctly wires to OutlookClient."""
        mock_outlook.get_email_full.return_value = {
            "id": "entry-001",
            "conversation_id": "CONV-ABC123",
        }
        mock_outlook.get_emails_by_conversation_id.return_value = [thread_email_1]
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            get_email_thread(
                email_id="entry-001",
                include_sent=True,
                include_dms=False,
                limit=25
            )
            
            mock_outlook.get_emails_by_conversation_id.assert_called_once_with(
                conversation_id="CONV-ABC123",
                include_sent=True,
                include_dms=False,
                limit=25
            )

    def test_thread_tools_handle_outlook_exceptions(self, mock_outlook):
        """Should handle Outlook COM exceptions gracefully."""
        mock_outlook.get_email_full.side_effect = Exception("COM error")
        
        with patch('effi_mail.tools.thread.outlook', mock_outlook):
            from effi_mail.tools.thread import get_email_thread
            
            result = get_email_thread(email_id="entry-001")
            data = json.loads(result)
            
            assert "error" in data
