"""Comprehensive tests for file_email_to_dms and batch_file_emails_to_dms tools.

Tests filing emails from Inbox/Sent Items to DMSforLegal folders.
Workflow: Copy email to DMS, add "Filed" category, mark as "effi:processed".
"""

import pytest
import json
from datetime import datetime, timedelta
from unittest.mock import Mock, MagicMock, patch, call
from contextlib import contextmanager
import sys
import os

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models import Email
from effi_mail import call_tool


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
# Fixtures
# ============================================================================

@pytest.fixture
def mock_message():
    """Create a mock Outlook message object."""
    msg = Mock()
    msg.Subject = "Contract Review Request"
    msg.SenderName = "John Smith"
    msg.SenderEmailAddress = "john.smith@acme.com"
    msg.ReceivedTime = datetime.now() - timedelta(days=1)
    msg.EntryID = "original-email-001"
    msg.Body = "Please review the attached contract."
    msg.HTMLBody = "<html><body>Please review the attached contract.</body></html>"
    msg.Categories = ""
    msg.ConversationID = "conv-001"
    msg.Attachments = Mock()
    msg.Attachments.Count = 1
    msg.Sender = Mock()
    msg.Sender.AddressEntryUserType = 1  # Not Exchange
    msg.PropertyAccessor = Mock()
    msg.PropertyAccessor.GetProperty = Mock(return_value="<msg-001@acme.com>")
    msg.Recipients = Mock()
    msg.Recipients.Count = 0
    msg.Save = Mock()
    
    # Copy returns a new message with new EntryID
    copied_msg = Mock()
    copied_msg.EntryID = "filed-email-001"
    copied_msg.Subject = msg.Subject
    copied_msg.ReceivedTime = msg.ReceivedTime
    copied_msg.Move = Mock(return_value=copied_msg)
    msg.Copy = Mock(return_value=copied_msg)
    
    return msg


@pytest.fixture
def mock_dms_folder_structure():
    """Create a mock DMS folder structure for filing tests."""
    
    def create_folder_mock(name, subfolders=None, items=None):
        """Helper to create a mock folder."""
        folder = Mock()
        folder.Name = name
        folder.Folders = Mock()
        
        if subfolders:
            folder.Folders.__iter__ = Mock(side_effect=lambda: iter(subfolders))
            folder.Folders.Count = len(subfolders)
            
            def get_folder(name):
                for sf in subfolders:
                    if sf.Name == name:
                        return sf
                raise Exception(f"Folder '{name}' not found")
            folder.Folders.__getitem__ = get_folder
        else:
            folder.Folders.__iter__ = Mock(side_effect=lambda: iter([]))
            folder.Folders.Count = 0
        
        if items:
            folder.Items = Mock()
            folder.Items.__iter__ = Mock(side_effect=lambda: iter(items))
            folder.Items.Count = len(items)
        else:
            folder.Items = Mock()
            folder.Items.__iter__ = Mock(side_effect=lambda: iter([]))
            folder.Items.Count = 0
        
        return folder
    
    # Build folder hierarchy for "Acme Corporation" / "Widget Agreement (12345)"
    emails_folder = create_folder_mock("Emails", items=[])
    admin_folder = create_folder_mock("Admin")
    documents_folder = create_folder_mock("Documents")
    
    matter1 = create_folder_mock(
        "Widget Agreement (12345)",
        subfolders=[admin_folder, documents_folder, emails_folder]
    )
    
    # Matter without Emails folder (should cause error)
    matter_no_emails = create_folder_mock(
        "Broken Matter (99999)",
        subfolders=[create_folder_mock("Admin"), create_folder_mock("Documents")]
    )
    
    client1 = create_folder_mock(
        "Acme Corporation",
        subfolders=[matter1, matter_no_emails]
    )
    
    # _My Matters folder
    my_matters = create_folder_mock("_My Matters", subfolders=[client1])
    
    # Root folder of DMSforLegal store
    root = create_folder_mock("DMSforLegal", subfolders=[my_matters])
    
    return {
        "root": root,
        "my_matters": my_matters,
        "emails_folder": emails_folder,
        "client": client1,
        "matter": matter1,
    }


@pytest.fixture
def mock_outlook_for_filing(mock_dms_folder_structure, mock_message):
    """Create a mock OutlookClient configured for filing tests."""
    mock = Mock()
    
    # DMS listing methods (for validation)
    mock.list_dms_clients = Mock(return_value=["Acme Corporation"])
    mock.list_dms_matters = Mock(return_value=[
        "Widget Agreement (12345)",
        "Broken Matter (99999)"
    ])
    
    # File email method (to be implemented)
    mock.file_email_to_dms = Mock(return_value={
        "success": True,
        "filed_entry_id": "filed-email-001",
        "subject": "Contract Review Request",
        "received_time": (datetime.now() - timedelta(days=1)).isoformat(),
        "filed_category": "Filed",
        "triage_status": "processed"
    })
    
    # Batch file method (to be implemented)
    mock.batch_file_emails_to_dms = Mock(return_value={
        "success": True,
        "filed_count": 3,
        "failed_count": 0,
        "filed_emails": [
            {"entry_id": "filed-001", "subject": "Email 1"},
            {"entry_id": "filed-002", "subject": "Email 2"},
            {"entry_id": "filed-003", "subject": "Email 3"},
        ],
        "failed_emails": []
    })
    
    # Triage methods
    mock.set_triage_status = Mock(return_value=True)
    
    # Category methods
    mock.add_category = Mock(return_value=True)
    
    return mock


# ============================================================================
# Unit Tests - OutlookClient.file_email_to_dms
# ============================================================================

class TestOutlookClientFileEmailToDMS:
    """Unit tests for OutlookClient.file_email_to_dms method."""
    
    def test_file_email_copies_to_dms_folder(self, mock_dms_folder_structure, mock_message):
        """Filing should copy email to DMS Emails folder."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        client._outlook = Mock()
        client._namespace = Mock()
        client._namespace.GetItemFromID = Mock(return_value=mock_message)
        
        # Mock DMS store access
        dms_store = Mock()
        dms_store.DisplayName = "DMSforLegal"
        dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
        
        stores = Mock()
        stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
        client._namespace.Stores = stores
        
        result = client.file_email_to_dms(
            email_id="original-email-001",
            client_name="Acme Corporation",
            matter_name="Widget Agreement (12345)"
        )
        
        # Verify email was copied
        mock_message.Copy.assert_called_once()
        
        # Verify copy was moved to Emails folder
        copied = mock_message.Copy.return_value
        copied.Move.assert_called_once_with(mock_dms_folder_structure["emails_folder"])
        
        # Verify result contains filed email info
        assert result["success"] is True
        assert result["filed_entry_id"] == "filed-email-001"
        assert result["subject"] == "Contract Review Request"
    
    def test_file_email_adds_filed_category(self, mock_dms_folder_structure, mock_message):
        """Filing should add 'Filed' category to original email."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        client._outlook = Mock()
        client._namespace = Mock()
        client._namespace.GetItemFromID = Mock(return_value=mock_message)
        
        # Mock DMS store access
        dms_store = Mock()
        dms_store.DisplayName = "DMSforLegal"
        dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
        
        stores = Mock()
        stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
        client._namespace.Stores = stores
        
        result = client.file_email_to_dms(
            email_id="original-email-001",
            client_name="Acme Corporation",
            matter_name="Widget Agreement (12345)"
        )
        
        # Verify Filed category was added
        assert "Filed" in mock_message.Categories
        mock_message.Save.assert_called()
    
    def test_file_email_marks_as_processed(self, mock_dms_folder_structure, mock_message):
        """Filing should mark original email as effi:processed."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        client._outlook = Mock()
        client._namespace = Mock()
        client._namespace.GetItemFromID = Mock(return_value=mock_message)
        
        # Mock DMS store access
        dms_store = Mock()
        dms_store.DisplayName = "DMSforLegal"
        dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
        
        stores = Mock()
        stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
        client._namespace.Stores = stores
        
        result = client.file_email_to_dms(
            email_id="original-email-001",
            client_name="Acme Corporation",
            matter_name="Widget Agreement (12345)"
        )
        
        # Verify triage status is in categories
        assert "effi:processed" in mock_message.Categories
        assert result["triage_status"] == "processed"
    
    def test_file_email_error_when_client_not_found(self, mock_dms_folder_structure, mock_message):
        """Filing should return error if client folder doesn't exist."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        client._outlook = Mock()
        client._namespace = Mock()
        client._namespace.GetItemFromID = Mock(return_value=mock_message)
        
        # Mock DMS store access
        dms_store = Mock()
        dms_store.DisplayName = "DMSforLegal"
        dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
        
        stores = Mock()
        stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
        client._namespace.Stores = stores
        
        result = client.file_email_to_dms(
            email_id="original-email-001",
            client_name="Nonexistent Client",
            matter_name="Some Matter"
        )
        
        assert result["success"] is False
        assert "client" in result["error"].lower() or "not found" in result["error"].lower()
    
    def test_file_email_error_when_matter_not_found(self, mock_dms_folder_structure, mock_message):
        """Filing should return error if matter folder doesn't exist."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        client._outlook = Mock()
        client._namespace = Mock()
        client._namespace.GetItemFromID = Mock(return_value=mock_message)
        
        # Mock DMS store access
        dms_store = Mock()
        dms_store.DisplayName = "DMSforLegal"
        dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
        
        stores = Mock()
        stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
        client._namespace.Stores = stores
        
        result = client.file_email_to_dms(
            email_id="original-email-001",
            client_name="Acme Corporation",
            matter_name="Nonexistent Matter"
        )
        
        assert result["success"] is False
        assert "matter" in result["error"].lower() or "not found" in result["error"].lower()
    
    def test_file_email_error_when_emails_folder_missing(self, mock_dms_folder_structure, mock_message):
        """Filing should return error if Emails subfolder doesn't exist."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        client._outlook = Mock()
        client._namespace = Mock()
        client._namespace.GetItemFromID = Mock(return_value=mock_message)
        
        # Mock DMS store access
        dms_store = Mock()
        dms_store.DisplayName = "DMSforLegal"
        dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
        
        stores = Mock()
        stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
        client._namespace.Stores = stores
        
        result = client.file_email_to_dms(
            email_id="original-email-001",
            client_name="Acme Corporation",
            matter_name="Broken Matter (99999)"
        )
        
        assert result["success"] is False
        assert "emails" in result["error"].lower() or "folder" in result["error"].lower()
    
    def test_file_email_error_when_email_not_found(self, mock_dms_folder_structure):
        """Filing should return error if email_id is invalid."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        client._outlook = Mock()
        client._namespace = Mock()
        client._namespace.GetItemFromID = Mock(side_effect=Exception("Item not found"))
        
        # Mock DMS store access
        dms_store = Mock()
        dms_store.DisplayName = "DMSforLegal"
        dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
        
        stores = Mock()
        stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
        client._namespace.Stores = stores
        
        result = client.file_email_to_dms(
            email_id="invalid-email-id",
            client_name="Acme Corporation",
            matter_name="Widget Agreement (12345)"
        )
        
        assert result["success"] is False
        assert "email" in result["error"].lower() or "not found" in result["error"].lower()


# ============================================================================
# Unit Tests - OutlookClient.batch_file_emails_to_dms
# ============================================================================

class TestOutlookClientBatchFileEmailsToDMS:
    """Unit tests for OutlookClient.batch_file_emails_to_dms method."""
    
    def test_batch_file_multiple_emails(self, mock_dms_folder_structure):
        """Batch filing should file multiple emails to same matter."""
        from outlook_client import OutlookClient
        
        # Create multiple mock messages
        messages = []
        for i in range(3):
            msg = Mock()
            msg.Subject = f"Email Subject {i}"
            msg.ReceivedTime = datetime.now() - timedelta(hours=i)
            msg.EntryID = f"original-{i}"
            msg.Categories = ""
            msg.Save = Mock()
            
            copied = Mock()
            copied.EntryID = f"filed-{i}"
            copied.Subject = msg.Subject
            copied.ReceivedTime = msg.ReceivedTime
            copied.Move = Mock(return_value=copied)
            msg.Copy = Mock(return_value=copied)
            
            messages.append(msg)
        
        client = OutlookClient()
        client._outlook = Mock()
        client._namespace = Mock()
        client._namespace.GetItemFromID = Mock(side_effect=messages)
        
        # Mock DMS store access
        dms_store = Mock()
        dms_store.DisplayName = "DMSforLegal"
        dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
        
        stores = Mock()
        stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
        client._namespace.Stores = stores
        
        result = client.batch_file_emails_to_dms(
            email_ids=["original-0", "original-1", "original-2"],
            client_name="Acme Corporation",
            matter_name="Widget Agreement (12345)"
        )
        
        assert result["success"] is True
        assert result["filed_count"] == 3
        assert result["failed_count"] == 0
        assert len(result["filed_emails"]) == 3
    
    def test_batch_file_validates_folder_once(self, mock_dms_folder_structure):
        """Batch filing should validate DMS folder exists once upfront."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        client._outlook = Mock()
        client._namespace = Mock()
        
        # Mock DMS store access
        dms_store = Mock()
        dms_store.DisplayName = "DMSforLegal"
        dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
        
        stores = Mock()
        stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
        client._namespace.Stores = stores
        
        result = client.batch_file_emails_to_dms(
            email_ids=["email-1", "email-2"],
            client_name="Nonexistent Client",
            matter_name="Some Matter"
        )
        
        # Should fail immediately without trying to process emails
        assert result["success"] is False
        assert "client" in result["error"].lower() or "not found" in result["error"].lower()
        # GetItemFromID should not have been called since folder validation failed
        client._namespace.GetItemFromID.assert_not_called()
    
    def test_batch_file_continues_on_single_email_error(self, mock_dms_folder_structure):
        """Batch filing should continue processing if one email fails."""
        from outlook_client import OutlookClient
        
        # First message succeeds, second fails, third succeeds
        msg1 = Mock()
        msg1.Subject = "Email 1"
        msg1.ReceivedTime = datetime.now()
        msg1.EntryID = "orig-1"
        msg1.Categories = ""
        msg1.Save = Mock()
        copied1 = Mock()
        copied1.EntryID = "filed-1"
        copied1.Subject = msg1.Subject
        copied1.ReceivedTime = msg1.ReceivedTime
        copied1.Move = Mock(return_value=copied1)
        msg1.Copy = Mock(return_value=copied1)
        
        msg3 = Mock()
        msg3.Subject = "Email 3"
        msg3.ReceivedTime = datetime.now()
        msg3.EntryID = "orig-3"
        msg3.Categories = ""
        msg3.Save = Mock()
        copied3 = Mock()
        copied3.EntryID = "filed-3"
        copied3.Subject = msg3.Subject
        copied3.ReceivedTime = msg3.ReceivedTime
        copied3.Move = Mock(return_value=copied3)
        msg3.Copy = Mock(return_value=copied3)
        
        def get_item(email_id):
            if email_id == "orig-1":
                return msg1
            elif email_id == "orig-2":
                raise Exception("Email not found")
            elif email_id == "orig-3":
                return msg3
        
        client = OutlookClient()
        client._outlook = Mock()
        client._namespace = Mock()
        client._namespace.GetItemFromID = Mock(side_effect=get_item)
        
        # Mock DMS store access
        dms_store = Mock()
        dms_store.DisplayName = "DMSforLegal"
        dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
        
        stores = Mock()
        stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
        client._namespace.Stores = stores
        
        result = client.batch_file_emails_to_dms(
            email_ids=["orig-1", "orig-2", "orig-3"],
            client_name="Acme Corporation",
            matter_name="Widget Agreement (12345)"
        )
        
        # Overall success is True if any succeeded
        assert result["filed_count"] == 2
        assert result["failed_count"] == 1
        assert len(result["failed_emails"]) == 1
        assert result["failed_emails"][0]["email_id"] == "orig-2"
    
    def test_batch_file_empty_list(self, mock_dms_folder_structure):
        """Batch filing with empty list should return appropriate response."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        client._outlook = Mock()
        client._namespace = Mock()
        
        # Mock DMS store access
        dms_store = Mock()
        dms_store.DisplayName = "DMSforLegal"
        dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
        
        stores = Mock()
        stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
        client._namespace.Stores = stores
        
        result = client.batch_file_emails_to_dms(
            email_ids=[],
            client_name="Acme Corporation",
            matter_name="Widget Agreement (12345)"
        )
        
        assert result["filed_count"] == 0
        assert result["failed_count"] == 0


# ============================================================================
# MCP Tool Tests - file_email_to_dms
# ============================================================================

class TestMCPFileEmailToDMS:
    """Tests for file_email_to_dms MCP tool."""
    
    @pytest.mark.asyncio
    async def test_file_email_tool_success(self, mock_outlook_for_filing):
        """file_email_to_dms tool should return success with filed email details."""
        with patch_outlook(mock_outlook_for_filing):
            result = await call_tool("file_email_to_dms", {
                "email_id": "test-email-001",
                "client": "Acme Corporation",
                "matter": "Widget Agreement (12345)"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert "filed_entry_id" in data
            assert "subject" in data
    
    @pytest.mark.asyncio
    async def test_file_email_tool_validates_client(self, mock_outlook_for_filing):
        """file_email_to_dms tool should validate client exists."""
        mock_outlook_for_filing.list_dms_clients = Mock(return_value=["Other Client"])
        mock_outlook_for_filing.file_email_to_dms = Mock(return_value={
            "success": False,
            "error": "Client 'Acme Corporation' not found in DMS"
        })
        
        with patch_outlook(mock_outlook_for_filing):
            result = await call_tool("file_email_to_dms", {
                "email_id": "test-email-001",
                "client": "Acme Corporation",
                "matter": "Widget Agreement (12345)"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is False
            assert "error" in data
    
    @pytest.mark.asyncio
    async def test_file_email_tool_validates_matter(self, mock_outlook_for_filing):
        """file_email_to_dms tool should validate matter exists."""
        mock_outlook_for_filing.list_dms_matters = Mock(return_value=["Other Matter"])
        mock_outlook_for_filing.file_email_to_dms = Mock(return_value={
            "success": False,
            "error": "Matter 'Widget Agreement (12345)' not found for client 'Acme Corporation'"
        })
        
        with patch_outlook(mock_outlook_for_filing):
            result = await call_tool("file_email_to_dms", {
                "email_id": "test-email-001",
                "client": "Acme Corporation",
                "matter": "Widget Agreement (12345)"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is False
            assert "error" in data
    
    @pytest.mark.asyncio
    async def test_file_email_tool_missing_parameters(self, mock_outlook_for_filing):
        """file_email_to_dms tool should require all parameters."""
        with patch_outlook(mock_outlook_for_filing):
            # Missing email_id
            result = await call_tool("file_email_to_dms", {
                "client": "Acme Corporation",
                "matter": "Widget Agreement (12345)"
            })
            
            data = json.loads(result[0].text)
            assert "error" in data


# ============================================================================
# MCP Tool Tests - batch_file_emails_to_dms
# ============================================================================

class TestMCPBatchFileEmailsToDMS:
    """Tests for batch_file_emails_to_dms MCP tool."""
    
    @pytest.mark.asyncio
    async def test_batch_file_tool_success(self, mock_outlook_for_filing):
        """batch_file_emails_to_dms tool should file multiple emails."""
        with patch_outlook(mock_outlook_for_filing):
            result = await call_tool("batch_file_emails_to_dms", {
                "email_ids": ["email-001", "email-002", "email-003"],
                "client": "Acme Corporation",
                "matter": "Widget Agreement (12345)"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is True
            assert data["filed_count"] == 3
            assert "filed_emails" in data
    
    @pytest.mark.asyncio
    async def test_batch_file_tool_partial_failure(self, mock_outlook_for_filing):
        """batch_file_emails_to_dms tool should report partial failures."""
        mock_outlook_for_filing.batch_file_emails_to_dms = Mock(return_value={
            "success": True,
            "filed_count": 2,
            "failed_count": 1,
            "filed_emails": [
                {"entry_id": "filed-001", "subject": "Email 1"},
                {"entry_id": "filed-003", "subject": "Email 3"},
            ],
            "failed_emails": [
                {"email_id": "email-002", "error": "Email not found"}
            ]
        })
        
        with patch_outlook(mock_outlook_for_filing):
            result = await call_tool("batch_file_emails_to_dms", {
                "email_ids": ["email-001", "email-002", "email-003"],
                "client": "Acme Corporation",
                "matter": "Widget Agreement (12345)"
            })
            
            data = json.loads(result[0].text)
            assert data["filed_count"] == 2
            assert data["failed_count"] == 1
            assert len(data["failed_emails"]) == 1
    
    @pytest.mark.asyncio
    async def test_batch_file_tool_validates_folder_upfront(self, mock_outlook_for_filing):
        """batch_file_emails_to_dms tool should validate folder before processing."""
        mock_outlook_for_filing.list_dms_clients = Mock(return_value=[])
        mock_outlook_for_filing.batch_file_emails_to_dms = Mock(return_value={
            "success": False,
            "error": "Client 'Acme Corporation' not found in DMS"
        })
        
        with patch_outlook(mock_outlook_for_filing):
            result = await call_tool("batch_file_emails_to_dms", {
                "email_ids": ["email-001", "email-002"],
                "client": "Acme Corporation",
                "matter": "Widget Agreement (12345)"
            })
            
            data = json.loads(result[0].text)
            assert data["success"] is False
            assert "error" in data
    
    @pytest.mark.asyncio
    async def test_batch_file_tool_empty_list(self, mock_outlook_for_filing):
        """batch_file_emails_to_dms tool should handle empty email list."""
        mock_outlook_for_filing.batch_file_emails_to_dms = Mock(return_value={
            "success": True,
            "filed_count": 0,
            "failed_count": 0,
            "filed_emails": [],
            "failed_emails": []
        })
        
        with patch_outlook(mock_outlook_for_filing):
            result = await call_tool("batch_file_emails_to_dms", {
                "email_ids": [],
                "client": "Acme Corporation",
                "matter": "Widget Agreement (12345)"
            })
            
            data = json.loads(result[0].text)
            assert data["filed_count"] == 0


# ============================================================================
# Integration Tests (require live Outlook)
# ============================================================================

class TestFileEmailIntegration:
    """Integration tests for email filing (require actual Outlook connection)."""
    
    @pytest.mark.skip(reason="Requires live Outlook connection")
    def test_live_file_email_to_dms(self):
        """Test actual email filing to DMS."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        
        # Get a recent email from Inbox
        inbox = client._namespace.GetDefaultFolder(6)  # olFolderInbox
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        
        if items.Count > 0:
            email = items[0]
            print(f"Would file: {email.Subject}")
            # Don't actually file in test
    
    @pytest.mark.skip(reason="Requires live Outlook connection")
    def test_live_dms_folder_structure(self):
        """Verify DMS folder structure exists."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        clients = client.list_dms_clients()
        
        print(f"DMS Clients: {clients}")
        
        if clients:
            matters = client.list_dms_matters(clients[0])
            print(f"Matters for {clients[0]}: {matters}")


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
