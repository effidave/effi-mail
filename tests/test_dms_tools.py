"""Comprehensive tests for DMS (DMSforLegal) tools.

Tests for read-only access to emails filed in the DMSforLegal Outlook store.
Structure: \\\\DMSforLegal\\_My Matters\\{Client}\\{Matter}\\Emails
"""

import pytest
import json
import asyncio
from datetime import datetime, timedelta
from unittest.mock import Mock, MagicMock, patch, PropertyMock
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
def mock_dms_folder_structure():
    """Create a mock DMS folder structure matching DMSforLegal layout."""
    
    def create_folder_mock(name, subfolders=None, items=None):
        """Helper to create a mock folder."""
        folder = Mock()
        folder.Name = name
        folder.Folders = Mock()
        
        if subfolders:
            # Use a lambda to create fresh iterator each time
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
            # Use a lambda to create fresh iterator each time
            folder.Items.__iter__ = Mock(side_effect=lambda: iter(items))
            folder.Items.Count = len(items)
            folder.Items.Sort = Mock()
            folder.Items.Restrict = Mock(return_value=Mock(
                __iter__=Mock(side_effect=lambda: iter(items)),
                Count=len(items)
            ))
        else:
            folder.Items = Mock()
            folder.Items.__iter__ = Mock(side_effect=lambda: iter([]))
            folder.Items.Count = 0
        
        return folder
    
    def create_message_mock(subject, sender, received_time, entry_id):
        """Helper to create a mock message."""
        msg = Mock()
        msg.Subject = subject
        msg.SenderName = sender
        msg.SenderEmailAddress = f"{sender.lower().replace(' ', '.')}@example.com"
        msg.ReceivedTime = received_time
        msg.EntryID = entry_id
        msg.Body = f"Body of {subject}"
        msg.HTMLBody = f"<html><body>{subject}</body></html>"
        msg.Categories = ""
        msg.ConversationID = "conv-001"
        msg.Attachments = Mock()
        msg.Attachments.Count = 0
        msg.Sender = Mock()
        msg.Sender.AddressEntryUserType = 1  # Not Exchange
        msg.PropertyAccessor = Mock()
        msg.PropertyAccessor.GetProperty = Mock(return_value=f"<msg-{entry_id}@example.com>")
        msg.Recipients = Mock()
        msg.Recipients.Count = 0
        return msg
    
    # Create sample messages
    messages = [
        create_message_mock(
            "Contract Review Request",
            "John Smith",
            datetime.now() - timedelta(days=1),
            "dms-email-001"
        ),
        create_message_mock(
            "RE: Contract Review Request",
            "David Sant",
            datetime.now() - timedelta(hours=12),
            "dms-email-002"
        ),
    ]
    
    # Build folder hierarchy
    # Client: "Acme Corporation" -> Matter: "Widget Agreement (12345)" -> Emails
    emails_folder = create_folder_mock("Emails", items=messages)
    admin_folder = create_folder_mock("Admin")
    documents_folder = create_folder_mock("Documents")
    
    matter1 = create_folder_mock(
        "Widget Agreement (12345)",
        subfolders=[admin_folder, documents_folder, emails_folder]
    )
    
    # Second matter with no emails
    matter2_emails = create_folder_mock("Emails", items=[])
    matter2 = create_folder_mock(
        "Consulting Services (12346)",
        subfolders=[create_folder_mock("Admin"), create_folder_mock("Documents"), matter2_emails]
    )
    
    client1 = create_folder_mock(
        "Acme Corporation",
        subfolders=[matter1, matter2]
    )
    
    # Second client
    client2_matter_emails = create_folder_mock("Emails", items=[
        create_message_mock(
            "NDA Discussion",
            "Jane Doe",
            datetime.now() - timedelta(days=3),
            "dms-email-003"
        )
    ])
    client2_matter = create_folder_mock(
        "IP License (23456)",
        subfolders=[create_folder_mock("Admin"), create_folder_mock("Documents"), client2_matter_emails]
    )
    client2 = create_folder_mock(
        "Beta Industries Ltd",
        subfolders=[client2_matter]
    )
    
    # _My Matters folder
    my_matters = create_folder_mock("_My Matters", subfolders=[client1, client2])
    
    # Root folder of DMSforLegal store
    root = create_folder_mock("DMSforLegal", subfolders=[my_matters])
    
    return {
        "root": root,
        "my_matters": my_matters,
        "clients": [client1, client2],
        "messages": messages,
    }


@pytest.fixture
def mock_outlook_with_dms(mock_dms_folder_structure):
    """Create a mock OutlookClient with DMS store access."""
    
    # Import here to avoid issues
    from outlook_client import OutlookClient
    
    client = OutlookClient()
    client._outlook = Mock()
    client._namespace = Mock()
    
    # Mock the Stores collection
    dms_store = Mock()
    dms_store.DisplayName = "DMSforLegal"
    dms_store.GetRootFolder = Mock(return_value=mock_dms_folder_structure["root"])
    
    # Stores collection mock - use side_effect for reusable iterator
    stores = Mock()
    stores.__iter__ = Mock(side_effect=lambda: iter([dms_store]))
    stores.Count = 1
    
    def get_store(name):
        if name == "DMSforLegal":
            return dms_store
        raise Exception(f"Store '{name}' not found")
    
    stores.__getitem__ = get_store
    client._namespace.Stores = stores
    
    return client


# ============================================================================
# Tests: OutlookClient DMS Methods
# ============================================================================

class TestDMSStoreMethods:
    """Tests for internal DMS store access methods."""
    
    def test_get_dms_store_returns_store(self, mock_outlook_with_dms):
        """_get_dms_store should return the DMSforLegal store."""
        store = mock_outlook_with_dms._get_dms_store()
        assert store is not None
        assert store.DisplayName == "DMSforLegal"
    
    def test_get_dms_store_not_found(self, mock_outlook_with_dms):
        """_get_dms_store should return None if store doesn't exist."""
        # Make stores empty
        mock_outlook_with_dms._namespace.Stores.__iter__ = Mock(return_value=iter([]))
        mock_outlook_with_dms._namespace.Stores.Count = 0
        
        store = mock_outlook_with_dms._get_dms_store()
        assert store is None
    
    def test_get_folder_by_path_valid(self, mock_outlook_with_dms):
        """_get_folder_by_path should navigate to nested folder."""
        folder = mock_outlook_with_dms._get_folder_by_path(
            "_My Matters\\Acme Corporation\\Widget Agreement (12345)\\Emails"
        )
        assert folder is not None
        assert folder.Name == "Emails"
    
    def test_get_folder_by_path_invalid(self, mock_outlook_with_dms):
        """_get_folder_by_path should return None for invalid path."""
        folder = mock_outlook_with_dms._get_folder_by_path(
            "_My Matters\\NonExistent Client\\Some Matter\\Emails"
        )
        assert folder is None


class TestListDMSClients:
    """Tests for list_dms_clients method."""
    
    def test_list_dms_clients_returns_client_names(self, mock_outlook_with_dms):
        """list_dms_clients should return all client folder names."""
        clients = mock_outlook_with_dms.list_dms_clients()
        
        assert len(clients) == 2
        assert "Acme Corporation" in clients
        assert "Beta Industries Ltd" in clients
    
    def test_list_dms_clients_empty_when_no_store(self, mock_outlook_with_dms):
        """list_dms_clients should return empty list if DMS store not found."""
        mock_outlook_with_dms._namespace.Stores.__iter__ = Mock(return_value=iter([]))
        
        clients = mock_outlook_with_dms.list_dms_clients()
        assert clients == []
    
    def test_list_dms_clients_sorted_alphabetically(self, mock_outlook_with_dms):
        """list_dms_clients should return clients sorted alphabetically."""
        clients = mock_outlook_with_dms.list_dms_clients()
        
        assert clients == sorted(clients)


class TestListDMSMatters:
    """Tests for list_dms_matters method."""
    
    def test_list_dms_matters_returns_matter_names(self, mock_outlook_with_dms):
        """list_dms_matters should return matter folder names for a client."""
        matters = mock_outlook_with_dms.list_dms_matters("Acme Corporation")
        
        assert len(matters) == 2
        assert "Widget Agreement (12345)" in matters
        assert "Consulting Services (12346)" in matters
    
    def test_list_dms_matters_client_not_found(self, mock_outlook_with_dms):
        """list_dms_matters should return empty list for non-existent client."""
        matters = mock_outlook_with_dms.list_dms_matters("NonExistent Client")
        
        assert matters == []
    
    def test_list_dms_matters_sorted_alphabetically(self, mock_outlook_with_dms):
        """list_dms_matters should return matters sorted alphabetically."""
        matters = mock_outlook_with_dms.list_dms_matters("Acme Corporation")
        
        assert matters == sorted(matters)


class TestGetDMSEmails:
    """Tests for get_dms_emails method."""
    
    def test_get_dms_emails_returns_emails(self, mock_outlook_with_dms):
        """get_dms_emails should return emails from matter's Emails folder."""
        emails = mock_outlook_with_dms.get_dms_emails(
            "Acme Corporation",
            "Widget Agreement (12345)"
        )
        
        assert len(emails) == 2
        assert all(isinstance(e, Email) for e in emails)
        assert emails[0].subject == "Contract Review Request"
    
    def test_get_dms_emails_with_limit(self, mock_outlook_with_dms):
        """get_dms_emails should respect limit parameter."""
        emails = mock_outlook_with_dms.get_dms_emails(
            "Acme Corporation",
            "Widget Agreement (12345)",
            limit=1
        )
        
        assert len(emails) == 1
    
    def test_get_dms_emails_empty_folder(self, mock_outlook_with_dms):
        """get_dms_emails should return empty list for matter with no emails."""
        emails = mock_outlook_with_dms.get_dms_emails(
            "Acme Corporation",
            "Consulting Services (12346)"
        )
        
        assert emails == []
    
    def test_get_dms_emails_client_not_found(self, mock_outlook_with_dms):
        """get_dms_emails should return empty list for non-existent client."""
        emails = mock_outlook_with_dms.get_dms_emails(
            "NonExistent Client",
            "Some Matter"
        )
        
        assert emails == []
    
    def test_get_dms_emails_matter_not_found(self, mock_outlook_with_dms):
        """get_dms_emails should return empty list for non-existent matter."""
        emails = mock_outlook_with_dms.get_dms_emails(
            "Acme Corporation",
            "NonExistent Matter"
        )
        
        assert emails == []


class TestSearchDMSEmails:
    """Tests for search_dms_emails method."""
    
    def test_search_dms_emails_by_client_only(self, mock_outlook_with_dms):
        """search_dms_emails with only client should search all matters."""
        emails = mock_outlook_with_dms.search_dms_emails(client="Acme Corporation")
        
        # Should find emails from both matters
        assert len(emails) == 2
    
    def test_search_dms_emails_by_client_and_matter(self, mock_outlook_with_dms):
        """search_dms_emails with client and matter should filter appropriately."""
        emails = mock_outlook_with_dms.search_dms_emails(
            client="Acme Corporation",
            matter="Widget Agreement (12345)"
        )
        
        assert len(emails) == 2
    
    def test_search_dms_emails_by_subject(self, mock_outlook_with_dms):
        """search_dms_emails should filter by subject text."""
        emails = mock_outlook_with_dms.search_dms_emails(
            client="Acme Corporation",
            subject_contains="Contract"
        )
        
        assert len(emails) >= 1
        assert all("Contract" in e.subject for e in emails)
    
    def test_search_dms_emails_across_all_clients(self, mock_outlook_with_dms):
        """search_dms_emails without client should search all clients."""
        emails = mock_outlook_with_dms.search_dms_emails()
        
        # Should find emails from all clients/matters
        assert len(emails) >= 3  # 2 from Acme + 1 from Beta
    
    def test_search_dms_emails_with_limit(self, mock_outlook_with_dms):
        """search_dms_emails should respect limit parameter."""
        emails = mock_outlook_with_dms.search_dms_emails(limit=1)
        
        assert len(emails) == 1


# ============================================================================
# Tests: MCP Server DMS Tools
# ============================================================================

class TestMCPListDMSClients:
    """Tests for list_dms_clients MCP tool."""
    
    @pytest.mark.asyncio
    async def test_list_dms_clients_tool(self, mock_outlook_with_dms):
        """list_dms_clients tool should return client list."""
        with patch_outlook(mock_outlook_with_dms):
            # call_tool imported at top
            
            result = await call_tool("list_dms_clients", {})
            
            assert len(result) == 1
            data = json.loads(result[0].text)
            assert "clients" in data
            assert len(data["clients"]) == 2
            assert "Acme Corporation" in data["clients"]
    
    @pytest.mark.asyncio
    async def test_list_dms_clients_tool_empty(self, mock_outlook_with_dms):
        """list_dms_clients tool should handle empty store gracefully."""
        mock_outlook_with_dms._namespace.Stores.__iter__ = Mock(return_value=iter([]))
        
        with patch_outlook(mock_outlook_with_dms):
            # call_tool imported at top
            
            result = await call_tool("list_dms_clients", {})
            
            data = json.loads(result[0].text)
            assert data["clients"] == []


class TestMCPListDMSMatters:
    """Tests for list_dms_matters MCP tool."""
    
    @pytest.mark.asyncio
    async def test_list_dms_matters_tool(self, mock_outlook_with_dms):
        """list_dms_matters tool should return matters for a client."""
        with patch_outlook(mock_outlook_with_dms):
            # call_tool imported at top
            
            result = await call_tool("list_dms_matters", {"client": "Acme Corporation"})
            
            data = json.loads(result[0].text)
            assert "matters" in data
            assert len(data["matters"]) == 2
    
    @pytest.mark.asyncio
    async def test_list_dms_matters_tool_missing_client_param(self, mock_outlook_with_dms):
        """list_dms_matters tool should require client parameter."""
        with patch_outlook(mock_outlook_with_dms):
            # call_tool imported at top
            
            result = await call_tool("list_dms_matters", {})
            
            data = json.loads(result[0].text)
            assert "error" in data


class TestMCPGetDMSEmails:
    """Tests for get_dms_emails MCP tool."""
    
    @pytest.mark.asyncio
    async def test_get_dms_emails_tool(self, mock_outlook_with_dms):
        """get_dms_emails tool should return emails for client/matter."""
        with patch_outlook(mock_outlook_with_dms):
            # call_tool imported at top
            
            result = await call_tool("get_dms_emails", {
                "client": "Acme Corporation",
                "matter": "Widget Agreement (12345)"
            })
            
            data = json.loads(result[0].text)
            assert "emails" in data
            assert len(data["emails"]) == 2
    
    @pytest.mark.asyncio
    async def test_get_dms_emails_tool_with_limit(self, mock_outlook_with_dms):
        """get_dms_emails tool should respect limit parameter."""
        with patch_outlook(mock_outlook_with_dms):
            # call_tool imported at top
            
            result = await call_tool("get_dms_emails", {
                "client": "Acme Corporation",
                "matter": "Widget Agreement (12345)",
                "limit": 1
            })
            
            data = json.loads(result[0].text)
            assert len(data["emails"]) == 1


class TestMCPSearchDMS:
    """Tests for search_dms MCP tool."""
    
    @pytest.mark.asyncio
    async def test_search_dms_tool_by_client(self, mock_outlook_with_dms):
        """search_dms tool should search by client."""
        with patch_outlook(mock_outlook_with_dms):
            # call_tool imported at top
            
            result = await call_tool("search_dms", {"client": "Acme Corporation"})
            
            data = json.loads(result[0].text)
            assert "emails" in data
            assert len(data["emails"]) >= 1
    
    @pytest.mark.asyncio
    async def test_search_dms_tool_by_subject(self, mock_outlook_with_dms):
        """search_dms tool should filter by subject."""
        with patch_outlook(mock_outlook_with_dms):
            # call_tool imported at top
            
            result = await call_tool("search_dms", {
                "client": "Acme Corporation",
                "subject_contains": "Contract"
            })
            
            data = json.loads(result[0].text)
            assert all("Contract" in e["subject"] for e in data["emails"])
    
    @pytest.mark.asyncio
    async def test_search_dms_tool_all_clients(self, mock_outlook_with_dms):
        """search_dms tool without client should search all."""
        with patch_outlook(mock_outlook_with_dms):
            # call_tool imported at top
            
            result = await call_tool("search_dms", {})
            
            data = json.loads(result[0].text)
            assert len(data["emails"]) >= 3  # Emails from multiple clients


# ============================================================================
# Integration Tests
# ============================================================================

class TestDMSIntegration:
    """Integration tests for DMS tools (require actual Outlook connection)."""
    
    @pytest.mark.skip(reason="Requires live Outlook connection")
    def test_live_dms_store_access(self):
        """Test actual DMS store access."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        clients = client.list_dms_clients()
        
        assert isinstance(clients, list)
        # Should have at least one client if DMS is configured
        print(f"Found {len(clients)} clients in DMS")
    
    @pytest.mark.skip(reason="Requires live Outlook connection")
    def test_live_dms_matter_listing(self):
        """Test actual matter listing."""
        from outlook_client import OutlookClient
        
        client = OutlookClient()
        clients = client.list_dms_clients()
        
        if clients:
            matters = client.list_dms_matters(clients[0])
            assert isinstance(matters, list)
            print(f"Client '{clients[0]}' has {len(matters)} matters")


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
