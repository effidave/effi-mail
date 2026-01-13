"""Tests for email ingestion module - pure functions only.

Tests cover:
- load_seen_ids / save_seen_ids: Persistent tracking of processed emails
- extract_new_content: Thread content separation

Note: build_email_markdown tests are skipped on non-Windows platforms
as the ingest module requires win32com.
"""

import pytest
import json
import tempfile
import sys
import os
from datetime import datetime
from pathlib import Path
import importlib.util

# Direct import without going through package __init__.py to avoid Windows dependencies
def import_module_from_path(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

# Get the project root
project_root = Path(__file__).parent.parent
ingestion_path = project_root / "effi_mail" / "ingestion"

# Import modules directly
storage = import_module_from_path("storage", ingestion_path / "storage.py")
thread_parser = import_module_from_path("thread_parser", ingestion_path / "thread_parser.py")

load_seen_ids = storage.load_seen_ids
save_seen_ids = storage.save_seen_ids
extract_new_content = thread_parser.extract_new_content

# Try to import build_email_markdown but skip if not on Windows
try:
    ingest = import_module_from_path("ingest", ingestion_path / "ingest.py")
    build_email_markdown = ingest.build_email_markdown
    HAS_WINDOWS_DEPS = True
except (ImportError, ModuleNotFoundError):
    HAS_WINDOWS_DEPS = False
    build_email_markdown = None


class TestSeenIdTracking:
    """Test seen ID persistence."""
    
    def test_empty_inbox_returns_empty_set(self):
        """Loading from non-existent _seen.json returns empty set."""
        with tempfile.TemporaryDirectory() as tmpdir:
            inbox_path = Path(tmpdir)
            seen = load_seen_ids(inbox_path)
            assert seen == set()
    
    def test_save_and_load_seen_ids(self):
        """Saving and loading seen IDs preserves the set."""
        with tempfile.TemporaryDirectory() as tmpdir:
            inbox_path = Path(tmpdir)
            
            # Save some IDs
            original_ids = {"id1", "id2", "id3"}
            save_seen_ids(inbox_path, original_ids)
            
            # Verify file exists
            seen_file = inbox_path / "_seen.json"
            assert seen_file.exists()
            
            # Load and verify
            loaded_ids = load_seen_ids(inbox_path)
            assert loaded_ids == original_ids
    
    def test_seen_ids_are_sorted_in_file(self):
        """Seen IDs are saved in sorted order for readability."""
        with tempfile.TemporaryDirectory() as tmpdir:
            inbox_path = Path(tmpdir)
            
            # Save in random order
            ids = {"id3", "id1", "id2"}
            save_seen_ids(inbox_path, ids)
            
            # Check file content
            seen_file = inbox_path / "_seen.json"
            data = json.loads(seen_file.read_text())
            assert data["seen_ids"] == sorted(list(ids))


class TestThreadContentExtraction:
    """Test extract_new_content function."""
    
    def test_no_quoted_content(self):
        """Simple email with no thread returns all content as new."""
        body = "This is a simple email with no quotes."
        new, quoted = extract_new_content(body)
        
        assert new == body
        assert quoted == ""
    
    def test_outlook_style_quote(self):
        """Outlook-style From/Sent/To header is detected."""
        body = """Thanks for the update.

From: John Smith
Sent: Monday, January 10, 2026 2:30 PM
To: You
Subject: RE: Project status

Here's the original message."""
        
        new, quoted = extract_new_content(body)
        
        assert new == "Thanks for the update."
        assert "Here's the original message" in quoted
    
    def test_on_date_wrote_quote(self):
        """'On ... wrote:' pattern is detected."""
        body = """Looks good!

On Mon, Jan 10, 2026, John wrote:
> This is the original message."""
        
        new, quoted = extract_new_content(body)
        
        assert new == "Looks good!"
        assert "This is the original message" in quoted
    
    def test_underscore_separator(self):
        """Underscore separator line is detected."""
        body = """New content here.

_____
Previous message below."""
        
        new, quoted = extract_new_content(body)
        
        assert new == "New content here."
        assert "Previous message below" in quoted


class TestEmailMarkdownFormatting:
    """Test build_email_markdown function."""
    
    @pytest.mark.skipif(not HAS_WINDOWS_DEPS, reason="Requires Windows dependencies")
    def test_basic_email_formatting(self):
        """Basic email is formatted with YAML frontmatter."""
        markdown = build_email_markdown(
            message_id="ABC123",
            received=datetime(2026, 1, 13, 10, 30, 0),
            from_address="john@example.com",
            to_addresses=["you@firm.com"],
            cc_addresses=[],
            subject="Test email",
            new_content="Hello world",
            quoted_content="",
        )
        
        # Check YAML frontmatter
        assert "---" in markdown
        assert "message_id: ABC123" in markdown
        assert "received: 2026-01-13T10:30:00" in markdown
        assert "from_address: john@example.com" in markdown
        assert "to_addresses:" in markdown
        assert "- you@firm.com" in markdown
        assert "subject: Test email" in markdown
        assert "state: received" in markdown
        
        # Check content
        assert "# New Content" in markdown
        assert "Hello world" in markdown
    
    @pytest.mark.skipif(not HAS_WINDOWS_DEPS, reason="Requires Windows dependencies")
    def test_email_with_quoted_content(self):
        """Email with quotes includes collapsible details section."""
        markdown = build_email_markdown(
            message_id="ABC123",
            received=datetime(2026, 1, 13, 10, 30, 0),
            from_address="john@example.com",
            to_addresses=["you@firm.com"],
            cc_addresses=[],
            subject="RE: Test",
            new_content="Thanks for your email.",
            quoted_content="> Your original message here.",
        )
        
        assert "<details>" in markdown
        assert "<summary>Previous messages in thread</summary>" in markdown
        assert "> Your original message here." in markdown
        assert "</details>" in markdown
    
    @pytest.mark.skipif(not HAS_WINDOWS_DEPS, reason="Requires Windows dependencies")
    def test_email_with_attachments(self):
        """Email with attachments includes attachment metadata."""
        attachments = [
            {
                "filename": "document.pdf",
                "local_path": "./attachments/document.pdf",
                "size_bytes": 1024
            }
        ]
        
        markdown = build_email_markdown(
            message_id="ABC123",
            received=datetime(2026, 1, 13, 10, 30, 0),
            from_address="john@example.com",
            to_addresses=["you@firm.com"],
            cc_addresses=["cc@example.com"],
            subject="With attachments",
            new_content="See attached.",
            quoted_content="",
            attachments=attachments,
        )
        
        assert "attachments:" in markdown
        assert "filename: document.pdf" in markdown
        assert "local_path: ./attachments/document.pdf" in markdown
    
    @pytest.mark.skipif(not HAS_WINDOWS_DEPS, reason="Requires Windows dependencies")
    def test_email_with_thread_id(self):
        """Email with thread_id includes it in metadata."""
        markdown = build_email_markdown(
            message_id="ABC123",
            received=datetime(2026, 1, 13, 10, 30, 0),
            from_address="john@example.com",
            to_addresses=["you@firm.com"],
            cc_addresses=[],
            subject="RE: Thread",
            new_content="Continuing thread.",
            quoted_content="",
            thread_id="THREAD456",
        )
        
        assert "thread_id: THREAD456" in markdown


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
