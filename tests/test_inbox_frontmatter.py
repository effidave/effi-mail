"""Tests for add_email_frontmatter tool and helper functions.

Tests cover:
- parse_yaml_frontmatter: Parsing existing YAML front matter
- format_yaml_frontmatter: Formatting dict to YAML front matter
- find_email_file_by_message_id: Searching inbox for email files
- add_email_frontmatter: Full integration with temp files
"""

import pytest
import tempfile
from pathlib import Path
import sys
import os

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from effi_mail.tools.inbox_frontmatter import (
    parse_yaml_frontmatter,
    format_yaml_frontmatter,
    find_email_file_by_id,
    add_email_frontmatter,
    DEFAULT_INBOX_PATH,
)


# ============================================================================
# Tests for parse_yaml_frontmatter
# ============================================================================

class TestParseYamlFrontmatter:
    """Test parse_yaml_frontmatter function."""
    
    def test_no_frontmatter(self):
        """Content without front matter returns empty dict."""
        content = "# Email: Test Subject\n\n**Date:** 2026-01-13"
        frontmatter, body = parse_yaml_frontmatter(content)
        assert frontmatter == {}
        assert body == content
    
    def test_simple_frontmatter(self):
        """Parse simple YAML front matter."""
        content = """---
client: Test Client
matter: Test Matter
filed: false
---
# Email: Test Subject"""
        frontmatter, body = parse_yaml_frontmatter(content)
        assert frontmatter["client"] == "Test Client"
        assert frontmatter["matter"] == "Test Matter"
        assert frontmatter["filed"] is False
        assert body.strip() == "# Email: Test Subject"
    
    def test_boolean_true(self):
        """Parse boolean true value."""
        content = """---
filed: true
---
Body"""
        frontmatter, body = parse_yaml_frontmatter(content)
        assert frontmatter["filed"] is True
    
    def test_boolean_false(self):
        """Parse boolean false value."""
        content = """---
filed: false
---
Body"""
        frontmatter, body = parse_yaml_frontmatter(content)
        assert frontmatter["filed"] is False
    
    def test_quoted_string(self):
        """Parse quoted string value."""
        content = """---
email_id: "<test@example.com>"
---
Body"""
        frontmatter, body = parse_yaml_frontmatter(content)
        assert frontmatter["email_id"] == "<test@example.com>"
    
    def test_single_quoted_string(self):
        """Parse single-quoted string value."""
        content = """---
client: 'Client: Special Name'
---
Body"""
        frontmatter, body = parse_yaml_frontmatter(content)
        assert frontmatter["client"] == "Client: Special Name"
    
    def test_preserves_body_content(self):
        """Body content is preserved exactly."""
        content = """---
filed: false
---
# Email: Test

**Date:** 2026-01-13

Body content here."""
        frontmatter, body = parse_yaml_frontmatter(content)
        assert "# Email: Test" in body
        assert "**Date:** 2026-01-13" in body
        assert "Body content here." in body


# ============================================================================
# Tests for format_yaml_frontmatter
# ============================================================================

class TestFormatYamlFrontmatter:
    """Test format_yaml_frontmatter function."""
    
    def test_empty_dict(self):
        """Empty dict returns empty string."""
        result = format_yaml_frontmatter({})
        assert result == ""
    
    def test_simple_string(self):
        """Format simple string value."""
        result = format_yaml_frontmatter({"client": "Test Client"})
        assert "---" in result
        assert "client: Test Client" in result
    
    def test_boolean_true(self):
        """Format boolean true as lowercase."""
        result = format_yaml_frontmatter({"filed": True})
        assert "filed: true" in result
    
    def test_boolean_false(self):
        """Format boolean false as lowercase."""
        result = format_yaml_frontmatter({"filed": False})
        assert "filed: false" in result
    
    def test_special_characters_quoted(self):
        """Strings with special characters are quoted."""
        result = format_yaml_frontmatter({"email_id": "<test@example.com>"})
        assert 'email_id: "<test@example.com>"' in result
    
    def test_colon_in_value_quoted(self):
        """Strings with colons are quoted."""
        result = format_yaml_frontmatter({"client": "Client: Special"})
        assert 'client: "Client: Special"' in result
    
    def test_none_value_skipped(self):
        """None values are not included."""
        result = format_yaml_frontmatter({"client": "Test", "matter": None})
        assert "client: Test" in result
        assert "matter" not in result
    
    def test_multiple_values(self):
        """Multiple values formatted correctly."""
        result = format_yaml_frontmatter({
            "email_id": "<test@example.com>",
            "client": "Test Client",
            "matter": "Test Matter",
            "filed": False
        })
        lines = result.split("\n")
        assert lines[0] == "---"
        assert lines[-1] == "---"
        assert 'email_id: "<test@example.com>"' in result
        assert "client: Test Client" in result
        assert "matter: Test Matter" in result
        assert "filed: false" in result


# ============================================================================
# Tests for find_email_file_by_id
# ============================================================================

class TestFindEmailFileById:
    """Test find_email_file_by_id function."""
    
    def test_finds_matching_file(self):
        """Find file containing the email ID."""
        with tempfile.TemporaryDirectory() as tmpdir:
            inbox = Path(tmpdir)
            
            # Create test file
            test_file = inbox / "test-email.md"
            test_file.write_text(
                "# Email: Test\n\n"
                "**Email ID:** 000000001234ABCD5678\n\n"
                "Body content",
                encoding="utf-8"
            )
            
            result = find_email_file_by_id(
                "000000001234ABCD5678",
                inbox_path=inbox
            )
            assert result == test_file
    
    def test_finds_file_in_subfolder(self):
        """Find file in nested subfolder."""
        with tempfile.TemporaryDirectory() as tmpdir:
            inbox = Path(tmpdir)
            subfolder = inbox / "client" / "2026-01"
            subfolder.mkdir(parents=True)
            
            # Create test file in subfolder
            test_file = subfolder / "test-email.md"
            test_file.write_text(
                "# Email: Test\n\n"
                "**Email ID:** 000000005678EFGH9012\n\n"
                "Body content",
                encoding="utf-8"
            )
            
            result = find_email_file_by_id(
                "000000005678EFGH9012",
                inbox_path=inbox
            )
            assert result == test_file
    
    def test_returns_none_when_not_found(self):
        """Return None when no matching file exists."""
        with tempfile.TemporaryDirectory() as tmpdir:
            inbox = Path(tmpdir)
            
            # Create file with different email ID
            test_file = inbox / "test-email.md"
            test_file.write_text(
                "# Email: Test\n\n"
                "**Email ID:** 000000001111AAAA2222\n\n"
                "Body content",
                encoding="utf-8"
            )
            
            result = find_email_file_by_id(
                "000000009999ZZZZ8888",
                inbox_path=inbox
            )
            assert result is None
    
    def test_returns_none_for_nonexistent_folder(self):
        """Return None when inbox folder doesn't exist."""
        result = find_email_file_by_id(
            "000000001234ABCD5678",
            inbox_path=Path("/nonexistent/path")
        )
        assert result is None


# ============================================================================
# Tests for add_email_frontmatter
# ============================================================================

class TestAddEmailFrontmatter:
    """Test add_email_frontmatter function with temporary files."""
    
    def test_adds_frontmatter_to_file_without(self):
        """Add front matter to a file that has none."""
        with tempfile.TemporaryDirectory() as tmpdir:
            inbox = Path(tmpdir)
            
            # Create test file without front matter
            test_file = inbox / "test-email.md"
            original_content = (
                "# Email: Test Subject\n\n"
                "**Date:** 2026-01-13\n"
                "**Email ID:** 000000001234ABCD5678EF\n\n"
                "Body content here."
            )
            test_file.write_text(original_content, encoding="utf-8")
            
            # Patch the default inbox path
            import effi_mail.tools.inbox_frontmatter as fm
            original_path = fm.DEFAULT_INBOX_PATH
            fm.DEFAULT_INBOX_PATH = inbox
            
            try:
                result = add_email_frontmatter(
                    email_id="000000001234ABCD5678EF",
                    client="Test Client",
                    matter="Test Matter",
                    filed=False
                )
                
                assert result["success"] is True
                assert result["file_path"] == str(test_file)
                assert result["frontmatter"]["client"] == "Test Client"
                assert result["frontmatter"]["matter"] == "Test Matter"
                assert result["frontmatter"]["filed"] is False
                
                # Verify file content
                content = test_file.read_text(encoding="utf-8")
                assert content.startswith("---")
                assert "client: Test Client" in content
                assert "matter: Test Matter" in content
                assert "filed: false" in content
                assert "# Email: Test Subject" in content
            finally:
                fm.DEFAULT_INBOX_PATH = original_path
    
    def test_updates_existing_frontmatter(self):
        """Update existing front matter preserving unchanged values."""
        with tempfile.TemporaryDirectory() as tmpdir:
            inbox = Path(tmpdir)
            
            # Create test file with existing front matter
            test_file = inbox / "test-email.md"
            original_content = (
                "---\n"
                "email_id: 000000002222BBBB3333CC\n"
                "client: Original Client\n"
                "matter: Original Matter\n"
                "filed: false\n"
                "---\n"
                "# Email: Test Subject\n\n"
                "**Email ID:** 000000002222BBBB3333CC\n\n"
                "Body content here."
            )
            test_file.write_text(original_content, encoding="utf-8")
            
            # Patch the default inbox path
            import effi_mail.tools.inbox_frontmatter as fm
            original_path = fm.DEFAULT_INBOX_PATH
            fm.DEFAULT_INBOX_PATH = inbox
            
            try:
                # Update only filed status
                result = add_email_frontmatter(
                    email_id="000000002222BBBB3333CC",
                    filed=True
                )
                
                assert result["success"] is True
                # Original values preserved
                assert result["frontmatter"]["client"] == "Original Client"
                assert result["frontmatter"]["matter"] == "Original Matter"
                # Updated value
                assert result["frontmatter"]["filed"] is True
                
                # Verify file content
                content = test_file.read_text(encoding="utf-8")
                assert "client: Original Client" in content
                assert "matter: Original Matter" in content
                assert "filed: true" in content
            finally:
                fm.DEFAULT_INBOX_PATH = original_path
    
    def test_overwrites_specific_values(self):
        """Overwrite specific values while preserving others."""
        with tempfile.TemporaryDirectory() as tmpdir:
            inbox = Path(tmpdir)
            
            # Create test file with existing front matter
            test_file = inbox / "test-email.md"
            original_content = (
                "---\n"
                "email_id: 000000004444DDDD5555EE\n"
                "client: Original Client\n"
                "matter: Original Matter\n"
                "filed: false\n"
                "---\n"
                "# Email: Test Subject\n\n"
                "**Email ID:** 000000004444DDDD5555EE\n\n"
                "Body content here."
            )
            test_file.write_text(original_content, encoding="utf-8")
            
            # Patch the default inbox path
            import effi_mail.tools.inbox_frontmatter as fm
            original_path = fm.DEFAULT_INBOX_PATH
            fm.DEFAULT_INBOX_PATH = inbox
            
            try:
                # Update client only
                result = add_email_frontmatter(
                    email_id="000000004444DDDD5555EE",
                    client="New Client"
                )
                
                assert result["success"] is True
                # Client updated
                assert result["frontmatter"]["client"] == "New Client"
                # Matter preserved
                assert result["frontmatter"]["matter"] == "Original Matter"
            finally:
                fm.DEFAULT_INBOX_PATH = original_path
    
    def test_returns_error_when_file_not_found(self):
        """Return error when no matching file exists."""
        with tempfile.TemporaryDirectory() as tmpdir:
            inbox = Path(tmpdir)
            
            # Patch the default inbox path
            import effi_mail.tools.inbox_frontmatter as fm
            original_path = fm.DEFAULT_INBOX_PATH
            fm.DEFAULT_INBOX_PATH = inbox
            
            try:
                result = add_email_frontmatter(
                    email_id="000000009999NOTFOUND",
                    client="Test Client"
                )
                
                assert result["success"] is False
                assert "No email file found" in result["message"]
                assert result["frontmatter"] == {}
            finally:
                fm.DEFAULT_INBOX_PATH = original_path


# ============================================================================
# Integration test with real inbox (skipped by default)
# ============================================================================

class TestIntegrationWithRealInbox:
    """Integration tests with real effi-work inbox folder.
    
    These tests are skipped by default. Run with:
        pytest tests/test_inbox_frontmatter.py -k "integration" --run-integration
    """
    
    @pytest.mark.skipif(
        not DEFAULT_INBOX_PATH.exists(),
        reason="Real inbox folder not available"
    )
    def test_can_search_real_inbox(self):
        """Verify search works on real inbox structure."""
        # Just verify the search runs without error
        # Don't actually modify any files
        result = find_email_file_by_id(
            "000000009999NONEXISTENT",
            inbox_path=DEFAULT_INBOX_PATH
        )
        # Should return None for non-existent ID
        assert result is None


# ============================================================================
# Run tests
# ============================================================================

if __name__ == "__main__":
    pytest.main([__file__, "-v"])
