"""Tests for file_email_to_workspace tool and helper functions.

Tests cover:
- parse_sender_name: Various sender formats
- slugify_subject: Subject line parsing with prefixes and reference numbers
- format_email_markdown: Email formatting
- file_email_to_workspace: Full integration with mocked Outlook
- extract_new_content_only: Stripping quoted replies
- find_existing_email_file: Deduplication
- file_thread_to_workspace: Thread-aware filing
"""

import pytest
import json
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
import sys
import os

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from effi_mail.tools.workspace_filing import (
    parse_sender_name,
    slugify_subject,
    html_to_plain_text,
    format_email_markdown,
    generate_email_filename,
    get_unique_filepath,
    file_email_to_workspace,
    # New functions
    extract_new_content_only,
    compute_body_hash,
    find_existing_email_file,
    detect_quote_modification,
    file_thread_to_workspace,
    format_email_markdown_new_content_only,
    fix_mojibake,
)


# ============================================================================
# Tests for parse_sender_name
# ============================================================================

class TestParseSenderName:
    """Test parse_sender_name function with various input formats."""
    
    def test_standard_display_name_with_email(self):
        """Standard format: Display Name <email@domain.com>"""
        result = parse_sender_name("Katie Brownridge <katie.brownridge@biorelate.com>")
        assert result == "katie-brownridge"
    
    def test_exchange_address(self):
        """Exchange format with /o=ExchangeLabs path."""
        result = parse_sender_name("David Sant </o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=xxx>")
        assert result == "david-sant"
    
    def test_email_only_no_display_name(self):
        """Just an email address, no display name."""
        result = parse_sender_name("john.smith@example.com")
        assert result == "john-smith"
    
    def test_email_with_underscores(self):
        """Email address with underscores."""
        result = parse_sender_name("jane_doe@company.com")
        assert result == "jane-doe"
    
    def test_empty_sender(self):
        """Empty sender should return 'unknown'."""
        result = parse_sender_name("")
        assert result == "unknown"
    
    def test_none_sender(self):
        """None sender should return 'unknown'."""
        result = parse_sender_name(None)
        assert result == "unknown"
    
    def test_display_name_with_special_chars(self):
        """Display name with special characters."""
        result = parse_sender_name("O'Brien, Mary <mary.obrien@corp.com>")
        assert result == "obrien-mary"
    
    def test_single_name_in_display(self):
        """Display name with just one name."""
        result = parse_sender_name("Support <support@company.com>")
        assert result == "support"
    
    def test_email_in_angle_brackets_only(self):
        """Email in angle brackets with no display name."""
        result = parse_sender_name("<notifications@github.com>")
        assert result == "notifications"
    
    def test_complex_exchange_with_name(self):
        """Exchange address where we can extract a name."""
        result = parse_sender_name("Alice Johnson </o=ExchangeLabs/ou=Exchange/cn=xxx>")
        assert result == "alice-johnson"


# ============================================================================
# Tests for slugify_subject
# ============================================================================

class TestSlugifySubject:
    """Test slugify_subject function."""
    
    def test_simple_subject(self):
        """Simple subject line."""
        result = slugify_subject("Contract Amendment Discussion")
        assert result == "contract-amendment-discussion"
    
    def test_remove_re_prefix(self):
        """Remove Re: prefix."""
        result = slugify_subject("Re: Meeting Tomorrow")
        assert result == "meeting-tomorrow"
    
    def test_remove_multiple_prefixes(self):
        """Remove multiple Re:/Fwd: prefixes."""
        result = slugify_subject("Re: Fwd: RE: Important Update")
        assert result == "important-update"
    
    def test_remove_reference_number(self):
        """Remove reference numbers like [HJ-xxx-xxx-xxx]."""
        result = slugify_subject("Re: Fwd: GDPR/Data Compliance [HJ-824-1841932017-289]")
        assert result == "gdpr-data-compliance"
    
    def test_remove_numeric_reference(self):
        """Remove numeric reference numbers."""
        result = slugify_subject("Invoice Payment [5165-1903107924-109]")
        assert result == "invoice-payment"
    
    def test_slash_to_hyphen(self):
        """Slashes should become hyphens."""
        result = slugify_subject("GDPR/Data Protection")
        assert result == "gdpr-data-protection"
    
    def test_empty_subject(self):
        """Empty subject should return 'no-subject'."""
        result = slugify_subject("")
        assert result == "no-subject"
    
    def test_none_subject(self):
        """None subject should return 'no-subject'."""
        result = slugify_subject(None)
        assert result == "no-subject"
    
    def test_truncation_long_subject(self):
        """Long subjects should be truncated to max 50 chars."""
        long_subject = "This is a very long subject line that should be truncated at a reasonable word boundary"
        result = slugify_subject(long_subject)
        assert len(result) <= 50
        # Should truncate at word boundary (hyphen) if possible
        assert "-" in result
    
    def test_special_characters_removed(self):
        """Special characters should be removed."""
        result = slugify_subject("Question: How to proceed? (Urgent!)")
        assert result == "question-how-to-proceed-urgent"
    
    def test_fw_prefix_removed(self):
        """FW: prefix should be removed."""
        result = slugify_subject("FW: Document for Review")
        assert result == "document-for-review"


# ============================================================================
# Tests for html_to_plain_text
# ============================================================================

class TestHtmlToPlainText:
    """Test HTML to plain text conversion."""
    
    def test_simple_html(self):
        """Simple HTML with paragraph."""
        html = "<p>Hello World</p>"
        result = html_to_plain_text(html)
        assert "Hello World" in result
    
    def test_br_tags(self):
        """BR tags should become newlines."""
        html = "Line 1<br>Line 2<br/>Line 3"
        result = html_to_plain_text(html)
        assert "Line 1\nLine 2\nLine 3" == result
    
    def test_html_entities(self):
        """HTML entities should be decoded."""
        html = "Tom &amp; Jerry &lt;friends&gt;"
        result = html_to_plain_text(html)
        assert result == "Tom & Jerry <friends>"
    
    def test_em_dash_replacement(self):
        """Em-dashes should be replaced with regular hyphens."""
        html = "Item 1 — Item 2 – Item 3"
        result = html_to_plain_text(html)
        assert result == "Item 1 - Item 2 - Item 3"
    
    def test_empty_html(self):
        """Empty HTML should return empty string."""
        result = html_to_plain_text("")
        assert result == ""
    
    def test_none_html(self):
        """None should return empty string."""
        result = html_to_plain_text(None)
        assert result == ""


# ============================================================================
# Tests for fix_mojibake
# ============================================================================

class TestFixMojibake:
    """Test UTF-8 mojibake correction from Outlook."""
    
    def test_left_double_quote(self):
        """Left double quote mojibake should be fixed."""
        # UTF-8 bytes E2 80 9C decoded as Windows-1252 produces these chars
        mojibake = 'He said \xe2\x80\x9chello'
        result = fix_mojibake(mojibake)
        assert '\u201c' in result or '"' in result  # Either proper quote or fixed
    
    def test_right_double_quote(self):
        """Right double quote mojibake should be fixed."""
        mojibake = 'hello\xe2\x80\x9d she replied'
        result = fix_mojibake(mojibake)
        assert '\u201d' in result or '"' in result
    
    def test_apostrophe_smart_quote(self):
        """Smart apostrophe mojibake should be fixed."""
        mojibake = 'don\xe2\x80\x99t'
        result = fix_mojibake(mojibake)
        assert '\u2019' in result or "'" in result
    
    def test_bullet_point(self):
        """Bullet point mojibake should be fixed."""
        mojibake = '\xe2\x80\xa2 First item'
        result = fix_mojibake(mojibake)
        assert '\u2022' in result or result.startswith('\u2022')
    
    def test_en_dash(self):
        """En-dash mojibake should be fixed."""
        mojibake = '2020\xe2\x80\x932025'
        result = fix_mojibake(mojibake)
        assert '\u2013' in result or '-' in result
    
    def test_em_dash(self):
        """Em-dash mojibake should be fixed."""
        mojibake = 'word\xe2\x80\x94another'
        result = fix_mojibake(mojibake)
        assert '\u2014' in result or '-' in result
    
    def test_ellipsis(self):
        """Ellipsis mojibake should be fixed."""
        mojibake = 'and so on\xe2\x80\xa6'
        result = fix_mojibake(mojibake)
        assert '\u2026' in result or '...' in result
    
    def test_empty_string(self):
        """Empty string should return empty."""
        assert fix_mojibake("") == ""
    
    def test_none_input(self):
        """None should return None."""
        assert fix_mojibake(None) is None
    
    def test_clean_text_unchanged(self):
        """Clean ASCII text should pass through unchanged."""
        text = "Hello, this is normal text without any special characters."
        result = fix_mojibake(text)
        assert result == text
    
    def test_proper_unicode_unchanged(self):
        """Proper Unicode should not be corrupted."""
        text = "He said \u201chello\u201d and she said \u2018hi\u2019"
        result = fix_mojibake(text)
        # Should preserve proper Unicode
        assert '\u201c' in result and '\u201d' in result


# ============================================================================
# Tests for format_email_markdown
# ============================================================================

class TestFormatEmailMarkdown:
    """Test email markdown formatting."""
    
    @pytest.fixture
    def sample_email(self):
        """Sample email dict for testing."""
        return {
            "subject": "Contract Review",
            "sender_name": "John Smith",
            "sender_email": "john.smith@company.com",
            "received_time": "2026-01-09T14:30:00",
            "body": "Please review the attached contract.",
            "html_body": "",
            "recipients_to": ["recipient@example.com"],
            "recipients_cc": [],
            "id": "EMAIL-ID-123",
            "internet_message_id": "<msg@company.com>",
            "attachments": []
        }
    
    def test_basic_formatting(self, sample_email):
        """Test basic email formatting."""
        result = format_email_markdown(sample_email)
        
        assert "# Email: Contract Review" in result
        assert "**Date:** 2026-01-09 14:30" in result
        assert "**From:** John Smith (john.smith@company.com)" in result
        assert "**To:** recipient@example.com" in result
        assert "**Email ID:** EMAIL-ID-123" in result
        assert "Please review the attached contract." in result
    
    def test_with_attachments(self, sample_email):
        """Test formatting with attachments."""
        sample_email["attachments"] = [
            {"name": "contract.docx", "size": 15360},
            {"name": "appendix.pdf", "size": 2097152}
        ]
        result = format_email_markdown(sample_email)
        
        assert "## Attachments" in result
        assert "contract.docx (15.0 KB)" in result
        assert "appendix.pdf (2.0 MB)" in result
    
    def test_no_attachments_section_when_empty(self, sample_email):
        """Attachments section should not appear if no attachments."""
        sample_email["attachments"] = []
        result = format_email_markdown(sample_email)
        
        assert "## Attachments" not in result
    
    def test_with_cc_recipients(self, sample_email):
        """Test formatting with CC recipients."""
        sample_email["recipients_cc"] = ["cc1@example.com", "cc2@example.com"]
        result = format_email_markdown(sample_email)
        
        assert "**CC:** cc1@example.com, cc2@example.com" in result
    
    def test_html_body_fallback(self, sample_email):
        """Test falling back to HTML body when plain text is empty."""
        sample_email["body"] = ""
        sample_email["html_body"] = "<p>HTML content here</p>"
        result = format_email_markdown(sample_email)
        
        assert "HTML content here" in result
    
    def test_em_dash_replacement_in_body(self, sample_email):
        """Em-dashes should be replaced in body."""
        sample_email["body"] = "Item 1 — Item 2"
        result = format_email_markdown(sample_email)
        
        assert "Item 1 - Item 2" in result
        assert "—" not in result


# ============================================================================
# Tests for generate_email_filename
# ============================================================================

class TestGenerateEmailFilename:
    """Test email filename generation."""
    
    @pytest.fixture
    def sample_email(self):
        """Sample email for filename testing."""
        return {
            "received_time": "2026-01-09T14:09:00",
            "sender_name": "Katie Brownridge",
            "sender_email": "katie@biorelate.com",
            "subject": "GDPR Compliance Discussion"
        }
    
    def test_basic_filename(self, sample_email):
        """Test basic filename generation."""
        result = generate_email_filename(sample_email)
        assert result == "2026-01-09-1409__katie-brownridge____gdpr-compliance-discussion.md"
    
    def test_with_custom_topic(self, sample_email):
        """Test with custom topic slug."""
        result = generate_email_filename(sample_email, topic_slug="custom-topic")
        assert result == "2026-01-09-1409__katie-brownridge____custom-topic.md"
    
    def test_filename_format_components(self, sample_email):
        """Verify filename format components."""
        result = generate_email_filename(sample_email)
        
        # Should have double underscore between timestamp and sender
        assert "__" in result
        # Should have quadruple underscore between sender and topic
        assert "____" in result
        # Should end with .md
        assert result.endswith(".md")


# ============================================================================
# Tests for get_unique_filepath
# ============================================================================

class TestGetUniqueFilepath:
    """Test unique filepath generation."""
    
    def test_no_conflict(self, tmp_path):
        """File doesn't exist, return original path."""
        result = get_unique_filepath(tmp_path, "test-file.md")
        assert result == tmp_path / "test-file.md"
    
    def test_with_conflict_adds_suffix(self, tmp_path):
        """Existing file gets -2 suffix."""
        (tmp_path / "test-file.md").write_text("existing content")
        
        result = get_unique_filepath(tmp_path, "test-file.md")
        assert result == tmp_path / "test-file-2.md"
    
    def test_multiple_conflicts(self, tmp_path):
        """Multiple existing files increment suffix."""
        (tmp_path / "test-file.md").write_text("existing 1")
        (tmp_path / "test-file-2.md").write_text("existing 2")
        (tmp_path / "test-file-3.md").write_text("existing 3")
        
        result = get_unique_filepath(tmp_path, "test-file.md")
        assert result == tmp_path / "test-file-4.md"


# ============================================================================
# Tests for file_email_to_workspace (integration)
# ============================================================================

class TestFileEmailToWorkspace:
    """Integration tests for file_email_to_workspace tool."""
    
    @pytest.fixture
    def mock_email_data(self):
        """Sample email data returned by get_email_full."""
        return {
            "id": "EMAIL-ID-12345",
            "subject": "Contract Review Request",
            "sender_name": "John Smith",
            "sender_email": "john.smith@acme.com",
            "received_time": "2026-01-09T14:30:00",
            "body": "Please review the attached contract.\n\nBest regards,\nJohn",
            "html_body": "",
            "recipients_to": ["david.sant@example.com"],
            "recipients_cc": [],
            "attachments": [
                {"name": "contract.docx", "size": 15360}
            ],
            "internet_message_id": "<msg-001@acme.com>",
            "conversation_id": "conv-001"
        }
    
    def test_successful_filing(self, mock_email_data, tmp_path):
        """Test successful email filing."""
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            mock_outlook.get_email_full.return_value = mock_email_data
            
            result = file_email_to_workspace(
                email_id="EMAIL-ID-12345",
                destination_folder=str(tmp_path)
            )
            
            result_data = json.loads(result)
            assert result_data["success"] is True
            assert "filename" in result_data
            assert "path" in result_data
            assert result_data["filename"].endswith(".md")
            
            # Verify file was created
            filepath = Path(result_data["path"])
            assert filepath.exists()
            
            # Verify content
            content = filepath.read_text(encoding="utf-8")
            assert "# Email: Contract Review Request" in content
            assert "John Smith" in content
            assert "contract.docx" in content
    
    def test_email_not_found(self, tmp_path):
        """Test handling of email not found."""
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            mock_outlook.get_email_full.return_value = None
            
            result = file_email_to_workspace(
                email_id="NONEXISTENT-ID",
                destination_folder=str(tmp_path)
            )
            
            result_data = json.loads(result)
            assert result_data["success"] is False
            assert "not found" in result_data["error"].lower()
    
    def test_missing_email_id(self, tmp_path):
        """Test error when email_id is missing."""
        result = file_email_to_workspace(
            email_id="",
            destination_folder=str(tmp_path)
        )
        
        result_data = json.loads(result)
        assert result_data["success"] is False
        assert "email_id" in result_data["error"]
    
    def test_missing_destination_folder(self):
        """Test error when destination_folder is missing."""
        result = file_email_to_workspace(
            email_id="EMAIL-ID-12345",
            destination_folder=""
        )
        
        result_data = json.loads(result)
        assert result_data["success"] is False
        assert "destination_folder" in result_data["error"]
    
    def test_creates_missing_folder(self, mock_email_data, tmp_path):
        """Test that missing destination folder is created."""
        nested_path = tmp_path / "nested" / "folder" / "path"
        
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            mock_outlook.get_email_full.return_value = mock_email_data
            
            result = file_email_to_workspace(
                email_id="EMAIL-ID-12345",
                destination_folder=str(nested_path)
            )
            
            result_data = json.loads(result)
            assert result_data["success"] is True
            assert nested_path.exists()
    
    def test_duplicate_filing_adds_suffix(self, mock_email_data, tmp_path):
        """Test filing same email twice adds -2 suffix."""
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            mock_outlook.get_email_full.return_value = mock_email_data
            
            # File first time
            result1 = file_email_to_workspace(
                email_id="EMAIL-ID-12345",
                destination_folder=str(tmp_path)
            )
            result1_data = json.loads(result1)
            
            # File second time
            result2 = file_email_to_workspace(
                email_id="EMAIL-ID-12345",
                destination_folder=str(tmp_path)
            )
            result2_data = json.loads(result2)
            
            # Both should succeed
            assert result1_data["success"] is True
            assert result2_data["success"] is True
            
            # Filenames should be different
            assert result1_data["filename"] != result2_data["filename"]
            assert "-2.md" in result2_data["filename"]
    
    def test_with_custom_topic_slug(self, mock_email_data, tmp_path):
        """Test filing with custom topic slug."""
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            mock_outlook.get_email_full.return_value = mock_email_data
            
            result = file_email_to_workspace(
                email_id="EMAIL-ID-12345",
                destination_folder=str(tmp_path),
                topic_slug="custom-topic"
            )
            
            result_data = json.loads(result)
            assert result_data["success"] is True
            assert "custom-topic" in result_data["filename"]
    
    def test_exchange_sender_format(self, mock_email_data, tmp_path):
        """Test with Exchange-style sender (outbound email)."""
        mock_email_data["sender_name"] = "David Sant"
        mock_email_data["sender_email"] = "/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=xxx"
        
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            mock_outlook.get_email_full.return_value = mock_email_data
            
            result = file_email_to_workspace(
                email_id="EMAIL-ID-12345",
                destination_folder=str(tmp_path)
            )
            
            result_data = json.loads(result)
            assert result_data["success"] is True
            assert "david-sant" in result_data["filename"]
    
    def test_forwarded_thread_subject(self, mock_email_data, tmp_path):
        """Test with forwarded thread subject containing prefixes and refs."""
        mock_email_data["subject"] = "Re: Fwd: GDPR/Data Compliance [HJ-824-1841932017-289]"
        
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            mock_outlook.get_email_full.return_value = mock_email_data
            
            result = file_email_to_workspace(
                email_id="EMAIL-ID-12345",
                destination_folder=str(tmp_path)
            )
            
            result_data = json.loads(result)
            assert result_data["success"] is True
            assert "gdpr-data-compliance" in result_data["filename"]
            # Should not contain reference number or prefixes
            assert "re-" not in result_data["filename"].lower()
            assert "fwd-" not in result_data["filename"].lower()
            assert "hj-" not in result_data["filename"].lower()
    
    def test_email_with_attachments_renders_correctly(self, mock_email_data, tmp_path):
        """Test that attachments are properly listed in markdown."""
        mock_email_data["attachments"] = [
            {"name": "document.docx", "size": 15360},
            {"name": "report.pdf", "size": 2097152}
        ]
        
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            mock_outlook.get_email_full.return_value = mock_email_data
            
            result = file_email_to_workspace(
                email_id="EMAIL-ID-12345",
                destination_folder=str(tmp_path)
            )
            
            result_data = json.loads(result)
            content = Path(result_data["path"]).read_text(encoding="utf-8")
            
            assert "## Attachments" in content
            assert "document.docx" in content
            assert "report.pdf" in content
    
    def test_forward_slashes_in_path(self, mock_email_data, tmp_path):
        """Test that returned path uses forward slashes."""
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            mock_outlook.get_email_full.return_value = mock_email_data
            
            result = file_email_to_workspace(
                email_id="EMAIL-ID-12345",
                destination_folder=str(tmp_path)
            )
            
            result_data = json.loads(result)
            # Path should use forward slashes
            assert "\\" not in result_data["path"]


# ============================================================================
# Tests for extract_new_content_only
# ============================================================================

class TestExtractNewContentOnly:
    """Test extract_new_content_only function for stripping quoted replies."""
    
    def test_simple_email_no_quotes(self):
        """Email with no quoted content returns full body."""
        body = "Hello,\n\nThis is a simple email.\n\nBest regards,\nJohn"
        result = extract_new_content_only(body)
        assert result == body.rstrip()
    
    def test_outlook_reply_separator(self):
        """Strip content after Outlook-style separator."""
        body = """Thanks for the update.

I'll review this tomorrow.

-----Original Message-----
From: Alice
Sent: Monday, January 9, 2026
To: Bob
Subject: Update

Here's the original message."""
        
        result = extract_new_content_only(body)
        assert "Thanks for the update" in result
        assert "I'll review this tomorrow" in result
        assert "Original Message" not in result
        assert "Here's the original message" not in result
    
    def test_from_sent_to_subject_block(self):
        """Strip content after From/Sent/To/Subject block."""
        body = """Got it, thanks!

From: John Smith
Sent: Monday, January 9, 2026 10:30 AM
To: Jane Doe
Subject: Re: Project Update

Original content here."""
        
        result = extract_new_content_only(body)
        assert "Got it, thanks!" in result
        assert "Original content here" not in result
    
    def test_gmail_style_on_wrote(self):
        """Strip content after 'On ... wrote:' pattern."""
        body = """Sounds good to me!

On Mon, Jan 9, 2026 at 10:30 AM, John Smith wrote:
> Here's the quoted content
> from the previous email."""
        
        result = extract_new_content_only(body)
        assert "Sounds good to me!" in result
        assert "quoted content" not in result
    
    def test_quoted_lines_with_angle_brackets(self):
        """Strip content starting with > quote markers."""
        body = """New content here.

> This is quoted
> More quoted content
> And more"""
        
        result = extract_new_content_only(body)
        assert "New content here" in result
        assert "This is quoted" not in result
    
    def test_underscore_separator(self):
        """Strip content after underscore separator."""
        body = """My response to your question.

_____________________________
Previous message content here."""
        
        result = extract_new_content_only(body)
        assert "My response" in result
        assert "Previous message" not in result
    
    def test_empty_body(self):
        """Empty body returns empty string."""
        assert extract_new_content_only("") == ""
        assert extract_new_content_only(None) == ""
    
    def test_signature_not_treated_as_separator(self):
        """Email signature (--) should not be treated as quote separator."""
        body = """Hello,

Please see the attached document.

--
John Smith
Manager"""
        
        result = extract_new_content_only(body)
        # The signature should still be included since -- alone isn't a reply separator
        assert "Please see the attached document" in result


# ============================================================================
# Tests for compute_body_hash
# ============================================================================

class TestComputeBodyHash:
    """Test compute_body_hash function."""
    
    def test_consistent_hash(self):
        """Same content produces same hash."""
        body = "Hello, this is a test email."
        hash1 = compute_body_hash(body)
        hash2 = compute_body_hash(body)
        assert hash1 == hash2
    
    def test_whitespace_normalization(self):
        """Different whitespace produces same hash."""
        body1 = "Hello,  this   is a test."
        body2 = "Hello, this is a test."
        assert compute_body_hash(body1) == compute_body_hash(body2)
    
    def test_case_normalization(self):
        """Case is normalized."""
        body1 = "Hello World"
        body2 = "hello world"
        assert compute_body_hash(body1) == compute_body_hash(body2)
    
    def test_different_content_different_hash(self):
        """Different content produces different hash."""
        body1 = "Hello World"
        body2 = "Goodbye World"
        assert compute_body_hash(body1) != compute_body_hash(body2)
    
    def test_empty_body(self):
        """Empty body returns empty string."""
        assert compute_body_hash("") == ""
        assert compute_body_hash(None) == ""


# ============================================================================
# Tests for find_existing_email_file
# ============================================================================

class TestFindExistingEmailFile:
    """Test find_existing_email_file function."""
    
    def test_find_by_internet_message_id(self, tmp_path):
        """Find file by Internet Message ID in content."""
        # Create a file with a message ID
        filepath = tmp_path / "2026-01-09-1000__john-smith____test.md"
        content = """# Email: Test

**Date:** 2026-01-09 10:00
**From:** John Smith (john@example.com)
**Internet Message ID:** <msg-12345@example.com>

Email body here."""
        filepath.write_text(content, encoding="utf-8")
        
        result = find_existing_email_file(
            tmp_path,
            internet_message_id="<msg-12345@example.com>"
        )
        
        assert result is not None
        assert result == filepath
    
    def test_find_by_filename_pattern(self, tmp_path):
        """Find file by timestamp and sender pattern."""
        filepath = tmp_path / "2026-01-09-1000__john-smith____something.md"
        filepath.write_text("content", encoding="utf-8")
        
        result = find_existing_email_file(
            tmp_path,
            sender_slug="john-smith",
            timestamp="2026-01-09-1000"
        )
        
        assert result is not None
        assert result == filepath
    
    def test_not_found(self, tmp_path):
        """Return None when no match found."""
        result = find_existing_email_file(
            tmp_path,
            internet_message_id="<nonexistent@example.com>"
        )
        
        assert result is None
    
    def test_folder_not_exists(self):
        """Return None when folder doesn't exist."""
        result = find_existing_email_file(
            Path("/nonexistent/folder"),
            internet_message_id="<msg@example.com>"
        )
        
        assert result is None


# ============================================================================
# Tests for detect_quote_modification
# ============================================================================

class TestDetectQuoteModification:
    """Test detect_quote_modification function."""
    
    def test_no_modification(self):
        """Detect when quote is unchanged."""
        original = "This is the original email content that is long enough to be meaningful."
        later = """Thanks for your email.

-----Original Message-----
This is the original email content that is long enough to be meaningful."""
        
        result = detect_quote_modification(original, later, "sender@example.com")
        assert result is False
    
    def test_with_modification(self):
        """Detect when quote is modified."""
        original = "This is the original email content with important details about the project."
        later = """Thanks for your email.

-----Original Message-----
This is REDACTED content with different details entirely changed from original."""
        
        result = detect_quote_modification(original, later, "sender@example.com")
        assert result is True
    
    def test_short_original(self):
        """Short original content returns False (can't meaningfully compare)."""
        original = "OK"
        later = """Got it.

> OK"""
        
        result = detect_quote_modification(original, later, "sender@example.com")
        assert result is False
    
    def test_empty_bodies(self):
        """Empty bodies return False."""
        assert detect_quote_modification("", "later", "sender") is False
        assert detect_quote_modification("original", "", "sender") is False


# ============================================================================
# Tests for format_email_markdown_new_content_only
# ============================================================================

class TestFormatEmailMarkdownNewContentOnly:
    """Test format_email_markdown_new_content_only function."""
    
    def test_strips_quoted_content(self):
        """Quoted content should be stripped from markdown output."""
        email = {
            "subject": "Re: Test",
            "sender_name": "John",
            "sender_email": "john@example.com",
            "received_time": "2026-01-09T10:00:00",
            "body": """Thanks for the update!

-----Original Message-----
From: Alice
Subject: Test

Original content here.""",
            "html_body": "",
            "recipients_to": ["alice@example.com"],
            "recipients_cc": [],
            "id": "ID-123",
            "internet_message_id": "<msg@example.com>",
            "attachments": []
        }
        
        result = format_email_markdown_new_content_only(email)
        
        assert "Thanks for the update!" in result
        assert "Original Message" not in result
        assert "Original content here" not in result


# ============================================================================
# Tests for file_thread_to_workspace (integration)
# ============================================================================

class TestFileThreadToWorkspace:
    """Integration tests for file_thread_to_workspace tool."""
    
    @pytest.fixture
    def mock_thread_emails(self):
        """Create a mock email thread (3 emails)."""
        return [
            {
                "id": "EMAIL-1",
                "subject": "Project Discussion",
                "sender_name": "Alice Client",
                "sender_email": "alice@client.com",
                "received_time": "2026-01-07T10:00:00",
                "body": "Hi, I wanted to discuss the project scope.",
                "html_body": "",
                "recipients_to": ["david@example.com"],
                "recipients_cc": [],
                "internet_message_id": "<msg-001@client.com>",
                "conversation_id": "CONV-123",
                "conversation_topic": "Project Discussion",
                "attachments": []
            },
            {
                "id": "EMAIL-2",
                "subject": "RE: Project Discussion",
                "sender_name": "David Sant",
                "sender_email": "david@example.com",
                "received_time": "2026-01-07T14:30:00",
                "body": """Thanks Alice, I've reviewed the scope.

-----Original Message-----
From: Alice Client
Hi, I wanted to discuss the project scope.""",
                "html_body": "",
                "recipients_to": ["alice@client.com"],
                "recipients_cc": [],
                "internet_message_id": "<msg-002@example.com>",
                "conversation_id": "CONV-123",
                "conversation_topic": "Project Discussion",
                "attachments": []
            },
            {
                "id": "EMAIL-3",
                "subject": "RE: Project Discussion",
                "sender_name": "Alice Client",
                "sender_email": "alice@client.com",
                "received_time": "2026-01-08T09:15:00",
                "body": """Great, let's schedule a call.

On Jan 7, 2026, David Sant wrote:
> Thanks Alice, I've reviewed the scope.""",
                "html_body": "",
                "recipients_to": ["david@example.com"],
                "recipients_cc": [],
                "internet_message_id": "<msg-003@client.com>",
                "conversation_id": "CONV-123",
                "conversation_topic": "Project Discussion",
                "attachments": []
            }
        ]
    
    @pytest.fixture
    def mock_email_objects(self, mock_thread_emails):
        """Create mock Email objects for get_emails_by_conversation_id."""
        emails = []
        for data in mock_thread_emails:
            email = Mock()
            email.id = data["id"]
            email.subject = data["subject"]
            email.sender_name = data["sender_name"]
            email.sender_email = data["sender_email"]
            email.received_time = datetime.fromisoformat(data["received_time"])
            email.body_preview = data["body"][:100]
            email.has_attachments = False
            email.folder_path = "Inbox"
            email.direction = "inbound"
            email.recipients_to = data["recipients_to"]
            email.recipients_cc = data["recipients_cc"]
            email.internet_message_id = data["internet_message_id"]
            emails.append(email)
        return emails
    
    def test_files_all_thread_emails(self, mock_thread_emails, mock_email_objects, tmp_path):
        """Test that all emails in thread are filed separately."""
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            # Mock get_email_full to return each email's data
            def get_email_full_side_effect(email_id):
                for email in mock_thread_emails:
                    if email["id"] == email_id:
                        return email
                return None
            
            mock_outlook.get_email_full.side_effect = get_email_full_side_effect
            mock_outlook.get_emails_by_conversation_id.return_value = mock_email_objects
            
            result = file_thread_to_workspace(
                email_id="EMAIL-1",
                destination_folder=str(tmp_path)
            )
            
            result_data = json.loads(result)
            
            assert result_data["success"] is True
            assert len(result_data["filed"]) == 3
            assert result_data["thread_info"]["total_emails"] == 3
            assert result_data["thread_info"]["filed_count"] == 3
            
            # Verify files were created
            files = list(tmp_path.glob("*.md"))
            assert len(files) == 3
    
    def test_skips_already_filed_emails(self, mock_thread_emails, mock_email_objects, tmp_path):
        """Test that already-filed emails are skipped."""
        # Pre-create a file for the first email
        existing_file = tmp_path / "2026-01-07-1000__alice-client____project-discussion.md"
        existing_file.write_text(
            "# Email: Project Discussion\n\n**Internet Message ID:** <msg-001@client.com>\n\nContent",
            encoding="utf-8"
        )
        
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            def get_email_full_side_effect(email_id):
                for email in mock_thread_emails:
                    if email["id"] == email_id:
                        return email
                return None
            
            mock_outlook.get_email_full.side_effect = get_email_full_side_effect
            mock_outlook.get_emails_by_conversation_id.return_value = mock_email_objects
            
            result = file_thread_to_workspace(
                email_id="EMAIL-1",
                destination_folder=str(tmp_path)
            )
            
            result_data = json.loads(result)
            
            assert result_data["success"] is True
            assert len(result_data["skipped"]) == 1
            assert result_data["skipped"][0]["reason"] == "already_exists"
            assert len(result_data["filed"]) == 2
    
    def test_strips_quotes_from_replies(self, mock_thread_emails, mock_email_objects, tmp_path):
        """Test that quoted content is stripped from each email."""
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            def get_email_full_side_effect(email_id):
                for email in mock_thread_emails:
                    if email["id"] == email_id:
                        return email
                return None
            
            mock_outlook.get_email_full.side_effect = get_email_full_side_effect
            mock_outlook.get_emails_by_conversation_id.return_value = mock_email_objects
            
            result = file_thread_to_workspace(
                email_id="EMAIL-1",
                destination_folder=str(tmp_path)
            )
            
            result_data = json.loads(result)
            assert result_data["success"] is True
            
            # Read the second email file and verify quotes are stripped
            for filed in result_data["filed"]:
                if "david-sant" in filed["filename"]:
                    content = Path(filed["path"]).read_text(encoding="utf-8")
                    assert "Thanks Alice, I've reviewed the scope" in content
                    assert "Original Message" not in content
    
    def test_missing_email_id_error(self, tmp_path):
        """Test error when email_id is missing."""
        result = file_thread_to_workspace(
            email_id="",
            destination_folder=str(tmp_path)
        )
        
        result_data = json.loads(result)
        assert result_data["success"] is False
        assert "email_id" in result_data["error"]
    
    def test_missing_destination_folder_error(self):
        """Test error when destination_folder is missing."""
        result = file_thread_to_workspace(
            email_id="EMAIL-1",
            destination_folder=""
        )
        
        result_data = json.loads(result)
        assert result_data["success"] is False
        assert "destination_folder" in result_data["error"]
    
    def test_creates_destination_folder(self, mock_thread_emails, mock_email_objects, tmp_path):
        """Test that missing destination folder is created."""
        nested_path = tmp_path / "nested" / "folder" / "path"
        
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            def get_email_full_side_effect(email_id):
                for email in mock_thread_emails:
                    if email["id"] == email_id:
                        return email
                return None
            
            mock_outlook.get_email_full.side_effect = get_email_full_side_effect
            mock_outlook.get_emails_by_conversation_id.return_value = mock_email_objects
            
            result = file_thread_to_workspace(
                email_id="EMAIL-1",
                destination_folder=str(nested_path)
            )
            
            result_data = json.loads(result)
            assert result_data["success"] is True
            assert nested_path.exists()
    
    def test_custom_topic_slug(self, mock_thread_emails, mock_email_objects, tmp_path):
        """Test using custom topic slug for filenames."""
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            def get_email_full_side_effect(email_id):
                for email in mock_thread_emails:
                    if email["id"] == email_id:
                        return email
                return None
            
            mock_outlook.get_email_full.side_effect = get_email_full_side_effect
            mock_outlook.get_emails_by_conversation_id.return_value = mock_email_objects
            
            result = file_thread_to_workspace(
                email_id="EMAIL-1",
                destination_folder=str(tmp_path),
                topic_slug="custom-topic"
            )
            
            result_data = json.loads(result)
            assert result_data["success"] is True
            
            # All filenames should have the custom topic
            for filed in result_data["filed"]:
                assert "custom-topic" in filed["filename"]
    
    def test_single_email_no_thread(self, tmp_path):
        """Test handling of single email not in a thread."""
        single_email = {
            "id": "SINGLE-EMAIL",
            "subject": "Standalone Email",
            "sender_name": "John",
            "sender_email": "john@example.com",
            "received_time": "2026-01-09T10:00:00",
            "body": "This is a standalone email.",
            "html_body": "",
            "recipients_to": ["recipient@example.com"],
            "recipients_cc": [],
            "internet_message_id": "<standalone@example.com>",
            "conversation_id": None,  # No thread
            "conversation_topic": None,
            "attachments": []
        }
        
        with patch('effi_mail.tools.workspace_filing.outlook') as mock_outlook:
            mock_outlook.get_email_full.return_value = single_email
            
            result = file_thread_to_workspace(
                email_id="SINGLE-EMAIL",
                destination_folder=str(tmp_path)
            )
            
            result_data = json.loads(result)
            assert result_data["success"] is True
            assert len(result_data["filed"]) == 1


# ============================================================================
# Run tests
# ============================================================================

if __name__ == "__main__":
    pytest.main([__file__, "-v"])
