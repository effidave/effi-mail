"""Email thread tracking tools for effi-mail MCP server.

Uses Exchange ConversationID to deterministically find related emails
across Inbox, Sent Items, and optionally DMS folders.
"""

import json
from typing import Optional

from effi_mail.helpers import retrieval, build_response_with_auto_file


def get_email_thread(
    email_id: str,
    include_sent: bool = True,
    include_dms: bool = False,
    limit: int = 50,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Get all emails in a conversation thread by email ID.
    
    Uses Exchange ConversationID to find all related emails across folders.
    Returns messages sorted chronologically with full metadata.
    
    Large results (>{auto_file_threshold} messages) are auto-saved to a cache file.
    Use force_inline=True to return full payload inline regardless of size.
    Use output_file to save results to a specific path.
    
    ⚠️ Results are LIMITED. Check 'results_truncated' in response to determine if more records exist.
    
    Args:
        email_id: EntryID or internet_message_id of any email in the thread
        include_sent: Include Sent Items folder in search (default: True)
        include_dms: Include DMS folders in search (default: False)
        limit: Maximum messages to return (default: 50)
        output_file: Path to save results to (optional)
        force_inline: Return full payload inline regardless of size (default False)
        auto_file_threshold: Auto-file results above this count (default 20)
    
    Returns:
        JSON with thread metadata and messages
    """
    try:
        # Get source email to extract ConversationID and ConversationTopic
        source_email = retrieval.get_email_full(email_id)
        
        if not source_email:
            return json.dumps({"error": f"Email not found: {email_id}"})
        
        conversation_id = source_email.get("conversation_id")
        conversation_topic = source_email.get("conversation_topic")
        
        if not conversation_id:
            return json.dumps({
                "error": "Email has no ConversationID - cannot retrieve thread"
            })
        
        if not conversation_topic:
            return json.dumps({
                "error": "Email has no ConversationTopic - cannot retrieve thread"
            })
        
        # Get all emails in the conversation with limit+1 to detect truncation
        # Note: We use ConversationTopic for filtering (ConversationID not filterable in Outlook)
        emails = retrieval.get_emails_by_conversation_id(
            conversation_id=conversation_id,
            conversation_topic=conversation_topic,
            include_sent=include_sent,
            include_dms=include_dms,
            limit=limit + 1
        )
        
        # Detect truncation
        was_truncated = len(emails) > limit
        emails = emails[:limit]
        
        # Sort chronologically (oldest first)
        emails.sort(key=lambda e: e.received_time)
        
        # Extract unique participants
        participants = set()
        for email in emails:
            participants.add(email.sender_email.lower())
            participants.update(r.lower() for r in email.recipients_to)
            participants.update(r.lower() for r in email.recipients_cc)
        
        # Build date range
        if emails:
            date_range = {
                "first": emails[0].received_time.isoformat(),
                "last": emails[-1].received_time.isoformat()
            }
        else:
            date_range = None
        
        # Format messages
        messages = []
        for email in emails:
            messages.append({
                "id": email.id,
                "subject": email.subject,
                "sender": email.sender_email,
                "direction": email.direction,
                "received": email.received_time.isoformat(),
                "folder": email.folder_path,
                "preview": email.body_preview[:200] if email.body_preview else "",
                "has_attachments": email.has_attachments,
            })
        
        return build_response_with_auto_file(
            data={
                "conversation_id": conversation_id,
                "participants": sorted(participants),
                "date_range": date_range,
                "messages": messages
            },
            items_key="messages",
            count=len(emails),
            limit=limit,
            was_truncated=was_truncated,
            total_available=None,
            output_file=output_file,
            force_inline=force_inline,
            auto_file_threshold=auto_file_threshold,
            cache_prefix="thread"
        )
        
    except Exception as e:
        return json.dumps({"error": f"Failed to retrieve thread: {str(e)}"})


def get_thread_locations(email_id: str) -> str:
    """Get lightweight thread info: just IDs, folders, and timestamps.
    
    Use this to quickly identify where thread messages are located
    before fetching full content with get_email_by_id.
    
    Args:
        email_id: EntryID or internet_message_id of any email in the thread
    
    Returns:
        JSON with conversation_id and list of locations
    """
    try:
        # Get source email
        source_email = retrieval.get_email_full(email_id)
        
        if not source_email:
            return json.dumps({"error": f"Email not found: {email_id}"})
        
        conversation_id = source_email.get("conversation_id")
        conversation_topic = source_email.get("conversation_topic")
        
        if not conversation_id:
            return json.dumps({
                "error": "Email has no ConversationID - cannot retrieve thread"
            })
        
        if not conversation_topic:
            return json.dumps({
                "error": "Email has no ConversationTopic - cannot retrieve thread"
            })
        
        # Get all emails in the conversation
        emails = retrieval.get_emails_by_conversation_id(
            conversation_id=conversation_id,
            conversation_topic=conversation_topic,
            include_sent=True,
            include_dms=False,
            limit=50
        )
        
        # Sort chronologically
        emails.sort(key=lambda e: e.received_time)
        
        # Build minimal location data
        locations = []
        for email in emails:
            locations.append({
                "id": email.id,
                "folder": email.folder_path,
                "direction": email.direction,
                "received": email.received_time.isoformat(),
            })
        
        return json.dumps({
            "conversation_id": conversation_id,
            "message_count": len(locations),
            "locations": locations
        }, indent=2)
        
    except Exception as e:
        return json.dumps({"error": f"Failed to retrieve thread locations: {str(e)}"})
