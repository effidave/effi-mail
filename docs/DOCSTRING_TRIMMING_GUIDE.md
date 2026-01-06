# Docstring Trimming Guide for MCP Tools

## Goal
Reduce token usage in tool schemas while preserving hints that prevent agent errors.

## What to DROP

### Args sections
Type hints already provide this info:
```python
# BEFORE
"""
Args:
    email_id: Email EntryID
    status: Triage status
"""

# AFTER - just delete it
```

### Returns sections
All tools return JSON - no need to repeat:
```python
# DROP THIS
"""
Returns:
    JSON string with success status
"""
```

### Obvious parameter descriptions
If the parameter name is clear, don't explain it:
- `limit` → obvious
- `days` → obvious  
- `email_id` → obvious

## What to KEEP

### Enum/allowed values
Agents will guess wrong without these:
```python
"""Set triage status ('processed', 'deferred', 'archived')."""
"""Category: Client, Internal, Marketing, Personal, or Spam."""
```

### Date formats
Prevent format guessing:
```python
"""Dates are YYYY-MM-DD format."""
```

### Case sensitivity
Prevents unnecessary exactness or retry loops:
```python
"""client_id is case-insensitive."""
```

### Non-obvious default behaviors
When defaults do something special:
```python
"""Saves to ./attachments/{domain}/{date}/{filename} if no path given."""
```

### Folder/path options
When there's a limited set:
```python
"""folder: 'Inbox' or 'Sent Items'."""
```

## Template

One-liner format:
```python
def tool_name(param1: str, param2: int = 30) -> str:
    """Brief action description. Key constraint or hint."""
```

## Examples

### Before (verbose)
```python
def triage_email(email_id: str, status: str) -> str:
    """Assign triage status to an email using Outlook categories.
    
    Status is stored in the email itself as an Outlook category.
    
    Args:
        email_id: Email EntryID
        status: Triage status - 'processed', 'deferred', or 'archived'
        
    Returns:
        JSON string with success/error status
    """
```

### After (trimmed)
```python
def triage_email(email_id: str, status: str) -> str:
    """Set triage status ('processed', 'deferred', 'archived') on an email."""
```

## Checklist

Before trimming, verify:
- [ ] Check if underlying code is case-insensitive (grep for `.lower()`)
- [ ] Identify enum-like parameters that need value hints
- [ ] Check for non-obvious default behaviors
- [ ] Look for date/format parameters

## Token savings

Expect ~50% reduction in docstring tokens:
- Verbose: ~80-100 tokens per tool
- Trimmed: ~40-50 tokens per tool
- Savings: ~750+ tokens across 15 tools
