"""Thread content extraction for email ingestion.

Separates new content from quoted replies in email threads.
"""

import re
from typing import Tuple


def extract_new_content(body: str) -> Tuple[str, str]:
    """Extract new content from email, separating from quoted thread.
    
    Args:
        body: Full email body text. May be ``None`` or an empty string,
            in which case no content is extracted.
        
    Returns:
        Tuple of (new_content, quoted_remainder). For ``None`` or empty
        input, both values will be empty strings (``"", "``).
    """
    if not body:
        return "", ""
    
    # Common patterns indicating quoted content
    patterns = [
        r'\n\s*On .+wrote:\s*\n',              # "On 10 Jan, John wrote:"
        r'\n\s*From:.+\nSent:.+\nTo:.+\n',     # Outlook-style header block
        r'\n\s*-{3,}\s*Original Message',      # "--- Original Message ---"
        r'\n_{5,}\n',                          # Underscore separator
    ]
    
    earliest_match = len(body)
    for pattern in patterns:
        match = re.search(pattern, body, re.IGNORECASE | re.MULTILINE)
        if match and match.start() < earliest_match:
            earliest_match = match.start()
    
    new_content = body[:earliest_match].strip()
    quoted = body[earliest_match:].strip() if earliest_match < len(body) else ""
    
    return new_content, quoted
