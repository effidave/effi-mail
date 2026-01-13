#!/usr/bin/env python
"""Email ingestion script.

Deterministic script to fetch and convert emails from Outlook to markdown
before triaging them. Saves emails to _inbox/ directory with YAML frontmatter.
"""

import argparse
import logging
import sys
from pathlib import Path

from effi_mail.ingestion import ingest_new_emails


def setup_logging(verbose: bool = False):
    """Configure logging output.
    
    Args:
        verbose: Enable debug logging
    """
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )


def main():
    """Run email ingestion."""
    parser = argparse.ArgumentParser(
        description="Ingest emails from Outlook to _inbox/ as markdown files"
    )
    parser.add_argument(
        "--folder", "-f",
        default="Inbox",
        help="Outlook folder to poll (default: Inbox). Examples: 'Sent Items', 'Projects/Active'"
    )
    parser.add_argument(
        "--limit", "-l",
        type=int,
        default=50,
        help="Maximum emails to process (default: 50)"
    )
    parser.add_argument(
        "--inbox-path",
        type=str,
        default=None,
        help="Path to _inbox directory (default: current_directory/_inbox)"
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose (debug) logging"
    )
    
    args = parser.parse_args()
    
    # Setup logging
    setup_logging(args.verbose)
    logger = logging.getLogger(__name__)
    
    # Determine inbox path
    if args.inbox_path:
        inbox_path = Path(args.inbox_path)
    else:
        # Default to _inbox in current directory
        inbox_path = Path.cwd() / "_inbox"
    
    logger.info(f"Starting email ingestion")
    logger.info(f"  Outlook folder: {args.folder}")
    logger.info(f"  Inbox path: {inbox_path}")
    logger.info(f"  Limit: {args.limit}")
    
    try:
        saved = ingest_new_emails(inbox_path, folder=args.folder, limit=args.limit)
        print(f"\n✓ Ingested {len(saved)} new emails from '{args.folder}'")
        print(f"  Saved to: {inbox_path}")
        
        if saved:
            print(f"\nFiles created:")
            for path in saved[:10]:  # Show first 10
                print(f"  - {path.name}")
            if len(saved) > 10:
                print(f"  ... and {len(saved) - 10} more")
        
        return 0
        
    except Exception as e:
        logger.error(f"Ingestion failed: {e}", exc_info=True)
        print(f"\n✗ Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
