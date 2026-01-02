"""JSON-based domain category management.

This module provides simple get/set access to domain_categories.json,
replacing the SQLite-based domain categorization.
"""

import json
from pathlib import Path
from typing import Optional, Dict, List

# Valid categories
VALID_CATEGORIES = {"Client", "Internal", "Marketing", "Personal", "Spam", "Uncategorized"}

# Default path to domain_categories.json (same directory as this module)
DEFAULT_JSON_PATH = Path(__file__).parent / "domain_categories.json"


def _load_categories(json_path: Path = DEFAULT_JSON_PATH) -> Dict[str, str]:
    """Load domain categories from JSON file."""
    if not json_path.exists():
        return {}
    
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return {}


def _save_categories(categories: Dict[str, str], json_path: Path = DEFAULT_JSON_PATH) -> None:
    """Save domain categories to JSON file."""
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(categories, f, indent=2)


def get_domain_category(domain: str, json_path: Path = DEFAULT_JSON_PATH) -> str:
    """
    Get the category for a domain.
    
    Args:
        domain: Domain name (e.g., 'example.com')
        json_path: Path to JSON file (optional)
        
    Returns:
        Category string or "Uncategorized" if domain not categorized
    """
    categories = _load_categories(json_path)
    # Case-insensitive lookup
    domain_lower = domain.lower()
    for d, cat in categories.items():
        if d.lower() == domain_lower:
            return cat
    return "Uncategorized"


def set_domain_category(domain: str, category: str, json_path: Path = DEFAULT_JSON_PATH) -> bool:
    """
    Set the category for a domain.
    
    Args:
        domain: Domain name (e.g., 'example.com')
        category: Category (Client, Internal, Marketing, Personal, Spam, Uncategorized)
        json_path: Path to JSON file (optional)
        
    Returns:
        True if successful, False if invalid category
    """
    if category not in VALID_CATEGORIES:
        return False
    
    categories = _load_categories(json_path)
    categories[domain] = category
    _save_categories(categories, json_path)
    return True


def get_all_domain_categories(json_path: Path = DEFAULT_JSON_PATH) -> Dict[str, str]:
    """
    Get all domain categories.
    
    Returns:
        Dict mapping domain -> category
    """
    return _load_categories(json_path)


def get_domains_by_category(category: str, json_path: Path = DEFAULT_JSON_PATH) -> List[str]:
    """
    Get all domains with a specific category.
    
    Args:
        category: Category to filter by
        
    Returns:
        List of domain names
    """
    categories = _load_categories(json_path)
    return [domain for domain, cat in categories.items() if cat == category]


def get_uncategorized_domains(known_domains: List[str], json_path: Path = DEFAULT_JSON_PATH) -> List[str]:
    """
    Get domains from a list that are not yet categorized.
    
    Args:
        known_domains: List of domain names to check
        json_path: Path to JSON file
        
    Returns:
        List of domains not in the categories file
    """
    categories = _load_categories(json_path)
    categorized_lower = {d.lower() for d in categories.keys()}
    return [d for d in known_domains if d.lower() not in categorized_lower]


def remove_domain_category(domain: str, json_path: Path = DEFAULT_JSON_PATH) -> bool:
    """
    Remove a domain from categories.
    
    Args:
        domain: Domain name to remove
        
    Returns:
        True if removed, False if not found
    """
    categories = _load_categories(json_path)
    domain_lower = domain.lower()
    
    to_remove = None
    for d in categories:
        if d.lower() == domain_lower:
            to_remove = d
            break
    
    if to_remove:
        del categories[to_remove]
        _save_categories(categories, json_path)
        return True
    return False
