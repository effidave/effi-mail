"""Migration script to import domain categories from JSON to database."""

import json
from pathlib import Path
from database import Database
from models import Domain, EmailCategory

def migrate_domain_categories(json_file: str = "domain_categories.json"):
    """Import domain categories from existing JSON file to database."""
    db = Database()
    
    json_path = Path(json_file)
    if not json_path.exists():
        print(f"No {json_file} found, skipping migration")
        return 0
    
    with open(json_path) as f:
        categories = json.load(f)
    
    count = 0
    for domain_name, category_str in categories.items():
        try:
            category = EmailCategory(category_str)
        except ValueError:
            category = EmailCategory.UNCATEGORIZED
        
        domain = Domain(
            name=domain_name,
            category=category,
            email_count=0,
            last_seen=None,
            sample_senders=[]
        )
        db.upsert_domain(domain)
        count += 1
        print(f"  Imported: {domain_name} -> {category.value}")
    
    print(f"\nMigrated {count} domain categories")
    return count


if __name__ == "__main__":
    migrate_domain_categories()
