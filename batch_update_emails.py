import json

# Read the JSON file
with open(r'c:\Users\DavidSant\effi-mail\Untitled-1.json', 'r', encoding='utf-8') as f:
    emails = json.load(f)

# Define updates for batch 3 (emails 19-32)
# Using EntryID as the unique identifier
updates = {
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003E14EE67A0000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "IT support issue",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003E14EE6790000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "IT support issue",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003DD92E2150000": {
        "proposedClient": "Didimo",
        "proposedMatter": "Consultancy Agreement Review",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003DD92E2140000": {
        "proposedClient": "Didimo",
        "proposedMatter": "Consultancy Agreement Review",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003DD92E2130000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "System notification",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003DD92E2110000": {
        "proposedClient": "Ko-fi",
        "proposedMatter": "Terms & Conditions",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003DAAC689D0000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "Call message relay",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003DAAC68960000": {
        "proposedClient": "Didimo",
        "proposedMatter": "Consultancy Agreement Review",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003DAAC68660000": {
        "proposedClient": "Extend",
        "proposedMatter": "Client tracking",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003D70BF4310000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "Call message relay",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003D70BF41D0000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "HR - CPD compliance",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003D70BF41B0000": {
        "proposedClient": "Extend",
        "proposedMatter": "Client tracking",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003E10E02E90000": {
        "proposedClient": "Streets Heaver Computer Systems Limited",
        "proposedMatter": "SaaS Template Agreement",
        "confidenceRating": 5
    }
}

# Apply updates
updated_count = 0
for email in emails:
    entry_id = email.get("EntryID")
    if entry_id in updates:
        email["proposedClient"] = updates[entry_id]["proposedClient"]
        email["proposedMatter"] = updates[entry_id]["proposedMatter"]
        email["confidenceRating"] = updates[entry_id]["confidenceRating"]
        updated_count += 1
        print(f"Updated email {updated_count}: {email['Subject'][:50]} -> {email['proposedClient']}")

# Save back to file
with open(r'c:\Users\DavidSant\effi-mail\Untitled-1.json', 'w', encoding='utf-8') as f:
    json.dump(emails, f, indent=4, ensure_ascii=False)

print(f"\nTotal emails updated: {updated_count}")
print("File saved successfully!")
