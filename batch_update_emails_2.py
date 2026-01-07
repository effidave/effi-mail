import json
import os

file_path = r'c:\Users\DavidSant\effi-mail\Untitled-1.json'

with open(file_path, 'r', encoding='utf-8') as f:
    data = json.load(f)

updates = {
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003D22472F40000": {
        "proposedClient": "Ko-fi Labs Limited",
        "proposedMatter": "Gold subscription model",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003D22472F30000": {
        "proposedClient": "Ko-fi Labs Limited",
        "proposedMatter": "Gold subscription model",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003D22472F00000": {
        "proposedClient": "Streets Heaver Computer Systems Limited",
        "proposedMatter": "SaaS Template Agreement",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003D22472ED0000": {
        "proposedClient": "Ko-fi Labs Limited",
        "proposedMatter": "Gold subscription model",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003CF3603C20000": {
        "proposedClient": "Ko-fi Labs Limited",
        "proposedMatter": "Gold subscription model",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003CF3603C00000": {
        "proposedClient": "Ko-fi Labs Limited",
        "proposedMatter": "Gold subscription model",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003CF3603AC0000": {
        "proposedClient": "Avallo Ltd",
        "proposedMatter": "General Advice / Engagement",
        "confidenceRating": 4
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003CF3603A10000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "PD/Knowledge Management",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003CA8458140000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "HR - Bonus Scheme",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C90C69170000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "PD/Knowledge Management",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C90C69150000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "PD/Knowledge Management",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C90C690F0000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "HR - Annual Leave",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C90C690E0000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "HR - Annual Leave",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C90C69070000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "Training/Conferences",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C90C69060000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "Training/Conferences",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C90C69050000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "Training/Conferences",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C7285C0E0000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "Precedents/Knowledge Management",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C7285C0D0000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "Precedents/Knowledge Management",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C5BFD54F0000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "Internal Communications",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C5BFD54A0000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "HR Update",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C496F9350000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "HR - Bonus Scheme",
        "confidenceRating": 5
    },
    "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340003C496F9300000": {
        "proposedClient": "Harper James (Internal)",
        "proposedMatter": "PD/Knowledge Management",
        "confidenceRating": 5
    }
}

updated_count = 0
for item in data:
    entry_id = item.get('EntryID')
    if entry_id in updates:
        item.update(updates[entry_id])
        updated_count += 1

with open(file_path, 'w', encoding='utf-8') as f:
    json.dump(data, f, indent=4)

print(f"Updated {updated_count} emails.")
