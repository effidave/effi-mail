---
name: organizeEmailsByTopic
description: Organize emails from a folder into topic-based subfolders using batch archiving
argument-hint: Source folder path and optional categorization preferences
---
Organize emails from the specified Outlook folder into topic-based subfolders.

## Process

1. **Search** the source folder to retrieve all emails (use a wide date range if needed)
2. **Analyze** the emails to identify logical groupings based on:
   - Sender domain (external clients/companies)
   - Subject line patterns (project names, matter references)
   - Recipient patterns (team distributions, internal vs external)
   - Content themes (HR, admin, client work, notifications)
3. **Propose categories** - group emails into meaningful topics such as:
   - Client/company names for external correspondence
   - Project or matter names
   - Internal topics (HR, team meetings, capacity, admin)
   - Notification types (budget alerts, system notifications)
4. **Create subfolders** under the source folder for each topic
5. **Batch archive** emails to their respective subfolders using `create_path=true` for new folders
6. **Report progress** with counts per folder and any failures

## Guidelines

- Process emails in parallel batches for efficiency
- Use descriptive folder names that will remain meaningful over time
- Group related emails even if from different senders (e.g., all emails about one client matter)
- Keep internal team communications separate from client work
- Create a catch-all folder for miscellaneous items if needed

## Expected Output

Summary of:
- Total emails organized
- New folders created
- Emails per folder
- Any errors encountered
