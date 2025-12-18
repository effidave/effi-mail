# Task Management with effigit

Use the **effigit** MCP server to manage tasks as GitHub issues. This provides persistent task tracking with full history and collaboration support.

## Setup

The context (e.g., `percayso-pdx`) is already set in `.effigit/context` and will be used automatically.

---

## Creating Tasks

Create tasks for planned work. Document is optional - create tasks before identifying specific documents:

```python
# Task without document (research, planning, etc.)
effigit.create_task(
    repo_path="c:\Users\DavidSant\effi-work",
    description="Find a precedent for data processing agreement"
)

# Task with document
effigit.create_task(
    repo_path="c:\Users\DavidSant\effi-work",
    description="Review liability caps in MSA",
    document="drafts/msa-v2.md"
)
```

---

## Viewing Tasks

```python
# List all tasks in current context
effigit.list_tasks(repo_path="c:\Users\DavidSant\effi-work")

# List tasks for a specific document
effigit.list_tasks_for_document(
    repo_path="c:\Users\DavidSant\effi-work",
    document="drafts/msa-v2.md"
)

# Get metadata for a specific task
effigit.show_task(repo_path="c:\Users\DavidSant\effi-work", issue=5)

# Get all comments on a task (for reading discussions, summarizing threads)
effigit.get_issue_comments(repo_path="c:\Users\DavidSant\effi-work", issue=5)
```

---

## Updating Tasks

Refine task requirements before starting work, or fix an incorrect document path:

```python
# Update title and description
effigit.update_task(
    repo_path="c:\Users\DavidSant\effi-work",
    issue=5,
    title="[Task] Review and amend liability caps",
    description="Focus on aggregate caps and exclusions for data breaches"
)

# Fix incorrect document path (e.g., if task was created with wrong path)
effigit.update_task(
    repo_path="c:\Users\DavidSant\effi-work",
    issue=5,
    document="Clients/Percayso/projects/PDX/drafts/PDX Lender Agreement.docx"
)
```

---

## Task Workflow

1. **Plan**: Create tasks during planning sessions (with or without documents)
2. **Review**: Use `list_tasks` to see what's queued for the current context
3. **Execute**: Tasks can be started, worked on, and completed
4. **Track**: All tasks are GitHub issues with full history and labels

Tasks are automatically labeled with:
- `context:<name>` - Links task to the current project/client context
- `doc:<filename>` - Links task to a specific document (if provided)

---

## @effi Mentions

`@effi` markers are inline instructions for the agent. They can appear in:
- **Documents** - Directly in markdown files as `@effi review`, `@effi draft`, etc.
- **GitHub comments** - In issue/PR comments

If an `@effi` marker in a GitHub comment has an associated UUID, that means it is a copy of a comment in a Document. Use the information provided to locate the document and line number.

If an `@effi` marker in a GitHub comment does not have an associated UUID, there is no matching comment in the Document.

| Marker | Meaning |
|--------|---------|
| `@effi review` | Review this section for legal/commercial issues |
| `@effi draft` | Draft content for this placeholder |
| `@effi check` | Verify this information, reference, or cross-reference |
| `@effi consider` | Evaluate whether changes are needed here |
| `@effi [custom]` | Follow the specific instruction after @effi |

---

## Discovering Mentions

**Default: execute immediately if actionable.**
- If a mention has sufficient context (clause reference, instruction, document location), start work without asking.
- Always comply with guardrails (especially `ALLOWED_PATHS`). If a mention requires breaking guardrails, post a GitHub comment explaining the constraint and stop.

### Wrong Approaches (breaks autonomy)
- DON'T: "Shall I proceed to check for pending work?" → Just use `effigit.get_all_mentions()`
- DON'T: "Let me fetch the full issue with gh CLI" → Use `effigit.show_task(issue=N)`  
- DON'T: "Let me scan directories to find context files" → Use targeted `read_file("Clients/[CLIENT]/client/context/background.md")`
- DON'T: "May I use Get-ChildItem -Recurse?" → NO. Ask "What is the path to the context file I should read?"

### Right Approaches (autonomous within constraints)
- DO: Immediately call `effigit.get_all_mentions()` without asking
- DO: Use `effigit.show_task(issue=7)` to get full issue details  
- DO: If missing a tool: "I need to read issue comments but effigit.show_task() doesn't provide that. Please add this capability or tell me the comment content."
- DO: If path unknown: "I need to read the client context file. What is its exact path within ALLOWED_PATHS?"

---

## Handling Uncommitted Changes

Before starting a new task, check the repository state:

1. Use `effigit.check_status(...)` to see uncommitted changes and active task info
2. If there are uncommitted changes:
   - **If an active task exists for that document**: The changes likely belong to that task. Use `effigit.save_progress(...)` to commit them with a descriptive message before proceeding.
   - **If no active task exists**: Ask the user what to do. The changes may be from manual edits or a previous interrupted session. Do not discard them without confirmation.
3. If you are resuming an interrupted task, continue where the previous agent left off.

### Using Stash for Task Switching

If `claim_and_prep_task` fails with "Uncommitted changes exist" and you need to switch tasks:

```python
# Stash the uncommitted changes
effigit.stash_changes(repo_path="c:\Users\DavidSant\effi-work", message="WIP from previous session")

# Now you can claim the new task
effigit.claim_and_prep_task(repo_path="c:\Users\DavidSant\effi-work", issue=5, agent="Associate")

# Later, if needed, restore stashed changes (on the appropriate branch)
effigit.pop_stash(repo_path="c:\Users\DavidSant\effi-work")
```

---

## Mandatory Workflow for @effi Mentions

Use ONLY these MCP tools for task management:

1. `effigit.get_all_mentions(...)` — Find pending work (NEVER use `gh issue list` or similar)
2. `effigit.notice_mention(...)` — Claim before starting (NEVER use `gh api` or similar)
3. `effigit.locate_mention(...)` — If UUID provided, find file/line (file must be within `ALLOWED_PATHS`)
4. Execute the action using available tools (read_file, replace_string_in_file, etc.)
5. `effigit.save_progress(...)` or `effigit.complete_task(...)` — Commit changes (converts docx→md, pushes)
6. `effigit.processed_mention(...)` — Post brief summary to the @effi comment (NEVER use `gh issue comment` or similar)

### Critical: Always Commit Before Responding

- **ALWAYS** use `effigit.save_progress(...)` or `effigit.complete_task(...)` to commit your work BEFORE calling `processed_mention(...)`.
- These functions commit changes, convert docx→md, and push to GitHub.
- If you respond with `processed_mention(...)` without committing first, all document changes will be LOST.
- The workflow is: edit → commit (save_progress or complete_task) → respond (processed_mention).

### If You Need Issue Context

- Use `effigit.show_task(...)` to get issue metadata (title, description, labels, status).
- Use `effigit.get_issue_comments(...)` to get the full comment thread for summarizing discussions.
- NEVER use `gh issue view`, `gh pr view`, or `gh api` commands.

### If the Mention Directs You Outside ALLOWED_PATHS

- Post GitHub comment via `effigit.processed_mention`: "[location] is outside allowed paths. Please copy [file] into [allowed folder] to proceed."
- Do not mark complete until unblocked.

---

## Example Workflow

```python
# Example workflow for document editing tasks
mentions = effigit.get_all_mentions(repo_path="c:\Users\DavidSant\effi-work")
effigit.notice_mention(repo_path="c:\Users\DavidSant\effi-work", comment_id=123456789)

# If you need issue metadata:
task_details = effigit.show_task(repo_path="c:\Users\DavidSant\effi-work", issue=7)

# If you need to read the full comment thread:
comments = effigit.get_issue_comments(repo_path="c:\Users\DavidSant\effi-work", issue=7)

# If UUID exists and within ALLOWED_PATHS:
location = effigit.locate_mention(repo_path="c:\Users\DavidSant\effi-work", uuid="abc123")

# If you need to read context files, use targeted reads:
context = read_file("Clients/[CLIENT]/projects/[PROJECT]/context/background.md")

# Execute the work (edit docx files)...

# CRITICAL: Commit changes BEFORE responding to the mention
# This saves your work and converts docx to markdown
effigit.complete_task(
    repo_path="c:\Users\DavidSant\effi-work",
    message="Added carve-out for gross negligence in clause 5.2"
)
# OR use save_progress if task isn't fully complete yet

# THEN respond to the @effi comment
effigit.processed_mention(
    repo_path="c:\Users\DavidSant\effi-work",
    comment_id=123456789,
    response="Added carve-out for gross negligence in clause 5.2. See PR for details."
)
```
