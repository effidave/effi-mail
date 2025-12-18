# Delegation Patterns

This document covers how to delegate tasks to subagents. This is primarily for LegalDirector agents who coordinate work across Associates and Partners.

---

## Choosing the Right Subagent

### When to Use Associate Subagents

Use **Associate** subagents for:
- **Document editing** based on specific, clear instructions
- **Mechanical tasks** with defined scope (e.g., "update clause 8.2 to add aggregate cap")
- **Localized changes** to specific sections or clauses
- Tasks where the work is clearly specified and doesn't require strategic judgment

### When to Use Partner Subagents

Use **Partner** subagents for:
- **Planning and strategy** - understanding overall aims and designing approach
- **Review and quality control** - assessing work against client objectives
- **Complex drafting** requiring judgment (e.g., abstracting common clauses from multiple schedules)
- **Considering overall context** - understanding commercial position, risk appetite, negotiation strategy
- **Legal analysis** - advising on risks, alternatives, and implications

---

## Two-Stage Pattern (Associate + Partner Review)

When multiple tasks need careful execution and review, use this delegation pattern.

### Stage 1: Associate Execution

Run a subagent to complete the editing task. The Associate must:

1. **Pre-analysis**: Ensure the baseline markdown is available:
   - If the docx has been committed, the markdown version will be automatically available (via GitHub workflow)
   - The markdown file has the same name as the docx but with .md extension (e.g., `drafts/msa-v2.docx` → `drafts/msa-v2.md`)
   - If working with a new/uncommitted file, commit the docx first to trigger automatic markdown conversion
   - Read the corresponding .md file for document structure and analysis
   - **Important**: The .md file is for reference only—all edits must be made to the .docx file

2. **Complete the task**: Make the required edits to the **.docx file** based on the task instructions.
   - Only access files specifically mentioned in the task
   - Do not explore or modify other files
   - **Critical**: Edit the .docx file, not the .md file. The markdown is for analysis only.
   - Use the .md version to understand structure and generate draft text, but apply all changes to the .docx

3. **Commit the changes**: After editing the docx, commit it to trigger markdown regeneration:
   ```python
   effigit.save_progress(repo_path="c:\Users\DavidSant\effi-work", message="Updated [clause/section] per task instructions")
   ```
   Or if the task is complete:
   ```python
   effigit.complete_task(repo_path="c:\Users\DavidSant\effi-work", message="Completed: [brief description of changes]")
   ```
   This will create a commit and, if using complete_task, open a pull request.
   
   **⚠️ NEVER use `close_issue` for document editing tasks** — it does NOT commit your work!

4. **Respond to the mention**: After committing, call `processed_mention` to post a summary:
   ```python
   effigit.processed_mention(repo_path="c:\Users\DavidSant\effi-work", comment_id=..., response="Summary of changes made")
   ```

5. **Add PR comment**: Post a useful comment on the pull request explaining:
   - What changes were made and why
   - Key considerations or decisions
   - Any relevant context for reviewers

6. **Report back**: Provide a clear summary of:
   - What sections/clauses were edited
   - What changes were made
   - Any issues or questions encountered

### Stage 2: Partner Review

Run a Partner subagent to perform quality control. The reviewing Partner must:

1. **Get updated markdown**: Ensure the edited docx has been committed to trigger automatic markdown conversion. The updated .md file (same filename as the docx, but .md extension) will reflect the changes.

2. **Compare states**: Review the baseline markdown vs the updated markdown to identify all changes.
   - Use git diff or read both versions to compare
   - The GitHub workflow automatically maintains markdown versions alongside docx files

3. **Quality check**:
   - Does the work properly complete the original task?
   - Are there any unintended changes beyond the task scope?
   - Are there any new risks for the client?
   - Is the drafting quality appropriate?

4. **Complete or return**:
   - If everything looks good: Add a PR comment with your review summary, then mark the task as complete
   - If issues found: Add a PR comment with specific feedback for revision

### Example: Two-Stage Delegation

```
LegalDirector:
  ↓
  → Run Associate subagent for Task #5 (simple editing task)
      Associate:
        1. Reads baseline drafts/msa-v2.md for structure/analysis
        2. Edits drafts/msa-v2.docx (the actual source file) per instructions
        3. Commits using effigit.complete_task() to create PR
        4. Posts PR comment explaining changes
        5. Reports summary of changes made
  ↓
  → Run Partner subagent for review
      Reviewing Partner:
        1. Reads updated drafts/msa-v2.md (auto-updated after commit)
        2. Compares baseline vs updated markdown
        3. Checks task completion and risk assessment
        4. Decides: Mark complete OR request revisions
```

### When to Use Two-Stage Pattern

- **ALL document editing tasks** - whether simple or complex
- Any task that results in changes to .docx files
- Complex drafting requiring judgment (Partner does the drafting, Partner reviews)
- Multiple related editing tasks requiring consistent quality control
- When client position needs careful protection

**For complex drafting:** Use Partner for execution (Stage 1) and Partner for review (Stage 2). The two-stage pattern ensures quality control even when the execution requires significant judgment.

---

## Direct Partner Delegation

For tasks that do **NOT involve document edits**, run a Partner subagent directly without the review stage.

### Example: Planning and Analysis

```
LegalDirector:
  ↓
  → Run Partner subagent for Task #8 (planning task)
      Partner:
        1. Read schedules 1-4 to identify common clauses
        2. Analyze patterns and recommend abstraction approach
        3. Report analysis and recommended strategy
```

### When to Delegate Directly to Partner

- **Planning work** (e.g., "Design approach for IP provisions") - no document edits yet
- **Strategic review or risk analysis** - advisory work, not drafting
- **Research tasks** - finding precedents, analyzing options
- **Analysis without implementation** - understanding commercial context, identifying issues

**If the task then leads to document edits**, switch to the two-stage pattern for execution + review.
