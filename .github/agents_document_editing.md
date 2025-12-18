# Document Editing Workflows

This document covers document capabilities, preparation, and best practices for document editing tasks.

---

## Document Capabilities

### Document Management
- Analyze files which you have been asked to review
- Parse document text, outlines, and structure
- Read text of specific clauses, sections, or whole documents - default to holding the whole document in context
- If a file contains instructions at the top, explaining the structure and any bespoke markup, you must read those instructions and can use them, if helpful

### Content Editing
- Add and delete paragraphs, headings, tables, and images
- Search and replace text with options for whole-word matching
- Insert content near specific text or clause numbers
- Edit text within specific formatting runs
- Replace whole clauses, if necessary - but lean towards edits to the existing clause wording

#### Text-Based Paragraph Insertion (RECOMMENDED)
Use `insert_paragraphs_near_text` for inserting paragraphs near specific text. This tool replaces the deprecated `insert_line_or_paragraph_near_text`.

**Benefits:** 
- Works for single or multiple paragraphs (just pass a list with one item for single)
- Returns para_ids enabling chaining operations
- Supports hierarchical numbering via `increase_level_by`
- 5-10x faster than legacy approach

**Example:**
```python
insert_paragraphs_near_text(
    filename="doc.docx",
    target_text="after this text",
    paragraphs=[
        {"text": "First clause", "style": "List Number"},
        {"text": "Subclause a", "style": "List Number", "increase_level_by": 1},
        {"text": "Subclause b", "style": "List Number", "increase_level_by": 1},
        {"text": "Second clause", "style": "List Number"}
    ],
    inherit_numbering=True
)
```

**Result in document:**
```
3. First clause
   3.1. Subclause a
   3.2. Subclause b
4. Second clause
```

**Key Parameters:**
- `paragraphs`: List of dicts with `text`, optional `style`, optional `increase_level_by`
- `inherit_numbering`: True to inherit numId from target location
- `increase_level_by`: Integer (0-9) to adjust numbering level for each paragraph

#### Para ID Workflow (RECOMMENDED for Precise Editing)

Every paragraph in a Word document has a unique `w14:paraId` (8-character hex string like `3DD8236A`). This is the **most stable and precise** way to target paragraphs for editing.

**Why use para_id instead of indices or text search?**
- Paragraph indices **shift** when you add/delete content
- Text search can match **multiple locations** or fail if text changes
- Para IDs are **stable** - they don't change when you edit nearby content

**Getting para_ids:**
1. From markdown exports: Each paragraph shows its para_id
2. From analysis artifacts: `blocks.jsonl` contains para_id for each block
3. From tool responses: All paragraph creation tools return the new para_id(s)

**Para ID return format from tools:**

| Scenario | Format |
|----------|--------|
| Single paragraph created | `... (para_id=3DD8236A)` |
| Multiple paragraphs created | `... para_ids: [3DD8236A, 4EE9347B, 5FF0458C]` |

**Example workflow - chaining operations using para_id:**
```python
# Step 1: Add a paragraph after clause 3.2
add_paragraph_after_clause(filename="doc.docx", clause_number="3.2", text="New clause text")
# Returns: "Paragraph added after clause '3.2'... (para_id=3DD8236A)"

# Step 2: Extract the para_id from the response
new_para_id = "3DD8236A"

# Step 3: Add more content after that paragraph using its para_id
insert_paragraphs_after_para_id(
    filename="doc.docx",
    para_id="3DD8236A",
    paragraphs=[
        {"text": "Subclause (a)", "increase_level_by": 1},
        {"text": "Subclause (b)", "increase_level_by": 1}
    ],
    inherit_numbering=True
)
# Returns: "Inserted 2 paragraph(s) after 3DD8236A... para_ids: [4EE9347B, 5FF0458C]"

# Step 4: Later, modify the first subclause by its para_id
replace_text_by_para_id(filename="doc.docx", para_id="4EE9347B", new_text="Updated subclause (a)")
```

**Available para_id tools:**
- `get_text_by_para_id` - Read paragraph content
- `replace_text_by_para_id` - Replace paragraph content
- `insert_paragraph_after_para_id` - Insert single paragraph after para_id
- `insert_paragraphs_after_para_id` - Batch insert with numbering/level support
- `delete_paragraph_by_para_id` - Delete by stable ID (no index shifting issues)

**Best practice:** Always extract and store para_ids from tool responses when you need to make subsequent edits to the same content.

### Formatting & Styling
- Create custom styles with specific fonts, colors, and formatting
- Apply formatting to text ranges (bold, italic, colors, underline)
- Set background highlighting on text
- Inspect paragraph runs for debugging formatting issues
- Match existing document styles when adding content

### Clause-Based Operations
- Add clauses, sub-clauses, defined terms, clause references, any other text
- Delete clauses, sub-clauses, defined terms, clause references, any other text
- Get text of clauses, sub-clauses, defined terms, clause references, any other text
- List all clause numbers in a document
- List all defined terms in a document
- Renumber clauses to follow conventions within the document
- Reorder defined terms
- Add paragraphs after specific clause numbers (e.g., "1.2.3", "5(a)")

**Note:** For replacing clause content, use para_id-based tools (`replace_text_by_para_id`) instead, as clause ordinals may not be unique in a document.

### Attachment Management
- Add paragraphs after attachments (Schedules, Annexes, Exhibits)
- Add multiple paragraphs to attachments
- Create new attachments after existing ones
- Identify attachment identifiers (e.g., "Schedule 3", "Annex B")

### Comment-Based Operations
- Add comments
- Edit the text of comments
- Delete comments
- Get text of comments
- List all comments in a document
- List all comments by author
- List all comments by prefix, e.g. "For [CLIENT]"
- Change the author of a comment by creating a new comment with the same location and the same text, checking it was created successfully, and only then deleting the original comment
- Extract comment status (active/resolved)

**Comments:**
- `add_comment_for_paragraph` – add a comment to a paragraph
- `get_all_comments` – extract all comments from document (including status: active/resolved)
- `get_comments_by_author` – filter comments by author name
- `get_comments_for_paragraph` – get comments on a specific paragraph
- `update_comment` – update the text of an existing comment

### Document Analysis
- Analyze numbering structure and hierarchy
- Extract document outline with clause numbers
- Get relationship metadata for document blocks
- Get numbering summary for the document

---

## Document Preparation

Before creating editing tasks, prepare the document you'll work on.

### Reviewing Precedents

Use `read_docx_as_md` to analyze a precedent before deciding to use it:

```python
# Returns markdown content as a string (no files created, no commit)
md_content = effigit.read_docx_as_md(
    repo_path="c:\Users\DavidSant\effi-work",
    filepath="Percayso/project/PDX/precedents/standard-contract.docx"
)
# Review the content to decide if this precedent is suitable
```

### Starting from a Precedent

Copy a chosen precedent into the drafts folder:

```python
effigit.copy_to_drafts(
    repo_path="c:\Users\DavidSant\effi-work",
    source="Percayso/project/PDX/precedents/standard-contract.docx",
    destination="Percayso/project/PDX/drafts/client-contract-draft.docx"
)
# Creates docx + md, commits, pushes - ready for editing tasks
```

### Starting from Blank Template

Create a new document with standard styles:

```python
effigit.create_blank_draft(
    repo_path="c:\Users\DavidSant\effi-work",
    destination="Percayso/project/PDX/drafts/new-agreement.docx"
)
# Creates docx + md from styled template, commits, pushes
```

**Note:** If the destination file already exists, the tool will refuse and ask for a different filename.

---

## Best Practices

### Always Verify Changes
After making edits:
- Use `get_document_outline` to verify structure changes
- Use `get_clause_text_by_ordinal` or `get_text_by_para_id` to verify content
- Report what you verified to the user

### Match Existing Formatting
When adding content:
- Inspect the document with `get_document_outline` to identify existing styles
- Use the same styles for new content (e.g., "Untitled subclause 1", "Title Clause")
- Maintain consistency with the document's formatting conventions

### Work Incrementally
- Make one logical change at a time
- Verify each change before proceeding to the next
- If a task has multiple steps, track progress using effigit task tools

### Handle Errors Gracefully
If a tool fails:
- Try an alternative approach (e.g., if clause-based fails, try text-based)
- Explain the issue clearly
- Suggest manual intervention if necessary

### Use Appropriate Tools
- **Clause numbers known?** Use clause-based tools (`add_paragraph_after_clause`)
- **Para ID known?** Use para_id-based tools (`insert_paragraphs_after_para_id`) - most precise
- **Text location known?** Use `insert_paragraphs_near_text` (works for single or multiple paragraphs)
- **Multiple similar edits?** Use `search_and_replace` with whole-word matching
- **Complex formatting?** Use `create_custom_style` for consistency
- **Repetitive tasks?** Create a Python script for automation and reusability

---

## Reporting Progress

When working on document editing tasks, provide clear reports:

### During Work
Report each completed step with verification results.

### After Completion
Summarize all changes made and their locations.

### Example Report Format
```
✔ Step 1: Added definition of "Equipment" to clause 1.1
  - Verified: Definition appears in document outline at paragraph 25

✔ Step 2: Updated 5 references to match new definition
  - Replaced "equipment" with "Equipment" in clauses 3.2, 4.1, 5.3, 7.2, 9.1
  - Verified: All replacements confirmed via search

✔ Step 3: Added explanatory comment
  - Comment added to paragraph 25
  - Status: Active
```

---

## Creating and Executing Python Scripts

When creating automation scripts:

1. Understand the task requirements (e.g., bulk changes, document analysis)
2. Use `create_file` to write the Python script with clear documentation
3. Include proper error handling and user feedback
4. Add usage instructions in comments or docstrings
5. Save to an appropriate location (e.g., `scripts/` directory)
6. Execute the script using `run_in_terminal` with proper command syntax
7. Report execution results to the user

Script execution workflow:

1. Create the script file with `create_file`
2. Run with `run_in_terminal`: `python scripts/your_script.py [arguments]`
3. Monitor the output for errors or completion messages
4. Verify results (e.g., check document changes, read output files)
5. Report success or troubleshoot errors

Common script patterns:
- **Document processing**: Parse and transform document content
- **Text analysis**: Extract and analyze document content
- **Bulk operations**: Process multiple documents or sections efficiently
- **JSON-driven edits**: Read instructions from JSON and apply systematically
