# effi-mail Enhancement TODO

This document tracks planned enhancements for effi-mail.

## High Priority

### FTS5 Full-Text Search
**Status:** Planned

Add SQLite FTS5 virtual table for efficient full-text search of email subjects and bodies.

**Benefits:**
- Fast subject/body searches instead of LIKE queries
- Ranking and scoring of search results
- Support for phrase queries and boolean operators

**Implementation Notes:**
- Create FTS5 virtual table: `CREATE VIRTUAL TABLE emails_fts USING fts5(subject, body_preview, content=emails, content_rowid=rowid)`
- Populate on sync with triggers
- Add `search_emails_fulltext` tool

---

## Medium Priority

### Subfolder Support for Direct Outlook Queries
**Status:** Planned

Add `include_subfolders` parameter to `search_outlook_direct` and related tools.

**Current Limitation:**
- Direct Outlook queries only search top-level folders (Inbox, Sent Items)
- Subfolders under Inbox are not searched

**Implementation Notes:**
- Recursively iterate folder.Folders collection
- May impact performance for deep folder hierarchies
- Consider depth limit parameter

---

### Batch Backfill of internet_message_id
**Status:** Optional

Script to backfill `internet_message_id` for existing database records.

**Current Approach:**
- Lazy backfill: populate when email is next accessed via Outlook
- This script would do a one-time batch update

**Implementation Notes:**
- Iterate emails in DB without internet_message_id
- Lookup each by EntryID in Outlook
- Extract PR_INTERNET_MESSAGE_ID and update record
- Handle case where EntryID is stale (email moved/deleted)

---

## Low Priority

### Recipient Domain Index Optimization
**Status:** Considered

Pre-compute and index recipient domains for faster queries.

**Current Implementation:**
- `recipient_domains` column stores comma-separated domains
- Queries use LIKE for matching

**Potential Improvement:**
- Separate `email_recipient_domains` junction table
- Proper indexing for exact matches
- Trade-off: More complex sync logic

---

## Completed

### internet_message_id Permanent Identifier
**Completed:** December 2025

Added `internet_message_id` field (RFC2822 Message-ID) as permanent email identifier that persists across folder moves.

### recipient_domains Computed Column
**Completed:** December 2025

Added `recipient_domains` column computed at sync time from To/CC recipients for efficient recipient-based queries.

### Client-Centric Search Tools
**Completed:** December 2025

Added 6 new tools:
- `search_emails_by_client` - DB query by client
- `search_outlook_by_client` - Direct Outlook query by client
- `search_outlook_direct` - Flexible direct Outlook query
- `get_email_by_id` - Fetch by EntryID or internet_message_id
- `sync_emails_by_client` - Targeted sync for client
- `sync_email_by_id` - Sync single email

### effi-work MCP Integration
**Completed:** December 2025

Refactored to use MCP protocol for effi-work communication:
- `get_client_identifiers_from_effi_work()` now calls effi-work via MCP client
- Removed `db.get_client_identifiers()` - domains now come from effi-work as single source of truth
- `db.search_emails_by_client()` now takes `domains` and `contact_emails` directly
- effi-work's `sync_clients_to_mail` tool is deprecated

---

## effi-work Integration Concerns

### Performance
- [ ] **Session overhead**: Each client lookup spawns a new effi-work subprocess via MCP. For bulk operations (e.g., triaging 100 emails), this could add significant latency
- [ ] **Consider session pooling**: Keep an effi-work session alive across multiple lookups in the same request

### Reliability
- [ ] **Error handling for effi-work unavailability**: Currently returns empty domains if effi-work fails. Should we fall back to a local cache?
- [ ] **Timeout handling**: No explicit timeout on MCP calls - a hung effi-work server could block effi-mail indefinitely

### Data Completeness
- [x] **Contact emails now parsed**: effi-work updated to use proper YAML parsing and now exposes `context.contact_emails` directly. effi-mail already handles this field.

### Testing
- [ ] **Integration test with real effi-work**: Current tests mock the MCP calls. Add a live integration test that actually calls effi-work

### Future Considerations
- [ ] **Versioned API contract**: If effi-work's tool response format changes, effi-mail will silently get wrong data. Consider adding a version check or schema validation
