# ADR 002: Google Docs Text-Anchored Comments Are Not Supported by API

**Date:** 2025-12-08

**Status:** Accepted

**Context:** We investigated adding functionality to create comments on specific text ranges within Google Documents (text-anchored comments) rather than just file-level comments.

## Decision

We will **NOT** implement text-anchored comment creation for Google Docs because it is **not possible** with the current Google APIs.

## Rationale

### What We Investigated

1. **Google Drive API** - Has a `comments.create()` method with an `anchor` parameter
   - The `anchor` parameter is explicitly **ignored** for Google Docs editor files
   - Per official documentation: "Google Workspace editor apps treat these comments as unanchored"
   - The anchor feature only works for non-editor files (PDFs, images, etc.)

2. **Google Docs API** - Has no comment-related Request types
   - No `CreateCommentRequest`, `AddCommentRequest`, or `InsertCommentRequest`
   - The Docs API only handles document structure, formatting, and content
   - Comments are handled exclusively via the Drive API

3. **Internal Implementation** - Google uses a proprietary "kix" anchor format
   - This format is not exposed to developers
   - Cannot be constructed or specified via any API

### Evidence

- **Official Documentation**: [Google Drive API - Manage Comments](https://developers.google.com/workspace/drive/api/guides/manage-comments)
  - Explicitly states anchor parameter is not supported for editor files

- **Stack Overflow**: [Multiple discussions from 2017-2024](https://stackoverflow.com/questions/41929652/is-it-possible-to-add-a-comment-attributed-to-specific-text-within-google-docs-u)
  - Developers have consistently confirmed this limitation

- **Issue Tracker**: [Issue 36763384](https://issuetracker.google.com/issues/36763384)
  - Feature request from 2016, still open with no resolution

### What IS Possible via Drive API

#### Currently Implemented ✅

1. **List/Read Comments** (`read_doc_comments`)
   - Retrieves all comments with replies
   - Gets author, content, timestamps, resolved status
   - Shows reply threads
   - **Fields we request:** `id, content, author, createdTime, modifiedTime, resolved, replies`

2. **Create File-Level Comments** (`create_doc_comment`)
   - Creates unanchored comments on the document
   - Comments appear in "All Comments" view
   - Not attached to specific text

3. **Reply to Comments** (`reply_to_comment`)
   - Add replies to existing comment threads
   - Works for both anchored (UI-created) and unanchored comments

4. **Resolve Comments** (`resolve_comment`)
   - Mark comments as resolved via special reply with `action: "resolve"`
   - Works for both anchored and unanchored comments

#### Available But NOT Yet Implemented 🔶

These operations are supported by the Drive API but we haven't implemented them:

1. **Update/Edit Comment Content** 
   - API: `comments.update()` or `comments.patch()`
   - Can modify comment content after creation
   - Can mark comments as resolved directly (without reply)
   - Can update other comment metadata

2. **Delete Comments**
   - API: `comments.delete()`
   - Permanently remove comments
   - Deleted comments are marked with `deleted: true` field

3. **Update/Edit Replies**
   - API: `replies.update()`
   - Modify reply content after posting
   - Change reply metadata

4. **Delete Replies**
   - API: `replies.delete()`
   - Remove specific replies from comment threads

5. **Read Additional Comment Fields**
   - `anchor` - For UI-created anchored comments, contains position data
   - `quotedFileContent` - The text that anchored comments reference
   - `deleted` - Whether comment has been deleted
   - `htmlContent` - HTML-formatted comment content
   - Additional author fields (photo URL, email, etc.)

6. **Reopen Resolved Comments**
   - Update comment with `resolved: false`
   - Via `comments.update()` or `comments.patch()`

### What is NOT Possible ❌

1. **Create Text-Anchored Comments**
   - Cannot specify start/end indices or text selection when creating
   - The `anchor` parameter in `comments.create()` is ignored for Google Docs
   - Only possible manually via Google Docs UI

2. **Modify Anchor Position**
   - Cannot change where an anchored comment points
   - Cannot convert file-level comment to anchored comment

3. **Create Anchored Comments via Docs API**
   - Docs API has no comment-related Request types
   - No `CreateCommentRequest`, `AddCommentRequest`, or similar

## Consequences

### Current Implementation

Our comment tools (`read_doc_comments`, `create_doc_comment`, `reply_to_comment`, `resolve_comment`):
- **Implement:** List, create (file-level), reply, resolve
- **Do NOT implement yet:** Update, delete, patch, reopen, reading anchor/quotedFileContent fields
- Use the Drive API for all comment operations
- Work consistently with Sheets and Slides (which have the same limitation)

### Potential Enhancements

We could add these tools based on available Drive API capabilities:

1. **`update_doc_comment`** - Edit comment content after creation
2. **`delete_doc_comment`** - Permanently remove comments
3. **`reopen_doc_comment`** - Unmark resolved comments
4. **`update_doc_reply`** - Edit reply content
5. **`delete_doc_reply`** - Remove specific replies
6. Enhanced `read_doc_comments` - Include `anchor` and `quotedFileContent` fields for UI-created anchored comments

These would be straightforward to implement following the existing pattern in `core/comments.py`.

### User Impact

Users who need text-anchored comments must:
1. Use the Google Docs UI manually to create anchored comments
2. Or create file-level comments via API and reference specific sections in text (e.g., "In paragraph 3, line 5...")
3. Can still read, reply to, and resolve manually-created anchored comments via our API tools

### Documentation

- Document this limitation in `PROMPT.md` and tool docstrings
- Close or mark as "blocked" any beads issues related to creating text-anchored comments
- Consider creating new issues for the unimplemented-but-possible features above
- Update this ADR if Google ever adds API support for creating anchored comments

## Alternatives Considered

### 1. Named Ranges as Comment Proxies
**Rejected** - Named ranges can mark text locations but:
- They're not comments (different UI, different purpose)
- Users would need to check two places for feedback
- No way to "resolve" a named range like a comment

### 2. Suggestions/Review Mode
**Rejected** - The Docs API has suggestion-related formatting (e.g., `SuggestTextStyleUpdate`) but:
- These are for tracking changes, not comments
- No way to programmatically enable suggestion mode
- Not the same workflow as comments

### 3. Third-Party Workarounds
**Rejected** - Some proposed:
- Using footnotes to simulate comments
- Using text highlighting with descriptions
- Using bookmarks with external mapping

All rejected because they don't provide the actual comment functionality users expect (threads, resolution, notifications, etc.)

## Future Considerations

If Google adds API support for text-anchored comments in the future:

1. **Monitor these resources:**
   - [Google Workspace API Release Notes](https://developers.google.com/workspace/releases)
   - Issue Tracker [36763384](https://issuetracker.google.com/issues/36763384)

2. **Implementation approach:**
   - Enhance `create_doc_comment` with optional range parameters
   - Leverage existing positioning modes (search, heading, index, etc.)
   - Update `read_doc_comments` to show anchor information
   - Add comprehensive tests

3. **Priority:** This would be a high-value feature given the limitation it addresses

## Related

- Beads issue `google_workspace_mcp-0f4f` (should be closed as "blocked by API limitation")
- `core/comments.py` - Generic comment factory for all Google Workspace apps
- `gdocs/docs_tools.py` lines 8006-8013 - Comment tool registration

## References

- [Google Drive API - Comments](https://developers.google.com/drive/api/reference/rest/v3/comments)
- [Google Drive API - Manage Comments Guide](https://developers.google.com/workspace/drive/api/guides/manage-comments)
- [Google Docs API - Request Types](https://developers.google.com/docs/api/reference/rest/v1/documents/request)
- [Stack Overflow Discussion](https://stackoverflow.com/questions/41929652/is-it-possible-to-add-a-comment-attributed-to-specific-text-within-google-docs-u)

