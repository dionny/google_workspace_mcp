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

### What IS Possible

✅ **Create file-level (unanchored) comments** - Comments appear in the document but not attached to specific text

✅ **Read existing anchored comments** - Comments created manually in the UI can be retrieved with their `anchor` and `quotedFileContent` fields

✅ **Reply to comments** - Both anchored and unanchored comments can receive replies

✅ **Resolve comments** - Both anchored and unanchored comments can be marked as resolved

### What is NOT Possible

❌ **Create comments on specific text ranges** - Cannot specify start/end indices or text selection

❌ **Anchor comments programmatically** - No way to target specific paragraphs, sentences, or words

## Consequences

### Current Implementation

Our comment tools (`read_doc_comments`, `create_doc_comment`, `reply_to_comment`, `resolve_comment`) will continue to:
- Create file-level comments only
- Use the Drive API for all comment operations
- Work consistently with Sheets and Slides (which have the same limitation)

### User Impact

Users who need text-anchored comments must:
1. Use the Google Docs UI manually
2. Or create file-level comments and reference specific sections in the comment text (e.g., "In paragraph 3, line 5...")

### Documentation

- Document this limitation in `PROMPT.md` and tool docstrings
- Close or mark as "blocked" any beads issues related to text-anchored comments
- Update this ADR if Google ever adds API support for this feature

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

