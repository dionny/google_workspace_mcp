#!/usr/bin/env python3
"""Script to add testing issues to beads from Dec 5 testing session."""
import json
import hashlib
from datetime import datetime, timezone

def generate_hash(content):
    return hashlib.sha256(content.encode()).hexdigest()

def make_issue(id_suffix, title, description, priority, issue_type, labels):
    issue_id = f"google_workspace_mcp-test-{id_suffix}"
    now = datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
    
    content = title + description
    content_hash = generate_hash(content)
    
    return {
        "id": issue_id,
        "content_hash": content_hash,
        "title": title,
        "description": description,
        "status": "open",
        "priority": priority,
        "issue_type": issue_type,
        "created_at": now,
        "updated_at": now,
        "labels": labels
    }

issues = [
    # BUGS
    make_issue("bug-1", 
        "modify_doc_text with empty text (delete) fails despite preview showing success",
        """## Problem
When using `modify_doc_text` with `search`, `position='replace'`, and `text=''` (empty string to delete), the preview mode shows success but actual execution fails.

## Steps to Reproduce
```bash
# Preview works:
modify_doc_text(doc_id, search="[TEST]", position="replace", text="", preview=True)
# Returns: {"preview": true, "would_modify": true, "operation": "replace", ...}

# Actual execution fails:
modify_doc_text(doc_id, search="[TEST]", position="replace", text="")
# Error: Invalid requests[1].insertText: Insert text requests must specify text to insert.
```

## Root Cause
The code is generating an `insertText` request with empty text instead of a `deleteContentRange` request.

## Impact
Users cannot delete text using `modify_doc_text` directly - must use `batch_modify_doc` with `delete_text` operation type instead.

## Suggested Fix
When `text=""` is provided with `position="replace"`, generate a `deleteContentRange` request instead of `insertText` + `deleteContentRange`.

## Workaround
Use `batch_modify_doc` with `{"type": "delete_text", "start_index": X, "end_index": Y}`""",
        0, "bug", ["google-docs", "critical", "api"]),

    make_issue("bug-2",
        "find_and_replace_doc lacks preview mode (inconsistent with modify_doc_text)",
        """## Problem
`find_and_replace_doc` does not support `preview=True` parameter while `modify_doc_text` does.

## Steps to Reproduce
```bash
find_and_replace_doc(doc_id, find_text="old", replace_text="new", preview=True)
# Error: got an unexpected keyword argument 'preview'
```

## Expected Behavior
Should support preview mode like `modify_doc_text` to show what would be replaced before executing.

## Impact
Users cannot safely preview global find/replace operations before committing them.

## Suggested Fix
Add `preview` parameter to `find_and_replace_doc` that returns count and locations of matches without making changes.""",
        1, "bug", ["google-docs", "consistency", "api"]),

    make_issue("bug-3",
        "find_and_replace_doc returns plain string instead of structured JSON",
        """## Problem
`find_and_replace_doc` returns a plain text string while other tools like `modify_doc_text` return structured JSON with `position_shift`, `affected_range`, etc.

## Current Response
```
"Replaced 3 occurrence(s) of 'old' with 'new' in document ABC. Link: ..."
```

## Expected Response
```json
{
  "success": true,
  "operation": "find_replace",
  "occurrences_replaced": 3,
  "find_text": "old",
  "replace_text": "new",
  "position_shift": -6,  // If replacements changed document length
  "affected_ranges": [...],  // Locations of replacements
  "document_link": "..."
}
```

## Impact
- Inconsistent API response format
- No way to track position shifts for follow-up operations
- Harder to parse programmatically""",
        2, "bug", ["google-docs", "consistency", "api"]),

    make_issue("bug-4",
        "Operation history not being recorded automatically",
        """## Problem
The `get_doc_operation_history` shows 0 operations even after multiple edits through `modify_doc_text` and `batch_modify_doc`.

## Steps to Reproduce
```bash
# Make several edits
modify_doc_text(doc_id, start_index=24, text="Test")
batch_modify_doc(doc_id, operations=[...])

# Check history
get_doc_operation_history(doc_id)
# Returns: {"operations": [], "total_operations": 0, "undoable_count": 0}
```

## Expected Behavior
Operations should be automatically recorded to enable undo functionality.

## Impact
- Undo functionality is unusable
- No way to track what operations have been performed

## Questions
- Is recording supposed to be automatic?
- Is there a flag to enable it?
- Do we need to call `record_doc_operation` manually?""",
        1, "bug", ["google-docs", "undo", "history"]),

    # MISSING FEATURES - Text Formatting
    make_issue("feat-1",
        "Add hyperlink support to modify_doc_text",
        """## Problem
Cannot add hyperlinks to text using `modify_doc_text`. The Google Docs API supports this via `updateTextStyle` with `link` field.

## Proposed API
```python
modify_doc_text(
    document_id="...",
    search="click here",
    position="replace",
    link="https://example.com"  # NEW: Add hyperlink
)

# Or add link while inserting:
modify_doc_text(
    document_id="...",
    start_index=100,
    text="Visit our site",
    link="https://example.com"
)
```

## Google Docs API Support
The `updateTextStyle` request supports:
```json
{
  "textStyle": {
    "link": {
      "url": "https://example.com"
    }
  }
}
```

## Use Cases
- Add links to text
- Create table of contents with internal links
- Insert linked references

## Priority
High - very common editing need""",
        1, "feature", ["google-docs", "formatting", "enhancement"]),

    make_issue("feat-2",
        "Add text color support (foreground and background)",
        """## Problem
Cannot change text color using `modify_doc_text`. Current formatting options are limited to: bold, italic, underline, font_size, font_family.

## Proposed API
```python
modify_doc_text(
    document_id="...",
    search="important",
    position="replace",
    foreground_color="#FF0000",  # Red text
    background_color="#FFFF00"   # Yellow highlight
)
```

## Google Docs API Support
The `updateTextStyle` request supports:
```json
{
  "textStyle": {
    "foregroundColor": {"color": {"rgbColor": {"red": 1.0, "green": 0.0, "blue": 0.0}}},
    "backgroundColor": {"color": {"rgbColor": {"red": 1.0, "green": 1.0, "blue": 0.0}}}
  }
}
```

## Use Cases
- Highlight important text
- Color-code content (errors in red, success in green)
- Create visual emphasis

## Format Options
Could accept: hex colors (#FF0000), rgb values, or named colors""",
        2, "feature", ["google-docs", "formatting", "enhancement"]),

    make_issue("feat-3",
        "Add strikethrough formatting option",
        """## Problem
Cannot apply strikethrough formatting using `modify_doc_text`.

## Proposed API
```python
modify_doc_text(
    document_id="...",
    search="deprecated",
    position="replace",
    strikethrough=True
)
```

## Google Docs API Support
The `updateTextStyle` request supports `strikethrough: true`.

## Use Cases
- Mark deleted/deprecated content
- Show revisions
- Track changes style editing""",
        3, "feature", ["google-docs", "formatting", "enhancement"]),

    # MISSING FEATURES - Paragraph Formatting
    make_issue("feat-4",
        "Add paragraph/heading style support",
        """## Problem
Cannot change paragraph style to heading (H1, H2, etc.) or vice versa. Current API only supports text-level formatting (bold, italic) not paragraph-level styles.

## Proposed API
```python
modify_doc_text(
    document_id="...",
    search="My New Section",
    position="replace",
    paragraph_style="HEADING_2"  # Convert to heading
)

# Or on insert:
modify_doc_text(
    document_id="...",
    start_index=100,
    text="New Heading\\n",
    paragraph_style="HEADING_1"
)
```

## Google Docs API Support
The `updateParagraphStyle` request supports:
- NORMAL_TEXT
- HEADING_1 through HEADING_6
- TITLE
- SUBTITLE

## Use Cases
- Create document structure programmatically
- Convert plain text to headings
- Reorganize document hierarchy""",
        2, "feature", ["google-docs", "formatting", "paragraph"]),

    make_issue("feat-5",
        "Add paragraph alignment support (center, right, justify)",
        """## Problem
Cannot change paragraph alignment. Text is always left-aligned by default.

## Proposed API
```python
modify_doc_text(
    document_id="...",
    search="Centered Title",
    position="replace",
    alignment="CENTER"  # or "LEFT", "RIGHT", "JUSTIFIED"
)
```

## Google Docs API Support
The `updateParagraphStyle` request supports:
- START (left for LTR)
- CENTER
- END (right for LTR)
- JUSTIFIED

## Use Cases
- Center titles and headings
- Right-align dates and signatures
- Justify body text""",
        3, "feature", ["google-docs", "formatting", "paragraph"]),

    # MISSING FEATURES - List/Structure Operations
    make_issue("feat-6",
        "Add convert-to-list functionality for existing text",
        """## Problem
`insert_doc_elements` can create new lists, but cannot convert existing paragraphs to bullet/numbered lists.

## Current Limitation
```python
# Can only INSERT a new list:
insert_doc_elements(doc_id, element_type="list", index=100, list_type="UNORDERED", text="Item 1")

# CANNOT convert existing text to list
```

## Proposed API
```python
# Convert existing paragraph(s) to list:
modify_doc_text(
    document_id="...",
    search="- Item one\\n- Item two",  # Or by range
    position="replace",
    convert_to_list="UNORDERED"  # or "ORDERED"
)
```

## Google Docs API Support
The `createParagraphBullets` request can apply bullets to existing paragraphs.

## Use Cases
- Convert dash-prefixed text to proper bullets
- Reformat imported content
- Toggle between list and paragraph format""",
        2, "feature", ["google-docs", "formatting", "lists"]),

    # CONSISTENCY ISSUES
    make_issue("consistency-1",
        "insert_doc_image requires explicit index (no location='end' option)",
        """## Problem
`insert_doc_image` requires explicit `index` parameter while `create_table_with_data` supports optional index (defaults to end).

## Current API
```python
# Table - index optional (appends to end):
create_table_with_data(doc_id, data)  # Works!

# Image - index required:
insert_doc_image(doc_id, image_source="...", index=???)  # Must specify
```

## Proposed Fix
Make `index` optional for `insert_doc_image`, default to end of document (like tables).

Also consider adding:
- `location="end"` / `location="start"` convenience parameter
- `after_heading="Section Name"` for heading-based positioning

## Impact
Users must call `get_doc_info` first to get document length before inserting images.""",
        2, "task", ["google-docs", "consistency", "dx"]),

    make_issue("consistency-2",
        "get_doc_operation_history doesn't follow user_google_email pattern",
        """## Problem
`get_doc_operation_history` does not accept `user_google_email` parameter, breaking the pattern of all other doc tools.

## Steps to Reproduce
```bash
get_doc_operation_history(doc_id, user_google_email="...")
# Error: got an unexpected keyword argument 'user_google_email'
```

## Expected
Should accept `user_google_email` like all other tools for consistency.

## Impact
Confusing when some tools require user_google_email and others don't.""",
        3, "bug", ["google-docs", "consistency", "api"]),
]

if __name__ == "__main__":
    # Append to issues.jsonl
    with open('.beads/issues.jsonl', 'a') as f:
        for issue in issues:
            f.write(json.dumps(issue) + '\n')

    print(f"Added {len(issues)} issues:")
    for issue in issues:
        print(f"  [{issue['issue_type']}] {issue['id']}: {issue['title']}")

