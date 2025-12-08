# 001. Auto-clean blank lines when converting to lists

Date: 2024-12-08

## Status

Accepted

## Context

When using Google Docs API to convert text ranges to lists (via `createParagraphBullets`), each paragraph (separated by `\n`) becomes a separate list item. This creates a problem when users naturally write text with blank lines for readability:

```
Goal 1
Success metric

Goal 2
Success metric

Goal 3
Success metric
```

When this text (with `\n\n` between goals) is converted to a list, the blank lines become empty list items:

1. Goal 1
2. Success metric
3. *(empty)*
4. Goal 2
5. Success metric
6. *(empty)*
7. Goal 3
8. Success metric

This creates an ugly user experience with numbered gaps (items 3, 6, etc. are blank).

During design review template testing, this issue was discovered when filling out Goals and Non-Goals sections. Users (both human and AI agents) naturally write text with blank lines for visual separation, not realizing these will become empty list items.

## Decision

Automatically remove consecutive blank lines (`\n\n` → `\n`) when text will be converted to a list.

**Implementation:**

1. **In `modify_doc_text`** (lines ~1833-1847 in `gdocs/docs_tools.py`):
   - Detect when `convert_to_list` parameter is specified with text insertion/replacement
   - Apply regex cleaning: `re.sub(r'\n\s*\n+', '\n', text)`
   - Log the number of blank lines removed

2. **In `batch_edit_doc`** (via `BatchOperationManager` in `gdocs/managers/batch_operation_manager.py`):
   - Added `_mark_list_conversion_operations()` preprocessing
   - Looks ahead to detect when `insert_text`/`replace_text` will be followed by `convert_to_list`
   - Marks operations with `_will_convert_to_list` flag
   - Added `_clean_text_for_list()` method
   - Applied during `_build_operation_request()` for marked operations

**Behavior:**
- Original: `"Goal 1\n\nGoal 2\n\nGoal 3"` → Items 1, 2 (empty), 3, 4 (empty), 5
- With cleaning: `"Goal 1\nGoal 2\nGoal 3"` → Items 1, 2, 3 (no empty items)

## Consequences

### Positive Consequences

- **Better UX**: No more ugly empty list items
- **Natural authoring**: Users can write text with blank lines for readability without thinking about list conversion
- **AI-friendly**: AI agents generating content don't need special logic to avoid blank lines
- **Consistent behavior**: Works the same for both `modify_doc_text` and `batch_edit_doc`
- **Non-intrusive**: Only triggers when `convert_to_list` is used

### Negative Consequences

- **Loss of control**: Users cannot intentionally create empty list items (edge case)
- **Implicit behavior**: Text is modified before insertion, which might surprise users who expect exact text preservation
- **Additional processing**: Adds regex operation to every list conversion (minimal performance impact)

### Mitigation

- Logging clearly shows when and how many blank lines were removed
- Documentation explains the automatic cleaning behavior
- Users who need empty list items can work around by using `insert_list_item` after initial creation

## Alternatives Considered

### 1. Don't auto-clean, require users to clean manually

**Rejected because:**
- Burdens every user/agent with implementation detail knowledge
- Easy to forget, leading to bad UX
- Defeats "it just works" principle

### 2. Warning in preview mode instead of auto-cleaning

**Rejected because:**
- Still requires user action to fix
- Doesn't solve the problem for non-preview mode
- More complex implementation for marginal benefit

### 3. Validation error when blank lines detected

**Rejected because:**
- Too strict, would break natural workflows
- Forces users to manually clean text
- Bad developer experience

### 4. Optional parameter `preserve_blank_lines=False`

**Considered but deferred:**
- Adds API complexity
- 99% of users will want cleaning behavior
- Can be added later if needed without breaking changes

## Related

- Beads Issues:
  - `google_workspace_mcp-cd06`: Original implementation for `modify_doc_text`
  - `google_workspace_mcp-6e14`: Implementation for `batch_edit_doc`
  - `google_workspace_mcp-9416`: `convert_to_list` support in `batch_edit_doc`
- Testing: Design review template fill-in stress test
- Git commits:
  - `93b559a`: Added blank line cleaning to `modify_doc_text`
  - (pending): Added blank line cleaning to `batch_edit_doc`

