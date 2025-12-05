#!/usr/bin/env python3
"""Script to add improvement issues to beads."""
import json
import hashlib
from datetime import datetime, timezone

def generate_hash(content):
    return hashlib.sha256(content.encode()).hexdigest()

def make_issue(id_suffix, title, description, priority, issue_type, labels):
    issue_id = f"google_workspace_mcp-66a1.{id_suffix}"
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
        "labels": labels,
        "dependencies": [{
            "issue_id": issue_id,
            "depends_on_id": "google_workspace_mcp-66a1",
            "type": "parent-child",
            "created_at": now,
            "created_by": "mbradshaw"
        }]
    }

issues = [
    make_issue(18, 
        "Add location convenience parameter to modify_doc_text",
        """## Problem
The most common editing operation - 'add text to the end of document' - requires two API calls:
1. Call `inspect_doc_structure` to get `total_length`
2. Call `modify_doc_text` with that index

This adds latency and complexity for the simplest operation.

## Solution
Add a `location` parameter to `modify_doc_text` that accepts semantic location shortcuts:

```python
# Current (2 calls):
structure = inspect_doc_structure(doc_id)
modify_doc_text(doc_id, start_index=structure['total_length'], text='New content')

# Proposed (1 call):
modify_doc_text(doc_id, location='end', text='New content')
modify_doc_text(doc_id, location='start', text='Prepend this')
```

## Supported Values
- `end` - append to document end (most common)
- `start` - insert at beginning (after first section break at index 1)

## Implementation
- Get document internally to calculate total_length
- Map location to appropriate start_index
- Maintain backward compatibility with explicit start_index

## Acceptance Criteria
- [ ] `location='end'` works without specifying start_index
- [ ] `location='start'` inserts at beginning
- [ ] Explicit start_index still works (backward compatible)
- [ ] Clear error if both location and start_index provided

## Files to Modify
- gdocs/docs_tools.py - modify_doc_text function

## Estimated Effort
4-6 hours""",
        1, "feature", ["enhancement", "google-docs", "dx", "quick-win"]),

    make_issue(19,
        "Auto-index for table creation (make index optional)",
        """## Problem
Creating a table requires calling `inspect_doc_structure` first to get a safe index. The tool documentation even says "YOU MUST CALL inspect_doc_structure FIRST" in caps. This is 3+ round trips minimum for a simple table.

## Solution
Make `index` parameter optional with smart defaults:

```python
# Current (requires pre-flight):
structure = inspect_doc_structure(doc_id)
create_table_with_data(doc_id, data, index=structure['total_length'])

# Proposed (auto-detect):
create_table_with_data(doc_id, data)  # Appends to end by default
create_table_with_data(doc_id, data, after_heading="Data Section")  # Or specify location
```

## Implementation
- If index not provided, fetch document and calculate safe insertion point
- Support `after_heading` parameter for heading-based positioning
- Maintain backward compatibility when index is explicitly provided

## Acceptance Criteria
- [ ] `create_table_with_data` works without `index` parameter
- [ ] Default behavior appends table to document end
- [ ] `after_heading` parameter positions table after specified heading
- [ ] Explicit `index` parameter still works (backward compatible)

## Files to Modify
- gdocs/docs_tools.py - create_table_with_data function
- gdocs/managers/table_operation_manager.py

## Estimated Effort
4-6 hours""",
        1, "feature", ["enhancement", "google-docs", "tables", "quick-win"]),

    make_issue(20,
        "Consolidate batch_update_doc and batch_modify_doc into single tool",
        """## Problem
There are two nearly-identical batch editing tools with confusing differences:
- `batch_update_doc` - uses operation types like `insert_text`, `delete_text`, `format_text`
- `batch_modify_doc` - uses types like `insert`, `delete`, `format` plus search support

Users don't know which to use, and the operation type names are inconsistent.

## Solution
Merge into a single `batch_edit_doc` tool that:
1. Supports BOTH operation type naming conventions (accept aliases)
2. Combines all features from both tools
3. Deprecates the old tools (keep as aliases initially)

```python
# Single unified tool:
batch_edit_doc(doc_id, operations=[
    {"type": "insert", "location": "end", "text": "..."},  # New style
    {"type": "insert_text", "index": 100, "text": "..."},  # Legacy style still works
    {"type": "replace", "search": "old", "text": "new"},   # Search-based
])
```

## Operation Type Mapping
| Legacy (batch_update_doc) | Modern (batch_modify_doc) | Unified |
|---------------------------|--------------------------|---------|
| insert_text | insert | insert (preferred) |
| delete_text | delete | delete (preferred) |
| replace_text | replace | replace (preferred) |
| format_text | format | format (preferred) |

## Acceptance Criteria
- [ ] Single `batch_edit_doc` tool accepts both naming conventions
- [ ] Search-based positioning works (from batch_modify_doc)
- [ ] Auto position adjustment works (from batch_modify_doc)
- [ ] Legacy tools remain as aliases with deprecation warning
- [ ] Documentation updated to recommend `batch_edit_doc`

## Files to Modify
- gdocs/docs_tools.py - merge tools, add unified tool
- gdocs/managers/batch_operation_manager.py

## Estimated Effort
6-8 hours""",
        2, "feature", ["enhancement", "google-docs", "consolidation"]),

    make_issue(21,
        "Merge inspect_doc_structure and get_doc_structure into single tool",
        """## Problem
Two tools provide overlapping document structure information:
- `inspect_doc_structure` - focused on safe insertion points, table info
- `get_doc_structure` - focused on hierarchical element view, headings

Users often need both types of info and must call two separate tools.

## Solution
Merge into a single `get_doc_info` tool with a detail level parameter:

```python
# Quick stats and safe indices (for table creation):
get_doc_info(doc_id, detail="summary")

# Full element hierarchy (for navigation):
get_doc_info(doc_id, detail="structure")

# Table-focused view:
get_doc_info(doc_id, detail="tables")

# Everything (current default):
get_doc_info(doc_id)  # Returns all info
```

## Detail Level Options
- `summary` - total_length, element counts, safe insertion indices (fast)
- `structure` - full hierarchical element view with positions
- `tables` - table-focused: dimensions, positions, cell data
- `headings` - just headings outline for navigation
- `all` (default) - everything combined

## Acceptance Criteria
- [ ] Single `get_doc_info` tool with `detail` parameter
- [ ] `detail='summary'` returns quick stats for table creation
- [ ] `detail='structure'` returns full element hierarchy
- [ ] Legacy tools remain as aliases
- [ ] Response format consistent across detail levels

## Files to Modify
- gdocs/docs_tools.py - merge tools
- gdocs/docs_structure.py - may need refactoring

## Estimated Effort
4-6 hours""",
        2, "feature", ["enhancement", "google-docs", "consolidation"]),

    make_issue(22,
        "Clarify distinction between find_and_replace_doc and modify_doc_text replace",
        """## Problem
Two ways to replace text with overlapping functionality:
- `find_and_replace_doc` - replaces ALL occurrences globally
- `modify_doc_text` with `search` + `position='replace'` - replaces ONE occurrence

Users pick the wrong tool, leading to unexpected results (replacing all when they wanted one, or vice versa).

## Solution
Make the distinction clearer through:

1. **Better naming** (consider renaming):
   - `find_and_replace_doc` → `replace_all_in_doc` (clearer it's global)
   - Or keep but add prominent documentation

2. **Clearer documentation** with comparison table:
   | Use Case | Tool |
   |----------|------|
   | Replace ALL occurrences | `find_and_replace_doc` |
   | Replace SPECIFIC occurrence | `modify_doc_text` with search+position |
   | Replace by index range | `modify_doc_text` with start/end_index |

3. **Warning in responses** when similar tool might be better:
   - If `find_and_replace_doc` replaces only 1 occurrence, suggest `modify_doc_text` for single replacements
   - If `modify_doc_text` replace is called multiple times for same text, suggest `find_and_replace_doc`

## Acceptance Criteria
- [ ] Tool documentation clearly explains when to use each
- [ ] Consider renaming `find_and_replace_doc` to `replace_all_in_doc`
- [ ] Response includes count of replacements made
- [ ] Add hint in response when other tool might be more appropriate

## Files to Modify
- gdocs/docs_tools.py - update docstrings, consider rename
- README or documentation

## Estimated Effort
2-4 hours""",
        3, "task", ["documentation", "google-docs", "dx"]),

    make_issue(23,
        "Add operation aliases to normalize batch operation types",
        """## Problem
batch_update_doc and batch_modify_doc use different operation type names:
- `insert_text` vs `insert`
- `delete_text` vs `delete`  
- `format_text` vs `format`
- `replace_text` vs `replace`

This is confusing and error-prone.

## Quick Fix (Before Full Consolidation)
Add alias mapping so BOTH naming conventions work in BOTH tools:

```python
OPERATION_ALIASES = {
    'insert_text': 'insert',
    'delete_text': 'delete',
    'format_text': 'format',
    'replace_text': 'replace',
    # Reverse aliases
    'insert': 'insert',
    'delete': 'delete',
    'format': 'format',
    'replace': 'replace',
}

def normalize_operation_type(op_type):
    return OPERATION_ALIASES.get(op_type, op_type)
```

## Acceptance Criteria
- [ ] Both tools accept both naming conventions
- [ ] `insert_text` and `insert` are equivalent
- [ ] Documentation updated to show preferred (shorter) names
- [ ] No breaking changes to existing integrations

## Files to Modify
- gdocs/managers/batch_operation_manager.py

## Estimated Effort
1-2 hours (quick fix)

## Related
This is a quick win toward the larger consolidation effort (issue .20)""",
        1, "task", ["enhancement", "google-docs", "quick-win"])
]

if __name__ == "__main__":
    # Append to issues.jsonl
    with open('.beads/issues.jsonl', 'a') as f:
        for issue in issues:
            f.write(json.dumps(issue) + '\n')

    print(f"Added {len(issues)} issues:")
    for issue in issues:
        print(f"  - {issue['id']}: {issue['title']}")

