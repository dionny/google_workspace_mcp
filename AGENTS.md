# Agent Guide for Google Workspace MCP

## Repository Information

**IMPORTANT**: This is a fork of [taylorwilsdon/google_workspace_mcp](https://github.com/taylorwilsdon/google_workspace_mcp) maintained by Indeed (dionny/google_workspace_mcp).

### Fork Status
- We are generally doing our own thing and developing features for Indeed's specific needs
- **We do NOT intend to push all changes upstream** to the original repository
- Some features may be Indeed-specific and not appropriate for the upstream project
- **By default, PRs should target `dionny/google_workspace_mcp:main`**, NOT `taylorwilsdon/google_workspace_mcp`

### Creating Pull Requests
```bash
# CORRECT: PR against our fork
gh pr create --repo dionny/google_workspace_mcp --base main --head <branch-name>

# INCORRECT: Do not PR against upstream by default
# gh pr create (without --repo flag may target upstream)
```

### When to Contribute Upstream
If you believe a feature or fix would benefit the broader community:
1. Discuss with the team first
2. Create a separate branch if needed
3. Ensure the feature is generic and well-documented
4. Submit PR to `taylorwilsdon/google_workspace_mcp` only after team approval

## Testing Tools

### Using tools_cli.py
Test individual MCP tools without running the server:

```bash
python tools_cli.py --tool modify_doc_text --document_id <doc_id> --user_google_email <email> --text "Hello" --location end
```

Use `--list` to see all available tools.

### Authentication
- Uses OAuth2 flow with credential store in `~/.cache/google_workspace_mcp/`
- Email parameter is required: `--user_google_email your@email.com`
- First run will trigger browser authentication

## Project Structure

```
gdocs/          - Google Docs tools (largest module)
  managers/     - Complex operation handlers (batch, tables, history, validation)
  docs_tools.py - Main tool definitions (~11K lines)
gcalendar/      - Calendar tools
gmail/          - Gmail tools
gdrive/         - Drive tools
gsheets/        - Sheets tools
gslides/        - Slides tools
gtasks/         - Tasks tools
gchat/          - Chat tools
gforms/         - Forms tools
gsearch/        - Custom Search tools
auth/           - OAuth and authentication
core/           - Server, config, utils
adrs/           - Architecture Decision Records
```

## Common Issues

### 1. Google Docs Text Operations
- **Index 0 is invalid** - Start from index 1 (first char after section break)
- **Blank lines create empty list items** - Use `convert_to_list` parameter, auto-cleaning handles this
- **Text inherits formatting** - Insertion after bold text inherits bold; clearing happens automatically

### 2. Batch Operations
- **Use batch_edit_doc** for multiple operations - more efficient than individual calls
- **Search-based positioning** - Prefer `search + position` over manual index calculation
- **Auto-adjustment** - Positions auto-adjust for sequential operations in batch
- **Table modifications** - Use `modify_table`, NOT batch_edit_doc (batch can only INSERT tables)

### 3. List Operations
- **List types** - Use "ORDERED"/"UNORDERED" or "numbered"/"bullet"
- **Blank line cleaning** - Automatically applied when `convert_to_list` is used
- **insert_list_item** - For adding single items to existing lists
- **append_to_list** - Specialized for appending to end of list

### 4. Error Messages
- **DocsErrorBuilder** - Structured error responses with codes like `MISSING_REQUIRED_PARAM`
- **ValidationManager** - Parameter validation with helpful error messages
- Check error `code` field for programmatic handling

## Design Patterns

### Tool Parameters
- **Multiple positioning modes** - location, index, search, heading, range
- **Preview mode** - Add `preview=True` to see what would change
- **Tab support** - Multi-tab documents need `tab_id` parameter

### Undo System
- **In-memory history** - Operations tracked per document
- **undo_doc_operation** - Undo last operation on document
- **Batch operations** - Single undo for entire batch via `batch_id`

## Code Style

### Testing
- Write tests in `tests/gdocs/` for gdocs features
- Use pytest: `pytest tests/gdocs/test_*.py`
- Integration tests go in `tests/gdocs/integration/`

### Commits
- Format: `module: Brief description`
- Reference beads issues in commit body
- Use ADRs for significant design decisions

### Beads Issue Tracker
```bash
bd create --title "..." --description "..." --type feature --labels gdocs --priority 2
bd update <id> --status closed --notes "..."
bd list --status open
```

## Documentation

- **ADRs** - Major decisions go in `adrs/` directory
- **Tool docstrings** - Comprehensive with examples in each tool function
- **README** - High-level overview and setup
- **PROMPT.md** - Detailed tool documentation for LLMs

## Quick Start for New Agents

1. Read `README.md` for project overview
2. Check `adrs/` for architectural context
3. Use `tools_cli.py --list` to explore available tools
4. Test changes with `tools_cli.py --tool <name> ...`
5. Run relevant tests before committing
6. Create ADR for significant design decisions
7. File beads issues for bugs/features

## Gotchas

- **Don't commit test files** from integration/ if another agent is working on them
- **pyproject.toml/uv.lock** changes are often from dependency updates, commit separately
- **Linting** - Always check and fix linter errors before committing
- **Timeout commands** - Use `timeout` for potentially long-running operations
- **Module imports** - Import from `gdocs.docs_helpers` for shared utilities

