# Testing Your MCP Tools Locally

You can test your MCP tools directly without running the full MCP server using the tools CLI.

## Quick Start

### List all available tools (87 tools)
```bash
uv run python tools_cli.py --list
```

### Get info about a specific tool
```bash
uv run python tools_cli.py --info search_docs
```

### Call a tool directly
```bash
uv run python tools_cli.py --tool search_docs \
  --query "test" \
  --user_google_email "your@email.com" \
  --page_size 5

# Boolean and integer values are auto-converted:
uv run python tools_cli.py --tool modify_doc_text \
  --document_id "your-doc-id" \
  --start_index 1 \
  --text "Hello" \
  --bold true \
  --font_size 14
```

### Interactive mode (REPL)
```bash
uv run python tools_cli.py --interactive

# Then use commands:
> list                    # List all tools
> info search_docs        # Get tool details
> call search_docs        # Call a tool (prompts for parameters)
> quit                    # Exit
```

## Configuration

Make sure your `.env` file has valid credentials:

```bash
USER_GOOGLE_EMAIL=your-email@example.com
GOOGLE_OAUTH_CLIENT_ID=your-client-id
GOOGLE_OAUTH_CLIENT_SECRET=your-client-secret
```

## Important: Use `uv run`

Always use `uv run python` to ensure the correct environment:

```bash
# ✅ Correct
uv run python tools_cli.py --list

# ❌ Wrong (will fail with import errors)
python tools_cli.py --list
```

## Tips

- Use `--verbose` flag for debug logging: `uv run python tools_cli.py --verbose --tool ...`
- The first time you call a tool, it will open a browser for OAuth authentication
- Subsequent calls use saved credentials

## How It Works

The tools CLI:
1. Imports all your tool modules
2. Registers them with the FastMCP server
3. Provides direct access to tool functions
4. Handles authentication automatically

This means you get:
- ✅ Fast iteration - no MCP protocol overhead
- ✅ Easy debugging - direct Python stack traces  
- ✅ Real authentication - uses your actual Google credentials
- ✅ All 87 tools available to test
