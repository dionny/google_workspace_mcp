# Google Docs Integration Tests

Welcome to the Google Docs integration test suite! These tests verify that the Google Docs MCP tools work correctly with real API calls.

## Quick Start

```bash
# Set your test email
export GOOGLE_TEST_EMAIL="your@email.com"

# Run all integration tests
uv run pytest tests/gdocs/integration/ -v

# Run with coverage
uv run pytest tests/gdocs/integration/ --cov=gdocs --cov=gdrive --cov-report=html
open htmlcov/index.html
```

## ⏱️ Test Duration

**~24 seconds** for 5 tests (~5 seconds per test)

## What These Tests Do

### ✅ Real API Calls
- Creates actual Google Docs in your account
- Makes real modifications (text, lists, formatting)
- Uses real authentication

### ✅ Automatic Cleanup
- Each test gets a fresh document
- Documents prefixed with `[TEST]` for easy identification
- **Always trashed after test** (pass or fail!)

### ✅ Test Isolation
- No state pollution between tests
- Each test runs independently
- Parallel execution safe

## Current Tests

1. **test_insert_text_at_start** - Insert text at document start
2. **test_create_simple_bullet_list** - Create bullet list
3. **test_create_numbered_list** - Create numbered list
4. **test_insert_text_at_end** - Insert text at document end
5. **test_apply_bold_formatting** - Apply bold formatting

## Coverage Stats

Run with coverage to see which code paths are tested:

```bash
uv run pytest tests/gdocs/integration/ \
    --cov=gdocs \
    --cov=gdrive \
    --cov-report=term-missing \
    --cov-report=html
```

**Current Coverage**: 17.04% overall
- See `INTEGRATION_COMPLETE.md` for detailed coverage breakdown
- HTML report in `htmlcov/index.html` shows line-by-line coverage

## Adding New Tests

```python
import pytest
import json
import gdocs.docs_tools as docs_tools_module

@pytest.mark.asyncio
@pytest.mark.integration
async def test_your_feature(user_google_email, test_document):
    """Test your feature description."""
    doc_id = test_document['document_id']
    
    # Call your tool
    result = await docs_tools_module.your_tool.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        # your parameters
    )
    
    # Verify result
    result_data = json.loads(result)
    assert result_data.get('success'), f"Should succeed: {result}"
    
    # Document auto-trashed by fixture!
```

## Fixtures Available

- `user_google_email` - Test email from environment
- `docs_service` - Authenticated Docs API service
- `drive_service` - Authenticated Drive API service  
- `test_document` - Fresh document (auto-created, auto-trashed)
- `populated_test_document` - Document with test content

## Tips

### Run Specific Test
```bash
uv run pytest tests/gdocs/integration/test_basic_operations.py::test_insert_text_at_start -v
```

### Run Only Integration Tests
```bash
uv run pytest -m integration -v
```

### Run in Parallel (faster)
```bash
uv pip install pytest-xdist
uv run pytest tests/gdocs/integration/ -n auto
```

### Debug Failed Test
```bash
# Show full output
uv run pytest tests/gdocs/integration/ -v -s

# Stop on first failure
uv run pytest tests/gdocs/integration/ -x
```

## Authentication

First run will open browser for OAuth2 authentication. Credentials cached in:
```
~/.google_workspace_mcp/credentials/
```

## Viewing Test Documents

While tests run, documents appear in your Google Drive with `[TEST]` prefix. They're automatically moved to trash after the test completes.

To keep test documents for debugging, modify the fixture in `conftest.py`.

## See Also

- `INTEGRATION_COMPLETE.md` - Full documentation with coverage details
- `INTEGRATION_SUCCESS.md` - Framework implementation notes
- `conftest.py` - Fixture definitions
