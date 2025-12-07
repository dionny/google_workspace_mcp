# Local Testing Guide for Google Workspace MCP

## Overview

This guide shows you multiple ways to test your MCP tools locally without the overhead of running the full MCP protocol. Pick the approach that best fits your needs.

## Testing Approaches (Ordered by Simplicity)

### 1. **Simple Direct Testing** ⭐ **RECOMMENDED FOR QUICK ITERATION**

**File:** `examples/simple_direct_test.py`

**Best for:** Quick testing during active development

**How it works:** Import and call your tool functions directly as regular Python async functions.

```python
# examples/simple_direct_test.py
from gdocs.docs_tools import search_docs

result = await search_docs(
    user_google_email="user@example.com",
    query="test",
    page_size=5
)
print(result)
```

**Usage:**
```bash
python examples/simple_direct_test.py
```

**Pros:**
- ✅ Fastest to run
- ✅ Minimal setup
- ✅ Easy to debug with pdb
- ✅ Perfect for TDD

**Cons:**
- ❌ Manual - need to edit the file for each test

---

### 2. **Quick Test Script** (Jupyter-style)

**File:** `examples/quick_test.py`

**Best for:** Iterative testing with frequent parameter changes

**How it works:** A script with multiple test cases you can uncomment/modify

```python
# Uncomment the test you want to run
# result = await search_docs(user_google_email=USER_EMAIL, query="test")
# print(result)
```

**Usage:**
```bash
python examples/quick_test.py
```

**Pros:**
- ✅ Fast iteration
- ✅ Easy to switch between tests
- ✅ Good for exploring APIs

**Cons:**
- ❌ Need to uncomment/comment code

---

### 3. **Test Harness (CLI)**

**File:** `test_harness.py`

**Best for:** Testing any tool with different parameters without editing code

**Usage:**
```bash
# List all tools
python test_harness.py --list

# Get info about a tool
python test_harness.py --info get_doc_content

# Call a tool
python test_harness.py --tool search_docs \
  --query "test" \
  --user_google_email "user@example.com" \
  --page_size 5

# Interactive mode
python test_harness.py --interactive
```

**Pros:**
- ✅ No code editing required
- ✅ Can test any tool
- ✅ Interactive REPL mode
- ✅ Great for exploratory testing

**Cons:**
- ❌ Slightly slower startup (loads all modules)
- ❌ More complex implementation

---

### 4. **PyTest Unit Tests**

**Files:** 
- `examples/test_docs_with_pytest.py` (example)
- `gdocs/test_*.py` (existing tests)

**Best for:** Automated testing, CI/CD, regression tests

**Usage:**
```bash
# Run all tests
pytest

# Run specific test file
pytest examples/test_docs_with_pytest.py -v

# Run specific test
pytest examples/test_docs_with_pytest.py::test_search_docs_returns_results -v
```

**Pros:**
- ✅ Automated
- ✅ Great for CI/CD
- ✅ Can use fixtures and mocks
- ✅ Test discovery

**Cons:**
- ❌ More boilerplate
- ❌ Requires pytest knowledge

---

## Setup

### 1. Environment Variables

Create/update your `.env` file:

```bash
# Required for authentication
USER_GOOGLE_EMAIL=your-email@example.com
GOOGLE_OAUTH_CLIENT_ID=your-client-id
GOOGLE_OAUTH_CLIENT_SECRET=your-client-secret

# Optional: For testing specific documents
TEST_DOC_ID=your-document-id-here

# Optional: Control logging
LOG_LEVEL=WARNING  # or DEBUG, INFO
```

### 2. Dependencies

Make sure you have dependencies installed:

```bash
uv sync
```

## Common Workflows

### Testing a New Feature

1. **Write your code** in the tool module (e.g., `gdocs/docs_tools.py`)

2. **Quick test** using simple direct test:
   ```bash
   # Edit examples/simple_direct_test.py to add your test
   python examples/simple_direct_test.py
   ```

3. **Iterate** until it works

4. **Write proper tests** using pytest:
   ```python
   # Add to examples/test_docs_with_pytest.py or create new file
   @pytest.mark.asyncio
   async def test_my_new_feature():
       result = await my_new_function(...)
       assert "expected" in result
   ```

5. **Run all tests** to ensure nothing broke:
   ```bash
   pytest
   ```

### Debugging an Issue

1. **Reproduce** with direct test:
   ```python
   # examples/simple_direct_test.py
   async def test_bug():
       # Add minimal reproduction case
       result = await problematic_function(...)
   ```

2. **Debug** with pdb:
   ```python
   import pdb; pdb.set_trace()
   result = await problematic_function(...)
   ```

3. **Or use VS Code debugger** - set breakpoints directly in your tool code

### Exploring the API

Use the test harness in interactive mode:

```bash
python test_harness.py --interactive

> list
> info search_docs
> call search_docs
  query: test
  page_size: 3
```

## Authentication

The first time you run a test that needs Google API access:

1. It will open a browser for OAuth authentication
2. Complete the OAuth flow
3. Credentials are saved locally
4. Subsequent runs use the saved credentials

For single-user testing, you can set:
```bash
MCP_SINGLE_USER_MODE=1
```

## Tips & Tricks

### 1. Reduce Logging Noise

```bash
export LOG_LEVEL=WARNING
# or in .env: LOG_LEVEL=WARNING
```

### 2. Test with Mock Data

```python
# Test helper functions without API calls
from gdocs.docs_structure import parse_document_structure

mock_doc = {
    'title': 'Test',
    'body': {'content': [...]}
}
result = parse_document_structure(mock_doc)
```

### 3. Performance Testing

```python
import time

async def perf_test():
    start = time.time()
    for i in range(10):
        await my_function(...)
    print(f"Average: {(time.time() - start) / 10:.2f}s")
```

### 4. Integration Testing

```python
async def test_workflow():
    # Test multiple tools together
    doc_id = await create_doc(...)
    await modify_doc_text(document_id=doc_id, ...)
    content = await get_doc_content(document_id=doc_id)
    assert "expected" in content
    await delete_doc(doc_id)  # cleanup
```

### 5. Use pytest fixtures for common setup

```python
@pytest.fixture
def user_email():
    return os.getenv('USER_GOOGLE_EMAIL')

@pytest.mark.asyncio
async def test_with_fixture(user_email):
    result = await my_function(user_google_email=user_email)
    ...
```

## Troubleshooting

### "Authentication failed"

Make sure your `.env` has valid credentials:
```bash
GOOGLE_OAUTH_CLIENT_ID=...
GOOGLE_OAUTH_CLIENT_SECRET=...
USER_GOOGLE_EMAIL=...
```

### "Module not found"

Run from the project root:
```bash
cd /path/to/google_workspace_mcp
python examples/simple_direct_test.py
```

### "Tool not found" (test harness)

The test harness loads tools dynamically. Use `--verbose` to see what's happening:
```bash
python test_harness.py --list --verbose
```

## Next Steps

1. **Start simple**: Use `examples/simple_direct_test.py` for your first test
2. **Iterate quickly**: Keep it open and modify as you develop
3. **Add structure**: Move to pytest when you have stable tests
4. **Automate**: Run pytest in CI/CD

## Examples

All example files are in the `examples/` directory:
- `simple_direct_test.py` - Simplest approach
- `quick_test.py` - Jupyter-style testing
- `test_docs_tools.py` - Example with specific scenarios
- `test_docs_with_pytest.py` - PyTest examples

Start with `simple_direct_test.py` and customize it for your needs!


