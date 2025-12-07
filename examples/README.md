# Testing Examples

This directory contains various testing approaches for the Google Workspace MCP server.

## Quick Start

### Absolute Simplest (Recommended First)

```bash
python examples/simple_direct_test.py
```

Edit the file to add your own tests - it's self-contained and easy to understand.

## Available Examples

### 1. `simple_direct_test.py` ⭐ **START HERE**

The simplest possible approach - import tools and call them directly.

**Perfect for:**
- Quick iteration during development
- Learning how the tools work
- Debugging specific issues

**Usage:**
```bash
python examples/simple_direct_test.py
```

### 2. `quick_test.py`

Jupyter-notebook style testing script with multiple test cases you can uncomment/modify.

**Perfect for:**
- Testing multiple scenarios
- Comparing different parameter combinations
- Exploring API behavior

**Usage:**
```bash
# Edit to uncomment the tests you want
python examples/quick_test.py
```

### 3. `test_docs_tools.py`

Structured examples showing specific testing scenarios for Google Docs tools.

**Perfect for:**
- Learning testing patterns
- Understanding async testing
- See examples of error handling

**Usage:**
```bash
python examples/test_docs_tools.py
```

### 4. `test_docs_with_pytest.py`

PyTest-based tests showing proper unit testing approach.

**Perfect for:**
- Automated testing
- CI/CD pipelines
- Regression testing

**Usage:**
```bash
pytest examples/test_docs_with_pytest.py -v
```

## Configuration

All examples read from your `.env` file:

```bash
# Required
USER_GOOGLE_EMAIL=your-email@example.com
GOOGLE_OAUTH_CLIENT_ID=your-client-id
GOOGLE_OAUTH_CLIENT_SECRET=your-client-secret

# Optional
TEST_DOC_ID=your-document-id-here
LOG_LEVEL=WARNING
```

## Testing Your Changes

The recommended workflow:

1. **Edit your tool code** (e.g., `gdocs/docs_tools.py`)

2. **Add a test** to `simple_direct_test.py`:
   ```python
   async def test_my_feature():
       from gdocs.docs_tools import my_new_function
       result = await my_new_function(...)
       print(result)
   ```

3. **Run it**:
   ```bash
   python examples/simple_direct_test.py
   ```

4. **Iterate** until it works

5. **Write proper tests** for CI/CD using pytest

## More Information

See the full testing guide: [`docs/TESTING_GUIDE.md`](../docs/TESTING_GUIDE.md)

## Tips

### Enable Debug Logging

```bash
export LOG_LEVEL=DEBUG
python examples/simple_direct_test.py
```

### Test with Debugger

```python
import pdb

async def test_my_feature():
    pdb.set_trace()  # Breakpoint here
    result = await my_function(...)
```

### Run from VS Code

1. Open example file
2. Set breakpoints
3. Run with debugger
4. Step through your code

## Common Issues

**"No output"**: Make sure you're running from project root
**"Auth failed"**: Check your `.env` file has valid credentials
**"Import error"**: Make sure dependencies are installed (`uv sync`)

## Contributing

When adding new features, please add example tests to this directory to help others understand how to use your code!



