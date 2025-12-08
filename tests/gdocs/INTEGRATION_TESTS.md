# Google Docs Integration Test Framework

## ✅ Complete! 

I've converted the gdocs tests from unit tests with mocks to integration tests that use real Google Docs API calls.

## What's Been Created

### 1. Test Infrastructure (`tests/gdocs/conftest.py`)
- **`test_document` fixture**: Creates a fresh `[TEST]` document for each test
- **Automatic cleanup**: Documents are **always** trashed after tests (pass or fail)
- **Service fixtures**: Provides authenticated Docs and Drive services
- **Helper functions**: `get_document()` and `batch_update()` for common operations

### 2. Integration Test Suites

#### `tests/gdocs/integration/test_list_operations.py`
Tests for list operations:
- ✅ Creating bullet and numbered lists
- ✅ Inserting items at start/end of lists
- ✅ Nested list items with indentation
- ✅ Mixed list types in same document
- ✅ Edge cases (empty items, etc.)

#### `tests/gdocs/integration/test_table_operations.py`
Tests for table operations:
- ✅ Creating tables with specified dimensions
- ✅ Inserting/deleting rows and columns
- ✅ Deleting entire tables
- ✅ Merging cells
- ✅ Formatting cells
- ✅ Multiple tables in one document

### 3. Configuration

#### `pyproject.toml`
Added pytest configuration:
- Test markers (unit, integration, asyncio)
- Asyncio auto-mode
- Coverage settings
- Output formatting

#### `tests/gdocs/integration/README.md`
Documentation for running tests and understanding the structure.

#### `run_integration_tests.sh`
Convenience script for running tests.

## How to Run

```bash
# 1. Set your test email
export GOOGLE_TEST_EMAIL="your-email@gmail.com"

# 2. Run all integration tests
pytest tests/gdocs/integration/ -v

# Or use the convenience script
./tests/gdocs/run_integration_tests.sh

# 3. Run specific tests
pytest tests/gdocs/integration/test_list_operations.py -v
pytest tests/gdocs/integration/test_table_operations.py::test_create_basic_table -v
```

## Key Features

### ✅ Test Document Naming
All test docs are prefixed with `[TEST]` so you can easily identify them in Drive:
- `[TEST] GDocs Integration Test abc12345`
- `[TEST] Populated Test Doc xyz67890`

### ✅ Guaranteed Cleanup
Every test document is trashed after completion, regardless of:
- Test passing or failing
- Exceptions raised
- Early test termination

The cleanup is in the fixture's `finally` block using pytest's yield pattern.

### ✅ Test Isolation
Each test gets a **fresh document**:
- No state pollution between tests
- Tests can run in any order
- Parallel test execution supported (future)

### ✅ Real API Behavior
Tests catch issues that mocks miss:
- Style inheritance bugs
- Index calculation after mutations
- List nesting edge cases
- Table cell merging quirks
- Multi-tab document handling

## Test Document Lifecycle

```
1. Fixture creates document with [TEST] prefix
2. Test receives document_id
3. Test performs operations
4. Test assertions verify results
5. Fixture trashes document (always runs)
6. Document auto-deletes from trash after 30 days
```

## Why Integration Tests?

Your codebase is perfect for integration tests because:

1. **Complex API**: Google Docs has subtle behaviors hard to mock
2. **State Management**: Document indices shift after operations
3. **Bug Discovery**: You've found bugs mocks didn't catch
4. **Real Behavior**: What matters for an MCP server is actual API behavior
5. **Test Surface**: Not huge - can afford slower tests

## Next Steps

You can now:
1. Run these tests to verify your current implementation
2. Add more test cases as you find edge cases
3. Convert other unit tests to integration tests as needed
4. Keep unit tests for pure logic functions (validators, parsers, etc.)

## Notes

- Tests require network connection and Google API access
- Documents are trashed, not permanently deleted (safer)
- You can manually empty trash to clean up immediately
- Test email must have Docs API access enabled

