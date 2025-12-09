"""
Integration tests for Google Docs tools.

These tests use real Google Docs API calls to verify end-to-end functionality.
Each test creates a fresh document and trashes it after completion.

To run these tests:
    export GOOGLE_TEST_EMAIL="your-email@gmail.com"
    pytest tests/gdocs/integration/ -v

Note: These tests are slower than unit tests but provide better coverage of
actual API behavior, including edge cases that mocks might miss.
"""
