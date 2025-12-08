#!/bin/bash
# Quick script to run integration tests for Google Docs

set -e

# Check for test email
if [ -z "$GOOGLE_TEST_EMAIL" ]; then
    echo "❌ Error: GOOGLE_TEST_EMAIL environment variable not set"
    echo ""
    echo "Usage:"
    echo "  export GOOGLE_TEST_EMAIL='your-email@gmail.com'"
    echo "  ./run_integration_tests.sh"
    exit 1
fi

echo "🧪 Running Google Docs Integration Tests"
echo "📧 Test email: $GOOGLE_TEST_EMAIL"
echo ""

# Install test dependencies if needed
if ! python -c "import pytest" 2>/dev/null; then
    echo "Installing pytest..."
    pip install pytest pytest-asyncio
fi

# Run the tests
echo "Running integration tests..."
pytest tests/gdocs/integration/ \
    -v \
    --tb=short \
    -m "not slow" \
    "$@"

echo ""
echo "✅ Tests complete!"
echo ""
echo "Note: Test documents are prefixed with [TEST] and moved to trash automatically."
echo "To clean up: Empty your Google Drive trash folder."

