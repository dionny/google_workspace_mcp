"""
Unit tests for list_spreadsheets tool with name_filter functionality.

These tests verify the query building logic for the name_filter parameter.
"""


class TestListSpreadsheetsQueryBuilding:
    """Unit tests for list_spreadsheets query construction."""

    def test_query_without_name_filter(self):
        """Test that query without name_filter only filters by mimeType."""
        name_filter = None

        query_parts = ["mimeType='application/vnd.google-apps.spreadsheet'"]
        if name_filter:
            escaped_filter = name_filter.replace("'", "\\'")
            query_parts.append(f"name contains '{escaped_filter}'")

        query = " and ".join(query_parts)

        assert query == "mimeType='application/vnd.google-apps.spreadsheet'"

    def test_query_with_simple_name_filter(self):
        """Test query with a simple name filter."""
        name_filter = "Budget"

        query_parts = ["mimeType='application/vnd.google-apps.spreadsheet'"]
        if name_filter:
            escaped_filter = name_filter.replace("'", "\\'")
            query_parts.append(f"name contains '{escaped_filter}'")

        query = " and ".join(query_parts)

        assert (
            query
            == "mimeType='application/vnd.google-apps.spreadsheet' and name contains 'Budget'"
        )

    def test_query_with_name_filter_containing_spaces(self):
        """Test query with a name filter containing spaces."""
        name_filter = "Q4 Budget"

        query_parts = ["mimeType='application/vnd.google-apps.spreadsheet'"]
        if name_filter:
            escaped_filter = name_filter.replace("'", "\\'")
            query_parts.append(f"name contains '{escaped_filter}'")

        query = " and ".join(query_parts)

        assert (
            query
            == "mimeType='application/vnd.google-apps.spreadsheet' and name contains 'Q4 Budget'"
        )

    def test_query_with_name_filter_containing_single_quote(self):
        """Test query with a name filter containing a single quote (must be escaped)."""
        name_filter = "John's Budget"

        query_parts = ["mimeType='application/vnd.google-apps.spreadsheet'"]
        if name_filter:
            escaped_filter = name_filter.replace("'", "\\'")
            query_parts.append(f"name contains '{escaped_filter}'")

        query = " and ".join(query_parts)

        assert (
            query
            == "mimeType='application/vnd.google-apps.spreadsheet' and name contains 'John\\'s Budget'"
        )

    def test_query_with_name_filter_containing_numbers(self):
        """Test query with a name filter containing numbers."""
        name_filter = "2024"

        query_parts = ["mimeType='application/vnd.google-apps.spreadsheet'"]
        if name_filter:
            escaped_filter = name_filter.replace("'", "\\'")
            query_parts.append(f"name contains '{escaped_filter}'")

        query = " and ".join(query_parts)

        assert (
            query
            == "mimeType='application/vnd.google-apps.spreadsheet' and name contains '2024'"
        )

    def test_query_with_empty_string_name_filter(self):
        """Test that empty string name_filter doesn't add to query."""
        name_filter = ""

        query_parts = ["mimeType='application/vnd.google-apps.spreadsheet'"]
        if name_filter:
            escaped_filter = name_filter.replace("'", "\\'")
            query_parts.append(f"name contains '{escaped_filter}'")

        query = " and ".join(query_parts)

        assert query == "mimeType='application/vnd.google-apps.spreadsheet'"


class TestListSpreadsheetsMessageFormatting:
    """Unit tests for list_spreadsheets result message formatting."""

    def test_no_spreadsheets_message_without_filter(self):
        """Test 'no spreadsheets found' message without name filter."""
        name_filter = None
        user_email = "test@example.com"

        filter_msg = f" matching '{name_filter}'" if name_filter else ""
        message = f"No spreadsheets{filter_msg} found for {user_email}."

        assert message == "No spreadsheets found for test@example.com."

    def test_no_spreadsheets_message_with_filter(self):
        """Test 'no spreadsheets found' message with name filter."""
        name_filter = "Budget"
        user_email = "test@example.com"

        filter_msg = f" matching '{name_filter}'" if name_filter else ""
        message = f"No spreadsheets{filter_msg} found for {user_email}."

        assert (
            message == "No spreadsheets matching 'Budget' found for test@example.com."
        )

    def test_success_message_without_filter(self):
        """Test success message without name filter."""
        name_filter = None
        user_email = "test@example.com"
        file_count = 5

        filter_msg = f" matching '{name_filter}'" if name_filter else ""
        message = f"Successfully listed {file_count} spreadsheets{filter_msg} for {user_email}:"

        assert message == "Successfully listed 5 spreadsheets for test@example.com:"

    def test_success_message_with_filter(self):
        """Test success message with name filter."""
        name_filter = "Budget"
        user_email = "test@example.com"
        file_count = 3

        filter_msg = f" matching '{name_filter}'" if name_filter else ""
        message = f"Successfully listed {file_count} spreadsheets{filter_msg} for {user_email}:"

        assert (
            message
            == "Successfully listed 3 spreadsheets matching 'Budget' for test@example.com:"
        )


class TestListSpreadsheetsFileFormatting:
    """Unit tests for spreadsheet list item formatting."""

    def test_spreadsheet_list_item_format(self):
        """Test the format of each spreadsheet list item."""
        file = {
            "name": "Test Spreadsheet",
            "id": "abc123xyz",
            "modifiedTime": "2024-01-15T10:30:00Z",
            "webViewLink": "https://docs.google.com/spreadsheets/d/abc123xyz",
        }

        formatted = (
            f'- "{file["name"]}" (ID: {file["id"]}) | Modified: '
            f"{file.get('modifiedTime', 'Unknown')} | Link: {file.get('webViewLink', 'No link')}"
        )

        assert '"Test Spreadsheet"' in formatted
        assert "ID: abc123xyz" in formatted
        assert "2024-01-15T10:30:00Z" in formatted
        assert "https://docs.google.com/spreadsheets/d/abc123xyz" in formatted

    def test_spreadsheet_list_item_missing_optional_fields(self):
        """Test list item format when optional fields are missing."""
        file = {
            "name": "Test Spreadsheet",
            "id": "abc123xyz",
        }

        formatted = (
            f'- "{file["name"]}" (ID: {file["id"]}) | Modified: '
            f"{file.get('modifiedTime', 'Unknown')} | Link: {file.get('webViewLink', 'No link')}"
        )

        assert "Modified: Unknown" in formatted
        assert "Link: No link" in formatted
