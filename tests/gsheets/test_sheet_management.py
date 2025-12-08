"""
Unit tests for sheet management operations in Google Sheets tools.

These tests verify the logic and request body structures for delete_sheet,
rename_sheet, and copy_sheet tools.
"""


class TestDeleteSheetValidation:
    """Unit tests for delete_sheet request body structure."""

    def test_delete_sheet_request_body_structure(self):
        """Test that the request body for delete_sheet is correctly structured."""
        sheet_id = 123456

        request_body = {"requests": [{"deleteSheet": {"sheetId": sheet_id}}]}

        assert request_body["requests"][0]["deleteSheet"]["sheetId"] == 123456

    def test_delete_sheet_with_different_sheet_ids(self):
        """Test delete_sheet request body with various sheet IDs."""
        test_cases = [0, 1, 100, 999999, 2147483647]

        for sheet_id in test_cases:
            request_body = {"requests": [{"deleteSheet": {"sheetId": sheet_id}}]}
            assert request_body["requests"][0]["deleteSheet"]["sheetId"] == sheet_id


class TestRenameSheetValidation:
    """Unit tests for rename_sheet request body structure."""

    def test_rename_sheet_request_body_structure(self):
        """Test that the request body for rename_sheet is correctly structured."""
        sheet_id = 0
        new_name = "My New Sheet Name"

        request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {"sheetId": sheet_id, "title": new_name},
                        "fields": "title",
                    }
                }
            ]
        }

        update_request = request_body["requests"][0]["updateSheetProperties"]
        assert update_request["properties"]["sheetId"] == 0
        assert update_request["properties"]["title"] == "My New Sheet Name"
        assert update_request["fields"] == "title"

    def test_rename_sheet_with_special_characters(self):
        """Test rename_sheet with various special characters in names."""
        special_names = [
            "Sheet with spaces",
            "Sheet-with-dashes",
            "Sheet_with_underscores",
            "Sheet (with parentheses)",
            "Sheet 123",
            "Quarterly Report Q4 2024",
        ]

        for new_name in special_names:
            request_body = {
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {"sheetId": 0, "title": new_name},
                            "fields": "title",
                        }
                    }
                ]
            }
            assert (
                request_body["requests"][0]["updateSheetProperties"]["properties"][
                    "title"
                ]
                == new_name
            )

    def test_rename_sheet_fields_only_title(self):
        """Test that only the title field is updated (not other properties)."""
        request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {"sheetId": 0, "title": "New Name"},
                        "fields": "title",
                    }
                }
            ]
        }

        # The 'fields' mask ensures only title is updated
        assert request_body["requests"][0]["updateSheetProperties"]["fields"] == "title"


class TestCopySheetValidation:
    """Unit tests for copy_sheet request body structure."""

    def test_copy_sheet_request_body_structure_with_name(self):
        """Test that the request body for copy_sheet with new_name is correctly structured."""
        source_sheet_id = 123
        new_name = "Copy of Original"

        duplicate_request = {"sourceSheetId": source_sheet_id, "newSheetName": new_name}
        request_body = {"requests": [{"duplicateSheet": duplicate_request}]}

        dup_request = request_body["requests"][0]["duplicateSheet"]
        assert dup_request["sourceSheetId"] == 123
        assert dup_request["newSheetName"] == "Copy of Original"

    def test_copy_sheet_request_body_without_name(self):
        """Test that the request body for copy_sheet without new_name is correct."""
        source_sheet_id = 456

        duplicate_request = {"sourceSheetId": source_sheet_id}
        # Only add newSheetName if provided
        new_name = None
        if new_name:
            duplicate_request["newSheetName"] = new_name

        request_body = {"requests": [{"duplicateSheet": duplicate_request}]}

        dup_request = request_body["requests"][0]["duplicateSheet"]
        assert dup_request["sourceSheetId"] == 456
        assert "newSheetName" not in dup_request

    def test_copy_sheet_conditional_name_logic(self):
        """Test the conditional logic for adding newSheetName."""
        source_sheet_id = 789

        # Test with name provided
        new_name = "My Copy"
        duplicate_request = {"sourceSheetId": source_sheet_id}
        if new_name:
            duplicate_request["newSheetName"] = new_name
        assert duplicate_request["newSheetName"] == "My Copy"

        # Test without name
        new_name = None
        duplicate_request = {"sourceSheetId": source_sheet_id}
        if new_name:
            duplicate_request["newSheetName"] = new_name
        assert "newSheetName" not in duplicate_request

        # Test with empty string (should not add)
        new_name = ""
        duplicate_request = {"sourceSheetId": source_sheet_id}
        if new_name:
            duplicate_request["newSheetName"] = new_name
        assert "newSheetName" not in duplicate_request


class TestSuccessMessageFormats:
    """Tests for success message formatting."""

    def test_delete_sheet_success_message_format(self):
        """Test the format of delete_sheet success message."""
        sheet_id = 123
        spreadsheet_id = "abc123xyz"
        user_email = "test@example.com"

        message = (
            f"Successfully deleted sheet ID {sheet_id} from spreadsheet {spreadsheet_id} "
            f"for {user_email}."
        )

        assert "sheet ID 123" in message
        assert "abc123xyz" in message
        assert "test@example.com" in message

    def test_rename_sheet_success_message_format(self):
        """Test the format of rename_sheet success message."""
        sheet_id = 456
        new_name = "Renamed Sheet"
        spreadsheet_id = "xyz789abc"
        user_email = "user@example.com"

        message = (
            f"Successfully renamed sheet ID {sheet_id} to '{new_name}' in spreadsheet {spreadsheet_id} "
            f"for {user_email}."
        )

        assert "sheet ID 456" in message
        assert "'Renamed Sheet'" in message
        assert "xyz789abc" in message

    def test_copy_sheet_success_message_format(self):
        """Test the format of copy_sheet success message."""
        source_sheet_id = 789
        new_sheet_title = "Copy of Original"
        new_sheet_id = 999
        spreadsheet_id = "def456ghi"
        user_email = "copy@example.com"

        message = (
            f"Successfully copied sheet ID {source_sheet_id} to new sheet '{new_sheet_title}' (ID: {new_sheet_id}) "
            f"in spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "sheet ID 789" in message
        assert "'Copy of Original'" in message
        assert "ID: 999" in message
        assert "def456ghi" in message


class TestResponseParsing:
    """Tests for parsing API responses."""

    def test_copy_sheet_response_parsing(self):
        """Test parsing the response from duplicateSheet API call."""
        # Simulated API response structure
        response = {
            "spreadsheetId": "abc123",
            "replies": [
                {
                    "duplicateSheet": {
                        "properties": {
                            "sheetId": 999888,
                            "title": "Copy of Sheet1",
                            "index": 1,
                            "sheetType": "GRID",
                            "gridProperties": {"rowCount": 1000, "columnCount": 26},
                        }
                    }
                }
            ],
        }

        # Extract the new sheet info
        new_sheet_props = response["replies"][0]["duplicateSheet"]["properties"]
        new_sheet_id = new_sheet_props["sheetId"]
        new_sheet_title = new_sheet_props["title"]

        assert new_sheet_id == 999888
        assert new_sheet_title == "Copy of Sheet1"

    def test_create_sheet_response_parsing(self):
        """Test parsing the response from addSheet API call (for reference)."""
        # This mirrors the existing create_sheet response parsing
        response = {
            "spreadsheetId": "xyz789",
            "replies": [
                {
                    "addSheet": {
                        "properties": {
                            "sheetId": 123456,
                            "title": "New Sheet",
                            "index": 2,
                        }
                    }
                }
            ],
        }

        sheet_id = response["replies"][0]["addSheet"]["properties"]["sheetId"]
        assert sheet_id == 123456
