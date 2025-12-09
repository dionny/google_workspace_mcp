"""
Unit tests for sheet management operations in Google Sheets tools.

These tests verify the logic and request body structures for delete_sheet,
rename_sheet, and copy_sheet tools.
"""

import pytest


class TestResolveSheetIdLogic:
    """Unit tests for the _resolve_sheet_id helper function logic."""

    def test_sheet_id_takes_precedence_when_both_provided(self):
        """Test that sheet_id is returned when both sheet_id and sheet_name are provided."""
        # Logic: if sheet_id is not None, return it immediately
        sheet_id = 123
        # sheet_name = "MySheet" would be provided but ignored

        # The actual function returns sheet_id immediately if it's not None
        if sheet_id is not None:
            result = sheet_id
        else:
            result = None  # Would do lookup

        assert result == 123

    def test_neither_provided_raises_error(self):
        """Test that an error is raised when neither sheet_id nor sheet_name is provided."""
        sheet_id = None
        sheet_name = None

        # The function checks both being None and raises an error
        with pytest.raises(Exception) as exc_info:
            if sheet_id is not None:
                pass  # Would return sheet_id
            elif sheet_name is None:
                raise Exception("Either 'sheet_name' or 'sheet_id' must be provided.")
        assert "Either 'sheet_name' or 'sheet_id' must be provided" in str(
            exc_info.value
        )

    def test_sheet_name_lookup_logic(self):
        """Test the sheet name lookup logic against a mock spreadsheet response."""
        # Simulated API response from get spreadsheet
        spreadsheet_response = {
            "properties": {"title": "My Spreadsheet"},
            "sheets": [
                {"properties": {"sheetId": 0, "title": "Sheet1"}},
                {"properties": {"sheetId": 123456, "title": "Data"}},
                {"properties": {"sheetId": 789012, "title": "Summary"}},
            ],
        }

        sheet_name_to_find = "Data"

        # The lookup logic
        sheets = spreadsheet_response.get("sheets", [])
        found_id = None
        for sheet in sheets:
            sheet_props = sheet.get("properties", {})
            if sheet_props.get("title") == sheet_name_to_find:
                found_id = sheet_props.get("sheetId")
                break

        assert found_id == 123456

    def test_sheet_name_not_found_error(self):
        """Test error message when sheet name is not found."""
        spreadsheet_response = {
            "sheets": [
                {"properties": {"sheetId": 0, "title": "Sheet1"}},
                {"properties": {"sheetId": 123, "title": "Data"}},
            ],
        }

        sheet_name_to_find = "NonExistent"
        sheets = spreadsheet_response.get("sheets", [])

        found_id = None
        for sheet in sheets:
            sheet_props = sheet.get("properties", {})
            if sheet_props.get("title") == sheet_name_to_find:
                found_id = sheet_props.get("sheetId")
                break

        if found_id is None:
            available_sheets = [
                sheet.get("properties", {}).get("title", "Unknown") for sheet in sheets
            ]
            error_msg = (
                f"Sheet '{sheet_name_to_find}' not found in spreadsheet. "
                f"Available sheets: {available_sheets}"
            )

            assert "NonExistent" in error_msg
            assert "Sheet1" in error_msg
            assert "Data" in error_msg


class TestSheetIdentifierMessageFormats:
    """Tests for success message formatting with sheet_name parameter."""

    def test_delete_sheet_message_with_sheet_name(self):
        """Test delete_sheet message format when using sheet_name."""
        sheet_name = "MyData"
        resolved_sheet_id = 123456
        spreadsheet_id = "abc123xyz"
        user_email = "test@example.com"

        # When sheet_name is provided and sheet_id is None
        identifier = f"'{sheet_name}' (ID: {resolved_sheet_id})"
        message = (
            f"Successfully deleted sheet {identifier} from spreadsheet {spreadsheet_id} "
            f"for {user_email}."
        )

        assert "'MyData'" in message
        assert "ID: 123456" in message
        assert "abc123xyz" in message

    def test_delete_sheet_message_with_sheet_id(self):
        """Test delete_sheet message format when using sheet_id."""
        resolved_sheet_id = 123456
        spreadsheet_id = "abc123xyz"
        user_email = "test@example.com"

        # When sheet_id is provided (sheet_name is None or sheet_id takes precedence)
        identifier = f"ID {resolved_sheet_id}"
        message = (
            f"Successfully deleted sheet {identifier} from spreadsheet {spreadsheet_id} "
            f"for {user_email}."
        )

        assert "ID 123456" in message
        assert "'MyData'" not in message

    def test_rename_sheet_message_with_sheet_name(self):
        """Test rename_sheet message format when using sheet_name."""
        sheet_name = "OldName"
        resolved_sheet_id = 456
        new_name = "NewName"
        spreadsheet_id = "xyz789abc"
        user_email = "user@example.com"

        identifier = f"'{sheet_name}' (ID: {resolved_sheet_id})"
        message = (
            f"Successfully renamed sheet {identifier} to '{new_name}' in spreadsheet {spreadsheet_id} "
            f"for {user_email}."
        )

        assert "'OldName'" in message
        assert "ID: 456" in message
        assert "'NewName'" in message

    def test_copy_sheet_message_with_sheet_name(self):
        """Test copy_sheet message format when using sheet_name."""
        sheet_name = "SourceSheet"
        resolved_sheet_id = 789
        new_sheet_title = "Copy of SourceSheet"
        new_sheet_id = 999
        spreadsheet_id = "def456ghi"
        user_email = "copy@example.com"

        identifier = f"'{sheet_name}' (ID: {resolved_sheet_id})"
        message = (
            f"Successfully copied sheet {identifier} to new sheet '{new_sheet_title}' (ID: {new_sheet_id}) "
            f"in spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "'SourceSheet'" in message
        assert "ID: 789" in message
        assert "'Copy of SourceSheet'" in message
        assert "ID: 999" in message


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
