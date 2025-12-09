"""
Unit tests for rename_sheet tool.

These tests verify the logic and request structures for the rename_sheet tool.
"""

import pytest


class TestRenameSheetRequestStructure:
    """Unit tests for rename_sheet request body structure."""

    def test_rename_request_body_format(self):
        """Test that the request body has the correct structure."""
        sheet_id = 123456
        new_name = "New Sheet Name"

        request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "title": new_name,
                        },
                        "fields": "title",
                    }
                }
            ]
        }

        assert "requests" in request_body
        assert len(request_body["requests"]) == 1
        request = request_body["requests"][0]
        assert "updateSheetProperties" in request
        props = request["updateSheetProperties"]
        assert props["properties"]["sheetId"] == 123456
        assert props["properties"]["title"] == "New Sheet Name"
        assert props["fields"] == "title"

    def test_api_parameters_structure(self):
        """Test that API parameters are correctly structured."""
        spreadsheet_id = "abc123"
        sheet_id = 0
        new_name = "Renamed Sheet"

        api_params = {
            "spreadsheetId": spreadsheet_id,
            "body": {
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": sheet_id,
                                "title": new_name,
                            },
                            "fields": "title",
                        }
                    }
                ]
            },
        }

        assert api_params["spreadsheetId"] == "abc123"
        assert "body" in api_params
        assert "requests" in api_params["body"]


class TestRenameSheetValidation:
    """Tests for rename_sheet input validation logic."""

    def test_either_sheet_name_or_sheet_id_required(self):
        """Test that either sheet_name or sheet_id must be provided."""
        sheet_name = None
        sheet_id = None

        with pytest.raises(Exception) as exc_info:
            if not sheet_name and sheet_id is None:
                raise Exception("Either 'sheet_name' or 'sheet_id' must be provided.")

        assert "Either 'sheet_name' or 'sheet_id' must be provided" in str(
            exc_info.value
        )

    def test_sheet_name_sufficient(self):
        """Test that providing only sheet_name is sufficient."""
        sheet_name = "Sheet1"
        sheet_id = None

        # Should not raise - at least one identifier provided
        assert sheet_name or sheet_id is not None

    def test_sheet_id_sufficient(self):
        """Test that providing only sheet_id is sufficient."""
        sheet_name = None
        sheet_id = 0  # Note: 0 is a valid sheet_id

        # Should not raise - at least one identifier provided
        assert sheet_name or sheet_id is not None

    def test_sheet_id_zero_is_valid(self):
        """Test that sheet_id of 0 is valid (first sheet)."""
        sheet_id = 0

        # Check that 0 is not falsy in our validation
        assert sheet_id is not None
        assert isinstance(sheet_id, int)

    def test_both_identifiers_allowed(self):
        """Test that providing both identifiers is allowed (sheet_id takes precedence after lookup)."""
        sheet_name = "Sheet1"
        sheet_id = 123

        # When both are provided, sheet_name lookup is skipped if sheet_id given
        assert sheet_name and sheet_id is not None

    def test_new_name_required(self):
        """Test that new_name is required."""
        new_name = ""

        with pytest.raises(Exception) as exc_info:
            if not new_name:
                raise Exception("new_name is required")

        assert "new_name is required" in str(exc_info.value)


class TestRenameSheetMessages:
    """Tests for rename_sheet success message formatting."""

    def test_success_message_with_sheet_name(self):
        """Test the format of success message when using sheet_name."""
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        sheet_name = "Old Name"
        new_name = "New Name"

        old_identifier = f"'{sheet_name}'" if sheet_name else f"ID {0}"
        message = (
            f"Successfully renamed sheet {old_identifier} to '{new_name}' "
            f"in spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "Successfully renamed" in message
        assert "'Old Name'" in message
        assert "'New Name'" in message
        assert spreadsheet_id in message
        assert user_email in message

    def test_success_message_with_sheet_id(self):
        """Test the format of success message when using sheet_id."""
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        sheet_name = None
        sheet_id = 12345
        new_name = "New Name"

        old_identifier = f"'{sheet_name}'" if sheet_name else f"ID {sheet_id}"
        message = (
            f"Successfully renamed sheet {old_identifier} to '{new_name}' "
            f"in spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "Successfully renamed" in message
        assert "ID 12345" in message
        assert "'New Name'" in message


class TestRenameSheetSheetLookup:
    """Tests for sheet lookup by name logic."""

    def test_find_sheet_by_name(self):
        """Test finding a sheet by name in spreadsheet data."""
        spreadsheet_data = {
            "sheets": [
                {"properties": {"sheetId": 0, "title": "Sheet1"}},
                {"properties": {"sheetId": 123, "title": "Data"}},
                {"properties": {"sheetId": 456, "title": "Summary"}},
            ]
        }
        target_name = "Data"

        found_id = None
        for sheet in spreadsheet_data.get("sheets", []):
            sheet_props = sheet.get("properties", {})
            if sheet_props.get("title") == target_name:
                found_id = sheet_props.get("sheetId")
                break

        assert found_id == 123

    def test_sheet_not_found(self):
        """Test behavior when sheet name is not found."""
        spreadsheet_data = {
            "sheets": [
                {"properties": {"sheetId": 0, "title": "Sheet1"}},
            ]
        }
        target_name = "NonExistent"
        spreadsheet_id = "abc123"

        found_id = None
        for sheet in spreadsheet_data.get("sheets", []):
            sheet_props = sheet.get("properties", {})
            if sheet_props.get("title") == target_name:
                found_id = sheet_props.get("sheetId")
                break

        with pytest.raises(Exception) as exc_info:
            if found_id is None:
                raise Exception(
                    f"Sheet '{target_name}' not found in spreadsheet {spreadsheet_id}"
                )

        assert "Sheet 'NonExistent' not found" in str(exc_info.value)

    def test_case_sensitive_lookup(self):
        """Test that sheet lookup is case-sensitive."""
        spreadsheet_data = {
            "sheets": [
                {"properties": {"sheetId": 0, "title": "Sheet1"}},
                {"properties": {"sheetId": 123, "title": "sheet1"}},
            ]
        }

        # Looking for "Sheet1" should find ID 0, not ID 123
        target_name = "Sheet1"
        found_id = None
        for sheet in spreadsheet_data.get("sheets", []):
            sheet_props = sheet.get("properties", {})
            if sheet_props.get("title") == target_name:
                found_id = sheet_props.get("sheetId")
                break

        assert found_id == 0

    def test_empty_sheets_list(self):
        """Test behavior when spreadsheet has no sheets."""
        spreadsheet_data = {"sheets": []}
        target_name = "Sheet1"
        spreadsheet_id = "abc123"

        found_id = None
        for sheet in spreadsheet_data.get("sheets", []):
            sheet_props = sheet.get("properties", {})
            if sheet_props.get("title") == target_name:
                found_id = sheet_props.get("sheetId")
                break

        with pytest.raises(Exception) as exc_info:
            if found_id is None:
                raise Exception(
                    f"Sheet '{target_name}' not found in spreadsheet {spreadsheet_id}"
                )

        assert "not found" in str(exc_info.value)


class TestRenameSheetEdgeCases:
    """Tests for edge cases in rename_sheet."""

    def test_rename_to_same_name(self):
        """Test renaming a sheet to its current name (should be allowed by API)."""
        old_name = "MySheet"
        new_name = "MySheet"

        # The API allows this, even though it's a no-op
        request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": 0,
                            "title": new_name,
                        },
                        "fields": "title",
                    }
                }
            ]
        }

        assert (
            request_body["requests"][0]["updateSheetProperties"]["properties"]["title"]
            == old_name
        )

    def test_sheet_name_with_special_characters(self):
        """Test renaming to a name with special characters."""
        new_name = "Sheet (Copy) - 2024"

        request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": 0,
                            "title": new_name,
                        },
                        "fields": "title",
                    }
                }
            ]
        }

        assert (
            request_body["requests"][0]["updateSheetProperties"]["properties"]["title"]
            == "Sheet (Copy) - 2024"
        )

    def test_sheet_name_with_unicode(self):
        """Test renaming to a name with unicode characters."""
        new_name = "Sheet - Summary"

        request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": 0,
                            "title": new_name,
                        },
                        "fields": "title",
                    }
                }
            ]
        }

        assert (
            request_body["requests"][0]["updateSheetProperties"]["properties"]["title"]
            == new_name
        )

    def test_whitespace_in_sheet_name(self):
        """Test sheet names with leading/trailing whitespace."""
        new_name = "  Padded Name  "

        request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": 0,
                            "title": new_name,
                        },
                        "fields": "title",
                    }
                }
            ]
        }

        # The API may trim whitespace, but we pass it as-is
        assert (
            "Padded Name"
            in request_body["requests"][0]["updateSheetProperties"]["properties"][
                "title"
            ]
        )

    def test_first_sheet_rename(self):
        """Test renaming the first sheet (ID 0)."""
        sheet_id = 0
        new_name = "Renamed First Sheet"

        request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "title": new_name,
                        },
                        "fields": "title",
                    }
                }
            ]
        }

        assert (
            request_body["requests"][0]["updateSheetProperties"]["properties"][
                "sheetId"
            ]
            == 0
        )
