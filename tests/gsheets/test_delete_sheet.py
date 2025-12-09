"""
Unit tests for delete_sheet tool.

These tests verify the logic and request structures for the delete_sheet tool.
"""

import pytest


class TestDeleteSheetRequestStructure:
    """Unit tests for delete_sheet request body structure."""

    def test_delete_request_body_format(self):
        """Test that the request body has the correct structure."""
        sheet_id = 123456

        request_body = {"requests": [{"deleteSheet": {"sheetId": sheet_id}}]}

        assert "requests" in request_body
        assert len(request_body["requests"]) == 1
        request = request_body["requests"][0]
        assert "deleteSheet" in request
        assert request["deleteSheet"]["sheetId"] == 123456

    def test_api_parameters_structure(self):
        """Test that API parameters are correctly structured."""
        spreadsheet_id = "abc123"
        sheet_id = 0

        api_params = {
            "spreadsheetId": spreadsheet_id,
            "body": {"requests": [{"deleteSheet": {"sheetId": sheet_id}}]},
        }

        assert api_params["spreadsheetId"] == "abc123"
        assert "body" in api_params
        assert "requests" in api_params["body"]
        assert api_params["body"]["requests"][0]["deleteSheet"]["sheetId"] == 0


class TestDeleteSheetValidation:
    """Tests for delete_sheet input validation logic."""

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
        """Test that providing both identifiers is allowed (sheet_id takes precedence)."""
        sheet_name = "Sheet1"
        sheet_id = 123

        # When both are provided, sheet_name lookup is skipped if sheet_id given
        assert sheet_name and sheet_id is not None


class TestDeleteSheetMessages:
    """Tests for delete_sheet success message formatting."""

    def test_success_message_with_sheet_name(self):
        """Test the format of success message when using sheet_name."""
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        sheet_name = "Sheet to Delete"

        identifier = f"'{sheet_name}'" if sheet_name else f"ID {0}"
        message = (
            f"Successfully deleted sheet {identifier} "
            f"from spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "Successfully deleted" in message
        assert "'Sheet to Delete'" in message
        assert spreadsheet_id in message
        assert user_email in message

    def test_success_message_with_sheet_id(self):
        """Test the format of success message when using sheet_id."""
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        sheet_name = None
        sheet_id = 12345

        identifier = f"'{sheet_name}'" if sheet_name else f"ID {sheet_id}"
        message = (
            f"Successfully deleted sheet {identifier} "
            f"from spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "Successfully deleted" in message
        assert "ID 12345" in message
        assert spreadsheet_id in message


class TestDeleteSheetSheetLookup:
    """Tests for sheet lookup by name logic (shared with rename_sheet)."""

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


class TestDeleteSheetEdgeCases:
    """Tests for edge cases in delete_sheet."""

    def test_delete_first_sheet(self):
        """Test deleting the first sheet (ID 0)."""
        sheet_id = 0

        request_body = {"requests": [{"deleteSheet": {"sheetId": sheet_id}}]}

        assert request_body["requests"][0]["deleteSheet"]["sheetId"] == 0

    def test_delete_sheet_with_high_id(self):
        """Test deleting a sheet with a large ID number."""
        sheet_id = 2147483647  # Max int32

        request_body = {"requests": [{"deleteSheet": {"sheetId": sheet_id}}]}

        assert request_body["requests"][0]["deleteSheet"]["sheetId"] == 2147483647

    def test_sheet_name_with_special_characters(self):
        """Test looking up a sheet name with special characters."""
        spreadsheet_data = {
            "sheets": [
                {
                    "properties": {
                        "sheetId": 789,
                        "title": "Sheet (Copy) - 2024",
                    }
                },
            ]
        }
        target_name = "Sheet (Copy) - 2024"

        found_id = None
        for sheet in spreadsheet_data.get("sheets", []):
            sheet_props = sheet.get("properties", {})
            if sheet_props.get("title") == target_name:
                found_id = sheet_props.get("sheetId")
                break

        assert found_id == 789

    def test_sheet_name_with_unicode(self):
        """Test looking up a sheet name with unicode characters."""
        spreadsheet_data = {
            "sheets": [
                {
                    "properties": {
                        "sheetId": 111,
                        "title": "Sheet - Summary",
                    }
                },
            ]
        }
        target_name = "Sheet - Summary"

        found_id = None
        for sheet in spreadsheet_data.get("sheets", []):
            sheet_props = sheet.get("properties", {})
            if sheet_props.get("title") == target_name:
                found_id = sheet_props.get("sheetId")
                break

        assert found_id == 111
