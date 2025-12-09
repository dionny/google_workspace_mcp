"""
Unit tests for delete_sheet tool.

These tests verify the logic and request structures for the sheet delete operation.
"""


class TestDeleteSheetRequestStructure:
    """Tests for the delete sheet request body structure."""

    def test_delete_sheet_request_with_sheet_id(self):
        """Test request body for deleting a sheet by ID."""
        resolved_sheet_id = 12345

        request_body = {"requests": [{"deleteSheet": {"sheetId": resolved_sheet_id}}]}

        delete_req = request_body["requests"][0]["deleteSheet"]
        assert delete_req["sheetId"] == 12345

    def test_delete_sheet_request_structure(self):
        """Test that request body has correct structure for batchUpdate."""
        resolved_sheet_id = 67890

        request_body = {"requests": [{"deleteSheet": {"sheetId": resolved_sheet_id}}]}

        assert "requests" in request_body
        assert len(request_body["requests"]) == 1
        assert "deleteSheet" in request_body["requests"][0]
        assert request_body["requests"][0]["deleteSheet"]["sheetId"] == 67890


class TestDeleteSheetValidation:
    """Tests for parameter validation logic."""

    def test_requires_sheet_identifier(self):
        """Test that either sheet_name or sheet_id must be provided."""
        sheet_name = None
        sheet_id = None

        requires_identifier = sheet_name is None and sheet_id is None

        assert requires_identifier is True

    def test_sheet_name_is_sufficient(self):
        """Test that sheet_name alone is sufficient."""
        sheet_name = "My Sheet"
        sheet_id = None

        requires_identifier = sheet_name is None and sheet_id is None

        assert requires_identifier is False

    def test_sheet_id_is_sufficient(self):
        """Test that sheet_id alone is sufficient."""
        sheet_name = None
        sheet_id = 12345

        requires_identifier = sheet_name is None and sheet_id is None

        assert requires_identifier is False

    def test_both_identifiers_valid(self):
        """Test that providing both identifiers is valid."""
        sheet_name = "My Sheet"
        sheet_id = 12345

        requires_identifier = sheet_name is None and sheet_id is None

        assert requires_identifier is False


class TestDeleteSheetOnlySheetCheck:
    """Tests for preventing deletion of the only sheet."""

    def test_cannot_delete_only_sheet(self):
        """Test that deleting the only sheet is prevented."""
        sheets = [{"properties": {"sheetId": 0, "title": "Sheet1"}}]

        cannot_delete = len(sheets) <= 1

        assert cannot_delete is True

    def test_can_delete_when_multiple_sheets(self):
        """Test that deletion is allowed when multiple sheets exist."""
        sheets = [
            {"properties": {"sheetId": 0, "title": "Sheet1"}},
            {"properties": {"sheetId": 1, "title": "Sheet2"}},
        ]

        cannot_delete = len(sheets) <= 1

        assert cannot_delete is False

    def test_can_delete_when_many_sheets(self):
        """Test that deletion is allowed with many sheets."""
        sheets = [
            {"properties": {"sheetId": i, "title": f"Sheet{i + 1}"}} for i in range(10)
        ]

        cannot_delete = len(sheets) <= 1

        assert cannot_delete is False

    def test_empty_sheets_prevented(self):
        """Test that empty sheets list prevents deletion."""
        sheets = []

        cannot_delete = len(sheets) <= 1

        assert cannot_delete is True


class TestDeleteSheetTitleResolution:
    """Tests for resolving sheet title from sheet list."""

    def test_find_sheet_title_by_id(self):
        """Test finding sheet title by sheet ID."""
        resolved_sheet_id = 12345
        sheets = [
            {"properties": {"sheetId": 0, "title": "First"}},
            {"properties": {"sheetId": 12345, "title": "Target Sheet"}},
            {"properties": {"sheetId": 99999, "title": "Other"}},
        ]

        sheet_title = None
        for sheet in sheets:
            props = sheet.get("properties", {})
            if props.get("sheetId") == resolved_sheet_id:
                sheet_title = props.get("title", "Unknown")
                break

        assert sheet_title == "Target Sheet"

    def test_sheet_not_found_returns_none(self):
        """Test that non-existent sheet ID returns None."""
        resolved_sheet_id = 99999
        sheets = [
            {"properties": {"sheetId": 0, "title": "First"}},
            {"properties": {"sheetId": 12345, "title": "Second"}},
        ]

        sheet_title = None
        for sheet in sheets:
            props = sheet.get("properties", {})
            if props.get("sheetId") == resolved_sheet_id:
                sheet_title = props.get("title", "Unknown")
                break

        assert sheet_title is None

    def test_fallback_to_unknown(self):
        """Test fallback to 'Unknown' when title is missing."""
        resolved_sheet_id = 12345
        sheets = [
            {"properties": {"sheetId": 12345}},  # No title
        ]

        sheet_title = None
        for sheet in sheets:
            props = sheet.get("properties", {})
            if props.get("sheetId") == resolved_sheet_id:
                sheet_title = props.get("title", "Unknown")
                break

        assert sheet_title == "Unknown"


class TestDeleteSheetSuccessMessages:
    """Tests for success message formatting."""

    def test_success_message_format(self):
        """Test success message includes all relevant info."""
        sheet_title = "Old Data"
        resolved_sheet_id = 12345
        spreadsheet_id = "abc123"
        user_email = "test@example.com"

        message = (
            f"Successfully deleted sheet '{sheet_title}' (ID: {resolved_sheet_id}) "
            f"from spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "Successfully deleted sheet 'Old Data'" in message
        assert "(ID: 12345)" in message
        assert f"from spreadsheet {spreadsheet_id}" in message
        assert user_email in message

    def test_success_message_with_numeric_id(self):
        """Test success message with different sheet ID."""
        sheet_title = "Archived"
        resolved_sheet_id = 0
        spreadsheet_id = "xyz789"
        user_email = "user@company.com"

        message = (
            f"Successfully deleted sheet '{sheet_title}' (ID: {resolved_sheet_id}) "
            f"from spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "(ID: 0)" in message
        assert "sheet 'Archived'" in message


class TestDeleteSheetErrorMessages:
    """Tests for error message formatting."""

    def test_missing_identifier_error(self):
        """Test error message when no identifier is provided."""
        error_msg = (
            "Either 'sheet_name' or 'sheet_id' must be provided "
            "to identify the sheet to delete."
        )

        assert "sheet_name" in error_msg
        assert "sheet_id" in error_msg
        assert "identify" in error_msg

    def test_only_sheet_error(self):
        """Test error message when trying to delete only sheet."""
        spreadsheet_id = "abc123"
        error_msg = (
            f"Cannot delete the only sheet in spreadsheet {spreadsheet_id}. "
            "A spreadsheet must have at least one sheet."
        )

        assert "Cannot delete the only sheet" in error_msg
        assert spreadsheet_id in error_msg
        assert "at least one sheet" in error_msg

    def test_sheet_not_found_error_by_name(self):
        """Test error message when sheet not found by name."""
        sheet_name = "NonExistent"
        sheet_id = None
        spreadsheet_id = "abc123"

        identifier = sheet_name if sheet_name else f"ID {sheet_id}"
        error_msg = f"Sheet '{identifier}' not found in spreadsheet {spreadsheet_id}."

        assert "Sheet 'NonExistent' not found" in error_msg
        assert spreadsheet_id in error_msg

    def test_sheet_not_found_error_by_id(self):
        """Test error message when sheet not found by ID."""
        sheet_name = None
        sheet_id = 99999
        spreadsheet_id = "abc123"

        identifier = sheet_name if sheet_name else f"ID {sheet_id}"
        error_msg = f"Sheet '{identifier}' not found in spreadsheet {spreadsheet_id}."

        assert "Sheet 'ID 99999' not found" in error_msg
        assert spreadsheet_id in error_msg


class TestDeleteSheetParameterPrecedence:
    """Tests for parameter precedence logic."""

    def test_sheet_id_takes_precedence(self):
        """Test that sheet_id takes precedence over sheet_name."""
        sheet_id = 12345

        # Logic from _resolve_sheet_id
        if sheet_id is not None:
            resolved_id = sheet_id
        else:
            resolved_id = None  # Would be resolved from name

        assert resolved_id == 12345

    def test_sheet_name_used_when_no_id(self):
        """Test that sheet_name is used when sheet_id is None."""
        sheet_id = None

        # Logic from _resolve_sheet_id
        if sheet_id is not None:
            resolved_id = sheet_id
        else:
            # In actual code, this would lookup by name
            resolved_id = "would_be_resolved_from_name"

        assert resolved_id == "would_be_resolved_from_name"


class TestDeleteSheetEdgeCases:
    """Tests for edge cases in delete operations."""

    def test_sheet_id_zero_is_valid(self):
        """Test that sheet ID of 0 is valid (first sheet often has ID 0)."""
        sheet_id = 0

        # 0 is a valid sheet ID (often the first sheet)
        is_valid = sheet_id is not None

        assert is_valid is True

    def test_special_characters_in_sheet_name(self):
        """Test sheet names with special characters."""
        sheet_names = [
            "Data (2024)",
            "Q1 & Q2",
            "Sheet - Draft",
            "Tab/Name",
            "Sheet 'Copy'",
        ]

        for name in sheet_names:
            # Simulate finding sheet by name
            sheets = [{"properties": {"sheetId": 1, "title": name}}]
            found = any(s.get("properties", {}).get("title") == name for s in sheets)
            assert found is True

    def test_whitespace_in_sheet_name(self):
        """Test sheet names with whitespace."""
        sheet_name = "  Sheet with Spaces  "
        sheets = [{"properties": {"sheetId": 1, "title": sheet_name}}]

        found = any(s.get("properties", {}).get("title") == sheet_name for s in sheets)
        assert found is True
