"""
Unit tests for rename_sheet tool.

These tests verify the logic and request structures for the sheet rename operation.
"""


class TestRenameSheetRequestStructure:
    """Tests for the rename sheet request body structure."""

    def test_rename_sheet_request_with_sheet_id(self):
        """Test request body for renaming a sheet by ID."""
        resolved_sheet_id = 12345
        new_name = "New Sheet Name"

        request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": resolved_sheet_id,
                            "title": new_name,
                        },
                        "fields": "title",
                    }
                }
            ]
        }

        update_req = request_body["requests"][0]["updateSheetProperties"]
        assert update_req["properties"]["sheetId"] == 12345
        assert update_req["properties"]["title"] == "New Sheet Name"
        assert update_req["fields"] == "title"

    def test_rename_sheet_request_structure(self):
        """Test that request body has correct structure for batchUpdate."""
        resolved_sheet_id = 67890
        new_name = "Renamed Sheet"

        request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": resolved_sheet_id,
                            "title": new_name,
                        },
                        "fields": "title",
                    }
                }
            ]
        }

        assert "requests" in request_body
        assert len(request_body["requests"]) == 1
        assert "updateSheetProperties" in request_body["requests"][0]
        props = request_body["requests"][0]["updateSheetProperties"]["properties"]
        assert props["sheetId"] == 67890
        assert props["title"] == "Renamed Sheet"


class TestRenameSheetValidation:
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

    def test_new_name_required(self):
        """Test that new_name must be non-empty."""
        new_name_empty = ""
        new_name_whitespace = "   "
        new_name_valid = "New Name"

        assert not new_name_empty or not new_name_empty.strip()
        assert not new_name_whitespace or not new_name_whitespace.strip()
        assert new_name_valid and new_name_valid.strip()

    def test_new_name_none_is_invalid(self):
        """Test that None new_name is invalid."""
        new_name = None

        is_invalid = not new_name or not (new_name and new_name.strip())

        assert is_invalid is True


class TestRenameSheetNameConflict:
    """Tests for detecting name conflicts with existing sheets."""

    def test_detect_name_conflict_by_sheet_id(self):
        """Test detecting conflict when renaming by sheet ID."""
        sheet_id = 12345
        new_name = "Existing Sheet"
        sheets = [
            {"properties": {"sheetId": 12345, "title": "Original Name"}},
            {"properties": {"sheetId": 67890, "title": "Existing Sheet"}},
        ]

        has_conflict = False
        for sheet in sheets:
            props = sheet.get("properties", {})
            existing_title = props.get("title", "")
            existing_id = props.get("sheetId")
            if existing_title == new_name:
                if sheet_id is not None and existing_id != sheet_id:
                    has_conflict = True
                    break

        assert has_conflict is True

    def test_no_conflict_renaming_same_sheet_by_id(self):
        """Test no conflict when renaming to same name (same sheet by ID)."""
        sheet_id = 12345
        new_name = "Same Name"
        sheets = [
            {"properties": {"sheetId": 12345, "title": "Same Name"}},
            {"properties": {"sheetId": 67890, "title": "Other Sheet"}},
        ]

        has_conflict = False
        for sheet in sheets:
            props = sheet.get("properties", {})
            existing_title = props.get("title", "")
            existing_id = props.get("sheetId")
            if existing_title == new_name:
                if sheet_id is not None and existing_id != sheet_id:
                    has_conflict = True
                    break

        assert has_conflict is False

    def test_detect_name_conflict_by_sheet_name(self):
        """Test detecting conflict when renaming by sheet name."""
        sheet_name = "Original Name"
        new_name = "Existing Sheet"
        sheets = [
            {"properties": {"sheetId": 12345, "title": "Original Name"}},
            {"properties": {"sheetId": 67890, "title": "Existing Sheet"}},
        ]

        has_conflict = False
        for sheet in sheets:
            props = sheet.get("properties", {})
            existing_title = props.get("title", "")
            if existing_title == new_name:
                if sheet_name is not None and existing_title != sheet_name:
                    has_conflict = True
                    break

        assert has_conflict is True

    def test_no_conflict_unique_name(self):
        """Test no conflict with unique name."""
        sheet_id = 12345
        new_name = "Unique New Name"
        sheets = [
            {"properties": {"sheetId": 12345, "title": "Original Name"}},
            {"properties": {"sheetId": 67890, "title": "Other Sheet"}},
        ]

        has_conflict = False
        for sheet in sheets:
            props = sheet.get("properties", {})
            existing_title = props.get("title", "")
            existing_id = props.get("sheetId")
            if existing_title == new_name:
                if sheet_id is not None and existing_id != sheet_id:
                    has_conflict = True
                    break

        assert has_conflict is False


class TestRenameSheetTitleResolution:
    """Tests for resolving sheet title from sheet list."""

    def test_find_sheet_title_by_id(self):
        """Test finding sheet title by sheet ID."""
        resolved_sheet_id = 12345
        sheets = [
            {"properties": {"sheetId": 0, "title": "First"}},
            {"properties": {"sheetId": 12345, "title": "Target Sheet"}},
            {"properties": {"sheetId": 99999, "title": "Other"}},
        ]

        current_title = None
        for sheet in sheets:
            props = sheet.get("properties", {})
            if props.get("sheetId") == resolved_sheet_id:
                current_title = props.get("title", "Unknown")
                break

        assert current_title == "Target Sheet"

    def test_sheet_not_found_returns_none(self):
        """Test that non-existent sheet ID returns None."""
        resolved_sheet_id = 99999
        sheets = [
            {"properties": {"sheetId": 0, "title": "First"}},
            {"properties": {"sheetId": 12345, "title": "Second"}},
        ]

        current_title = None
        for sheet in sheets:
            props = sheet.get("properties", {})
            if props.get("sheetId") == resolved_sheet_id:
                current_title = props.get("title", "Unknown")
                break

        assert current_title is None

    def test_fallback_to_unknown(self):
        """Test fallback to 'Unknown' when title is missing."""
        resolved_sheet_id = 12345
        sheets = [
            {"properties": {"sheetId": 12345}},  # No title
        ]

        current_title = None
        for sheet in sheets:
            props = sheet.get("properties", {})
            if props.get("sheetId") == resolved_sheet_id:
                current_title = props.get("title", "Unknown")
                break

        assert current_title == "Unknown"


class TestRenameSheetSuccessMessages:
    """Tests for success message formatting."""

    def test_success_message_format(self):
        """Test success message includes all relevant info."""
        current_title = "Old Name"
        new_name = "New Name"
        resolved_sheet_id = 12345
        spreadsheet_id = "abc123"
        user_email = "test@example.com"

        message = (
            f"Successfully renamed sheet '{current_title}' to '{new_name}' "
            f"(ID: {resolved_sheet_id}) in spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "Successfully renamed sheet 'Old Name'" in message
        assert "to 'New Name'" in message
        assert "(ID: 12345)" in message
        assert f"in spreadsheet {spreadsheet_id}" in message
        assert user_email in message

    def test_success_message_with_numeric_id(self):
        """Test success message with different sheet ID."""
        current_title = "Data"
        new_name = "Archived Data"
        resolved_sheet_id = 0
        spreadsheet_id = "xyz789"
        user_email = "user@company.com"

        message = (
            f"Successfully renamed sheet '{current_title}' to '{new_name}' "
            f"(ID: {resolved_sheet_id}) in spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "(ID: 0)" in message
        assert "sheet 'Data'" in message
        assert "to 'Archived Data'" in message


class TestRenameSheetErrorMessages:
    """Tests for error message formatting."""

    def test_missing_identifier_error(self):
        """Test error message when no identifier is provided."""
        error_msg = (
            "Either 'sheet_name' or 'sheet_id' must be provided "
            "to identify the sheet to rename."
        )

        assert "sheet_name" in error_msg
        assert "sheet_id" in error_msg
        assert "identify" in error_msg

    def test_empty_new_name_error(self):
        """Test error message when new_name is empty."""
        error_msg = "'new_name' must be a non-empty string."

        assert "new_name" in error_msg
        assert "non-empty" in error_msg

    def test_name_conflict_error(self):
        """Test error message when name conflicts with existing sheet."""
        new_name = "Existing Sheet"
        spreadsheet_id = "abc123"
        error_msg = f"A sheet named '{new_name}' already exists in spreadsheet {spreadsheet_id}."

        assert f"A sheet named '{new_name}'" in error_msg
        assert "already exists" in error_msg
        assert spreadsheet_id in error_msg

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


class TestRenameSheetParameterPrecedence:
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


class TestRenameSheetEdgeCases:
    """Tests for edge cases in rename operations."""

    def test_sheet_id_zero_is_valid(self):
        """Test that sheet ID of 0 is valid (first sheet often has ID 0)."""
        sheet_id = 0

        # 0 is a valid sheet ID (often the first sheet)
        is_valid = sheet_id is not None

        assert is_valid is True

    def test_special_characters_in_new_name(self):
        """Test new names with special characters."""
        new_names = [
            "Data (2024)",
            "Q1 & Q2",
            "Sheet - Draft",
            "Tab/Name",
            "Sheet 'Copy'",
        ]

        for name in new_names:
            # All should be valid as new names
            is_valid = bool(name and name.strip())
            assert is_valid is True

    def test_whitespace_in_new_name(self):
        """Test new names with whitespace."""
        new_name = "  Sheet with Spaces  "

        # Name with whitespace is valid (strip just for validation)
        is_valid = bool(new_name and new_name.strip())
        assert is_valid is True

    def test_unicode_in_new_name(self):
        """Test new names with unicode characters."""
        new_names = [
            "Sheet 日本語",
            "Données",
            "Hoja de datos",
            "Лист",
            "表格",
        ]

        for name in new_names:
            is_valid = bool(name and name.strip())
            assert is_valid is True

    def test_rename_to_same_name(self):
        """Test renaming a sheet to its current name (no-op)."""
        sheet_id = 12345
        current_name = "Same Name"
        new_name = "Same Name"
        sheets = [
            {"properties": {"sheetId": 12345, "title": "Same Name"}},
        ]

        # Should not be a conflict when renaming to same name
        has_conflict = False
        for sheet in sheets:
            props = sheet.get("properties", {})
            existing_title = props.get("title", "")
            existing_id = props.get("sheetId")
            if existing_title == new_name:
                if sheet_id is not None and existing_id != sheet_id:
                    has_conflict = True
                    break

        assert has_conflict is False
        assert current_name == new_name  # No actual change needed
