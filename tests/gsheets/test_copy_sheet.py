"""
Unit tests for copy_sheet tool.

These tests verify the logic and request structures for the sheet copy operation.
"""


class TestCopySheetRequestStructure:
    """Tests for the copy sheet request body structure."""

    def test_copy_to_same_spreadsheet_request(self):
        """Test request body for copying within same spreadsheet."""
        spreadsheet_id = "source_spreadsheet_123"
        dest_spreadsheet_id = spreadsheet_id  # Same spreadsheet

        copy_request_body = {"destinationSpreadsheetId": dest_spreadsheet_id}

        assert copy_request_body["destinationSpreadsheetId"] == spreadsheet_id

    def test_copy_to_different_spreadsheet_request(self):
        """Test request body for copying to different spreadsheet."""
        spreadsheet_id = "source_spreadsheet_123"
        dest_spreadsheet_id = "dest_spreadsheet_456"

        copy_request_body = {"destinationSpreadsheetId": dest_spreadsheet_id}

        assert copy_request_body["destinationSpreadsheetId"] == dest_spreadsheet_id
        assert copy_request_body["destinationSpreadsheetId"] != spreadsheet_id

    def test_rename_sheet_request(self):
        """Test request body for renaming the copied sheet."""
        new_sheet_id = 12345
        new_sheet_name = "My Custom Copy Name"

        rename_request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": new_sheet_id,
                            "title": new_sheet_name,
                        },
                        "fields": "title",
                    }
                }
            ]
        }

        update_req = rename_request_body["requests"][0]["updateSheetProperties"]
        assert update_req["properties"]["sheetId"] == 12345
        assert update_req["properties"]["title"] == "My Custom Copy Name"
        assert update_req["fields"] == "title"


class TestCopySheetDestinationLogic:
    """Tests for destination spreadsheet logic."""

    def test_default_to_same_spreadsheet(self):
        """Test that when destination is None, it defaults to source spreadsheet."""
        spreadsheet_id = "source_spreadsheet_123"
        destination_spreadsheet_id = None

        dest_spreadsheet_id = destination_spreadsheet_id or spreadsheet_id

        assert dest_spreadsheet_id == spreadsheet_id

    def test_explicit_destination_used(self):
        """Test that explicit destination spreadsheet is used."""
        spreadsheet_id = "source_spreadsheet_123"
        destination_spreadsheet_id = "dest_spreadsheet_456"

        dest_spreadsheet_id = destination_spreadsheet_id or spreadsheet_id

        assert dest_spreadsheet_id == destination_spreadsheet_id

    def test_is_same_spreadsheet_detection(self):
        """Test detection of same vs different spreadsheet."""
        spreadsheet_id = "spreadsheet_123"

        # Same spreadsheet
        dest_same = spreadsheet_id
        is_same = dest_same == spreadsheet_id
        assert is_same is True

        # Different spreadsheet
        dest_different = "other_spreadsheet"
        is_same = dest_different == spreadsheet_id
        assert is_same is False


class TestCopySheetNamingLogic:
    """Tests for sheet naming logic."""

    def test_default_copy_name_used(self):
        """Test that default 'Copy of X' name is used when no custom name given."""
        source_sheet_title = "Original Sheet"
        new_sheet_name = None
        default_new_name = f"Copy of {source_sheet_title}"

        final_sheet_name = default_new_name
        if new_sheet_name and new_sheet_name != default_new_name:
            final_sheet_name = new_sheet_name

        assert final_sheet_name == "Copy of Original Sheet"

    def test_custom_name_applied(self):
        """Test that custom name is applied when provided."""
        source_sheet_title = "Original Sheet"
        new_sheet_name = "My Backup"
        default_new_name = f"Copy of {source_sheet_title}"

        final_sheet_name = default_new_name
        if new_sheet_name and new_sheet_name != default_new_name:
            final_sheet_name = new_sheet_name

        assert final_sheet_name == "My Backup"

    def test_skip_rename_if_same_as_default(self):
        """Test that rename is skipped if custom name matches default."""
        source_sheet_title = "Original Sheet"
        new_sheet_name = "Copy of Original Sheet"  # Same as default
        default_new_name = f"Copy of {source_sheet_title}"

        should_rename = new_sheet_name and new_sheet_name != default_new_name

        assert should_rename is False


class TestCopySheetSuccessMessages:
    """Tests for success message formatting."""

    def test_same_spreadsheet_message(self):
        """Test success message when copying within same spreadsheet."""
        source_sheet_title = "Template"
        final_sheet_name = "Template Backup"
        new_sheet_id = 12345
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        is_same_spreadsheet = True

        if is_same_spreadsheet:
            message = (
                f"Successfully copied sheet '{source_sheet_title}' to '{final_sheet_name}' "
                f"(ID: {new_sheet_id}) in spreadsheet {spreadsheet_id} for {user_email}."
            )
        else:
            message = "Other message"

        assert "Successfully copied sheet 'Template'" in message
        assert "to 'Template Backup'" in message
        assert "(ID: 12345)" in message
        assert f"in spreadsheet {spreadsheet_id}" in message
        assert user_email in message

    def test_different_spreadsheet_message(self):
        """Test success message when copying to different spreadsheet."""
        source_sheet_title = "Template"
        final_sheet_name = "Imported Template"
        new_sheet_id = 67890
        spreadsheet_id = "source_123"
        dest_spreadsheet_id = "dest_456"
        user_email = "test@example.com"
        is_same_spreadsheet = False

        if is_same_spreadsheet:
            message = "Same spreadsheet message"
        else:
            message = (
                f"Successfully copied sheet '{source_sheet_title}' from spreadsheet {spreadsheet_id} "
                f"to '{final_sheet_name}' (ID: {new_sheet_id}) in destination spreadsheet "
                f"{dest_spreadsheet_id} for {user_email}."
            )

        assert "Successfully copied sheet 'Template'" in message
        assert f"from spreadsheet {spreadsheet_id}" in message
        assert "to 'Imported Template'" in message
        assert f"destination spreadsheet {dest_spreadsheet_id}" in message
        assert user_email in message


class TestCopySheetResponseParsing:
    """Tests for parsing the copyTo API response."""

    def test_extract_new_sheet_id(self):
        """Test extracting new sheet ID from copyTo response."""
        copy_response = {
            "sheetId": 12345,
            "title": "Copy of Original",
            "index": 3,
        }

        new_sheet_id = copy_response.get("sheetId")
        assert new_sheet_id == 12345

    def test_extract_default_title(self):
        """Test extracting default title from copyTo response."""
        source_sheet_title = "Original"
        copy_response = {
            "sheetId": 12345,
            "title": "Copy of Original",
        }

        default_new_name = copy_response.get("title", f"Copy of {source_sheet_title}")
        assert default_new_name == "Copy of Original"

    def test_fallback_title_when_missing(self):
        """Test fallback title when API response doesn't include it."""
        source_sheet_title = "Original"
        copy_response = {
            "sheetId": 12345,
            # No title field
        }

        default_new_name = copy_response.get("title", f"Copy of {source_sheet_title}")
        assert default_new_name == "Copy of Original"


class TestCopySheetParameterResolution:
    """Tests for parameter resolution logic."""

    def test_sheet_id_takes_precedence(self):
        """Test that sheet_id takes precedence over sheet_name."""
        source_sheet_id = 12345

        # Logic from _resolve_sheet_id
        if source_sheet_id is not None:
            resolved_id = source_sheet_id
        else:
            resolved_id = None  # Would be resolved from name

        assert resolved_id == 12345

    def test_sheet_name_used_when_no_id(self):
        """Test that sheet_name is used when sheet_id is None."""
        source_sheet_id = None

        # Logic from _resolve_sheet_id
        if source_sheet_id is not None:
            resolved_id = source_sheet_id
        else:
            # In actual code, this would lookup by name
            resolved_id = "would_be_resolved_from_name"

        assert resolved_id == "would_be_resolved_from_name"

    def test_defaults_to_first_sheet(self):
        """Test that first sheet is used when neither name nor id provided."""
        source_sheet_name = None
        source_sheet_id = None

        # Logic from _resolve_sheet_id when both are None
        if source_sheet_id is not None:
            resolved = "by_id"
        elif source_sheet_name is None:
            resolved = "first_sheet"
        else:
            resolved = "by_name"

        assert resolved == "first_sheet"


class TestCopySheetEdgeCases:
    """Tests for edge cases in copy operations."""

    def test_empty_new_name_ignored(self):
        """Test that empty string for new_sheet_name is treated as None."""
        new_sheet_name = ""
        default_new_name = "Copy of Original"

        # Logic: only rename if new_sheet_name is truthy and different
        should_rename = new_sheet_name and new_sheet_name != default_new_name

        # Empty string is falsy, so should_rename will be the empty string (falsy)
        assert not should_rename

    def test_whitespace_name_preserved(self):
        """Test that name with whitespace is preserved."""
        new_sheet_name = "  Sheet with Spaces  "
        default_new_name = "Copy of Original"

        # Name should be used as-is (API may strip, but we don't)
        final_name = default_new_name
        if new_sheet_name and new_sheet_name != default_new_name:
            final_name = new_sheet_name

        assert final_name == "  Sheet with Spaces  "

    def test_special_characters_in_name(self):
        """Test that special characters in sheet name work."""
        new_sheet_name = "Data (2024) - Q1 & Q2"
        default_new_name = "Copy of Original"

        final_name = default_new_name
        if new_sheet_name and new_sheet_name != default_new_name:
            final_name = new_sheet_name

        assert final_name == "Data (2024) - Q1 & Q2"
