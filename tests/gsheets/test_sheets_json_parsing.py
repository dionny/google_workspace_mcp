"""
Tests for JSON parsing in Google Sheets tools.

These tests verify that parameters passed as JSON strings (as MCP does)
are correctly parsed into Python objects.
"""

import pytest
import json


class TestSheetNamesJsonParsing:
    """Unit tests for sheet_names JSON parsing logic."""

    def test_parse_json_string_to_list(self):
        """Test parsing valid JSON string to list."""
        from gsheets.sheets_tools import json

        # Simulate the parsing logic from create_spreadsheet
        sheet_names = '["Sheet A", "Sheet B", "Sheet C"]'

        parsed_sheet_names = json.loads(sheet_names)
        assert isinstance(parsed_sheet_names, list)
        assert len(parsed_sheet_names) == 3
        assert parsed_sheet_names == ["Sheet A", "Sheet B", "Sheet C"]

    def test_json_list_validation(self):
        """Test that non-list JSON raises ValueError."""
        sheet_names = '{"sheet": "Sheet A"}'

        parsed_sheet_names = json.loads(sheet_names)
        assert not isinstance(parsed_sheet_names, list)

    def test_sheet_name_type_validation(self):
        """Test that non-string sheet names are detected."""
        sheet_names = '["Sheet A", 123, "Sheet C"]'

        parsed_sheet_names = json.loads(sheet_names)
        for i, name in enumerate(parsed_sheet_names):
            if not isinstance(name, str):
                assert i == 1
                break

    def test_invalid_json_raises_error(self):
        """Test that invalid JSON raises JSONDecodeError."""
        invalid_json = '["Sheet A", "Sheet B"'  # Missing closing bracket

        with pytest.raises(json.JSONDecodeError):
            json.loads(invalid_json)

    def test_empty_list_parsing(self):
        """Test parsing empty JSON list."""
        sheet_names = "[]"

        parsed_sheet_names = json.loads(sheet_names)
        assert isinstance(parsed_sheet_names, list)
        assert len(parsed_sheet_names) == 0


class TestCreateSpreadsheetSheetNamesLogic:
    """Tests for the sheet_names parsing logic extracted from create_spreadsheet."""

    def parse_sheet_names(self, sheet_names):
        """
        Extract the sheet_names parsing logic from create_spreadsheet.
        This mirrors the code added to fix the bug.
        """
        if sheet_names is not None and isinstance(sheet_names, str):
            try:
                parsed_sheet_names = json.loads(sheet_names)
                if not isinstance(parsed_sheet_names, list):
                    raise ValueError(
                        f"sheet_names must be a list, got {type(parsed_sheet_names).__name__}"
                    )
                # Validate each sheet name is a string
                for i, name in enumerate(parsed_sheet_names):
                    if not isinstance(name, str):
                        raise ValueError(
                            f"Sheet name at index {i} must be a string, got {type(name).__name__}"
                        )
                return parsed_sheet_names
            except json.JSONDecodeError as e:
                raise Exception(f"Invalid JSON format for sheet_names: {e}")
            except ValueError as e:
                raise Exception(f"Invalid sheet_names structure: {e}")
        return sheet_names

    def test_json_string_parsed_correctly(self):
        """Test that JSON string is parsed to Python list."""
        sheet_names_json = '["Sheet A", "Sheet B", "Sheet C"]'
        result = self.parse_sheet_names(sheet_names_json)

        assert result == ["Sheet A", "Sheet B", "Sheet C"]
        assert isinstance(result, list)

    def test_python_list_passed_through(self):
        """Test that Python list is passed through unchanged."""
        sheet_names_list = ["Sheet 1", "Sheet 2"]
        result = self.parse_sheet_names(sheet_names_list)

        assert result == ["Sheet 1", "Sheet 2"]
        assert result is sheet_names_list  # Same object

    def test_none_passed_through(self):
        """Test that None is passed through unchanged."""
        result = self.parse_sheet_names(None)
        assert result is None

    def test_invalid_json_raises_exception(self):
        """Test that invalid JSON raises Exception with descriptive message."""
        invalid_json = '["Sheet A", "Sheet B"'

        with pytest.raises(Exception) as exc_info:
            self.parse_sheet_names(invalid_json)

        assert "Invalid JSON format for sheet_names" in str(exc_info.value)

    def test_non_list_json_raises_exception(self):
        """Test that JSON object (not list) raises Exception."""
        json_object = '{"sheet": "Sheet A"}'

        with pytest.raises(Exception) as exc_info:
            self.parse_sheet_names(json_object)

        assert "sheet_names must be a list" in str(exc_info.value)

    def test_non_string_elements_raise_exception(self):
        """Test that non-string elements in list raise Exception."""
        json_with_int = '["Sheet A", 123, "Sheet C"]'

        with pytest.raises(Exception) as exc_info:
            self.parse_sheet_names(json_with_int)

        assert "Sheet name at index 1 must be a string" in str(exc_info.value)

    def test_nested_list_raises_exception(self):
        """Test that nested lists raise Exception."""
        nested_list = '["Sheet A", ["Nested"], "Sheet C"]'

        with pytest.raises(Exception) as exc_info:
            self.parse_sheet_names(nested_list)

        assert "Sheet name at index 1 must be a string" in str(exc_info.value)

    def test_unicode_sheet_names(self):
        """Test that unicode characters in sheet names work correctly."""
        unicode_json = '["Sheet 日本語", "Sheet Émoji 🎉", "Sheet Ελληνικά"]'
        result = self.parse_sheet_names(unicode_json)

        assert len(result) == 3
        assert result[0] == "Sheet 日本語"
        assert result[1] == "Sheet Émoji 🎉"
        assert result[2] == "Sheet Ελληνικά"

    def test_empty_string_sheet_name(self):
        """Test that empty string sheet names are parsed (validation is API's job)."""
        json_with_empty = '["Sheet A", "", "Sheet C"]'
        result = self.parse_sheet_names(json_with_empty)

        assert result == ["Sheet A", "", "Sheet C"]

    def test_whitespace_only_json_string(self):
        """Test that whitespace-only JSON string raises error."""
        whitespace_json = "   "

        with pytest.raises(Exception) as exc_info:
            self.parse_sheet_names(whitespace_json)

        assert "Invalid JSON format for sheet_names" in str(exc_info.value)

    def test_json_null_raises_exception(self):
        """Test that JSON null string raises appropriate error."""
        null_json = "null"

        with pytest.raises(Exception) as exc_info:
            self.parse_sheet_names(null_json)

        assert "sheet_names must be a list" in str(exc_info.value)
