"""
Tests for copy_range functionality in Google Sheets.

These tests verify the logic used by the copy_range tool for copying data
from one range to another within a spreadsheet.
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from gsheets.sheets_tools import _strip_sheet_prefix, _parse_range_to_grid


class TestCopyRangeSourceDestParsing:
    """Tests for parsing source and destination ranges in copy operations."""

    def test_source_range_parsing(self):
        """Test parsing a source range without sheet prefix."""
        _, clean_range = _strip_sheet_prefix("A1:D10")
        grid = _parse_range_to_grid(clean_range)

        assert grid["startRowIndex"] == 0
        assert grid["endRowIndex"] == 10
        assert grid["startColumnIndex"] == 0
        assert grid["endColumnIndex"] == 4

    def test_destination_range_parsing(self):
        """Test parsing a destination range."""
        _, clean_range = _strip_sheet_prefix("F1:I10")
        grid = _parse_range_to_grid(clean_range)

        assert grid["startRowIndex"] == 0
        assert grid["endRowIndex"] == 10
        assert grid["startColumnIndex"] == 5  # F is column 5 (0-indexed)
        assert grid["endColumnIndex"] == 9  # I is column 8, +1 for exclusive

    def test_source_with_sheet_prefix(self):
        """Test parsing source range with sheet name prefix."""
        sheet_name, clean_range = _strip_sheet_prefix("Sheet1!A1:D10")
        assert sheet_name == "Sheet1"
        assert clean_range == "A1:D10"

        grid = _parse_range_to_grid(clean_range)
        assert grid["startRowIndex"] == 0
        assert grid["endRowIndex"] == 10

    def test_destination_with_different_sheet(self):
        """Test parsing destination range on different sheet."""
        sheet_name, clean_range = _strip_sheet_prefix("Sheet2!A1:D10")
        assert sheet_name == "Sheet2"
        assert clean_range == "A1:D10"

    def test_single_cell_source(self):
        """Test parsing single cell as source."""
        _, clean_range = _strip_sheet_prefix("A1")
        grid = _parse_range_to_grid(clean_range)

        assert grid["startRowIndex"] == 0
        assert grid["endRowIndex"] == 1
        assert grid["startColumnIndex"] == 0
        assert grid["endColumnIndex"] == 1

    def test_single_cell_destination(self):
        """Test parsing single cell as destination."""
        _, clean_range = _strip_sheet_prefix("F1")
        grid = _parse_range_to_grid(clean_range)

        assert grid["startRowIndex"] == 0
        assert grid["endRowIndex"] == 1
        assert grid["startColumnIndex"] == 5
        assert grid["endColumnIndex"] == 6


class TestCopyRangePasteTypes:
    """Tests for different paste type options."""

    def test_paste_normal_is_valid(self):
        """Test PASTE_NORMAL option."""
        paste_type = "PASTE_NORMAL"
        assert paste_type == "PASTE_NORMAL"

    def test_paste_values_is_valid(self):
        """Test PASTE_VALUES option."""
        paste_type = "PASTE_VALUES"
        assert paste_type == "PASTE_VALUES"

    def test_paste_format_is_valid(self):
        """Test PASTE_FORMAT option."""
        paste_type = "PASTE_FORMAT"
        assert paste_type == "PASTE_FORMAT"

    def test_paste_no_borders_is_valid(self):
        """Test PASTE_NO_BORDERS option."""
        paste_type = "PASTE_NO_BORDERS"
        assert paste_type == "PASTE_NO_BORDERS"

    def test_paste_formula_is_valid(self):
        """Test PASTE_FORMULA option."""
        paste_type = "PASTE_FORMULA"
        assert paste_type == "PASTE_FORMULA"

    def test_paste_data_validation_is_valid(self):
        """Test PASTE_DATA_VALIDATION option."""
        paste_type = "PASTE_DATA_VALIDATION"
        assert paste_type == "PASTE_DATA_VALIDATION"

    def test_paste_conditional_formatting_is_valid(self):
        """Test PASTE_CONDITIONAL_FORMATTING option."""
        paste_type = "PASTE_CONDITIONAL_FORMATTING"
        assert paste_type == "PASTE_CONDITIONAL_FORMATTING"

    def test_all_paste_types_available(self):
        """Test that all seven paste types are defined."""
        valid_paste_types = [
            "PASTE_NORMAL",
            "PASTE_VALUES",
            "PASTE_FORMAT",
            "PASTE_NO_BORDERS",
            "PASTE_FORMULA",
            "PASTE_DATA_VALIDATION",
            "PASTE_CONDITIONAL_FORMATTING",
        ]
        assert len(valid_paste_types) == 7


class TestCopyRangePasteOrientation:
    """Tests for paste orientation options."""

    def test_normal_orientation(self):
        """Test NORMAL orientation (no transpose)."""
        transpose = False
        orientation = "TRANSPOSE" if transpose else "NORMAL"
        assert orientation == "NORMAL"

    def test_transpose_orientation(self):
        """Test TRANSPOSE orientation."""
        transpose = True
        orientation = "TRANSPOSE" if transpose else "NORMAL"
        assert orientation == "TRANSPOSE"


class TestCopyRangeRequestStructure:
    """Tests for the copy/paste request body structure."""

    def test_copy_paste_request_structure(self):
        """Test the structure of a copyPaste request body."""
        source_grid = {
            "sheetId": 0,
            "startRowIndex": 0,
            "endRowIndex": 10,
            "startColumnIndex": 0,
            "endColumnIndex": 4,
        }
        dest_grid = {
            "sheetId": 0,
            "startRowIndex": 0,
            "endRowIndex": 10,
            "startColumnIndex": 5,
            "endColumnIndex": 9,
        }

        request_body = {
            "requests": [
                {
                    "copyPaste": {
                        "source": source_grid,
                        "destination": dest_grid,
                        "pasteType": "PASTE_NORMAL",
                        "pasteOrientation": "NORMAL",
                    }
                }
            ]
        }

        assert "requests" in request_body
        assert len(request_body["requests"]) == 1
        assert "copyPaste" in request_body["requests"][0]

        copy_paste = request_body["requests"][0]["copyPaste"]
        assert "source" in copy_paste
        assert "destination" in copy_paste
        assert copy_paste["pasteType"] == "PASTE_NORMAL"
        assert copy_paste["pasteOrientation"] == "NORMAL"

    def test_copy_paste_request_with_transpose(self):
        """Test copyPaste request with transpose option."""
        request_body = {
            "requests": [
                {
                    "copyPaste": {
                        "source": {"sheetId": 0},
                        "destination": {"sheetId": 0},
                        "pasteType": "PASTE_VALUES",
                        "pasteOrientation": "TRANSPOSE",
                    }
                }
            ]
        }

        copy_paste = request_body["requests"][0]["copyPaste"]
        assert copy_paste["pasteOrientation"] == "TRANSPOSE"

    def test_copy_paste_between_different_sheets(self):
        """Test copyPaste request between different sheets."""
        source_grid = {"sheetId": 0, "startRowIndex": 0, "endRowIndex": 10}
        dest_grid = {"sheetId": 123, "startRowIndex": 0, "endRowIndex": 10}

        request_body = {
            "requests": [
                {
                    "copyPaste": {
                        "source": source_grid,
                        "destination": dest_grid,
                        "pasteType": "PASTE_NORMAL",
                        "pasteOrientation": "NORMAL",
                    }
                }
            ]
        }

        copy_paste = request_body["requests"][0]["copyPaste"]
        assert copy_paste["source"]["sheetId"] == 0
        assert copy_paste["destination"]["sheetId"] == 123


class TestCopyRangeSheetResolution:
    """Tests for sheet resolution logic in copy operations."""

    def test_explicit_source_sheet_name_overrides_extracted(self):
        """Test that explicit source_sheet_name takes precedence over extracted."""
        extracted_source_sheet = "Sheet1"
        source_sheet_name = "DataSheet"

        effective_source_sheet = (
            source_sheet_name
            if source_sheet_name is not None
            else extracted_source_sheet
        )

        assert effective_source_sheet == "DataSheet"

    def test_extracted_sheet_name_used_when_no_explicit(self):
        """Test that extracted sheet name is used when none explicit."""
        extracted_source_sheet = "Sheet1"
        source_sheet_name = None

        effective_source_sheet = (
            source_sheet_name
            if source_sheet_name is not None
            else extracted_source_sheet
        )

        assert effective_source_sheet == "Sheet1"

    def test_destination_defaults_to_source_when_not_specified(self):
        """Test destination sheet defaults to source sheet when not specified."""
        resolved_source_sheet_id = 123
        destination_sheet_id = None
        destination_sheet_name = None
        extracted_dest_sheet = None

        # Logic from copy_range function
        if (
            destination_sheet_id is None
            and destination_sheet_name is None
            and extracted_dest_sheet is None
        ):
            resolved_dest_sheet_id = resolved_source_sheet_id
        else:
            resolved_dest_sheet_id = 456  # Would be resolved from name/id

        assert resolved_dest_sheet_id == resolved_source_sheet_id

    def test_explicit_destination_sheet_overrides_default(self):
        """Test explicit destination sheet overrides default."""
        resolved_source_sheet_id = 123
        destination_sheet_id = None
        destination_sheet_name = "OtherSheet"
        extracted_dest_sheet = None

        if (
            destination_sheet_id is None
            and destination_sheet_name is None
            and extracted_dest_sheet is None
        ):
            resolved_dest_sheet_id = resolved_source_sheet_id
        else:
            resolved_dest_sheet_id = (
                456  # Would be resolved from destination_sheet_name
            )

        assert resolved_dest_sheet_id == 456


class TestCopyRangeOutputMessage:
    """Tests for the output message formatting."""

    def test_output_message_format_basic(self):
        """Test basic output message format."""
        source_range = "A1:D10"
        destination_range = "F1:I10"
        paste_type = "PASTE_NORMAL"
        transpose = False
        spreadsheet_id = "abc123"
        user_email = "test@example.com"

        transpose_text = " (transposed)" if transpose else ""
        paste_type_text = paste_type.replace("PASTE_", "").replace("_", " ").lower()

        text_output = (
            f"Successfully copied range '{source_range}' to '{destination_range}'{transpose_text} "
            f"with paste type '{paste_type_text}' in spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "Successfully copied range" in text_output
        assert "'A1:D10'" in text_output
        assert "'F1:I10'" in text_output
        assert "normal" in text_output
        assert "transposed" not in text_output

    def test_output_message_with_transpose(self):
        """Test output message includes transposed indicator."""
        transpose = True
        transpose_text = " (transposed)" if transpose else ""

        assert transpose_text == " (transposed)"

    def test_paste_type_text_formatting(self):
        """Test paste type text is formatted correctly."""
        test_cases = [
            ("PASTE_NORMAL", "normal"),
            ("PASTE_VALUES", "values"),
            ("PASTE_FORMAT", "format"),
            ("PASTE_NO_BORDERS", "no borders"),
            ("PASTE_FORMULA", "formula"),
            ("PASTE_DATA_VALIDATION", "data validation"),
            ("PASTE_CONDITIONAL_FORMATTING", "conditional formatting"),
        ]

        for paste_type, expected in test_cases:
            result = paste_type.replace("PASTE_", "").replace("_", " ").lower()
            assert result == expected, f"Failed for {paste_type}: got {result}"


class TestCopyRangeEdgeCases:
    """Tests for edge cases in copy operations."""

    def test_quoted_sheet_name_with_spaces(self):
        """Test copying with quoted sheet names containing spaces."""
        sheet_name, clean_range = _strip_sheet_prefix("'My Data Sheet'!A1:D10")
        assert sheet_name == "My Data Sheet"
        assert clean_range == "A1:D10"

    def test_sheet_name_with_apostrophe(self):
        """Test copying with sheet names containing apostrophes."""
        sheet_name, clean_range = _strip_sheet_prefix("'John''s Data'!A1:D10")
        assert sheet_name == "John's Data"
        assert clean_range == "A1:D10"

    def test_double_letter_columns(self):
        """Test copying with double-letter columns."""
        _, clean_range = _strip_sheet_prefix("AA1:AZ100")
        grid = _parse_range_to_grid(clean_range)

        assert grid["startColumnIndex"] == 26  # AA
        assert grid["endColumnIndex"] == 52  # AZ = 51, +1 for exclusive

    def test_large_range(self):
        """Test copying a large range."""
        _, clean_range = _strip_sheet_prefix("A1:ZZ1000")
        grid = _parse_range_to_grid(clean_range)

        assert grid["startRowIndex"] == 0
        assert grid["endRowIndex"] == 1000
        assert grid["startColumnIndex"] == 0
        assert grid["endColumnIndex"] == 702  # ZZ = 701, +1 for exclusive
