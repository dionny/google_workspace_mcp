"""
Tests for _strip_sheet_prefix and _parse_range_to_grid functions.

These tests verify that range strings with sheet name prefixes are correctly
parsed, fixing the issue where tools like format_cells, merge_cells, etc.
would fail when range_name included a sheet name prefix (e.g., 'SheetName!A1:B2').
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from gsheets.sheets_tools import _strip_sheet_prefix, _parse_range_to_grid


class TestStripSheetPrefix:
    """Tests for _strip_sheet_prefix function."""

    def test_no_sheet_prefix(self):
        """Test range without sheet prefix returns None and original range."""
        sheet_name, clean_range = _strip_sheet_prefix("A1:D10")
        assert sheet_name is None
        assert clean_range == "A1:D10"

    def test_single_cell_no_prefix(self):
        """Test single cell without sheet prefix."""
        sheet_name, clean_range = _strip_sheet_prefix("B2")
        assert sheet_name is None
        assert clean_range == "B2"

    def test_simple_sheet_prefix(self):
        """Test simple sheet name prefix is extracted."""
        sheet_name, clean_range = _strip_sheet_prefix("DataSheet!B2:B3")
        assert sheet_name == "DataSheet"
        assert clean_range == "B2:B3"

    def test_simple_sheet_prefix_single_cell(self):
        """Test simple sheet name prefix with single cell."""
        sheet_name, clean_range = _strip_sheet_prefix("Sheet1!A1")
        assert sheet_name == "Sheet1"
        assert clean_range == "A1"

    def test_quoted_sheet_name_with_space(self):
        """Test quoted sheet name with space is extracted and unquoted."""
        sheet_name, clean_range = _strip_sheet_prefix("'My Data Sheet'!A1:C5")
        assert sheet_name == "My Data Sheet"
        assert clean_range == "A1:C5"

    def test_quoted_sheet_name_with_escaped_apostrophe(self):
        """Test quoted sheet name with escaped apostrophe (doubled quotes)."""
        sheet_name, clean_range = _strip_sheet_prefix("'John''s Data'!A1:D10")
        assert sheet_name == "John's Data"
        assert clean_range == "A1:D10"

    def test_quoted_sheet_name_with_multiple_escaped_apostrophes(self):
        """Test quoted sheet name with multiple escaped apostrophes."""
        sheet_name, clean_range = _strip_sheet_prefix("'It''s John''s Sheet'!B2:C3")
        assert sheet_name == "It's John's Sheet"
        assert clean_range == "B2:C3"

    def test_sheet_name_with_numbers(self):
        """Test sheet name containing numbers."""
        sheet_name, clean_range = _strip_sheet_prefix("Data2024!A1:Z100")
        assert sheet_name == "Data2024"
        assert clean_range == "A1:Z100"

    def test_sheet_name_with_underscore(self):
        """Test sheet name containing underscores."""
        sheet_name, clean_range = _strip_sheet_prefix("my_data_sheet!A1:B2")
        assert sheet_name == "my_data_sheet"
        assert clean_range == "A1:B2"

    def test_quoted_sheet_name_with_exclamation(self):
        """Test quoted sheet name containing exclamation mark."""
        sheet_name, clean_range = _strip_sheet_prefix("'Important!'!A1:A10")
        assert sheet_name == "Important!"
        assert clean_range == "A1:A10"

    def test_quoted_sheet_name_with_colon(self):
        """Test quoted sheet name containing colon."""
        sheet_name, clean_range = _strip_sheet_prefix("'Data:2024'!A1:D10")
        assert sheet_name == "Data:2024"
        assert clean_range == "A1:D10"

    def test_empty_string(self):
        """Test empty string returns None and empty string."""
        sheet_name, clean_range = _strip_sheet_prefix("")
        assert sheet_name is None
        assert clean_range == ""


class TestParseRangeToGridWithSheetPrefix:
    """Tests for _parse_range_to_grid with sheet name prefixes."""

    def test_basic_range(self):
        """Test basic range without sheet prefix."""
        result = _parse_range_to_grid("A1:D10")
        assert result == {
            "startRowIndex": 0,
            "endRowIndex": 10,
            "startColumnIndex": 0,
            "endColumnIndex": 4,
        }

    def test_range_with_sheet_prefix(self):
        """Test that sheet prefix is stripped before parsing."""
        result = _parse_range_to_grid("DataSheet!B2:C5")
        assert result == {
            "startRowIndex": 1,
            "endRowIndex": 5,
            "startColumnIndex": 1,
            "endColumnIndex": 3,
        }

    def test_range_with_quoted_sheet_prefix(self):
        """Test that quoted sheet prefix is stripped before parsing."""
        result = _parse_range_to_grid("'My Sheet'!A1:B2")
        assert result == {
            "startRowIndex": 0,
            "endRowIndex": 2,
            "startColumnIndex": 0,
            "endColumnIndex": 2,
        }

    def test_single_cell_with_sheet_prefix(self):
        """Test single cell with sheet prefix."""
        result = _parse_range_to_grid("Sheet1!C3")
        assert result == {
            "startRowIndex": 2,
            "endRowIndex": 3,
            "startColumnIndex": 2,
            "endColumnIndex": 3,
        }

    def test_range_with_escaped_apostrophe_sheet(self):
        """Test range with escaped apostrophe in sheet name."""
        result = _parse_range_to_grid("'John''s Data'!A1:D10")
        assert result == {
            "startRowIndex": 0,
            "endRowIndex": 10,
            "startColumnIndex": 0,
            "endColumnIndex": 4,
        }

    def test_wide_range_with_double_letter_columns(self):
        """Test range with double-letter columns and sheet prefix."""
        result = _parse_range_to_grid("Data!AA1:AB100")
        assert result == {
            "startRowIndex": 0,
            "endRowIndex": 100,
            "startColumnIndex": 26,  # AA = 26
            "endColumnIndex": 28,  # AB = 27, +1 for exclusive end
        }
