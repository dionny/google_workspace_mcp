"""
Tests for cell notes functionality in Google Sheets.

These tests verify:
1. Cell reference parsing (A1 notation to indices)
2. Column index to letter conversion
3. Unit tests for the helper functions
"""

import pytest

# Import the helper functions directly from the module
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from gsheets.sheets_tools import (
    _parse_cell_reference,
    _column_index_to_letter,
    _strip_sheet_prefix,
)


class TestParseCellReference:
    """Tests for _parse_cell_reference function."""

    def test_simple_cell_a1(self):
        """Test parsing A1."""
        row, col = _parse_cell_reference("A1")
        assert row == 0
        assert col == 0

    def test_simple_cell_b2(self):
        """Test parsing B2."""
        row, col = _parse_cell_reference("B2")
        assert row == 1
        assert col == 1

    def test_cell_z26(self):
        """Test parsing Z26 - last single-letter column."""
        row, col = _parse_cell_reference("Z26")
        assert row == 25
        assert col == 25

    def test_cell_aa1(self):
        """Test parsing AA1 - first double-letter column."""
        row, col = _parse_cell_reference("AA1")
        assert row == 0
        assert col == 26

    def test_cell_ab10(self):
        """Test parsing AB10."""
        row, col = _parse_cell_reference("AB10")
        assert row == 9
        assert col == 27

    def test_cell_az100(self):
        """Test parsing AZ100."""
        row, col = _parse_cell_reference("AZ100")
        assert row == 99
        assert col == 51  # A=0-25, AA-AZ=26-51

    def test_lowercase_cell(self):
        """Test that lowercase cell references are parsed correctly."""
        row, col = _parse_cell_reference("a1")
        assert row == 0
        assert col == 0

    def test_mixed_case_cell(self):
        """Test that mixed case cell references are parsed correctly."""
        row, col = _parse_cell_reference("aB2")
        assert row == 1
        assert col == 27

    def test_cell_with_whitespace(self):
        """Test that whitespace is stripped."""
        row, col = _parse_cell_reference("  A1  ")
        assert row == 0
        assert col == 0

    def test_invalid_cell_no_number(self):
        """Test that cell without number raises ValueError."""
        with pytest.raises(ValueError, match="Invalid cell reference"):
            _parse_cell_reference("A")

    def test_invalid_cell_no_letters(self):
        """Test that cell without letters raises ValueError."""
        with pytest.raises(ValueError, match="Invalid cell reference"):
            _parse_cell_reference("123")

    def test_invalid_cell_special_chars(self):
        """Test that cell with special chars raises ValueError."""
        with pytest.raises(ValueError, match="Invalid cell reference"):
            _parse_cell_reference("A1!")

    def test_invalid_cell_empty(self):
        """Test that empty string raises ValueError."""
        with pytest.raises(ValueError, match="Invalid cell reference"):
            _parse_cell_reference("")


class TestColumnIndexToLetter:
    """Tests for _column_index_to_letter function."""

    def test_column_a(self):
        """Test index 0 -> A."""
        assert _column_index_to_letter(0) == "A"

    def test_column_b(self):
        """Test index 1 -> B."""
        assert _column_index_to_letter(1) == "B"

    def test_column_z(self):
        """Test index 25 -> Z."""
        assert _column_index_to_letter(25) == "Z"

    def test_column_aa(self):
        """Test index 26 -> AA."""
        assert _column_index_to_letter(26) == "AA"

    def test_column_ab(self):
        """Test index 27 -> AB."""
        assert _column_index_to_letter(27) == "AB"

    def test_column_az(self):
        """Test index 51 -> AZ."""
        assert _column_index_to_letter(51) == "AZ"

    def test_column_ba(self):
        """Test index 52 -> BA."""
        assert _column_index_to_letter(52) == "BA"

    def test_column_aaa(self):
        """Test index 702 -> AAA (26 + 26*26 = 702)."""
        assert _column_index_to_letter(702) == "AAA"


class TestRoundTrip:
    """Tests for round-trip conversion between cell reference and indices."""

    def test_roundtrip_a1(self):
        """Test A1 round-trip."""
        row, col = _parse_cell_reference("A1")
        letter = _column_index_to_letter(col)
        assert f"{letter}{row + 1}" == "A1"

    def test_roundtrip_z100(self):
        """Test Z100 round-trip."""
        row, col = _parse_cell_reference("Z100")
        letter = _column_index_to_letter(col)
        assert f"{letter}{row + 1}" == "Z100"

    def test_roundtrip_aa1(self):
        """Test AA1 round-trip."""
        row, col = _parse_cell_reference("AA1")
        letter = _column_index_to_letter(col)
        assert f"{letter}{row + 1}" == "AA1"

    def test_roundtrip_ab50(self):
        """Test AB50 round-trip."""
        row, col = _parse_cell_reference("AB50")
        letter = _column_index_to_letter(col)
        assert f"{letter}{row + 1}" == "AB50"


class TestCellNoteSheetPrefixExtraction:
    """
    Tests for cell note functions' sheet prefix extraction.

    These tests verify the logic used in update_cell_note and clear_cell_note
    to extract sheet names from cell references like 'SheetName!A1'.
    """

    def test_cell_with_simple_sheet_prefix(self):
        """Test extracting sheet name from 'SheetName!A1'."""
        cell = "DataSheet!A1"
        cell_sheet_prefix, clean_cell = _strip_sheet_prefix(cell)
        assert cell_sheet_prefix == "DataSheet"
        assert clean_cell == "A1"
        # Verify the clean cell can be parsed
        row, col = _parse_cell_reference(clean_cell)
        assert row == 0
        assert col == 0

    def test_cell_with_quoted_sheet_prefix(self):
        """Test extracting sheet name with spaces from quoted format."""
        cell = "'My Sheet'!B2"
        cell_sheet_prefix, clean_cell = _strip_sheet_prefix(cell)
        assert cell_sheet_prefix == "My Sheet"
        assert clean_cell == "B2"
        row, col = _parse_cell_reference(clean_cell)
        assert row == 1
        assert col == 1

    def test_cell_with_escaped_apostrophe_in_sheet_name(self):
        """Test extracting sheet name with escaped apostrophe."""
        cell = "'John''s Data'!C3"
        cell_sheet_prefix, clean_cell = _strip_sheet_prefix(cell)
        assert cell_sheet_prefix == "John's Data"
        assert clean_cell == "C3"
        row, col = _parse_cell_reference(clean_cell)
        assert row == 2
        assert col == 2

    def test_cell_without_sheet_prefix(self):
        """Test cell without sheet prefix returns None for sheet name."""
        cell = "A1"
        cell_sheet_prefix, clean_cell = _strip_sheet_prefix(cell)
        assert cell_sheet_prefix is None
        assert clean_cell == "A1"
        row, col = _parse_cell_reference(clean_cell)
        assert row == 0
        assert col == 0

    def test_cell_with_double_letter_column(self):
        """Test cell with sheet prefix and double-letter column."""
        cell = "Sheet1!AA100"
        cell_sheet_prefix, clean_cell = _strip_sheet_prefix(cell)
        assert cell_sheet_prefix == "Sheet1"
        assert clean_cell == "AA100"
        row, col = _parse_cell_reference(clean_cell)
        assert row == 99
        assert col == 26

    def test_cell_with_exclamation_in_sheet_name(self):
        """Test quoted sheet name containing exclamation mark."""
        cell = "'Important!'!D4"
        cell_sheet_prefix, clean_cell = _strip_sheet_prefix(cell)
        assert cell_sheet_prefix == "Important!"
        assert clean_cell == "D4"
        row, col = _parse_cell_reference(clean_cell)
        assert row == 3
        assert col == 3
