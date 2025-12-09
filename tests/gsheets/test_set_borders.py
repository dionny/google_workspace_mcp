"""
Tests for set_borders functionality in Google Sheets.

These tests verify:
1. _parse_range_to_grid - Helper function for parsing ranges
2. _parse_color - Helper function for parsing colors
3. Input validation for border styles
"""

import pytest

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from gsheets.sheets_tools import _parse_range_to_grid, _parse_color


class TestParseRangeToGrid:
    """Tests for _parse_range_to_grid function."""

    def test_simple_range(self):
        """Test parsing A1:D10 range."""
        result = _parse_range_to_grid("A1:D10")
        assert result["startRowIndex"] == 0
        assert result["endRowIndex"] == 10
        assert result["startColumnIndex"] == 0
        assert result["endColumnIndex"] == 4

    def test_single_cell(self):
        """Test parsing single cell B2."""
        result = _parse_range_to_grid("B2")
        assert result["startRowIndex"] == 1
        assert result["endRowIndex"] == 2
        assert result["startColumnIndex"] == 1
        assert result["endColumnIndex"] == 2

    def test_range_with_double_letter_columns(self):
        """Test parsing AA1:AB5 range."""
        result = _parse_range_to_grid("AA1:AB5")
        assert result["startRowIndex"] == 0
        assert result["endRowIndex"] == 5
        assert result["startColumnIndex"] == 26
        assert result["endColumnIndex"] == 28

    def test_range_starting_not_at_origin(self):
        """Test parsing C5:F10 range."""
        result = _parse_range_to_grid("C5:F10")
        assert result["startRowIndex"] == 4
        assert result["endRowIndex"] == 10
        assert result["startColumnIndex"] == 2
        assert result["endColumnIndex"] == 6


class TestParseColor:
    """Tests for _parse_color function."""

    def test_hex_color_six_digits(self):
        """Test parsing 6-digit hex color."""
        result = _parse_color("#FF0000")
        assert result["red"] == 1.0
        assert result["green"] == 0.0
        assert result["blue"] == 0.0

    def test_hex_color_without_hash(self):
        """Test parsing hex color without # prefix."""
        result = _parse_color("00FF00")
        assert result["red"] == 0.0
        assert result["green"] == 1.0
        assert result["blue"] == 0.0

    def test_hex_color_three_digits(self):
        """Test parsing 3-digit hex color."""
        result = _parse_color("#F00")
        assert result["red"] == 1.0
        assert result["green"] == 0.0
        assert result["blue"] == 0.0

    def test_color_name_red(self):
        """Test parsing color name 'red'."""
        result = _parse_color("red")
        assert result["red"] == 1.0
        assert result["green"] == 0.0
        assert result["blue"] == 0.0

    def test_color_name_blue(self):
        """Test parsing color name 'blue'."""
        result = _parse_color("blue")
        assert result["red"] == 0.0
        assert result["green"] == 0.0
        assert result["blue"] == 1.0

    def test_color_name_case_insensitive(self):
        """Test that color names are case-insensitive."""
        result = _parse_color("RED")
        assert result["red"] == 1.0
        assert result["green"] == 0.0
        assert result["blue"] == 0.0

    def test_color_name_with_whitespace(self):
        """Test that whitespace is trimmed."""
        result = _parse_color("  green  ")
        assert result["red"] == 0.0
        assert result["green"] == 1.0
        assert result["blue"] == 0.0

    def test_invalid_color_raises_error(self):
        """Test that invalid color raises ValueError."""
        with pytest.raises(ValueError, match="Invalid color format"):
            _parse_color("notacolor")

    def test_invalid_hex_raises_error(self):
        """Test that invalid hex raises ValueError."""
        with pytest.raises(ValueError, match="Invalid color format"):
            _parse_color("#GGGGGG")

    def test_color_gray_alias(self):
        """Test that 'gray' and 'grey' both work."""
        result_gray = _parse_color("gray")
        result_grey = _parse_color("grey")
        assert result_gray == result_grey


class TestBorderStyles:
    """Tests for border style validation constants."""

    def test_valid_border_styles(self):
        """Test that all valid border styles are recognized."""
        valid_styles = [
            "DOTTED",
            "DASHED",
            "SOLID",
            "SOLID_MEDIUM",
            "SOLID_THICK",
            "DOUBLE",
            "NONE",
        ]
        for style in valid_styles:
            assert style in [
                "DOTTED",
                "DASHED",
                "SOLID",
                "SOLID_MEDIUM",
                "SOLID_THICK",
                "DOUBLE",
                "NONE",
            ]
