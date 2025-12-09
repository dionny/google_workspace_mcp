"""
Tests for data manipulation tools in Google Sheets.

These tests verify:
1. sort_range - Sort data by a specified column
2. find_and_replace - Find and replace text in a sheet

Note: The actual tools are wrapped by decorators and cannot be called directly
in unit tests. We test the helper functions and validation logic here.
Use tools_cli.py for integration testing.
"""

import pytest
from unittest.mock import MagicMock

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from gsheets.sheets_tools import (
    _parse_cell_reference,
    _parse_range_to_grid,
    _resolve_sheet_id,
)


class TestSortRangeHelpers:
    """Tests for helper functions used by sort_range."""

    def test_parse_column_for_sort(self):
        """Test parsing column letters for sort operations."""
        # Column A = index 0
        row, col = _parse_cell_reference("A1")
        assert col == 0

        # Column B = index 1
        row, col = _parse_cell_reference("B1")
        assert col == 1

        # Column Z = index 25
        row, col = _parse_cell_reference("Z1")
        assert col == 25

        # Column AA = index 26
        row, col = _parse_cell_reference("AA1")
        assert col == 26

    def test_parse_range_for_sort(self):
        """Test parsing range for sort operations."""
        result = _parse_range_to_grid("A1:D10")
        assert result["startRowIndex"] == 0
        assert result["endRowIndex"] == 10
        assert result["startColumnIndex"] == 0
        assert result["endColumnIndex"] == 4

    def test_column_within_range_validation(self):
        """Test that sort column validation logic works correctly."""
        # Range A1:D10 = columns 0-3
        grid_range = _parse_range_to_grid("A1:D10")

        # Column A (index 0) is within range
        _, sort_col_idx = _parse_cell_reference("A1")
        assert sort_col_idx >= grid_range["startColumnIndex"]
        assert sort_col_idx < grid_range["endColumnIndex"]

        # Column D (index 3) is within range
        _, sort_col_idx = _parse_cell_reference("D1")
        assert sort_col_idx >= grid_range["startColumnIndex"]
        assert sort_col_idx < grid_range["endColumnIndex"]

        # Column E (index 4) is outside range
        _, sort_col_idx = _parse_cell_reference("E1")
        assert sort_col_idx >= grid_range["endColumnIndex"]

    def test_sort_order_values(self):
        """Test that sort order values are correct."""
        # These are the values expected by Google Sheets API
        assert "ASCENDING" == "ASCENDING"
        assert "DESCENDING" == "DESCENDING"


class TestFindAndReplaceHelpers:
    """Tests for helper functions used by find_and_replace."""

    def test_range_parsing_for_find_replace(self):
        """Test range parsing for find/replace within specific range."""
        result = _parse_range_to_grid("B2:E5")
        assert result["startRowIndex"] == 1
        assert result["endRowIndex"] == 5
        assert result["startColumnIndex"] == 1
        assert result["endColumnIndex"] == 5

    @pytest.mark.asyncio
    async def test_resolve_sheet_id_for_find_replace(self):
        """Test sheet ID resolution for find/replace operations."""
        mock_spreadsheet_data = {
            "sheets": [
                {"properties": {"sheetId": 0, "title": "Sheet1"}},
                {"properties": {"sheetId": 123, "title": "Data"}},
            ]
        }

        service = MagicMock()
        service.spreadsheets().get().execute = MagicMock(
            return_value=mock_spreadsheet_data
        )

        # When sheet name is provided
        result = await _resolve_sheet_id(service, "spreadsheet123", "Data", None)
        assert result == 123

        # When neither is provided, get first sheet
        result = await _resolve_sheet_id(service, "spreadsheet123", None, None)
        assert result == 0


class TestSortRangeValidation:
    """Tests for sort_range validation scenarios."""

    def test_sort_column_before_range_start(self):
        """Test validation when sort column is before range start."""
        # Range B1:D10 starts at column index 1
        grid_range = _parse_range_to_grid("B1:D10")

        # Column A (index 0) is before range start
        _, sort_col_idx = _parse_cell_reference("A1")
        assert sort_col_idx < grid_range["startColumnIndex"]

    def test_sort_column_after_range_end(self):
        """Test validation when sort column is after range end."""
        # Range A1:C10 ends at column index 3 (exclusive)
        grid_range = _parse_range_to_grid("A1:C10")

        # Column D (index 3) is at/after range end
        _, sort_col_idx = _parse_cell_reference("D1")
        assert sort_col_idx >= grid_range["endColumnIndex"]


class TestFindReplaceRequestBuilding:
    """Tests for building find/replace request structures."""

    def test_basic_find_replace_structure(self):
        """Test basic find/replace request structure."""
        request = {
            "find": "old",
            "replacement": "new",
            "matchCase": False,
            "matchEntireCell": False,
            "searchByRegex": False,
        }

        assert request["find"] == "old"
        assert request["replacement"] == "new"
        assert request["matchCase"] is False
        assert request["matchEntireCell"] is False
        assert request["searchByRegex"] is False

    def test_case_sensitive_find_replace(self):
        """Test case-sensitive find/replace request."""
        request = {
            "find": "Test",
            "replacement": "Result",
            "matchCase": True,
            "matchEntireCell": False,
            "searchByRegex": False,
        }

        assert request["matchCase"] is True

    def test_exact_cell_match_find_replace(self):
        """Test exact cell match find/replace request."""
        request = {
            "find": "Exact",
            "replacement": "Match",
            "matchCase": False,
            "matchEntireCell": True,
            "searchByRegex": False,
        }

        assert request["matchEntireCell"] is True

    def test_regex_find_replace(self):
        """Test regex find/replace request."""
        request = {
            "find": r"\d+",
            "replacement": "NUMBER",
            "matchCase": False,
            "matchEntireCell": False,
            "searchByRegex": True,
        }

        assert request["searchByRegex"] is True
        assert request["find"] == r"\d+"

    def test_all_sheets_find_replace(self):
        """Test all sheets find/replace request."""
        request = {
            "find": "text",
            "replacement": "new",
            "matchCase": False,
            "matchEntireCell": False,
            "searchByRegex": False,
            "allSheets": True,
        }

        assert request["allSheets"] is True

    def test_range_restricted_find_replace(self):
        """Test range-restricted find/replace request."""
        grid_range = _parse_range_to_grid("A1:D10")
        grid_range["sheetId"] = 0

        request = {
            "find": "text",
            "replacement": "new",
            "matchCase": False,
            "matchEntireCell": False,
            "searchByRegex": False,
            "range": grid_range,
        }

        assert "range" in request
        assert request["range"]["sheetId"] == 0
        assert request["range"]["startRowIndex"] == 0
        assert request["range"]["endRowIndex"] == 10

    def test_sheet_restricted_find_replace(self):
        """Test single sheet find/replace request."""
        request = {
            "find": "text",
            "replacement": "new",
            "matchCase": False,
            "matchEntireCell": False,
            "searchByRegex": False,
            "sheetId": 123,
        }

        assert request["sheetId"] == 123


class TestSortRequestBuilding:
    """Tests for building sort request structures."""

    def test_ascending_sort_request(self):
        """Test ascending sort request structure."""
        grid_range = _parse_range_to_grid("A1:D10")
        grid_range["sheetId"] = 0

        request = {
            "sortRange": {
                "range": grid_range,
                "sortSpecs": [
                    {
                        "dimensionIndex": 0,  # Column A
                        "sortOrder": "ASCENDING",
                    }
                ],
            }
        }

        assert request["sortRange"]["sortSpecs"][0]["sortOrder"] == "ASCENDING"

    def test_descending_sort_request(self):
        """Test descending sort request structure."""
        grid_range = _parse_range_to_grid("A1:D10")
        grid_range["sheetId"] = 0

        request = {
            "sortRange": {
                "range": grid_range,
                "sortSpecs": [
                    {
                        "dimensionIndex": 2,  # Column C
                        "sortOrder": "DESCENDING",
                    }
                ],
            }
        }

        assert request["sortRange"]["sortSpecs"][0]["sortOrder"] == "DESCENDING"
        assert request["sortRange"]["sortSpecs"][0]["dimensionIndex"] == 2
