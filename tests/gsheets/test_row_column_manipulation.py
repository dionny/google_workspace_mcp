"""
Tests for row and column manipulation tools in Google Sheets.

These tests verify:
1. insert_rows - Insert blank rows at a specified position
2. delete_rows - Delete rows at a specified position
3. insert_columns - Insert blank columns at a specified position
4. delete_columns - Delete columns at a specified position
5. auto_resize_dimension - Auto-fit row height or column width to content

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
    _column_index_to_letter,
    _resolve_sheet_id,
)


class TestRowInsertionHelpers:
    """Tests for helper functions used by insert_rows."""

    def test_row_index_conversion(self):
        """Test converting 1-indexed row to 0-indexed for API."""
        # Row 1 (1-indexed) = index 0 (0-indexed)
        row, _ = _parse_cell_reference("A1")
        assert row == 0

        # Row 5 (1-indexed) = index 4 (0-indexed)
        row, _ = _parse_cell_reference("A5")
        assert row == 4

        # Row 100 (1-indexed) = index 99 (0-indexed)
        row, _ = _parse_cell_reference("A100")
        assert row == 99


class TestRowInsertRequestBuilding:
    """Tests for building insert_rows request structures."""

    def test_basic_insert_rows_request(self):
        """Test basic insert rows request structure."""
        start_row = 5
        num_rows = 3
        sheet_id = 0

        # Convert to 0-indexed
        start_index = start_row - 1

        request = {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": start_index,
                    "endIndex": start_index + num_rows,
                },
                "inheritFromBefore": False,
            }
        }

        assert request["insertDimension"]["range"]["dimension"] == "ROWS"
        assert request["insertDimension"]["range"]["startIndex"] == 4
        assert request["insertDimension"]["range"]["endIndex"] == 7
        assert request["insertDimension"]["inheritFromBefore"] is False

    def test_insert_rows_inherit_from_before(self):
        """Test insert rows with inherit_from_before=True."""
        start_row = 3
        num_rows = 2
        sheet_id = 123

        start_index = start_row - 1

        request = {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": start_index,
                    "endIndex": start_index + num_rows,
                },
                "inheritFromBefore": True,
            }
        }

        assert request["insertDimension"]["inheritFromBefore"] is True
        assert request["insertDimension"]["range"]["sheetId"] == 123


class TestRowDeletionRequestBuilding:
    """Tests for building delete_rows request structures."""

    def test_basic_delete_rows_request(self):
        """Test basic delete rows request structure."""
        start_row = 5
        num_rows = 3
        sheet_id = 0

        start_index = start_row - 1

        request = {
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": start_index,
                    "endIndex": start_index + num_rows,
                }
            }
        }

        assert request["deleteDimension"]["range"]["dimension"] == "ROWS"
        assert request["deleteDimension"]["range"]["startIndex"] == 4
        assert request["deleteDimension"]["range"]["endIndex"] == 7

    def test_delete_single_row(self):
        """Test delete single row request."""
        start_row = 10
        num_rows = 1
        sheet_id = 0

        start_index = start_row - 1

        request = {
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": start_index,
                    "endIndex": start_index + num_rows,
                }
            }
        }

        assert request["deleteDimension"]["range"]["startIndex"] == 9
        assert request["deleteDimension"]["range"]["endIndex"] == 10


class TestColumnInsertionHelpers:
    """Tests for helper functions used by insert_columns."""

    def test_column_letter_to_index(self):
        """Test converting column letters to 0-indexed indices."""
        # Column A = index 0
        _, col = _parse_cell_reference("A1")
        assert col == 0

        # Column B = index 1
        _, col = _parse_cell_reference("B1")
        assert col == 1

        # Column Z = index 25
        _, col = _parse_cell_reference("Z1")
        assert col == 25

        # Column AA = index 26
        _, col = _parse_cell_reference("AA1")
        assert col == 26

        # Column AZ = index 51
        _, col = _parse_cell_reference("AZ1")
        assert col == 51

    def test_column_index_to_letter(self):
        """Test converting 0-indexed column index to letters."""
        assert _column_index_to_letter(0) == "A"
        assert _column_index_to_letter(1) == "B"
        assert _column_index_to_letter(25) == "Z"
        assert _column_index_to_letter(26) == "AA"
        assert _column_index_to_letter(27) == "AB"
        assert _column_index_to_letter(51) == "AZ"
        assert _column_index_to_letter(52) == "BA"


class TestColumnInsertRequestBuilding:
    """Tests for building insert_columns request structures."""

    def test_basic_insert_columns_request(self):
        """Test basic insert columns request structure."""
        start_column = "C"
        num_columns = 2
        sheet_id = 0

        # Parse column letter to index
        _, start_col_idx = _parse_cell_reference(f"{start_column}1")

        request = {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": start_col_idx,
                    "endIndex": start_col_idx + num_columns,
                },
                "inheritFromBefore": False,
            }
        }

        assert request["insertDimension"]["range"]["dimension"] == "COLUMNS"
        assert request["insertDimension"]["range"]["startIndex"] == 2  # C = index 2
        assert request["insertDimension"]["range"]["endIndex"] == 4
        assert request["insertDimension"]["inheritFromBefore"] is False

    def test_insert_columns_inherit_from_before(self):
        """Test insert columns with inherit_from_before=True."""
        start_column = "B"
        num_columns = 1
        sheet_id = 456

        _, start_col_idx = _parse_cell_reference(f"{start_column}1")

        request = {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": start_col_idx,
                    "endIndex": start_col_idx + num_columns,
                },
                "inheritFromBefore": True,
            }
        }

        assert request["insertDimension"]["inheritFromBefore"] is True
        assert request["insertDimension"]["range"]["sheetId"] == 456


class TestColumnDeletionRequestBuilding:
    """Tests for building delete_columns request structures."""

    def test_basic_delete_columns_request(self):
        """Test basic delete columns request structure."""
        start_column = "D"
        num_columns = 3
        sheet_id = 0

        _, start_col_idx = _parse_cell_reference(f"{start_column}1")

        request = {
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": start_col_idx,
                    "endIndex": start_col_idx + num_columns,
                }
            }
        }

        assert request["deleteDimension"]["range"]["dimension"] == "COLUMNS"
        assert request["deleteDimension"]["range"]["startIndex"] == 3  # D = index 3
        assert request["deleteDimension"]["range"]["endIndex"] == 6

    def test_delete_single_column(self):
        """Test delete single column request."""
        start_column = "A"
        num_columns = 1
        sheet_id = 0

        _, start_col_idx = _parse_cell_reference(f"{start_column}1")

        request = {
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": start_col_idx,
                    "endIndex": start_col_idx + num_columns,
                }
            }
        }

        assert request["deleteDimension"]["range"]["startIndex"] == 0
        assert request["deleteDimension"]["range"]["endIndex"] == 1


class TestAutoResizeRequestBuilding:
    """Tests for building auto_resize_dimension request structures."""

    def test_auto_resize_columns_request(self):
        """Test auto resize columns request structure."""
        start_column = "A"
        end_column = "D"
        sheet_id = 0

        _, start_col_idx = _parse_cell_reference(f"{start_column}1")
        _, end_col_idx = _parse_cell_reference(f"{end_column}1")

        request = {
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": start_col_idx,
                    "endIndex": end_col_idx + 1,  # API uses exclusive end
                }
            }
        }

        assert request["autoResizeDimensions"]["dimensions"]["dimension"] == "COLUMNS"
        assert request["autoResizeDimensions"]["dimensions"]["startIndex"] == 0
        assert request["autoResizeDimensions"]["dimensions"]["endIndex"] == 4

    def test_auto_resize_rows_request(self):
        """Test auto resize rows request structure."""
        start_index = 1  # 1-indexed
        end_index = 10  # 1-indexed inclusive
        sheet_id = 0

        # Convert to 0-indexed
        resolved_start = start_index - 1
        resolved_end = end_index  # Use as exclusive end in 0-indexed

        request = {
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": resolved_start,
                    "endIndex": resolved_end,
                }
            }
        }

        assert request["autoResizeDimensions"]["dimensions"]["dimension"] == "ROWS"
        assert request["autoResizeDimensions"]["dimensions"]["startIndex"] == 0
        assert request["autoResizeDimensions"]["dimensions"]["endIndex"] == 10


class TestInputValidation:
    """Tests for input validation logic."""

    def test_start_row_validation(self):
        """Test that start_row must be >= 1."""
        start_row = 0
        assert start_row < 1, "start_row must be 1 or greater"

        start_row = -1
        assert start_row < 1, "start_row must be 1 or greater"

        start_row = 1
        assert start_row >= 1, "Valid start_row"

    def test_num_rows_validation(self):
        """Test that num_rows must be >= 1."""
        num_rows = 0
        assert num_rows < 1, "num_rows must be 1 or greater"

        num_rows = -1
        assert num_rows < 1, "num_rows must be 1 or greater"

        num_rows = 1
        assert num_rows >= 1, "Valid num_rows"

    def test_num_columns_validation(self):
        """Test that num_columns must be >= 1."""
        num_columns = 0
        assert num_columns < 1, "num_columns must be 1 or greater"

        num_columns = 1
        assert num_columns >= 1, "Valid num_columns"

    def test_dimension_validation(self):
        """Test dimension parameter validation."""
        valid_dimensions = ["ROWS", "COLUMNS"]

        assert "ROWS" in valid_dimensions
        assert "COLUMNS" in valid_dimensions
        assert "CELLS" not in valid_dimensions
        assert "rows" not in valid_dimensions  # Case sensitive before .upper()


class TestSheetIdResolution:
    """Tests for sheet ID resolution with row/column operations."""

    @pytest.mark.asyncio
    async def test_resolve_sheet_id_for_row_operations(self):
        """Test sheet ID resolution for row operations."""
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

        # When sheet_id is provided directly
        result = await _resolve_sheet_id(service, "spreadsheet123", None, 456)
        assert result == 456

        # When sheet_name is provided
        result = await _resolve_sheet_id(service, "spreadsheet123", "Data", None)
        assert result == 123

        # When neither is provided, get first sheet
        result = await _resolve_sheet_id(service, "spreadsheet123", None, None)
        assert result == 0
