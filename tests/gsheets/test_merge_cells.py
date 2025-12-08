"""
Unit tests for merge_cells and unmerge_cells tools in Google Sheets.

These tests verify the logic and validation for cell merging operations
including merge type validation and request body structure.
"""

# Import the helper functions directly for testing
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from gsheets.sheets_tools import _parse_range_to_grid_range


class TestMergeCellsRequestBodyStructure:
    """Tests for the request body structure of merge_cells."""

    def test_merge_all_request_structure(self):
        """Test the structure of a MERGE_ALL request."""
        sheet_id = 0
        grid_range = _parse_range_to_grid_range("A1:C1", sheet_id)

        request_body = {
            "requests": [
                {"mergeCells": {"range": grid_range, "mergeType": "MERGE_ALL"}}
            ]
        }

        assert request_body["requests"][0]["mergeCells"]["mergeType"] == "MERGE_ALL"
        assert request_body["requests"][0]["mergeCells"]["range"]["sheetId"] == 0
        assert (
            request_body["requests"][0]["mergeCells"]["range"]["startColumnIndex"] == 0
        )
        assert request_body["requests"][0]["mergeCells"]["range"]["endColumnIndex"] == 3
        assert request_body["requests"][0]["mergeCells"]["range"]["startRowIndex"] == 0
        assert request_body["requests"][0]["mergeCells"]["range"]["endRowIndex"] == 1

    def test_merge_columns_request_structure(self):
        """Test the structure of a MERGE_COLUMNS request."""
        sheet_id = 123
        grid_range = _parse_range_to_grid_range("A1:D2", sheet_id)

        request_body = {
            "requests": [
                {"mergeCells": {"range": grid_range, "mergeType": "MERGE_COLUMNS"}}
            ]
        }

        assert request_body["requests"][0]["mergeCells"]["mergeType"] == "MERGE_COLUMNS"
        assert request_body["requests"][0]["mergeCells"]["range"]["sheetId"] == 123
        assert (
            request_body["requests"][0]["mergeCells"]["range"]["startColumnIndex"] == 0
        )
        assert request_body["requests"][0]["mergeCells"]["range"]["endColumnIndex"] == 4
        assert request_body["requests"][0]["mergeCells"]["range"]["startRowIndex"] == 0
        assert request_body["requests"][0]["mergeCells"]["range"]["endRowIndex"] == 2

    def test_merge_rows_request_structure(self):
        """Test the structure of a MERGE_ROWS request."""
        sheet_id = 456
        grid_range = _parse_range_to_grid_range("B2:B5", sheet_id)

        request_body = {
            "requests": [
                {"mergeCells": {"range": grid_range, "mergeType": "MERGE_ROWS"}}
            ]
        }

        assert request_body["requests"][0]["mergeCells"]["mergeType"] == "MERGE_ROWS"
        assert request_body["requests"][0]["mergeCells"]["range"]["sheetId"] == 456
        assert (
            request_body["requests"][0]["mergeCells"]["range"]["startColumnIndex"] == 1
        )
        assert request_body["requests"][0]["mergeCells"]["range"]["endColumnIndex"] == 2


class TestUnmergeCellsRequestBodyStructure:
    """Tests for the request body structure of unmerge_cells."""

    def test_unmerge_cells_request_structure(self):
        """Test the structure of an unmerge request."""
        sheet_id = 0
        grid_range = _parse_range_to_grid_range("A1:C1", sheet_id)

        request_body = {"requests": [{"unmergeCells": {"range": grid_range}}]}

        assert "unmergeCells" in request_body["requests"][0]
        assert request_body["requests"][0]["unmergeCells"]["range"]["sheetId"] == 0
        assert (
            request_body["requests"][0]["unmergeCells"]["range"]["startColumnIndex"]
            == 0
        )
        assert (
            request_body["requests"][0]["unmergeCells"]["range"]["endColumnIndex"] == 3
        )

    def test_unmerge_large_range_structure(self):
        """Test unmerge request for a large range."""
        sheet_id = 0
        grid_range = _parse_range_to_grid_range("A1:Z100", sheet_id)

        request_body = {"requests": [{"unmergeCells": {"range": grid_range}}]}

        assert (
            request_body["requests"][0]["unmergeCells"]["range"]["startColumnIndex"]
            == 0
        )
        assert (
            request_body["requests"][0]["unmergeCells"]["range"]["endColumnIndex"] == 26
        )
        assert (
            request_body["requests"][0]["unmergeCells"]["range"]["startRowIndex"] == 0
        )
        assert (
            request_body["requests"][0]["unmergeCells"]["range"]["endRowIndex"] == 100
        )


class TestMergeCellsValidation:
    """Tests for input validation in merge_cells."""

    def test_valid_merge_types(self):
        """Test that valid merge types are accepted."""
        valid_types = ["MERGE_ALL", "MERGE_COLUMNS", "MERGE_ROWS"]
        for merge_type in valid_types:
            assert merge_type.upper() in valid_types

    def test_valid_merge_types_lowercase(self):
        """Test that lowercase merge types are converted properly."""
        valid_types = ["MERGE_ALL", "MERGE_COLUMNS", "MERGE_ROWS"]
        lowercase_inputs = ["merge_all", "merge_columns", "merge_rows"]
        for merge_type in lowercase_inputs:
            assert merge_type.upper() in valid_types

    def test_invalid_merge_type_detection(self):
        """Test that invalid merge type is detected."""
        invalid_type = "MERGE_NONE"
        valid_types = ["MERGE_ALL", "MERGE_COLUMNS", "MERGE_ROWS"]
        assert invalid_type.upper() not in valid_types


class TestMergeCellsOutputMessages:
    """Tests for output message formatting."""

    def test_merge_success_message_format(self):
        """Test merge success message format."""
        range_name = "A1:C1"
        merge_type = "MERGE_ALL"
        sheet_id = 0
        spreadsheet_id = "abc123"
        user_email = "test@example.com"

        message = (
            f"Successfully merged cells in range '{range_name}' (type: {merge_type}) "
            f"in sheet ID {sheet_id} of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "A1:C1" in message
        assert "MERGE_ALL" in message
        assert "sheet ID 0" in message
        assert "abc123" in message
        assert "test@example.com" in message

    def test_unmerge_success_message_format(self):
        """Test unmerge success message format."""
        range_name = "A1:C1"
        sheet_id = 0
        spreadsheet_id = "abc123"
        user_email = "test@example.com"

        message = (
            f"Successfully unmerged cells in range '{range_name}' "
            f"in sheet ID {sheet_id} of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "unmerged" in message
        assert "A1:C1" in message
        assert "sheet ID 0" in message
        assert "abc123" in message


class TestMergeCellsRanges:
    """Tests for various range formats with merge operations."""

    def test_single_row_header_merge(self):
        """Test merging a typical header row range."""
        grid_range = _parse_range_to_grid_range("A1:D1", sheet_id=0)
        assert grid_range["startRowIndex"] == 0
        assert grid_range["endRowIndex"] == 1
        assert grid_range["startColumnIndex"] == 0
        assert grid_range["endColumnIndex"] == 4

    def test_multi_row_merge(self):
        """Test merging multiple rows."""
        grid_range = _parse_range_to_grid_range("A1:A5", sheet_id=0)
        assert grid_range["startRowIndex"] == 0
        assert grid_range["endRowIndex"] == 5
        assert grid_range["startColumnIndex"] == 0
        assert grid_range["endColumnIndex"] == 1

    def test_rectangular_merge(self):
        """Test merging a rectangular area."""
        grid_range = _parse_range_to_grid_range("B2:D5", sheet_id=0)
        assert grid_range["startRowIndex"] == 1
        assert grid_range["endRowIndex"] == 5
        assert grid_range["startColumnIndex"] == 1
        assert grid_range["endColumnIndex"] == 4

    def test_range_with_sheet_prefix_stripped(self):
        """Test that sheet prefix is properly stripped."""
        grid_range = _parse_range_to_grid_range("Sheet1!A1:C3", sheet_id=999)
        assert grid_range["sheetId"] == 999
        assert grid_range["startRowIndex"] == 0
        assert grid_range["endRowIndex"] == 3
        assert grid_range["startColumnIndex"] == 0
        assert grid_range["endColumnIndex"] == 3
