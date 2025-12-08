"""
Unit tests for batch_update_sheet tool in Google Sheets.

These tests verify the logic and validation for batch update operations,
including request body structure and operation building functions.
"""

import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from gsheets.sheets_tools import (
    _build_format_cells_request,
    _build_freeze_panes_request,
    _build_set_column_width_request,
    _build_set_row_height_request,
    _build_merge_cells_request,
    _build_unmerge_cells_request,
    _build_sort_range_request,
)


class TestBuildFormatCellsRequest:
    """Tests for _build_format_cells_request helper function."""

    def test_basic_bold_formatting(self):
        """Test building a bold formatting request."""
        op = {"type": "format_cells", "sheet_id": 0, "range": "A1:C1", "bold": True}
        request, summary = _build_format_cells_request(op, 0)

        assert "repeatCell" in request
        assert request["repeatCell"]["range"]["sheetId"] == 0
        assert (
            request["repeatCell"]["cell"]["userEnteredFormat"]["textFormat"]["bold"]
            is True
        )
        assert "userEnteredFormat.textFormat.bold" in request["repeatCell"]["fields"]
        assert "bold=True" in summary

    def test_background_color_formatting(self):
        """Test building a background color request."""
        op = {
            "type": "format_cells",
            "sheet_id": 0,
            "range": "A1:D1",
            "background_color": "#E0E0E0",
        }
        request, summary = _build_format_cells_request(op, 0)

        bg_color = request["repeatCell"]["cell"]["userEnteredFormat"]["backgroundColor"]
        assert bg_color["red"] == pytest.approx(224 / 255.0)
        assert bg_color["green"] == pytest.approx(224 / 255.0)
        assert bg_color["blue"] == pytest.approx(224 / 255.0)
        assert "backgroundColor=#E0E0E0" in summary

    def test_multiple_formatting_options(self):
        """Test building a request with multiple formatting options."""
        op = {
            "type": "format_cells",
            "sheet_id": 5,
            "range": "B2:F10",
            "bold": True,
            "font_size": 14,
            "horizontal_alignment": "CENTER",
            "background_color": "#FFFF00",
        }
        request, summary = _build_format_cells_request(op, 0)

        cell = request["repeatCell"]["cell"]["userEnteredFormat"]
        assert cell["textFormat"]["bold"] is True
        assert cell["textFormat"]["fontSize"] == 14
        assert cell["horizontalAlignment"] == "CENTER"
        assert "backgroundColor" in cell

    def test_number_format(self):
        """Test building a number format request."""
        op = {
            "type": "format_cells",
            "sheet_id": 0,
            "range": "B2:B100",
            "number_format_type": "CURRENCY",
            "number_format_pattern": "$#,##0.00",
        }
        request, summary = _build_format_cells_request(op, 0)

        num_format = request["repeatCell"]["cell"]["userEnteredFormat"]["numberFormat"]
        assert num_format["type"] == "CURRENCY"
        assert num_format["pattern"] == "$#,##0.00"
        assert "numType=CURRENCY" in summary

    def test_missing_sheet_id_raises_error(self):
        """Test that missing sheet_id raises ValueError."""
        op = {"type": "format_cells", "range": "A1:C1", "bold": True}
        with pytest.raises(ValueError) as excinfo:
            _build_format_cells_request(op, 0)
        assert "sheet_id" in str(excinfo.value)

    def test_missing_range_raises_error(self):
        """Test that missing range raises ValueError."""
        op = {"type": "format_cells", "sheet_id": 0, "bold": True}
        with pytest.raises(ValueError) as excinfo:
            _build_format_cells_request(op, 0)
        assert "range" in str(excinfo.value)

    def test_no_formatting_options_raises_error(self):
        """Test that no formatting options raises ValueError."""
        op = {"type": "format_cells", "sheet_id": 0, "range": "A1:C1"}
        with pytest.raises(ValueError) as excinfo:
            _build_format_cells_request(op, 0)
        assert "at least one formatting option" in str(excinfo.value)

    def test_invalid_horizontal_alignment_raises_error(self):
        """Test that invalid horizontal alignment raises ValueError."""
        op = {
            "type": "format_cells",
            "sheet_id": 0,
            "range": "A1:C1",
            "horizontal_alignment": "JUSTIFY",
        }
        with pytest.raises(ValueError) as excinfo:
            _build_format_cells_request(op, 0)
        assert "horizontal_alignment" in str(excinfo.value)


class TestBuildFreezePanesRequest:
    """Tests for _build_freeze_panes_request helper function."""

    def test_freeze_one_row(self):
        """Test freezing one row."""
        op = {"type": "freeze_panes", "sheet_id": 0, "frozen_row_count": 1}
        request, summary = _build_freeze_panes_request(op, 0)

        assert "updateSheetProperties" in request
        props = request["updateSheetProperties"]["properties"]
        assert props["sheetId"] == 0
        assert props["gridProperties"]["frozenRowCount"] == 1
        assert props["gridProperties"]["frozenColumnCount"] == 0
        assert "1 row(s)" in summary

    def test_freeze_rows_and_columns(self):
        """Test freezing both rows and columns."""
        op = {
            "type": "freeze_panes",
            "sheet_id": 5,
            "frozen_row_count": 2,
            "frozen_column_count": 1,
        }
        request, summary = _build_freeze_panes_request(op, 0)

        props = request["updateSheetProperties"]["properties"]
        assert props["gridProperties"]["frozenRowCount"] == 2
        assert props["gridProperties"]["frozenColumnCount"] == 1
        assert "2 row(s)" in summary
        assert "1 column(s)" in summary

    def test_missing_sheet_id_raises_error(self):
        """Test that missing sheet_id raises ValueError."""
        op = {"type": "freeze_panes", "frozen_row_count": 1}
        with pytest.raises(ValueError) as excinfo:
            _build_freeze_panes_request(op, 0)
        assert "sheet_id" in str(excinfo.value)

    def test_negative_frozen_row_count_raises_error(self):
        """Test that negative frozen_row_count raises ValueError."""
        op = {"type": "freeze_panes", "sheet_id": 0, "frozen_row_count": -1}
        with pytest.raises(ValueError) as excinfo:
            _build_freeze_panes_request(op, 0)
        assert "frozen_row_count" in str(excinfo.value)


class TestBuildSetColumnWidthRequest:
    """Tests for _build_set_column_width_request helper function."""

    def test_set_column_width(self):
        """Test setting column width."""
        op = {
            "type": "set_column_width",
            "sheet_id": 0,
            "start_index": 0,
            "end_index": 4,
            "width": 150,
        }
        request, summary = _build_set_column_width_request(op, 0)

        assert "updateDimensionProperties" in request
        dim = request["updateDimensionProperties"]
        assert dim["range"]["sheetId"] == 0
        assert dim["range"]["dimension"] == "COLUMNS"
        assert dim["range"]["startIndex"] == 0
        assert dim["range"]["endIndex"] == 4
        assert dim["properties"]["pixelSize"] == 150
        assert "150px" in summary

    def test_missing_width_raises_error(self):
        """Test that missing width raises ValueError."""
        op = {
            "type": "set_column_width",
            "sheet_id": 0,
            "start_index": 0,
            "end_index": 4,
        }
        with pytest.raises(ValueError) as excinfo:
            _build_set_column_width_request(op, 0)
        assert "width" in str(excinfo.value)

    def test_invalid_end_index_raises_error(self):
        """Test that end_index <= start_index raises ValueError."""
        op = {
            "type": "set_column_width",
            "sheet_id": 0,
            "start_index": 5,
            "end_index": 3,
            "width": 100,
        }
        with pytest.raises(ValueError) as excinfo:
            _build_set_column_width_request(op, 0)
        assert "end_index" in str(excinfo.value)

    def test_zero_width_raises_error(self):
        """Test that width <= 0 raises ValueError."""
        op = {
            "type": "set_column_width",
            "sheet_id": 0,
            "start_index": 0,
            "end_index": 4,
            "width": 0,
        }
        with pytest.raises(ValueError) as excinfo:
            _build_set_column_width_request(op, 0)
        assert "width" in str(excinfo.value)


class TestBuildSetRowHeightRequest:
    """Tests for _build_set_row_height_request helper function."""

    def test_set_row_height(self):
        """Test setting row height."""
        op = {
            "type": "set_row_height",
            "sheet_id": 0,
            "start_index": 0,
            "end_index": 1,
            "height": 40,
        }
        request, summary = _build_set_row_height_request(op, 0)

        assert "updateDimensionProperties" in request
        dim = request["updateDimensionProperties"]
        assert dim["range"]["dimension"] == "ROWS"
        assert dim["properties"]["pixelSize"] == 40
        assert "40px" in summary

    def test_missing_height_raises_error(self):
        """Test that missing height raises ValueError."""
        op = {
            "type": "set_row_height",
            "sheet_id": 0,
            "start_index": 0,
            "end_index": 1,
        }
        with pytest.raises(ValueError) as excinfo:
            _build_set_row_height_request(op, 0)
        assert "height" in str(excinfo.value)


class TestBuildMergeCellsRequest:
    """Tests for _build_merge_cells_request helper function."""

    def test_merge_all_cells(self):
        """Test merging all cells in a range."""
        op = {"type": "merge_cells", "sheet_id": 0, "range": "A1:C1"}
        request, summary = _build_merge_cells_request(op, 0)

        assert "mergeCells" in request
        assert request["mergeCells"]["mergeType"] == "MERGE_ALL"
        assert "MERGE_ALL" in summary

    def test_merge_columns(self):
        """Test merging cells by columns."""
        op = {
            "type": "merge_cells",
            "sheet_id": 0,
            "range": "A1:D2",
            "merge_type": "MERGE_COLUMNS",
        }
        request, summary = _build_merge_cells_request(op, 0)

        assert request["mergeCells"]["mergeType"] == "MERGE_COLUMNS"

    def test_invalid_merge_type_raises_error(self):
        """Test that invalid merge_type raises ValueError."""
        op = {
            "type": "merge_cells",
            "sheet_id": 0,
            "range": "A1:C1",
            "merge_type": "INVALID",
        }
        with pytest.raises(ValueError) as excinfo:
            _build_merge_cells_request(op, 0)
        assert "merge_type" in str(excinfo.value)


class TestBuildUnmergeCellsRequest:
    """Tests for _build_unmerge_cells_request helper function."""

    def test_unmerge_cells(self):
        """Test unmerging cells."""
        op = {"type": "unmerge_cells", "sheet_id": 0, "range": "A1:C1"}
        request, summary = _build_unmerge_cells_request(op, 0)

        assert "unmergeCells" in request
        assert "range" in request["unmergeCells"]
        assert "unmerge_cells" in summary

    def test_missing_range_raises_error(self):
        """Test that missing range raises ValueError."""
        op = {"type": "unmerge_cells", "sheet_id": 0}
        with pytest.raises(ValueError) as excinfo:
            _build_unmerge_cells_request(op, 0)
        assert "range" in str(excinfo.value)


class TestBuildSortRangeRequest:
    """Tests for _build_sort_range_request helper function."""

    def test_single_column_sort(self):
        """Test sorting by a single column."""
        op = {
            "type": "sort_range",
            "sheet_id": 0,
            "range": "A1:D100",
            "sort_specs": [{"column_index": 0, "order": "ASCENDING"}],
        }
        request, summary = _build_sort_range_request(op, 0)

        assert "sortRange" in request
        sort_specs = request["sortRange"]["sortSpecs"]
        assert len(sort_specs) == 1
        assert sort_specs[0]["dimensionIndex"] == 0
        assert sort_specs[0]["sortOrder"] == "ASCENDING"
        assert "col0 ASCENDING" in summary

    def test_multi_column_sort(self):
        """Test sorting by multiple columns."""
        op = {
            "type": "sort_range",
            "sheet_id": 0,
            "range": "A1:D100",
            "sort_specs": [
                {"column_index": 0, "order": "ASCENDING"},
                {"column_index": 2, "order": "DESCENDING"},
            ],
        }
        request, summary = _build_sort_range_request(op, 0)

        sort_specs = request["sortRange"]["sortSpecs"]
        assert len(sort_specs) == 2
        assert sort_specs[1]["dimensionIndex"] == 2
        assert sort_specs[1]["sortOrder"] == "DESCENDING"

    def test_missing_sort_specs_raises_error(self):
        """Test that missing sort_specs raises ValueError."""
        op = {"type": "sort_range", "sheet_id": 0, "range": "A1:D100"}
        with pytest.raises(ValueError) as excinfo:
            _build_sort_range_request(op, 0)
        assert "sort_specs" in str(excinfo.value)

    def test_invalid_sort_order_raises_error(self):
        """Test that invalid sort order raises ValueError."""
        op = {
            "type": "sort_range",
            "sheet_id": 0,
            "range": "A1:D100",
            "sort_specs": [{"column_index": 0, "order": "INVALID"}],
        }
        with pytest.raises(ValueError) as excinfo:
            _build_sort_range_request(op, 0)
        assert "order" in str(excinfo.value)

    def test_negative_column_index_raises_error(self):
        """Test that negative column_index raises ValueError."""
        op = {
            "type": "sort_range",
            "sheet_id": 0,
            "range": "A1:D100",
            "sort_specs": [{"column_index": -1, "order": "ASCENDING"}],
        }
        with pytest.raises(ValueError) as excinfo:
            _build_sort_range_request(op, 0)
        assert "column_index" in str(excinfo.value)


class TestBatchUpdateRequestBodyStructure:
    """Tests for the combined request body structure."""

    def test_multiple_operations_structure(self):
        """Test that multiple operations create proper combined request."""
        ops = [
            {"type": "format_cells", "sheet_id": 0, "range": "A1:D1", "bold": True},
            {"type": "freeze_panes", "sheet_id": 0, "frozen_row_count": 1},
            {
                "type": "set_column_width",
                "sheet_id": 0,
                "start_index": 0,
                "end_index": 4,
                "width": 150,
            },
        ]

        requests = []
        for i, op in enumerate(ops):
            if op["type"] == "format_cells":
                request, _ = _build_format_cells_request(op, i)
            elif op["type"] == "freeze_panes":
                request, _ = _build_freeze_panes_request(op, i)
            elif op["type"] == "set_column_width":
                request, _ = _build_set_column_width_request(op, i)
            requests.append(request)

        request_body = {"requests": requests}

        assert len(request_body["requests"]) == 3
        assert "repeatCell" in request_body["requests"][0]
        assert "updateSheetProperties" in request_body["requests"][1]
        assert "updateDimensionProperties" in request_body["requests"][2]
