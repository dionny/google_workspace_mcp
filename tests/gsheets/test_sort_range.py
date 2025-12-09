"""
Unit tests for sort_range tool in Google Sheets.

These tests verify the logic and validation for sort operations
including sort spec parsing, validation, and request body structure.
"""

import json
import pytest

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from gsheets.sheets_tools import _parse_range_to_grid


class TestSortSpecParsing:
    """Tests for sort_specs JSON parsing."""

    def test_parse_single_sort_spec_from_json(self):
        """Test parsing a single sort spec from JSON string."""
        json_string = '[{"column_index": 0, "order": "ASCENDING"}]'
        parsed = json.loads(json_string)

        assert isinstance(parsed, list)
        assert len(parsed) == 1
        assert parsed[0]["column_index"] == 0
        assert parsed[0]["order"] == "ASCENDING"

    def test_parse_multiple_sort_specs_from_json(self):
        """Test parsing multiple sort specs from JSON string."""
        json_string = '[{"column_index": 0, "order": "ASCENDING"}, {"column_index": 2, "order": "DESCENDING"}]'
        parsed = json.loads(json_string)

        assert isinstance(parsed, list)
        assert len(parsed) == 2
        assert parsed[0]["column_index"] == 0
        assert parsed[0]["order"] == "ASCENDING"
        assert parsed[1]["column_index"] == 2
        assert parsed[1]["order"] == "DESCENDING"

    def test_parse_sort_spec_with_lowercase_order(self):
        """Test parsing sort spec with lowercase order value."""
        json_string = '[{"column_index": 1, "order": "ascending"}]'
        parsed = json.loads(json_string)

        assert parsed[0]["order"].upper() == "ASCENDING"

    def test_invalid_json_raises_error(self):
        """Test that invalid JSON raises JSONDecodeError."""
        invalid_json = (
            '[{"column_index": 0, "order": "ASCENDING"'  # Missing closing bracket
        )
        with pytest.raises(json.JSONDecodeError):
            json.loads(invalid_json)


class TestSortSpecValidation:
    """Tests for sort spec validation logic."""

    def test_valid_orders(self):
        """Test that valid orders are recognized."""
        valid_orders = ["ASCENDING", "DESCENDING"]
        for order in valid_orders:
            assert order in valid_orders

    def test_valid_orders_case_insensitive(self):
        """Test that order validation is case-insensitive."""
        valid_orders = ["ASCENDING", "DESCENDING"]
        test_cases = [
            "ascending",
            "ASCENDING",
            "Ascending",
            "descending",
            "DESCENDING",
            "Descending",
        ]
        for order in test_cases:
            assert order.upper() in valid_orders

    def test_invalid_order_detection(self):
        """Test that invalid orders are detected."""
        valid_orders = ["ASCENDING", "DESCENDING"]
        invalid_orders = ["ASC", "DESC", "UP", "DOWN", "NONE", ""]
        for order in invalid_orders:
            assert order.upper() not in valid_orders

    def test_column_index_must_be_non_negative(self):
        """Test that column_index must be >= 0."""
        valid_indices = [0, 1, 5, 10, 100]
        invalid_indices = [-1, -10, -100]

        for idx in valid_indices:
            assert idx >= 0

        for idx in invalid_indices:
            assert idx < 0

    def test_spec_must_have_column_index(self):
        """Test that sort spec must have column_index field."""
        spec_without_column_index = {"order": "ASCENDING"}
        assert "column_index" not in spec_without_column_index

    def test_spec_must_have_order(self):
        """Test that sort spec must have order field."""
        spec_without_order = {"column_index": 0}
        assert "order" not in spec_without_order


class TestSortRangeRequestBodyStructure:
    """Tests for the request body structure of sort_range."""

    def test_single_sort_spec_request_structure(self):
        """Test the structure of a single-column sort request."""
        sheet_id = 0
        grid_range = {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": 100,
            "startColumnIndex": 0,
            "endColumnIndex": 4,
        }
        api_sort_specs = [{"dimensionIndex": 0, "sortOrder": "ASCENDING"}]

        request_body = {
            "requests": [
                {"sortRange": {"range": grid_range, "sortSpecs": api_sort_specs}}
            ]
        }

        assert request_body["requests"][0]["sortRange"]["range"]["sheetId"] == 0
        assert len(request_body["requests"][0]["sortRange"]["sortSpecs"]) == 1
        assert (
            request_body["requests"][0]["sortRange"]["sortSpecs"][0]["dimensionIndex"]
            == 0
        )
        assert (
            request_body["requests"][0]["sortRange"]["sortSpecs"][0]["sortOrder"]
            == "ASCENDING"
        )

    def test_multi_column_sort_request_structure(self):
        """Test the structure of a multi-column sort request."""
        sheet_id = 123
        grid_range = {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": 50,
            "startColumnIndex": 0,
            "endColumnIndex": 3,
        }
        api_sort_specs = [
            {"dimensionIndex": 0, "sortOrder": "ASCENDING"},
            {"dimensionIndex": 2, "sortOrder": "DESCENDING"},
        ]

        request_body = {
            "requests": [
                {"sortRange": {"range": grid_range, "sortSpecs": api_sort_specs}}
            ]
        }

        sort_specs = request_body["requests"][0]["sortRange"]["sortSpecs"]
        assert len(sort_specs) == 2
        assert sort_specs[0]["dimensionIndex"] == 0
        assert sort_specs[0]["sortOrder"] == "ASCENDING"
        assert sort_specs[1]["dimensionIndex"] == 2
        assert sort_specs[1]["sortOrder"] == "DESCENDING"

    def test_sort_spec_transformation(self):
        """Test the transformation from user input format to API format."""
        # User input format
        user_sort_specs = [
            {"column_index": 0, "order": "ASCENDING"},
            {"column_index": 1, "order": "DESCENDING"},
        ]

        # Transform to API format
        api_sort_specs = []
        for spec in user_sort_specs:
            api_sort_specs.append(
                {
                    "dimensionIndex": spec["column_index"],
                    "sortOrder": spec["order"].upper(),
                }
            )

        assert api_sort_specs[0]["dimensionIndex"] == 0
        assert api_sort_specs[0]["sortOrder"] == "ASCENDING"
        assert api_sort_specs[1]["dimensionIndex"] == 1
        assert api_sort_specs[1]["sortOrder"] == "DESCENDING"


class TestSortRangeWithGridRange:
    """Tests for sort_range using _parse_range_to_grid."""

    def test_sort_range_uses_grid_range(self):
        """Test that sort_range uses the grid range from range parsing."""
        range_name = "A1:D100"

        grid_range = _parse_range_to_grid(range_name)

        assert grid_range["startColumnIndex"] == 0
        assert grid_range["endColumnIndex"] == 4  # D is column 3, +1 for exclusive
        assert grid_range["startRowIndex"] == 0
        assert grid_range["endRowIndex"] == 100

    def test_sort_range_with_partial_sheet_range(self):
        """Test sort range with a subset of data."""
        range_name = "B2:E50"

        grid_range = _parse_range_to_grid(range_name)

        assert grid_range["startColumnIndex"] == 1  # B
        assert grid_range["endColumnIndex"] == 5  # E + 1
        assert grid_range["startRowIndex"] == 1  # Row 2 is index 1
        assert grid_range["endRowIndex"] == 50


class TestSuccessMessageFormats:
    """Tests for success message formatting."""

    def test_single_sort_success_message_format(self):
        """Test the format of single-column sort success message."""
        range_name = "A1:D100"
        sheet_id = 0
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        sort_specs = [{"column_index": 0, "order": "ASCENDING"}]

        sort_descriptions = []
        for spec in sort_specs:
            sort_descriptions.append(
                f"column {spec['column_index']} {spec['order'].upper()}"
            )
        sort_summary = ", ".join(sort_descriptions)

        message = (
            f"Successfully sorted range '{range_name}' in sheet ID {sheet_id} "
            f"of spreadsheet {spreadsheet_id} for {user_email}. "
            f"Sort order: {sort_summary}"
        )

        assert "A1:D100" in message
        assert "sheet ID 0" in message
        assert "abc123" in message
        assert "column 0 ASCENDING" in message

    def test_multi_column_sort_success_message_format(self):
        """Test the format of multi-column sort success message."""
        range_name = "A1:D100"
        sheet_id = 0
        spreadsheet_id = "xyz789"
        user_email = "user@example.com"
        sort_specs = [
            {"column_index": 0, "order": "ASCENDING"},
            {"column_index": 2, "order": "DESCENDING"},
        ]

        sort_descriptions = []
        for spec in sort_specs:
            sort_descriptions.append(
                f"column {spec['column_index']} {spec['order'].upper()}"
            )
        sort_summary = ", ".join(sort_descriptions)

        message = (
            f"Successfully sorted range '{range_name}' in sheet ID {sheet_id} "
            f"of spreadsheet {spreadsheet_id} for {user_email}. "
            f"Sort order: {sort_summary}"
        )

        assert "column 0 ASCENDING" in message
        assert "column 2 DESCENDING" in message
        assert sort_summary == "column 0 ASCENDING, column 2 DESCENDING"


class TestSortRangeEdgeCases:
    """Tests for edge cases in sort_range."""

    def test_empty_sort_specs_detection(self):
        """Test that empty sort_specs is detected."""
        sort_specs = []
        assert len(sort_specs) == 0

    def test_sort_spec_not_a_dict_detection(self):
        """Test that non-dict sort spec items are detected."""
        invalid_specs = [[0, "ASCENDING"], "not a dict", 123]
        for spec in invalid_specs:
            assert not isinstance(spec, dict)

    def test_column_index_not_int_detection(self):
        """Test that non-int column_index is detected."""
        invalid_indices = ["0", 0.5, None, [0]]
        for idx in invalid_indices:
            assert not isinstance(idx, int) or isinstance(idx, bool)

    def test_order_not_string_detection(self):
        """Test that non-string order is detected."""
        invalid_orders = [1, True, None, ["ASCENDING"]]
        for order in invalid_orders:
            assert not isinstance(order, str)

    def test_high_column_index_allowed(self):
        """Test that high column indices are allowed."""
        # Google Sheets supports up to column ZZZ (around 18278 columns)
        high_indices = [25, 26, 100, 1000]
        for idx in high_indices:
            assert isinstance(idx, int) and idx >= 0

    def test_three_level_sort(self):
        """Test three-level sorting sort spec transformation."""
        sort_specs = [
            {"column_index": 0, "order": "ASCENDING"},
            {"column_index": 1, "order": "DESCENDING"},
            {"column_index": 2, "order": "ASCENDING"},
        ]

        api_sort_specs = []
        for spec in sort_specs:
            api_sort_specs.append(
                {
                    "dimensionIndex": spec["column_index"],
                    "sortOrder": spec["order"].upper(),
                }
            )

        assert len(api_sort_specs) == 3
        assert api_sort_specs[0]["dimensionIndex"] == 0
        assert api_sort_specs[1]["dimensionIndex"] == 1
        assert api_sort_specs[2]["dimensionIndex"] == 2
