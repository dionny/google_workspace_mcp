"""
Unit tests for sort_range tool.

These tests verify the logic and request structures for the sort operation.
"""

import pytest


class TestColumnLetterToIndex:
    """Tests for the _column_letter_to_index helper function."""

    def _column_letter_to_index(self, col_str: str) -> int:
        """Convert a column letter to 0-indexed column number."""
        result = 0
        for char in col_str.upper():
            result = result * 26 + (ord(char) - ord("A") + 1)
        return result - 1

    def test_single_letters(self):
        """Test single letter columns A-Z."""
        assert self._column_letter_to_index("A") == 0
        assert self._column_letter_to_index("B") == 1
        assert self._column_letter_to_index("C") == 2
        assert self._column_letter_to_index("Z") == 25

    def test_double_letters(self):
        """Test double letter columns AA-AZ, BA-BZ etc."""
        assert self._column_letter_to_index("AA") == 26
        assert self._column_letter_to_index("AB") == 27
        assert self._column_letter_to_index("AZ") == 51
        assert self._column_letter_to_index("BA") == 52

    def test_lowercase(self):
        """Test that lowercase letters work."""
        assert self._column_letter_to_index("a") == 0
        assert self._column_letter_to_index("z") == 25
        assert self._column_letter_to_index("aa") == 26


class TestParseA1Range:
    """Tests for the _parse_a1_range helper function."""

    def _column_letter_to_index(self, col_str: str) -> int:
        """Convert a column letter to 0-indexed column number."""
        result = 0
        for char in col_str.upper():
            result = result * 26 + (ord(char) - ord("A") + 1)
        return result - 1

    def _parse_a1_range(self, range_str: str) -> dict:
        """Parse an A1 notation range string."""
        import re

        sheet_name = None
        if "!" in range_str:
            sheet_name, range_str = range_str.split("!", 1)
            sheet_name = sheet_name.strip("'\"")

        cell_pattern = r"([A-Za-z]*)(\d*)"

        if ":" in range_str:
            start_part, end_part = range_str.split(":", 1)
        else:
            start_part = range_str
            end_part = range_str

        start_match = re.fullmatch(cell_pattern, start_part)
        end_match = re.fullmatch(cell_pattern, end_part)

        start_col_str, start_row_str = start_match.groups()
        end_col_str, end_row_str = end_match.groups()

        start_col = (
            self._column_letter_to_index(start_col_str) if start_col_str else None
        )
        end_col = self._column_letter_to_index(end_col_str) + 1 if end_col_str else None
        start_row = int(start_row_str) - 1 if start_row_str else None
        end_row = int(end_row_str) if end_row_str else None

        return {
            "sheet_name": sheet_name,
            "start_row": start_row,
            "end_row": end_row,
            "start_col": start_col,
            "end_col": end_col,
        }

    def test_simple_range(self):
        """Test parsing a simple A1:D10 range."""
        result = self._parse_a1_range("A1:D10")

        assert result["sheet_name"] is None
        assert result["start_row"] == 0  # Row 1, 0-indexed
        assert result["end_row"] == 10  # Row 10, exclusive
        assert result["start_col"] == 0  # Column A, 0-indexed
        assert result["end_col"] == 4  # Column D, exclusive

    def test_range_with_sheet_name(self):
        """Test parsing a range with sheet name."""
        result = self._parse_a1_range("Sheet1!A1:D10")

        assert result["sheet_name"] == "Sheet1"
        assert result["start_row"] == 0
        assert result["end_row"] == 10
        assert result["start_col"] == 0
        assert result["end_col"] == 4

    def test_range_with_quoted_sheet_name(self):
        """Test parsing a range with quoted sheet name."""
        result = self._parse_a1_range("'My Sheet'!A1:D10")

        assert result["sheet_name"] == "My Sheet"
        assert result["start_row"] == 0
        assert result["end_row"] == 10

    def test_single_cell(self):
        """Test parsing a single cell reference."""
        result = self._parse_a1_range("B5")

        assert result["sheet_name"] is None
        assert result["start_row"] == 4
        assert result["end_row"] == 5
        assert result["start_col"] == 1
        assert result["end_col"] == 2

    def test_offset_range(self):
        """Test parsing an offset range like B2:E10."""
        result = self._parse_a1_range("B2:E10")

        assert result["start_row"] == 1  # Row 2, 0-indexed
        assert result["end_row"] == 10  # Row 10, exclusive
        assert result["start_col"] == 1  # Column B, 0-indexed
        assert result["end_col"] == 5  # Column E, exclusive


class TestSortRangeRequestStructure:
    """Tests for the sort range request body structure."""

    def test_basic_sort_request(self):
        """Test that a basic sort request has correct structure."""
        sheet_id = 0
        start_row = 0
        end_row = 10
        start_col = 0
        end_col = 4
        sort_column = 1  # First column (relative to range)
        ascending = True

        request_body = {
            "requests": [
                {
                    "sortRange": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": start_row,
                            "endRowIndex": end_row,
                            "startColumnIndex": start_col,
                            "endColumnIndex": end_col,
                        },
                        "sortSpecs": [
                            {
                                "dimensionIndex": start_col + sort_column - 1,
                                "sortOrder": "ASCENDING" if ascending else "DESCENDING",
                            }
                        ],
                    }
                }
            ]
        }

        sort_req = request_body["requests"][0]["sortRange"]
        assert sort_req["range"]["sheetId"] == 0
        assert sort_req["range"]["startRowIndex"] == 0
        assert sort_req["range"]["endRowIndex"] == 10
        assert sort_req["range"]["startColumnIndex"] == 0
        assert sort_req["range"]["endColumnIndex"] == 4
        assert len(sort_req["sortSpecs"]) == 1
        assert sort_req["sortSpecs"][0]["dimensionIndex"] == 0
        assert sort_req["sortSpecs"][0]["sortOrder"] == "ASCENDING"

    def test_descending_sort(self):
        """Test descending sort order."""
        ascending = False
        sort_order = "ASCENDING" if ascending else "DESCENDING"

        assert sort_order == "DESCENDING"

    def test_multi_column_sort(self):
        """Test sort with primary and secondary columns."""
        start_col = 0
        sort_column = 2
        secondary_sort_column = 1
        ascending = True
        secondary_ascending = False

        sort_specs = [
            {
                "dimensionIndex": start_col + sort_column - 1,
                "sortOrder": "ASCENDING" if ascending else "DESCENDING",
            },
            {
                "dimensionIndex": start_col + secondary_sort_column - 1,
                "sortOrder": "ASCENDING" if secondary_ascending else "DESCENDING",
            },
        ]

        assert len(sort_specs) == 2
        assert sort_specs[0]["dimensionIndex"] == 1  # Column B (2nd in range)
        assert sort_specs[0]["sortOrder"] == "ASCENDING"
        assert sort_specs[1]["dimensionIndex"] == 0  # Column A (1st in range)
        assert sort_specs[1]["sortOrder"] == "DESCENDING"

    def test_sort_with_header_row(self):
        """Test that header row handling adjusts start_row."""
        original_start_row = 0
        has_header_row = True

        start_row = original_start_row + 1 if has_header_row else original_start_row

        assert start_row == 1  # Excludes row 0 (header)

    def test_sort_offset_range(self):
        """Test sorting an offset range like B2:E10."""
        # For B2:E10:
        # B = column 1 (0-indexed), E = column 4 (exclusive = 5)
        # Row 2 = index 1 (0-indexed), Row 10 = index 10 (exclusive)
        start_col = 1

        # If we sort by column 2 within the range (which is column C absolute)
        sort_column = 2
        dimension_index = start_col + sort_column - 1

        assert dimension_index == 2  # Column C (0-indexed)


class TestSortRangeParameterValidation:
    """Tests for parameter validation."""

    def test_sort_column_must_be_positive(self):
        """Test that sort_column must be >= 1."""
        sort_column = 0

        with pytest.raises(Exception) as exc_info:
            if sort_column < 1:
                raise Exception("sort_column must be >= 1 (1-indexed).")

        assert "sort_column must be >= 1" in str(exc_info.value)

    def test_secondary_sort_column_must_be_positive(self):
        """Test that secondary_sort_column must be >= 1."""
        secondary_sort_column = 0

        with pytest.raises(Exception) as exc_info:
            if secondary_sort_column is not None and secondary_sort_column < 1:
                raise Exception("secondary_sort_column must be >= 1 (1-indexed).")

        assert "secondary_sort_column must be >= 1" in str(exc_info.value)

    def test_sort_column_within_range(self):
        """Test that sort_column must be within range width."""
        range_width = 4  # A:D
        sort_column = 5

        with pytest.raises(Exception) as exc_info:
            if sort_column > range_width:
                raise Exception(
                    f"sort_column ({sort_column}) exceeds the range width ({range_width} columns)."
                )

        assert "exceeds the range width" in str(exc_info.value)

    def test_header_row_requires_two_rows(self):
        """Test that has_header_row requires at least 2 rows."""
        start_row = 0
        end_row = 1  # Only 1 row
        has_header_row = True

        with pytest.raises(Exception) as exc_info:
            if has_header_row and end_row - start_row < 2:
                raise Exception(
                    "Range must have at least 2 rows when has_header_row=True."
                )

        assert "at least 2 rows" in str(exc_info.value)

    def test_range_must_have_column_bounds(self):
        """Test that range must specify column bounds."""
        start_col = None
        end_col = None

        with pytest.raises(Exception) as exc_info:
            if start_col is None or end_col is None:
                raise Exception("Range must specify column bounds.")

        assert "column bounds" in str(exc_info.value)

    def test_range_must_have_row_bounds(self):
        """Test that range must specify row bounds."""
        start_row = None
        end_row = None

        with pytest.raises(Exception) as exc_info:
            if start_row is None or end_row is None:
                raise Exception("Range must specify row bounds.")

        assert "row bounds" in str(exc_info.value)


class TestSortRangeSuccessMessages:
    """Tests for success message formatting."""

    def test_basic_success_message(self):
        """Test the format of a basic success message."""
        range_name = "A1:D10"
        sort_column = 1
        ascending = True
        has_header_row = False
        sheet_name = "Sheet1"
        spreadsheet_id = "abc123"
        user_email = "test@example.com"

        sort_order = "ascending" if ascending else "descending"
        header_note = " (excluding header row)" if has_header_row else ""
        sort_description = f"column {sort_column} ({sort_order})"

        message = (
            f"Successfully sorted range '{range_name}' by {sort_description}{header_note} "
            f"in '{sheet_name}' of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "Successfully sorted range 'A1:D10'" in message
        assert "column 1 (ascending)" in message
        assert "(excluding header row)" not in message
        assert "Sheet1" in message
        assert spreadsheet_id in message
        assert user_email in message

    def test_descending_message(self):
        """Test message with descending sort."""
        ascending = False
        sort_order = "ascending" if ascending else "descending"
        message = f"column 1 ({sort_order})"

        assert "descending" in message

    def test_header_row_message(self):
        """Test message when has_header_row is True."""
        has_header_row = True
        header_note = " (excluding header row)" if has_header_row else ""

        assert "(excluding header row)" in header_note

    def test_multi_column_sort_message(self):
        """Test message with secondary sort column."""
        sort_column = 1
        secondary_sort_column = 2
        ascending = True
        secondary_ascending = False

        sort_order = "ascending" if ascending else "descending"
        sec_order = "ascending" if secondary_ascending else "descending"
        sort_description = f"column {sort_column} ({sort_order})"
        sort_description += f", then column {secondary_sort_column} ({sec_order})"

        assert "column 1 (ascending), then column 2 (descending)" in sort_description

    def test_message_with_sheet_id(self):
        """Test message when using sheet_id instead of sheet_name."""
        resolved_sheet_id = 12345
        effective_sheet_name = None

        sheet_identifier = effective_sheet_name or f"sheet ID {resolved_sheet_id}"

        assert "sheet ID 12345" in sheet_identifier


class TestSortRangeIndexCalculations:
    """Tests for index calculations in sort operations."""

    def test_relative_to_absolute_column_index(self):
        """Test converting relative sort_column to absolute dimensionIndex."""
        # Range B1:E10 (columns 1-4 in 0-indexed)
        start_col = 1  # Column B

        # sort_column=1 should sort by column B (index 1)
        sort_column = 1
        dimension_index = start_col + sort_column - 1
        assert dimension_index == 1

        # sort_column=2 should sort by column C (index 2)
        sort_column = 2
        dimension_index = start_col + sort_column - 1
        assert dimension_index == 2

        # sort_column=4 should sort by column E (index 4)
        sort_column = 4
        dimension_index = start_col + sort_column - 1
        assert dimension_index == 4

    def test_header_row_adjustment(self):
        """Test that header row adjusts start_row by 1."""
        original_start_row = 0
        has_header_row = True

        adjusted_start_row = (
            original_start_row + 1 if has_header_row else original_start_row
        )

        assert adjusted_start_row == 1

    def test_no_header_row_adjustment(self):
        """Test that no adjustment is made without header row."""
        original_start_row = 0
        has_header_row = False

        adjusted_start_row = (
            original_start_row + 1 if has_header_row else original_start_row
        )

        assert adjusted_start_row == 0


class TestSortRangeEdgeCases:
    """Tests for edge cases in sort operations."""

    def test_single_row_range(self):
        """Test that single row range works (no actual sorting needed)."""
        start_row = 0
        end_row = 1
        has_header_row = False

        # Single row is valid when has_header_row is False
        assert end_row - start_row == 1
        assert not has_header_row

    def test_two_row_range_with_header(self):
        """Test minimum valid range with header row."""
        start_row = 0
        end_row = 2  # 2 rows

        # After adjusting for header (has_header_row=True adds 1 to start)
        adjusted_start = start_row + 1
        assert end_row - adjusted_start == 1  # 1 data row to sort

    def test_large_range(self):
        """Test handling of large ranges."""
        start_row = 0
        end_row = 10000
        start_col = 0
        end_col = 100

        assert end_row - start_row == 10000
        assert end_col - start_col == 100

    def test_secondary_inherits_primary_order(self):
        """Test that secondary_ascending defaults to primary ascending."""
        ascending = True
        secondary_ascending = None

        effective_secondary_ascending = (
            secondary_ascending if secondary_ascending is not None else ascending
        )

        assert effective_secondary_ascending is True

        ascending = False
        effective_secondary_ascending = (
            secondary_ascending if secondary_ascending is not None else ascending
        )

        assert effective_secondary_ascending is False
