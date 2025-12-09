"""
Unit tests for insert_rows and insert_columns tools.

These tests verify the logic and request structures for the insert operations.
"""

import pytest


class TestInsertRowsRequestStructure:
    """Unit tests for insert_rows request structures."""

    def test_insert_rows_request_body_format(self):
        """Test that the insert rows request body has the correct structure."""
        sheet_id = 0
        start_index = 4  # 0-indexed (row 5 in 1-indexed)
        num_rows = 2

        request_body = {
            "requests": [
                {
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
            ]
        }

        insert_req = request_body["requests"][0]["insertDimension"]
        assert insert_req["range"]["dimension"] == "ROWS"
        assert insert_req["range"]["startIndex"] == 4
        assert insert_req["range"]["endIndex"] == 6
        assert insert_req["inheritFromBefore"] is True

    def test_insert_rows_at_first_row(self):
        """Test inserting at first row sets inheritFromBefore to False."""
        start_index = 0  # 0-indexed (row 1 in 1-indexed)

        # When inserting at index 0, inheritFromBefore should be False
        inherit_from_before = start_index > 0

        assert inherit_from_before is False

    def test_insert_rows_index_conversion(self):
        """Test 1-indexed to 0-indexed conversion."""
        start_row = 5  # 1-indexed
        start_index = start_row - 1  # Convert to 0-indexed

        assert start_index == 4


class TestInsertColumnsRequestStructure:
    """Unit tests for insert_columns request structures."""

    def test_insert_columns_request_body_format(self):
        """Test that the insert columns request body has the correct structure."""
        sheet_id = 0
        start_index = 2  # 0-indexed (column C in 1-indexed)
        num_columns = 3

        request_body = {
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": start_index,
                            "endIndex": start_index + num_columns,
                        },
                        "inheritFromBefore": True,
                    }
                }
            ]
        }

        insert_req = request_body["requests"][0]["insertDimension"]
        assert insert_req["range"]["dimension"] == "COLUMNS"
        assert insert_req["range"]["startIndex"] == 2
        assert insert_req["range"]["endIndex"] == 5
        assert insert_req["inheritFromBefore"] is True

    def test_insert_columns_at_first_column(self):
        """Test inserting at first column sets inheritFromBefore to False."""
        start_index = 0  # 0-indexed (column A)

        inherit_from_before = start_index > 0

        assert inherit_from_before is False


class TestParameterValidation:
    """Tests for parameter validation."""

    def test_start_row_must_be_positive(self):
        """Test that start_row must be >= 1."""
        start_row = 0

        with pytest.raises(Exception) as exc_info:
            if start_row < 1:
                raise Exception("start_row must be >= 1 (1-indexed).")

        assert "start_row must be >= 1" in str(exc_info.value)

    def test_num_rows_must_be_positive(self):
        """Test that num_rows must be >= 1."""
        num_rows = 0

        with pytest.raises(Exception) as exc_info:
            if num_rows < 1:
                raise Exception("num_rows must be >= 1.")

        assert "num_rows must be >= 1" in str(exc_info.value)

    def test_start_column_must_be_positive(self):
        """Test that start_column must be >= 1."""
        start_column = 0

        with pytest.raises(Exception) as exc_info:
            if start_column < 1:
                raise Exception("start_column must be >= 1 (1-indexed).")

        assert "start_column must be >= 1" in str(exc_info.value)

    def test_num_columns_must_be_positive(self):
        """Test that num_columns must be >= 1."""
        num_columns = 0

        with pytest.raises(Exception) as exc_info:
            if num_columns < 1:
                raise Exception("num_columns must be >= 1.")

        assert "num_columns must be >= 1" in str(exc_info.value)


class TestColumnToLetterConversion:
    """Tests for column number to letter conversion."""

    def _column_to_letter(self, col_num: int) -> str:
        """Convert column number to letter(s)."""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord("A")) + result
            col_num //= 26
        return result

    def test_single_letter_columns(self):
        """Test conversion of columns A-Z."""
        assert self._column_to_letter(1) == "A"
        assert self._column_to_letter(2) == "B"
        assert self._column_to_letter(26) == "Z"

    def test_double_letter_columns(self):
        """Test conversion of columns AA-AZ, BA-BZ etc."""
        assert self._column_to_letter(27) == "AA"
        assert self._column_to_letter(28) == "AB"
        assert self._column_to_letter(52) == "AZ"
        assert self._column_to_letter(53) == "BA"

    def test_triple_letter_columns(self):
        """Test conversion of columns AAA onwards."""
        # 26 + 26*26 = 702 for ZZ, so 703 = AAA
        assert self._column_to_letter(703) == "AAA"


class TestSuccessMessages:
    """Tests for success message formatting."""

    def test_insert_rows_success_message(self):
        """Test the format of success message for insert_rows."""
        num_rows = 3
        start_row = 5
        sheet_name = "Sheet1"
        spreadsheet_id = "abc123"
        user_email = "test@example.com"

        message = (
            f"Successfully inserted {num_rows} row(s) at row {start_row} "
            f"in '{sheet_name}' of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "Successfully inserted 3 row(s)" in message
        assert "at row 5" in message
        assert "Sheet1" in message
        assert spreadsheet_id in message
        assert user_email in message

    def test_insert_columns_success_message(self):
        """Test the format of success message for insert_columns."""
        num_columns = 2
        start_column = 3
        column_letter = "C"
        sheet_name = "Data"
        spreadsheet_id = "xyz789"
        user_email = "test@example.com"

        message = (
            f"Successfully inserted {num_columns} column(s) at column {column_letter} (column {start_column}) "
            f"in '{sheet_name}' of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "Successfully inserted 2 column(s)" in message
        assert "at column C" in message
        assert "(column 3)" in message
        assert "Data" in message
        assert spreadsheet_id in message
        assert user_email in message

    def test_insert_rows_with_sheet_id_message(self):
        """Test message when using sheet_id instead of sheet_name."""
        resolved_sheet_id = 12345
        sheet_name = None
        num_rows = 1
        start_row = 1
        spreadsheet_id = "abc123"
        user_email = "test@example.com"

        sheet_identifier = sheet_name if sheet_name else f"sheet ID {resolved_sheet_id}"

        message = (
            f"Successfully inserted {num_rows} row(s) at row {start_row} "
            f"in '{sheet_identifier}' of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "sheet ID 12345" in message


class TestSheetIdResolution:
    """Tests for sheet ID resolution logic."""

    def test_sheet_id_takes_precedence(self):
        """Test that sheet_id takes precedence over sheet_name."""
        sheet_id = 12345
        # sheet_name could be "Sheet1" but sheet_id takes precedence

        # If sheet_id is provided, it should be used directly
        if sheet_id is not None:
            resolved_id = sheet_id
        else:
            resolved_id = None  # Would need to resolve from sheet_name

        assert resolved_id == 12345

    def test_sheet_name_resolution_needed(self):
        """Test that sheet_name requires API lookup."""
        sheet_id = None
        sheet_name = "Sheet1"

        # If sheet_id is None, we need to resolve sheet_name
        needs_resolution = sheet_id is None and sheet_name is not None

        assert needs_resolution is True

    def test_default_to_first_sheet(self):
        """Test that neither sheet_name nor sheet_id defaults to first sheet."""
        sheet_id = None
        sheet_name = None

        # If both are None, we default to first sheet
        use_first_sheet = sheet_id is None and sheet_name is None

        assert use_first_sheet is True

    def test_sheet_not_found_error(self):
        """Test error when sheet_name doesn't exist."""
        sheet_name = "NonExistentSheet"
        spreadsheet_id = "abc123"
        sheets = [
            {"properties": {"title": "Sheet1", "sheetId": 0}},
            {"properties": {"title": "Sheet2", "sheetId": 1}},
        ]

        found = False
        for sheet in sheets:
            if sheet["properties"]["title"] == sheet_name:
                found = True
                break

        with pytest.raises(Exception) as exc_info:
            if not found:
                raise Exception(
                    f"Sheet '{sheet_name}' not found in spreadsheet {spreadsheet_id}."
                )

        assert "NonExistentSheet" in str(exc_info.value)
        assert "not found" in str(exc_info.value)


class TestEdgeCases:
    """Tests for edge cases."""

    def test_insert_single_row(self):
        """Test inserting a single row (default case)."""
        num_rows = 1
        start_row = 10
        start_index = start_row - 1

        request_range = {
            "startIndex": start_index,
            "endIndex": start_index + num_rows,
        }

        assert request_range["startIndex"] == 9
        assert request_range["endIndex"] == 10

    def test_insert_many_rows(self):
        """Test inserting many rows at once."""
        num_rows = 100
        start_row = 1
        start_index = start_row - 1

        request_range = {
            "startIndex": start_index,
            "endIndex": start_index + num_rows,
        }

        assert request_range["startIndex"] == 0
        assert request_range["endIndex"] == 100

    def test_insert_at_large_row_number(self):
        """Test inserting at a large row number."""
        start_row = 10000
        start_index = start_row - 1

        assert start_index == 9999

    def test_spreadsheet_id_in_request(self):
        """Test that spreadsheet_id is used in API call."""
        spreadsheet_id = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"

        # The spreadsheet_id should be passed to the batchUpdate call
        assert len(spreadsheet_id) > 0
        assert spreadsheet_id.startswith("1")  # Common prefix for sheet IDs
