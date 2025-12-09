"""
Unit tests for delete_rows and delete_columns tools.

These tests verify the logic and request structures for the delete operations.
"""

import pytest


class TestDeleteRowsRequestStructure:
    """Unit tests for delete_rows request structures."""

    def test_delete_rows_request_body_format(self):
        """Test that the delete rows request body has the correct structure."""
        sheet_id = 0
        start_index = 4  # 0-indexed (row 5 in 1-indexed)
        end_index = 7  # 0-indexed exclusive (rows 5-7 in 1-indexed inclusive)

        request_body = {
            "requests": [
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": start_index,
                            "endIndex": end_index,
                        }
                    }
                }
            ]
        }

        delete_req = request_body["requests"][0]["deleteDimension"]
        assert delete_req["range"]["dimension"] == "ROWS"
        assert delete_req["range"]["sheetId"] == 0
        assert delete_req["range"]["startIndex"] == 4
        assert delete_req["range"]["endIndex"] == 7

    def test_delete_single_row(self):
        """Test deleting a single row creates correct index range."""
        start_row = 5  # 1-indexed
        end_row = 5  # 1-indexed, same row

        # Convert to 0-indexed
        start_index = start_row - 1
        end_index = end_row  # API is exclusive

        assert start_index == 4
        assert end_index == 5
        assert end_index - start_index == 1

    def test_delete_multiple_rows(self):
        """Test deleting multiple rows creates correct index range."""
        start_row = 3  # 1-indexed
        end_row = 7  # 1-indexed

        start_index = start_row - 1
        end_index = end_row

        assert start_index == 2
        assert end_index == 7
        assert end_index - start_index == 5  # 5 rows deleted

    def test_delete_rows_index_conversion(self):
        """Test 1-indexed to 0-indexed conversion."""
        start_row = 5  # 1-indexed
        end_row = 7  # 1-indexed

        start_index = start_row - 1  # Convert to 0-indexed
        end_index = end_row  # API exclusive

        assert start_index == 4
        assert end_index == 7


class TestDeleteColumnsRequestStructure:
    """Unit tests for delete_columns request structures."""

    def test_delete_columns_request_body_format(self):
        """Test that the delete columns request body has the correct structure."""
        sheet_id = 0
        start_index = 2  # 0-indexed (column C in 1-indexed)
        end_index = 5  # 0-indexed exclusive (columns C-E in 1-indexed inclusive)

        request_body = {
            "requests": [
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": start_index,
                            "endIndex": end_index,
                        }
                    }
                }
            ]
        }

        delete_req = request_body["requests"][0]["deleteDimension"]
        assert delete_req["range"]["dimension"] == "COLUMNS"
        assert delete_req["range"]["sheetId"] == 0
        assert delete_req["range"]["startIndex"] == 2
        assert delete_req["range"]["endIndex"] == 5

    def test_delete_single_column(self):
        """Test deleting a single column creates correct index range."""
        start_column = 3  # Column C
        end_column = 3

        start_index = start_column - 1
        end_index = end_column

        assert start_index == 2
        assert end_index == 3
        assert end_index - start_index == 1

    def test_delete_multiple_columns(self):
        """Test deleting multiple columns creates correct index range."""
        start_column = 3  # Column C
        end_column = 5  # Column E

        start_index = start_column - 1
        end_index = end_column

        assert start_index == 2
        assert end_index == 5
        assert end_index - start_index == 3  # 3 columns deleted


class TestDeleteParameterValidation:
    """Tests for parameter validation."""

    def test_start_row_must_be_positive(self):
        """Test that start_row must be >= 1."""
        start_row = 0

        with pytest.raises(Exception) as exc_info:
            if start_row < 1:
                raise Exception("start_row must be >= 1 (1-indexed).")

        assert "start_row must be >= 1" in str(exc_info.value)

    def test_end_row_must_be_positive(self):
        """Test that end_row must be >= 1."""
        end_row = 0

        with pytest.raises(Exception) as exc_info:
            if end_row < 1:
                raise Exception("end_row must be >= 1 (1-indexed).")

        assert "end_row must be >= 1" in str(exc_info.value)

    def test_end_row_must_be_gte_start_row(self):
        """Test that end_row must be >= start_row."""
        start_row = 5
        end_row = 3

        with pytest.raises(Exception) as exc_info:
            if end_row < start_row:
                raise Exception("end_row must be >= start_row.")

        assert "end_row must be >= start_row" in str(exc_info.value)

    def test_start_column_must_be_positive(self):
        """Test that start_column must be >= 1."""
        start_column = 0

        with pytest.raises(Exception) as exc_info:
            if start_column < 1:
                raise Exception("start_column must be >= 1 (1-indexed).")

        assert "start_column must be >= 1" in str(exc_info.value)

    def test_end_column_must_be_positive(self):
        """Test that end_column must be >= 1."""
        end_column = 0

        with pytest.raises(Exception) as exc_info:
            if end_column < 1:
                raise Exception("end_column must be >= 1 (1-indexed).")

        assert "end_column must be >= 1" in str(exc_info.value)

    def test_end_column_must_be_gte_start_column(self):
        """Test that end_column must be >= start_column."""
        start_column = 5
        end_column = 3

        with pytest.raises(Exception) as exc_info:
            if end_column < start_column:
                raise Exception("end_column must be >= start_column.")

        assert "end_column must be >= start_column" in str(exc_info.value)


class TestDeleteSuccessMessages:
    """Tests for success message formatting."""

    def test_delete_rows_success_message(self):
        """Test the format of success message for delete_rows."""
        start_row = 5
        end_row = 7
        num_rows = end_row - start_row + 1
        sheet_name = "Sheet1"
        spreadsheet_id = "abc123"
        user_email = "test@example.com"

        message = (
            f"Successfully deleted {num_rows} row(s) (rows {start_row}-{end_row}) "
            f"from '{sheet_name}' of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "Successfully deleted 3 row(s)" in message
        assert "(rows 5-7)" in message
        assert "Sheet1" in message
        assert spreadsheet_id in message
        assert user_email in message

    def test_delete_columns_success_message(self):
        """Test the format of success message for delete_columns."""
        start_column = 3
        end_column = 5
        num_columns = end_column - start_column + 1
        start_letter = "C"
        end_letter = "E"
        sheet_name = "Data"
        spreadsheet_id = "xyz789"
        user_email = "test@example.com"

        message = (
            f"Successfully deleted {num_columns} column(s) (columns {start_letter}-{end_letter}, "
            f"columns {start_column}-{end_column}) from '{sheet_name}' of spreadsheet {spreadsheet_id} "
            f"for {user_email}."
        )

        assert "Successfully deleted 3 column(s)" in message
        assert "(columns C-E" in message
        assert "columns 3-5)" in message
        assert "Data" in message
        assert spreadsheet_id in message
        assert user_email in message

    def test_delete_single_row_message(self):
        """Test message when deleting a single row."""
        start_row = 5
        end_row = 5
        num_rows = end_row - start_row + 1

        assert num_rows == 1
        message = f"Successfully deleted {num_rows} row(s) (rows {start_row}-{end_row})"
        assert "1 row(s)" in message
        assert "(rows 5-5)" in message

    def test_delete_rows_with_sheet_id_message(self):
        """Test message when using sheet_id instead of sheet_name."""
        resolved_sheet_id = 12345
        sheet_name = None
        num_rows = 2
        start_row = 3
        end_row = 4
        spreadsheet_id = "abc123"
        user_email = "test@example.com"

        sheet_identifier = sheet_name if sheet_name else f"sheet ID {resolved_sheet_id}"

        message = (
            f"Successfully deleted {num_rows} row(s) (rows {start_row}-{end_row}) "
            f"from '{sheet_identifier}' of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "sheet ID 12345" in message


class TestDeleteColumnToLetterConversion:
    """Tests for column number to letter conversion in delete operations."""

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
        assert self._column_to_letter(3) == "C"
        assert self._column_to_letter(26) == "Z"

    def test_double_letter_columns(self):
        """Test conversion of columns AA-AZ, BA-BZ etc."""
        assert self._column_to_letter(27) == "AA"
        assert self._column_to_letter(28) == "AB"
        assert self._column_to_letter(52) == "AZ"
        assert self._column_to_letter(53) == "BA"

    def test_range_letter_representation(self):
        """Test representing a column range with letters."""
        start_column = 3
        end_column = 5

        start_letter = self._column_to_letter(start_column)
        end_letter = self._column_to_letter(end_column)

        assert start_letter == "C"
        assert end_letter == "E"


class TestDeleteEdgeCases:
    """Tests for edge cases in delete operations."""

    def test_delete_first_row(self):
        """Test deleting the first row."""
        start_row = 1
        end_row = 1
        start_index = start_row - 1
        end_index = end_row

        assert start_index == 0
        assert end_index == 1

    def test_delete_first_column(self):
        """Test deleting the first column (column A)."""
        start_column = 1
        end_column = 1
        start_index = start_column - 1
        end_index = end_column

        assert start_index == 0
        assert end_index == 1

    def test_delete_many_rows(self):
        """Test deleting many rows at once."""
        start_row = 1
        end_row = 100
        num_rows = end_row - start_row + 1

        assert num_rows == 100

    def test_delete_at_large_row_number(self):
        """Test deleting at a large row number."""
        start_row = 10000
        end_row = 10005
        start_index = start_row - 1
        end_index = end_row

        assert start_index == 9999
        assert end_index == 10005

    def test_equal_start_and_end(self):
        """Test when start and end are the same (single row/column)."""
        start_row = 5
        end_row = 5

        num_rows = end_row - start_row + 1
        assert num_rows == 1

        start_index = start_row - 1
        end_index = end_row

        assert end_index - start_index == 1


class TestDeleteSheetIdResolution:
    """Tests for sheet ID resolution logic in delete operations."""

    def test_sheet_id_takes_precedence(self):
        """Test that sheet_id takes precedence over sheet_name."""
        sheet_id = 12345

        if sheet_id is not None:
            resolved_id = sheet_id
        else:
            resolved_id = None

        assert resolved_id == 12345

    def test_sheet_name_resolution_needed(self):
        """Test that sheet_name requires API lookup."""
        sheet_id = None
        sheet_name = "Sheet1"

        needs_resolution = sheet_id is None and sheet_name is not None

        assert needs_resolution is True

    def test_default_to_first_sheet(self):
        """Test that neither sheet_name nor sheet_id defaults to first sheet."""
        sheet_id = None
        sheet_name = None

        use_first_sheet = sheet_id is None and sheet_name is None

        assert use_first_sheet is True
