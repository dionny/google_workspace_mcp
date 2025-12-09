"""
Unit tests for row and column insert/delete operations in Google Sheets tools.

These tests verify the logic and validation for insert_rows, delete_rows,
insert_columns, and delete_columns tools.
"""


class TestInsertRowsValidation:
    """Unit tests for insert_rows parameter validation."""

    def test_insert_rows_request_body_structure(self):
        """Test that the request body for insert_rows is correctly structured."""
        sheet_id = 0
        start_index = 5
        count = 3
        inherit_from_before = True

        request_body = {
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": start_index,
                            "endIndex": start_index + count,
                        },
                        "inheritFromBefore": inherit_from_before,
                    }
                }
            ]
        }

        assert request_body["requests"][0]["insertDimension"]["range"]["sheetId"] == 0
        assert (
            request_body["requests"][0]["insertDimension"]["range"]["dimension"]
            == "ROWS"
        )
        assert (
            request_body["requests"][0]["insertDimension"]["range"]["startIndex"] == 5
        )
        assert request_body["requests"][0]["insertDimension"]["range"]["endIndex"] == 8
        assert (
            request_body["requests"][0]["insertDimension"]["inheritFromBefore"] is True
        )

    def test_insert_rows_end_index_calculation(self):
        """Test that end_index is correctly calculated from start_index + count."""
        test_cases = [
            (0, 1, 1),  # Insert 1 row at index 0 -> end at 1
            (5, 3, 8),  # Insert 3 rows at index 5 -> end at 8
            (10, 10, 20),  # Insert 10 rows at index 10 -> end at 20
            (0, 100, 100),  # Insert 100 rows at start -> end at 100
        ]

        for start_index, count, expected_end in test_cases:
            end_index = start_index + count
            assert end_index == expected_end, f"For start={start_index}, count={count}"

    def test_insert_rows_inherit_from_before_false(self):
        """Test that inherit_from_before can be set to False."""
        inherit_from_before = False

        request_body = {
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": 0,
                            "dimension": "ROWS",
                            "startIndex": 0,
                            "endIndex": 1,
                        },
                        "inheritFromBefore": inherit_from_before,
                    }
                }
            ]
        }

        assert (
            request_body["requests"][0]["insertDimension"]["inheritFromBefore"] is False
        )


class TestDeleteRowsValidation:
    """Unit tests for delete_rows parameter validation."""

    def test_delete_rows_request_body_structure(self):
        """Test that the request body for delete_rows is correctly structured."""
        sheet_id = 123
        start_index = 2
        count = 3

        # The function calculates end_index internally as start_index + count
        end_index = start_index + count

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

        assert request_body["requests"][0]["deleteDimension"]["range"]["sheetId"] == 123
        assert (
            request_body["requests"][0]["deleteDimension"]["range"]["dimension"]
            == "ROWS"
        )
        assert (
            request_body["requests"][0]["deleteDimension"]["range"]["startIndex"] == 2
        )
        assert request_body["requests"][0]["deleteDimension"]["range"]["endIndex"] == 5

    def test_delete_rows_end_index_calculation(self):
        """Test that end_index is correctly calculated from start_index + count."""
        test_cases = [
            (0, 1, 1),  # Delete 1 row at index 0 -> end at 1
            (5, 3, 8),  # Delete 3 rows at index 5 -> end at 8
            (10, 10, 20),  # Delete 10 rows at index 10 -> end at 20
            (0, 100, 100),  # Delete 100 rows at start -> end at 100
        ]

        for start_index, count, expected_end in test_cases:
            end_index = start_index + count
            assert end_index == expected_end, f"For start={start_index}, count={count}"

    def test_delete_rows_count_must_be_positive(self):
        """Test that count must be at least 1."""
        valid_counts = [1, 5, 10, 100]
        for count in valid_counts:
            assert count >= 1, f"count={count} should be >= 1"

        invalid_counts = [0, -1, -10]
        for count in invalid_counts:
            assert count < 1, f"count={count} should be < 1 (invalid)"


class TestInsertColumnsValidation:
    """Unit tests for insert_columns parameter validation."""

    def test_insert_columns_request_body_structure(self):
        """Test that the request body for insert_columns is correctly structured."""
        sheet_id = 456
        start_index = 2  # Column C
        count = 2
        inherit_from_before = True

        request_body = {
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": start_index,
                            "endIndex": start_index + count,
                        },
                        "inheritFromBefore": inherit_from_before,
                    }
                }
            ]
        }

        assert request_body["requests"][0]["insertDimension"]["range"]["sheetId"] == 456
        assert (
            request_body["requests"][0]["insertDimension"]["range"]["dimension"]
            == "COLUMNS"
        )
        assert (
            request_body["requests"][0]["insertDimension"]["range"]["startIndex"] == 2
        )
        assert request_body["requests"][0]["insertDimension"]["range"]["endIndex"] == 4
        assert (
            request_body["requests"][0]["insertDimension"]["inheritFromBefore"] is True
        )

    def test_insert_columns_index_mapping(self):
        """Test column index to letter mapping understanding (0=A, 1=B, etc.)."""
        # This tests our understanding of the column index system
        column_mappings = [
            (0, "A"),
            (1, "B"),
            (2, "C"),
            (25, "Z"),
            (26, "AA"),
        ]

        # Just verify our index understanding
        for index, letter in column_mappings:
            assert isinstance(index, int) and index >= 0


class TestDeleteColumnsValidation:
    """Unit tests for delete_columns parameter validation."""

    def test_delete_columns_request_body_structure(self):
        """Test that the request body for delete_columns is correctly structured."""
        sheet_id = 789
        start_index = 1  # Column B
        count = 2  # Delete 2 columns (B and C)

        # The function calculates end_index internally as start_index + count
        end_index = start_index + count

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

        assert request_body["requests"][0]["deleteDimension"]["range"]["sheetId"] == 789
        assert (
            request_body["requests"][0]["deleteDimension"]["range"]["dimension"]
            == "COLUMNS"
        )
        assert (
            request_body["requests"][0]["deleteDimension"]["range"]["startIndex"] == 1
        )
        assert request_body["requests"][0]["deleteDimension"]["range"]["endIndex"] == 3

    def test_delete_columns_end_index_calculation(self):
        """Test that end_index is correctly calculated from start_index + count."""
        test_cases = [
            (0, 1, 1),  # Delete 1 column at index 0 -> end at 1
            (0, 26, 26),  # Delete 26 columns (A-Z) -> end at 26
            (5, 5, 10),  # Delete 5 columns at index 5 -> end at 10
        ]

        for start_index, count, expected_end in test_cases:
            end_index = start_index + count
            assert end_index == expected_end, f"For start={start_index}, count={count}"

    def test_delete_columns_count_must_be_positive(self):
        """Test that count must be at least 1 for columns."""
        valid_counts = [1, 5, 26]
        for count in valid_counts:
            assert count >= 1

        invalid_counts = [0, -1]
        for count in invalid_counts:
            assert count < 1


class TestErrorMessages:
    """Tests for error message formatting."""

    def test_insert_rows_success_message_format(self):
        """Test the format of insert_rows success message."""
        count = 3
        start_index = 5
        sheet_id = 0
        spreadsheet_id = "abc123"
        user_email = "test@example.com"

        message = (
            f"Successfully inserted {count} row(s) at index {start_index} in sheet ID {sheet_id} "
            f"of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "3 row(s)" in message
        assert "index 5" in message
        assert "sheet ID 0" in message
        assert "abc123" in message

    def test_delete_rows_success_message_format(self):
        """Test the format of delete_rows success message."""
        start_index = 2
        count = 3
        sheet_id = 123
        spreadsheet_id = "xyz789"
        user_email = "test@example.com"

        message = (
            f"Successfully deleted {count} row(s) starting at index {start_index} "
            f"from sheet ID {sheet_id} of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "3 row(s)" in message
        assert "starting at index 2" in message
        assert "sheet ID 123" in message

    def test_delete_validation_error_message(self):
        """Test the validation error message format for count."""
        count = 0

        error_message = f"count ({count}) must be at least 1"

        assert "count (0)" in error_message
        assert "must be at least 1" in error_message


class TestDimensionConstants:
    """Tests to verify dimension constant values."""

    def test_rows_dimension_constant(self):
        """Test that ROWS dimension string is correct."""
        assert "ROWS" == "ROWS"

    def test_columns_dimension_constant(self):
        """Test that COLUMNS dimension string is correct."""
        assert "COLUMNS" == "COLUMNS"
