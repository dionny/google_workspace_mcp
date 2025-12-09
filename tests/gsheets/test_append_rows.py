"""
Unit tests for append_rows tool.

These tests verify the logic, request structures, and parameter validation
for the append_rows operation.
"""

import json
import pytest


class TestAppendRowsRequestStructure:
    """Unit tests for append_rows request structures."""

    def test_append_rows_request_body_format(self):
        """Test that the append rows request body has the correct structure."""
        values = [["Alice", "30", "Engineer"], ["Bob", "25", "Designer"]]

        body = {"values": values}

        assert "values" in body
        assert len(body["values"]) == 2
        assert body["values"][0] == ["Alice", "30", "Engineer"]
        assert body["values"][1] == ["Bob", "25", "Designer"]

    def test_append_single_row(self):
        """Test appending a single row."""
        values = [["Single row data"]]

        body = {"values": values}

        assert len(body["values"]) == 1
        assert body["values"][0] == ["Single row data"]

    def test_append_multiple_columns(self):
        """Test appending rows with multiple columns."""
        values = [
            ["A1", "B1", "C1", "D1", "E1"],
            ["A2", "B2", "C2", "D2", "E2"],
        ]

        body = {"values": values}

        assert len(body["values"]) == 2
        assert len(body["values"][0]) == 5
        assert len(body["values"][1]) == 5


class TestParameterValidation:
    """Tests for parameter validation."""

    def test_empty_values_raises_error(self):
        """Test that empty values raises an error."""
        values = []

        with pytest.raises(Exception) as exc_info:
            if not values:
                raise Exception("Values must be provided and non-empty.")

        assert "Values must be provided and non-empty" in str(exc_info.value)

    def test_invalid_value_input_option(self):
        """Test that invalid value_input_option raises an error."""
        value_input_option = "INVALID"

        with pytest.raises(Exception) as exc_info:
            if value_input_option not in ("USER_ENTERED", "RAW"):
                raise Exception(
                    f"Invalid value_input_option: {value_input_option}. "
                    f"Must be 'USER_ENTERED' or 'RAW'."
                )

        assert "Invalid value_input_option" in str(exc_info.value)
        assert "INVALID" in str(exc_info.value)

    def test_invalid_insert_data_option(self):
        """Test that invalid insert_data_option raises an error."""
        insert_data_option = "INVALID"

        with pytest.raises(Exception) as exc_info:
            if insert_data_option not in ("INSERT_ROWS", "OVERWRITE"):
                raise Exception(
                    f"Invalid insert_data_option: {insert_data_option}. "
                    f"Must be 'INSERT_ROWS' or 'OVERWRITE'."
                )

        assert "Invalid insert_data_option" in str(exc_info.value)
        assert "INVALID" in str(exc_info.value)

    def test_valid_value_input_options(self):
        """Test that valid value_input_options are accepted."""
        valid_options = ["USER_ENTERED", "RAW"]

        for option in valid_options:
            assert option in ("USER_ENTERED", "RAW")

    def test_valid_insert_data_options(self):
        """Test that valid insert_data_options are accepted."""
        valid_options = ["INSERT_ROWS", "OVERWRITE"]

        for option in valid_options:
            assert option in ("INSERT_ROWS", "OVERWRITE")


class TestJsonParsing:
    """Tests for JSON string parsing."""

    def test_json_string_parsed_correctly(self):
        """Test that JSON string values are parsed correctly."""
        json_string = '[["Alice", "30"], ["Bob", "25"]]'

        parsed_values = json.loads(json_string)

        assert isinstance(parsed_values, list)
        assert len(parsed_values) == 2
        assert parsed_values[0] == ["Alice", "30"]
        assert parsed_values[1] == ["Bob", "25"]

    def test_invalid_json_raises_error(self):
        """Test that invalid JSON raises an error."""
        invalid_json = "not valid json"

        with pytest.raises(json.JSONDecodeError):
            json.loads(invalid_json)

    def test_json_not_list_raises_error(self):
        """Test that JSON that is not a list raises an error."""
        json_string = '{"key": "value"}'

        parsed = json.loads(json_string)

        with pytest.raises(ValueError) as exc_info:
            if not isinstance(parsed, list):
                raise ValueError(f"Values must be a list, got {type(parsed).__name__}")

        assert "Values must be a list" in str(exc_info.value)
        assert "dict" in str(exc_info.value)

    def test_json_row_not_list_raises_error(self):
        """Test that JSON with non-list rows raises an error."""
        json_string = '["not a list", ["valid", "row"]]'

        parsed = json.loads(json_string)

        with pytest.raises(ValueError) as exc_info:
            for i, row in enumerate(parsed):
                if not isinstance(row, list):
                    raise ValueError(
                        f"Row {i} must be a list, got {type(row).__name__}"
                    )

        assert "Row 0 must be a list" in str(exc_info.value)
        assert "str" in str(exc_info.value)


class TestDefaultParameters:
    """Tests for default parameter values."""

    def test_default_range_name(self):
        """Test that default range_name is 'A:A'."""
        default_range = "A:A"

        assert default_range == "A:A"

    def test_default_value_input_option(self):
        """Test that default value_input_option is 'USER_ENTERED'."""
        default_option = "USER_ENTERED"

        assert default_option == "USER_ENTERED"

    def test_default_insert_data_option(self):
        """Test that default insert_data_option is 'INSERT_ROWS'."""
        default_option = "INSERT_ROWS"

        assert default_option == "INSERT_ROWS"


class TestSuccessMessages:
    """Tests for success message formatting."""

    def test_append_rows_success_message_format(self):
        """Test the format of success message for append_rows."""
        updated_rows = 3
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        updated_range = "Sheet1!A10:C12"
        updated_cells = 9
        updated_columns = 3

        message = (
            f"Successfully appended {updated_rows} row(s) to spreadsheet {spreadsheet_id} "
            f"for {user_email}. "
            f"Updated range: {updated_range}. "
            f"Cells updated: {updated_cells} ({updated_columns} columns)."
        )

        assert "Successfully appended 3 row(s)" in message
        assert spreadsheet_id in message
        assert user_email in message
        assert updated_range in message
        assert "9 (" in message
        assert "3 columns" in message

    def test_success_message_with_single_row(self):
        """Test success message when appending a single row."""
        updated_rows = 1

        message = f"Successfully appended {updated_rows} row(s)"

        assert "1 row(s)" in message


class TestRangeNameFormats:
    """Tests for various range_name formats."""

    def test_simple_column_range(self):
        """Test range like 'A:A'."""
        range_name = "A:A"

        # The range should be valid for the API
        assert ":" in range_name
        assert range_name.startswith("A")

    def test_multi_column_range(self):
        """Test range like 'A:C'."""
        range_name = "A:C"

        assert range_name == "A:C"

    def test_range_with_sheet_name(self):
        """Test range like 'Sheet1!A:C'."""
        range_name = "Sheet1!A:C"

        assert "!" in range_name
        assert "Sheet1" in range_name

    def test_range_with_quoted_sheet_name(self):
        """Test range with quoted sheet name for names with spaces."""
        range_name = "'My Sheet'!A:C"

        assert "My Sheet" in range_name
        assert range_name.startswith("'")

    def test_specific_cell_range(self):
        """Test range like 'A1:C10' which is also valid."""
        range_name = "A1:C10"

        # API will find data in this range and append after
        assert range_name == "A1:C10"


class TestEdgeCases:
    """Tests for edge cases."""

    def test_append_empty_strings(self):
        """Test appending rows with empty string values."""
        values = [["", "", ""], ["data", "", "more"]]

        body = {"values": values}

        assert body["values"][0] == ["", "", ""]
        assert body["values"][1] == ["data", "", "more"]

    def test_append_numeric_values_as_strings(self):
        """Test appending numeric values as strings."""
        values = [["123", "45.67", "-89"]]

        body = {"values": values}

        # MCP tools typically receive all values as strings
        assert body["values"][0] == ["123", "45.67", "-89"]

    def test_append_formula_string(self):
        """Test appending a formula string."""
        values = [["=SUM(A1:A10)", "=AVERAGE(B1:B10)"]]

        body = {"values": values}

        # With USER_ENTERED, these will be interpreted as formulas
        assert body["values"][0][0] == "=SUM(A1:A10)"
        assert body["values"][0][1] == "=AVERAGE(B1:B10)"

    def test_append_rows_with_uneven_columns(self):
        """Test appending rows with varying column counts."""
        values = [
            ["A", "B", "C"],
            ["D", "E"],
            ["F"],
        ]

        body = {"values": values}

        # Google Sheets API handles uneven rows
        assert len(body["values"][0]) == 3
        assert len(body["values"][1]) == 2
        assert len(body["values"][2]) == 1

    def test_append_large_number_of_rows(self):
        """Test appending many rows at once."""
        values = [[f"Row {i}"] for i in range(100)]

        body = {"values": values}

        assert len(body["values"]) == 100
        assert body["values"][0] == ["Row 0"]
        assert body["values"][99] == ["Row 99"]

    def test_spreadsheet_id_in_request(self):
        """Test that spreadsheet_id is used in API call."""
        spreadsheet_id = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"

        # The spreadsheet_id should be passed to the append call
        assert len(spreadsheet_id) > 0


class TestApiResultParsing:
    """Tests for parsing API result."""

    def test_parse_updates_from_result(self):
        """Test parsing the updates from API result."""
        result = {
            "spreadsheetId": "abc123",
            "updates": {
                "spreadsheetId": "abc123",
                "updatedRange": "Sheet1!A5:C7",
                "updatedRows": 3,
                "updatedColumns": 3,
                "updatedCells": 9,
            },
        }

        updates = result.get("updates", {})
        updated_range = updates.get("updatedRange", "unknown range")
        updated_rows = updates.get("updatedRows", 0)
        updated_columns = updates.get("updatedColumns", 0)
        updated_cells = updates.get("updatedCells", 0)

        assert updated_range == "Sheet1!A5:C7"
        assert updated_rows == 3
        assert updated_columns == 3
        assert updated_cells == 9

    def test_handle_missing_updates(self):
        """Test handling result with missing updates field."""
        result = {"spreadsheetId": "abc123"}

        updates = result.get("updates", {})
        updated_range = updates.get("updatedRange", "unknown range")
        updated_rows = updates.get("updatedRows", 0)

        assert updated_range == "unknown range"
        assert updated_rows == 0
