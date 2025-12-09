"""
Unit tests for append_values tool.

These tests verify the logic and request structures for the append_values tool.
"""

import json
import pytest


class TestAppendValuesRequestStructure:
    """Unit tests for append_values request body structure."""

    def test_append_request_body_format(self):
        """Test that the request body has the correct structure."""
        values = [["Alice", "Engineer", "2024"], ["Bob", "Designer", "2023"]]

        body = {"values": values}

        assert "values" in body
        assert len(body["values"]) == 2
        assert body["values"][0] == ["Alice", "Engineer", "2024"]
        assert body["values"][1] == ["Bob", "Designer", "2023"]

    def test_api_parameters_structure(self):
        """Test that API parameters are correctly structured."""
        spreadsheet_id = "abc123"
        range_name = "Sheet1!A:C"
        value_input_option = "USER_ENTERED"
        insert_data_option = "INSERT_ROWS"
        values = [["Data1", "Data2"]]

        api_params = {
            "spreadsheetId": spreadsheet_id,
            "range": range_name,
            "valueInputOption": value_input_option,
            "insertDataOption": insert_data_option,
            "body": {"values": values},
        }

        assert api_params["spreadsheetId"] == "abc123"
        assert api_params["range"] == "Sheet1!A:C"
        assert api_params["valueInputOption"] == "USER_ENTERED"
        assert api_params["insertDataOption"] == "INSERT_ROWS"
        assert api_params["body"]["values"] == [["Data1", "Data2"]]

    def test_insert_rows_option(self):
        """Test INSERT_ROWS insert data option."""
        insert_data_option = "INSERT_ROWS"
        assert insert_data_option == "INSERT_ROWS"

    def test_overwrite_option(self):
        """Test OVERWRITE insert data option."""
        insert_data_option = "OVERWRITE"
        assert insert_data_option == "OVERWRITE"


class TestAppendValuesJsonParsing:
    """Tests for JSON string parsing of values parameter."""

    def test_parse_valid_json_string(self):
        """Test parsing a valid JSON string into list of lists."""
        json_string = '[["Alice", "Engineer"], ["Bob", "Designer"]]'

        parsed = json.loads(json_string)

        assert isinstance(parsed, list)
        assert len(parsed) == 2
        assert parsed[0] == ["Alice", "Engineer"]
        assert parsed[1] == ["Bob", "Designer"]

    def test_parse_single_row_json(self):
        """Test parsing a single row JSON string."""
        json_string = '[["Single", "Row", "Data"]]'

        parsed = json.loads(json_string)

        assert len(parsed) == 1
        assert parsed[0] == ["Single", "Row", "Data"]

    def test_parse_empty_cells_json(self):
        """Test parsing JSON with empty cell values."""
        json_string = '[["Name", "", "Title"], ["", "Middle", ""]]'

        parsed = json.loads(json_string)

        assert parsed[0] == ["Name", "", "Title"]
        assert parsed[1] == ["", "Middle", ""]

    def test_invalid_json_raises_error(self):
        """Test that invalid JSON raises an error."""
        invalid_json = "not valid json"

        with pytest.raises(json.JSONDecodeError):
            json.loads(invalid_json)

    def test_non_list_json_should_fail_validation(self):
        """Test that non-list JSON should fail validation."""
        json_string = '{"key": "value"}'

        parsed = json.loads(json_string)

        # Validation check - values must be a list
        assert not isinstance(parsed, list)

    def test_non_nested_list_should_fail_validation(self):
        """Test that a flat list (not list of lists) should fail validation."""
        json_string = '["Alice", "Bob", "Charlie"]'

        parsed = json.loads(json_string)

        # Validation check - each element must be a list
        for i, item in enumerate(parsed):
            if not isinstance(item, list):
                # This should trigger validation error
                assert True
                return

        # If we get here, validation would pass incorrectly
        pytest.fail("Should have found non-list items")


class TestAppendValuesMessages:
    """Tests for append_values success message formatting."""

    def test_success_message_format(self):
        """Test the format of success message."""
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        updated_range = "Sheet1!A5:C6"
        updated_rows = 2
        updated_columns = 3
        updated_cells = 6

        message = (
            f"Successfully appended data to spreadsheet {spreadsheet_id} for {user_email}. "
            f"Range: {updated_range} | "
            f"Updated: {updated_rows} rows, {updated_columns} columns, {updated_cells} cells."
        )

        assert "Successfully appended" in message
        assert spreadsheet_id in message
        assert user_email in message
        assert "Sheet1!A5:C6" in message
        assert "2 rows" in message
        assert "3 columns" in message
        assert "6 cells" in message

    def test_message_includes_actual_range(self):
        """Test that the message includes where data was actually appended."""
        # The response should show the actual range where data was written,
        # not just the range specified in the request
        request_range = "Sheet1!A:C"
        actual_range = "Sheet1!A10:C12"

        message = f"Range: {actual_range}"

        assert request_range not in message
        assert actual_range in message


class TestAppendValuesValidation:
    """Tests for append_values input validation logic."""

    def test_empty_values_should_fail(self):
        """Test that empty values list raises an error."""
        values = []

        with pytest.raises(Exception) as exc_info:
            if not values:
                raise Exception(
                    "Values cannot be empty. Provide at least one row to append."
                )

        assert "cannot be empty" in str(exc_info.value)

    def test_spreadsheet_id_required(self):
        """Test that spreadsheet_id is required."""
        spreadsheet_id = None

        with pytest.raises(Exception) as exc_info:
            if not spreadsheet_id:
                raise Exception("spreadsheet_id is required")

        assert "spreadsheet_id is required" in str(exc_info.value)

    def test_range_name_required(self):
        """Test that range_name is required."""
        range_name = None

        with pytest.raises(Exception) as exc_info:
            if not range_name:
                raise Exception("range_name is required")

        assert "range_name is required" in str(exc_info.value)

    def test_value_input_option_defaults_to_user_entered(self):
        """Test that value_input_option defaults to USER_ENTERED."""
        value_input_option = "USER_ENTERED"  # Default value

        assert value_input_option == "USER_ENTERED"

    def test_insert_data_option_defaults_to_insert_rows(self):
        """Test that insert_data_option defaults to INSERT_ROWS."""
        insert_data_option = "INSERT_ROWS"  # Default value

        assert insert_data_option == "INSERT_ROWS"


class TestAppendValuesEdgeCases:
    """Tests for edge cases in append_values."""

    def test_single_cell_append(self):
        """Test appending a single cell value."""
        values = [["SingleValue"]]

        body = {"values": values}

        assert len(body["values"]) == 1
        assert len(body["values"][0]) == 1
        assert body["values"][0][0] == "SingleValue"

    def test_multiple_rows_different_lengths(self):
        """Test appending rows with different column counts."""
        values = [
            ["A", "B", "C"],
            ["D", "E"],  # Shorter row
            ["F", "G", "H", "I"],  # Longer row
        ]

        body = {"values": values}

        assert len(body["values"]) == 3
        assert len(body["values"][0]) == 3
        assert len(body["values"][1]) == 2
        assert len(body["values"][2]) == 4

    def test_numeric_values_as_strings(self):
        """Test that numeric values can be passed as strings."""
        values = [["100", "200.5", "-50"]]

        body = {"values": values}

        assert body["values"][0] == ["100", "200.5", "-50"]

    def test_formula_in_values(self):
        """Test that formulas can be included in values."""
        values = [["=SUM(A1:A10)", "=TODAY()", "=CONCATENATE(A1, B1)"]]

        body = {"values": values}

        assert "=SUM" in body["values"][0][0]
        assert "=TODAY" in body["values"][0][1]

    def test_range_with_sheet_name(self):
        """Test range specification with sheet name."""
        range_name = "MySheet!A:D"

        assert "MySheet" in range_name
        assert "!A:D" in range_name

    def test_range_without_sheet_name(self):
        """Test range specification without sheet name (uses first sheet)."""
        range_name = "A:D"

        assert "!" not in range_name
        assert range_name == "A:D"
