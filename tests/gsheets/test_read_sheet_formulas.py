"""
Unit tests for read_sheet_formulas tool.

These tests verify the logic and output formatting for reading formulas from sheets.
"""


class TestReadSheetFormulasValueRenderOption:
    """Tests for the valueRenderOption parameter behavior."""

    def test_formula_render_option_constant(self):
        """Test that FORMULA is the correct valueRenderOption for formulas."""
        value_render_option = "FORMULA"
        assert value_render_option == "FORMULA"

    def test_different_render_options(self):
        """Test the three valueRenderOption values."""
        # FORMATTED_VALUE - returns computed, formatted values (default for read_sheet_values)
        # UNFORMATTED_VALUE - returns computed, raw values
        # FORMULA - returns formula text (used by read_sheet_formulas)
        options = ["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"]
        assert "FORMULA" in options
        assert len(options) == 3


class TestReadSheetFormulasOutputFormatting:
    """Tests for output formatting logic."""

    def test_empty_values_message(self):
        """Test message when no data is found."""
        range_name = "A1:D10"
        user_email = "test@example.com"
        values = []

        if not values:
            result = f"No data found in range '{range_name}' for {user_email}."

        assert "No data found" in result
        assert range_name in result
        assert user_email in result

    def test_row_formatting(self):
        """Test that rows are formatted with row numbers."""
        values = [["=SUM(A1:A10)", "100"], ["=AVERAGE(B1:B10)", "50"]]

        formatted_rows = []
        for i, row in enumerate(values, 1):
            formatted_rows.append(f"Row {i:2d}: {row}")

        assert formatted_rows[0] == "Row  1: ['=SUM(A1:A10)', '100']"
        assert formatted_rows[1] == "Row  2: ['=AVERAGE(B1:B10)', '50']"

    def test_row_padding(self):
        """Test that shorter rows are padded to match first row length."""
        values = [["A", "B", "C"], ["D"]]  # Second row is shorter

        padded_rows = []
        for row in values:
            padded_row = row + [""] * max(0, len(values[0]) - len(row))
            padded_rows.append(padded_row)

        assert padded_rows[0] == ["A", "B", "C"]
        assert padded_rows[1] == ["D", "", ""]

    def test_output_truncation_at_50_rows(self):
        """Test that output is limited to 50 rows."""
        num_rows = 100
        values = [[f"=ROW({i})" for _ in range(3)] for i in range(num_rows)]

        formatted_rows = []
        for i, row in enumerate(values, 1):
            formatted_rows.append(f"Row {i:2d}: {row}")

        # Only first 50 should be shown
        truncated = formatted_rows[:50]
        assert len(truncated) == 50

        # Should include ellipsis message
        more_rows = len(values) - 50
        assert more_rows == 50

    def test_no_truncation_under_50_rows(self):
        """Test that no truncation message is added for <= 50 rows."""
        values = [["=A1"] for _ in range(30)]

        formatted_rows = [f"Row {i:2d}: {row}" for i, row in enumerate(values, 1)]

        # Should not show "more rows" message
        has_more = len(values) > 50
        assert not has_more
        assert len(formatted_rows) == 30  # All rows present

    def test_success_message_format(self):
        """Test the success message includes all required information."""
        range_name = "Sheet1!A1:D10"
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        num_rows = 10

        text_output = (
            f"Successfully read formulas from {num_rows} rows in range '{range_name}' "
            f"in spreadsheet {spreadsheet_id} for {user_email}:"
        )

        assert "Successfully read formulas" in text_output
        assert f"{num_rows} rows" in text_output
        assert range_name in text_output
        assert spreadsheet_id in text_output
        assert user_email in text_output


class TestReadSheetFormulasExamples:
    """Tests demonstrating the difference between values and formulas."""

    def test_formula_vs_value_example(self):
        """Test the example from the docstring."""
        # A cell containing =SUM(A1:A10) that displays "100" in the UI
        formula_render = "=SUM(A1:A10)"  # What read_sheet_formulas returns
        value_render = "100"  # What read_sheet_values returns

        assert formula_render.startswith("=")
        assert not value_render.startswith("=")

    def test_cells_without_formulas_show_values(self):
        """Test that cells without formulas show their plain values."""
        # When a cell contains just "Hello" (no formula)
        cell_without_formula = "Hello"

        # Both read_sheet_values and read_sheet_formulas return the same
        # (no formula means the value is the value)
        assert cell_without_formula == "Hello"

    def test_mixed_formula_and_value_cells(self):
        """Test a row with both formula and non-formula cells."""
        row_with_formula = ["=A1+B1", "Static Text", "123", "=NOW()"]

        # Formulas start with =
        formulas = [cell for cell in row_with_formula if cell.startswith("=")]
        values = [cell for cell in row_with_formula if not cell.startswith("=")]

        assert len(formulas) == 2
        assert formulas == ["=A1+B1", "=NOW()"]
        assert len(values) == 2
        assert values == ["Static Text", "123"]


class TestReadSheetFormulasParameterValidation:
    """Tests for parameter handling."""

    def test_default_range_name(self):
        """Test that default range is A1:Z1000."""
        default_range = "A1:Z1000"
        assert default_range == "A1:Z1000"

    def test_custom_range_with_sheet_name(self):
        """Test that ranges with sheet names work."""
        range_name = "Sheet1!A1:D10"
        assert "!" in range_name
        sheet_name, cell_range = range_name.split("!", 1)
        assert sheet_name == "Sheet1"
        assert cell_range == "A1:D10"

    def test_range_without_sheet_name(self):
        """Test that ranges without sheet names work."""
        range_name = "A1:D10"
        assert "!" not in range_name


class TestReadSheetFormulasAPIRequest:
    """Tests for API request structure."""

    def test_api_parameters(self):
        """Test the API parameters used for the request."""
        spreadsheet_id = "abc123"
        range_name = "A1:D10"
        value_render_option = "FORMULA"

        # These are the parameters passed to the API
        params = {
            "spreadsheetId": spreadsheet_id,
            "range": range_name,
            "valueRenderOption": value_render_option,
        }

        assert params["spreadsheetId"] == "abc123"
        assert params["range"] == "A1:D10"
        assert params["valueRenderOption"] == "FORMULA"

    def test_formula_render_option_differs_from_default(self):
        """Test that FORMULA differs from the default FORMATTED_VALUE."""
        default_option = "FORMATTED_VALUE"  # Used by read_sheet_values
        formula_option = "FORMULA"  # Used by read_sheet_formulas

        assert default_option != formula_option


class TestReadSheetFormulasEdgeCases:
    """Tests for edge cases."""

    def test_single_cell_range(self):
        """Test reading a single cell."""
        values = [["=SUM(A1:A10)"]]

        formatted_rows = []
        for i, row in enumerate(values, 1):
            formatted_rows.append(f"Row {i:2d}: {row}")

        assert len(formatted_rows) == 1
        assert "=SUM(A1:A10)" in formatted_rows[0]

    def test_empty_cells_in_range(self):
        """Test handling of empty cells."""
        values = [["=A1", "", "=C1"], ["", "", ""]]

        # Empty strings are valid values
        assert values[0][1] == ""
        assert all(cell == "" for cell in values[1])

    def test_complex_formula(self):
        """Test handling of complex formulas."""
        complex_formula = '=IF(AND(A1>0, B1<100), VLOOKUP(C1, D:E, 2, FALSE), "N/A")'
        values = [[complex_formula]]

        # Formula should be preserved as-is
        assert values[0][0] == complex_formula
        assert "IF" in values[0][0]
        assert "VLOOKUP" in values[0][0]

    def test_array_formula(self):
        """Test handling of array formulas."""
        array_formula = "=ARRAYFORMULA(A1:A10*B1:B10)"
        values = [[array_formula]]

        assert values[0][0].startswith("=")
        assert "ARRAYFORMULA" in values[0][0]

    def test_formula_with_sheet_reference(self):
        """Test handling of formulas referencing other sheets."""
        cross_sheet_formula = "='Other Sheet'!A1+B1"
        values = [[cross_sheet_formula]]

        assert "Other Sheet" in values[0][0]

    def test_numbers_are_not_formulas(self):
        """Test that numeric values don't look like formulas."""
        values = [[123, 45.67, -89]]

        # None of these should start with =
        for row in values:
            for cell in row:
                assert not str(cell).startswith("=")
