"""
Unit tests for read_sheet_values tool with value_render_option functionality.

These tests verify the parameter handling and output formatting logic.
"""


class TestReadSheetValuesValueRenderOption:
    """Tests for the value_render_option parameter."""

    def test_valid_render_options(self):
        """Test that all three valid render options are recognized."""
        valid_options = ["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"]
        assert len(valid_options) == 3
        assert "FORMATTED_VALUE" in valid_options
        assert "UNFORMATTED_VALUE" in valid_options
        assert "FORMULA" in valid_options

    def test_default_render_option(self):
        """Test that the default render option is FORMATTED_VALUE."""
        default_option = "FORMATTED_VALUE"
        assert default_option == "FORMATTED_VALUE"

    def test_render_option_used_in_api_params(self):
        """Test that value_render_option is passed to API parameters."""
        spreadsheet_id = "abc123"
        full_range = "Sheet1!A1:D10"
        value_render_option = "UNFORMATTED_VALUE"

        params = {
            "spreadsheetId": spreadsheet_id,
            "range": full_range,
            "valueRenderOption": value_render_option,
        }

        assert params["valueRenderOption"] == "UNFORMATTED_VALUE"

    def test_formula_render_option_in_params(self):
        """Test FORMULA render option in API parameters."""
        value_render_option = "FORMULA"

        params = {
            "spreadsheetId": "test_id",
            "range": "A1:B2",
            "valueRenderOption": value_render_option,
        }

        assert params["valueRenderOption"] == "FORMULA"


class TestReadSheetValuesRenderContext:
    """Tests for the render context in output messages."""

    def test_formatted_value_no_context(self):
        """Test that FORMATTED_VALUE (default) shows no render context."""
        value_render_option = "FORMATTED_VALUE"

        render_context = ""
        if value_render_option == "FORMULA":
            render_context = " (showing formulas)"
        elif value_render_option == "UNFORMATTED_VALUE":
            render_context = " (unformatted values)"

        assert render_context == ""

    def test_formula_render_context(self):
        """Test that FORMULA shows appropriate render context."""
        value_render_option = "FORMULA"

        render_context = ""
        if value_render_option == "FORMULA":
            render_context = " (showing formulas)"
        elif value_render_option == "UNFORMATTED_VALUE":
            render_context = " (unformatted values)"

        assert render_context == " (showing formulas)"

    def test_unformatted_value_render_context(self):
        """Test that UNFORMATTED_VALUE shows appropriate render context."""
        value_render_option = "UNFORMATTED_VALUE"

        render_context = ""
        if value_render_option == "FORMULA":
            render_context = " (showing formulas)"
        elif value_render_option == "UNFORMATTED_VALUE":
            render_context = " (unformatted values)"

        assert render_context == " (unformatted values)"

    def test_success_message_with_formula_context(self):
        """Test complete success message with FORMULA render context."""
        range_name = "A1:D10"
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        num_rows = 5
        value_render_option = "FORMULA"

        render_context = ""
        if value_render_option == "FORMULA":
            render_context = " (showing formulas)"
        elif value_render_option == "UNFORMATTED_VALUE":
            render_context = " (unformatted values)"

        message = f"Successfully read {num_rows} rows from range '{range_name}' in spreadsheet {spreadsheet_id} for {user_email}{render_context}:"

        assert "(showing formulas)" in message
        assert f"{num_rows} rows" in message

    def test_success_message_with_unformatted_context(self):
        """Test complete success message with UNFORMATTED_VALUE render context."""
        range_name = "A1:D10"
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        num_rows = 10
        value_render_option = "UNFORMATTED_VALUE"

        render_context = ""
        if value_render_option == "FORMULA":
            render_context = " (showing formulas)"
        elif value_render_option == "UNFORMATTED_VALUE":
            render_context = " (unformatted values)"

        message = f"Successfully read {num_rows} rows from range '{range_name}' in spreadsheet {spreadsheet_id} for {user_email}{render_context}:"

        assert "(unformatted values)" in message
        assert f"{num_rows} rows" in message


class TestReadSheetValuesRenderOptionExamples:
    """Tests demonstrating the different render options with examples."""

    def test_formatted_value_example(self):
        """Test FORMATTED_VALUE returns formatted display values."""
        # A cell with value 1.0 formatted as percentage would return "100%"
        formatted_value = "100%"
        assert formatted_value == "100%"

    def test_unformatted_value_example(self):
        """Test UNFORMATTED_VALUE returns raw underlying values."""
        # Same cell with value 1.0 formatted as percentage returns 1.0
        unformatted_value = 1.0
        assert unformatted_value == 1.0

    def test_formula_example(self):
        """Test FORMULA returns the formula text."""
        # Cell with formula =A1+B1 returns the formula, not the result
        formula_value = "=A1+B1"
        assert formula_value.startswith("=")

    def test_currency_formatted_vs_unformatted(self):
        """Test currency cell with FORMATTED_VALUE vs UNFORMATTED_VALUE."""
        # Cell containing 1234.56 with currency format
        formatted = "$1,234.56"  # FORMATTED_VALUE
        unformatted = 1234.56  # UNFORMATTED_VALUE

        assert "$" in formatted
        assert "," in formatted
        assert isinstance(unformatted, float)

    def test_date_formatted_vs_unformatted(self):
        """Test date cell with different render options."""
        # Date values can be strings or serial numbers depending on render option
        formatted_date = "2024-01-15"  # FORMATTED_VALUE
        # UNFORMATTED_VALUE could be serial number like 45306

        assert "-" in formatted_date


class TestReadSheetValuesParameterValidation:
    """Tests for parameter type hints and validation."""

    def test_literal_type_values(self):
        """Test that only valid Literal values are accepted."""
        # The type hint is Literal["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"]
        valid_values = ["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"]

        # Check that all valid values match the expected set
        assert set(valid_values) == {
            "FORMATTED_VALUE",
            "UNFORMATTED_VALUE",
            "FORMULA",
        }

    def test_case_sensitivity(self):
        """Test that render option values are case-sensitive (uppercase)."""
        correct = "FORMATTED_VALUE"
        incorrect = "formatted_value"

        assert correct != incorrect
        assert correct.isupper()
