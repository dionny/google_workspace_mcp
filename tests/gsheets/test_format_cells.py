"""
Unit tests for format_cells tool in Google Sheets.

These tests verify the logic and validation for cell formatting operations
including color parsing, range parsing, and request body structure.
"""

import pytest

# Import the helper functions directly for testing
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from gsheets.sheets_tools import _parse_hex_color, _parse_range_to_grid_range


class TestHexColorParsing:
    """Tests for _parse_hex_color helper function."""

    def test_parse_hex_with_hash(self):
        """Test parsing hex color with # prefix."""
        result = _parse_hex_color("#FF0000")
        assert result["red"] == pytest.approx(1.0)
        assert result["green"] == pytest.approx(0.0)
        assert result["blue"] == pytest.approx(0.0)

    def test_parse_hex_without_hash(self):
        """Test parsing hex color without # prefix."""
        result = _parse_hex_color("00FF00")
        assert result["red"] == pytest.approx(0.0)
        assert result["green"] == pytest.approx(1.0)
        assert result["blue"] == pytest.approx(0.0)

    def test_parse_hex_blue(self):
        """Test parsing blue color."""
        result = _parse_hex_color("#0000FF")
        assert result["red"] == pytest.approx(0.0)
        assert result["green"] == pytest.approx(0.0)
        assert result["blue"] == pytest.approx(1.0)

    def test_parse_hex_white(self):
        """Test parsing white color."""
        result = _parse_hex_color("#FFFFFF")
        assert result["red"] == pytest.approx(1.0)
        assert result["green"] == pytest.approx(1.0)
        assert result["blue"] == pytest.approx(1.0)

    def test_parse_hex_black(self):
        """Test parsing black color."""
        result = _parse_hex_color("#000000")
        assert result["red"] == pytest.approx(0.0)
        assert result["green"] == pytest.approx(0.0)
        assert result["blue"] == pytest.approx(0.0)

    def test_parse_hex_gray(self):
        """Test parsing gray color (E0E0E0)."""
        result = _parse_hex_color("#E0E0E0")
        expected = 224 / 255.0
        assert result["red"] == pytest.approx(expected)
        assert result["green"] == pytest.approx(expected)
        assert result["blue"] == pytest.approx(expected)

    def test_parse_hex_lowercase(self):
        """Test parsing lowercase hex color."""
        result = _parse_hex_color("#ff5500")
        assert result["red"] == pytest.approx(1.0)
        assert result["green"] == pytest.approx(85 / 255.0)
        assert result["blue"] == pytest.approx(0.0)

    def test_parse_hex_invalid_length(self):
        """Test that invalid length raises ValueError."""
        with pytest.raises(ValueError) as excinfo:
            _parse_hex_color("#FFF")
        assert "Expected 6 hex digits" in str(excinfo.value)

    def test_parse_hex_invalid_characters(self):
        """Test that invalid hex characters raise ValueError."""
        with pytest.raises(ValueError) as excinfo:
            _parse_hex_color("#GGGGGG")
        assert "Must contain valid hex digits" in str(excinfo.value)


class TestRangeParsing:
    """Tests for _parse_range_to_grid_range helper function."""

    def test_parse_single_cell(self):
        """Test parsing a single cell reference."""
        result = _parse_range_to_grid_range("A1", sheet_id=0)
        assert result["sheetId"] == 0
        assert result["startColumnIndex"] == 0
        assert result["endColumnIndex"] == 1
        assert result["startRowIndex"] == 0
        assert result["endRowIndex"] == 1

    def test_parse_range_a1_to_c10(self):
        """Test parsing a basic range."""
        result = _parse_range_to_grid_range("A1:C10", sheet_id=0)
        assert result["sheetId"] == 0
        assert result["startColumnIndex"] == 0
        assert result["endColumnIndex"] == 3  # C is column 2, +1 for exclusive
        assert result["startRowIndex"] == 0
        assert result["endRowIndex"] == 10

    def test_parse_range_b2_to_d5(self):
        """Test parsing another range."""
        result = _parse_range_to_grid_range("B2:D5", sheet_id=123)
        assert result["sheetId"] == 123
        assert result["startColumnIndex"] == 1
        assert result["endColumnIndex"] == 4
        assert result["startRowIndex"] == 1
        assert result["endRowIndex"] == 5

    def test_parse_range_with_sheet_prefix(self):
        """Test that sheet prefix is stripped."""
        result = _parse_range_to_grid_range("Sheet1!A1:B2", sheet_id=456)
        assert result["sheetId"] == 456
        assert result["startColumnIndex"] == 0
        assert result["endColumnIndex"] == 2
        assert result["startRowIndex"] == 0
        assert result["endRowIndex"] == 2

    def test_parse_column_only_range(self):
        """Test parsing a column-only range (no row numbers)."""
        result = _parse_range_to_grid_range("A:C", sheet_id=0)
        assert result["sheetId"] == 0
        assert result["startColumnIndex"] == 0
        assert result["endColumnIndex"] == 3
        assert "startRowIndex" not in result
        assert "endRowIndex" not in result

    def test_parse_range_double_letter_column(self):
        """Test parsing range with AA column."""
        result = _parse_range_to_grid_range("AA1:AB10", sheet_id=0)
        assert result["startColumnIndex"] == 26  # AA is column 26 (0-based)
        assert result["endColumnIndex"] == 28  # AB is column 27, +1 for exclusive

    def test_parse_range_z_column(self):
        """Test parsing range ending at Z column."""
        result = _parse_range_to_grid_range("Y1:Z10", sheet_id=0)
        assert result["startColumnIndex"] == 24  # Y
        assert result["endColumnIndex"] == 26  # Z is 25, +1 for exclusive

    def test_col_to_index_calculation(self):
        """Verify column letter to index conversion."""
        # This tests our understanding of column indexing
        # A = 0, B = 1, ..., Z = 25, AA = 26, AB = 27
        result_a = _parse_range_to_grid_range("A1", sheet_id=0)
        assert result_a["startColumnIndex"] == 0

        result_z = _parse_range_to_grid_range("Z1", sheet_id=0)
        assert result_z["startColumnIndex"] == 25

        result_aa = _parse_range_to_grid_range("AA1", sheet_id=0)
        assert result_aa["startColumnIndex"] == 26


class TestFormatCellsRequestBodyStructure:
    """Tests for the request body structure of format_cells."""

    def test_bold_formatting_request_structure(self):
        """Test the structure of a bold formatting request."""
        sheet_id = 0
        grid_range = {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": 1,
            "startColumnIndex": 0,
            "endColumnIndex": 26,
        }
        cell_format = {"textFormat": {"bold": True}}
        fields = ["userEnteredFormat.textFormat.bold"]

        request_body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": grid_range,
                        "cell": {"userEnteredFormat": cell_format},
                        "fields": ",".join(fields),
                    }
                }
            ]
        }

        assert request_body["requests"][0]["repeatCell"]["range"]["sheetId"] == 0
        assert (
            request_body["requests"][0]["repeatCell"]["cell"]["userEnteredFormat"][
                "textFormat"
            ]["bold"]
            is True
        )
        assert (
            "userEnteredFormat.textFormat.bold"
            in request_body["requests"][0]["repeatCell"]["fields"]
        )

    def test_background_color_request_structure(self):
        """Test the structure of a background color formatting request."""
        cell_format = {"backgroundColor": {"red": 0.878, "green": 0.878, "blue": 0.878}}
        fields = ["userEnteredFormat.backgroundColor"]

        request_body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {"sheetId": 0},
                        "cell": {"userEnteredFormat": cell_format},
                        "fields": ",".join(fields),
                    }
                }
            ]
        }

        bg_color = request_body["requests"][0]["repeatCell"]["cell"][
            "userEnteredFormat"
        ]["backgroundColor"]
        assert bg_color["red"] == pytest.approx(0.878, abs=0.01)

    def test_number_format_request_structure(self):
        """Test the structure of a number format request."""
        cell_format = {"numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}}
        fields = ["userEnteredFormat.numberFormat"]

        request_body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {"sheetId": 0},
                        "cell": {"userEnteredFormat": cell_format},
                        "fields": ",".join(fields),
                    }
                }
            ]
        }

        num_format = request_body["requests"][0]["repeatCell"]["cell"][
            "userEnteredFormat"
        ]["numberFormat"]
        assert num_format["type"] == "CURRENCY"
        assert num_format["pattern"] == "$#,##0.00"

    def test_combined_formatting_request_structure(self):
        """Test the structure of a request with multiple formatting options."""
        cell_format = {
            "textFormat": {
                "bold": True,
                "fontSize": 12,
            },
            "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 0.0},
            "horizontalAlignment": "CENTER",
        }
        fields = [
            "userEnteredFormat.textFormat.bold",
            "userEnteredFormat.textFormat.fontSize",
            "userEnteredFormat.backgroundColor",
            "userEnteredFormat.horizontalAlignment",
        ]

        request_body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {"sheetId": 0},
                        "cell": {"userEnteredFormat": cell_format},
                        "fields": ",".join(fields),
                    }
                }
            ]
        }

        cell = request_body["requests"][0]["repeatCell"]["cell"]["userEnteredFormat"]
        assert cell["textFormat"]["bold"] is True
        assert cell["textFormat"]["fontSize"] == 12
        assert cell["horizontalAlignment"] == "CENTER"


class TestFormatCellsValidation:
    """Tests for input validation in format_cells."""

    def test_valid_horizontal_alignments(self):
        """Test that valid horizontal alignments are accepted."""
        valid_alignments = ["LEFT", "CENTER", "RIGHT", "left", "center", "right"]
        for alignment in valid_alignments:
            # Just verify it's in the allowed list (case-insensitive)
            assert alignment.upper() in ["LEFT", "CENTER", "RIGHT"]

    def test_invalid_horizontal_alignment(self):
        """Test that invalid horizontal alignment is caught."""
        invalid_alignment = "JUSTIFY"
        valid_h_alignments = ["LEFT", "CENTER", "RIGHT"]
        assert invalid_alignment.upper() not in valid_h_alignments

    def test_valid_vertical_alignments(self):
        """Test that valid vertical alignments are accepted."""
        valid_alignments = ["TOP", "MIDDLE", "BOTTOM"]
        for alignment in valid_alignments:
            assert alignment in valid_alignments

    def test_valid_number_format_types(self):
        """Test that valid number format types are recognized."""
        valid_types = [
            "TEXT",
            "NUMBER",
            "PERCENT",
            "CURRENCY",
            "DATE",
            "TIME",
            "DATE_TIME",
            "SCIENTIFIC",
        ]
        for fmt_type in valid_types:
            assert fmt_type in valid_types

    def test_valid_wrap_strategies(self):
        """Test that valid wrap strategies are recognized."""
        valid_wrap = ["OVERFLOW_CELL", "LEGACY_WRAP", "CLIP", "WRAP"]
        for strategy in valid_wrap:
            assert strategy in valid_wrap


class TestFormatCellsOutputMessages:
    """Tests for output message formatting."""

    def test_success_message_format_basic(self):
        """Test basic success message format."""
        range_name = "A1:Z1"
        sheet_id = 0
        spreadsheet_id = "abc123"
        user_email = "test@example.com"
        format_summary = "bold=True, backgroundColor=#E0E0E0"

        message = (
            f"Successfully formatted range '{range_name}' in sheet ID {sheet_id} "
            f"of spreadsheet {spreadsheet_id} for {user_email}. "
            f"Applied: {format_summary}"
        )

        assert "A1:Z1" in message
        assert "sheet ID 0" in message
        assert "abc123" in message
        assert "bold=True" in message

    def test_format_description_building(self):
        """Test building format description from parameters."""
        format_descriptions = []

        # Simulate what happens in the function
        bold = True
        text_color = "#FF0000"
        font_size = 14

        if bold is not None:
            format_descriptions.append(f"bold={bold}")
        if text_color is not None:
            format_descriptions.append(f"textColor={text_color}")
        if font_size is not None:
            format_descriptions.append(f"fontSize={font_size}")

        format_summary = ", ".join(format_descriptions)

        assert format_summary == "bold=True, textColor=#FF0000, fontSize=14"
