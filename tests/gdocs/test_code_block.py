"""
Unit tests for code_block formatting functionality in modify_doc_text.

The code_block parameter is a convenience feature that applies monospace font
(Courier New) and light gray background (#f5f5f5) for code-style formatting.
"""
import pytest
from gdocs.docs_helpers import (
    create_format_text_request,
    build_text_style,
)


class TestCodeBlockFormattingHelpers:
    """Tests for code block formatting helpers."""

    def test_create_format_text_request_with_code_block_defaults(self):
        """Test that code_block style uses Courier New and gray background."""
        # Simulating what modify_doc_text does when code_block=True
        font_family = "Courier New"  # code_block default
        background_color = "#f5f5f5"  # code_block default

        request = create_format_text_request(
            start_index=100,
            end_index=150,
            font_family=font_family,
            background_color=background_color
        )

        assert 'updateTextStyle' in request
        style_req = request['updateTextStyle']

        # Check range
        assert style_req['range']['startIndex'] == 100
        assert style_req['range']['endIndex'] == 150

        # Check code block formatting
        text_style = style_req['textStyle']
        assert text_style['weightedFontFamily']['fontFamily'] == "Courier New"
        assert 'backgroundColor' in text_style

        # Check fields
        fields = style_req['fields']
        assert 'weightedFontFamily' in fields
        assert 'backgroundColor' in fields

    def test_build_text_style_monospace_font(self):
        """Test building text style with monospace font for code."""
        text_style, fields = build_text_style(font_family="Courier New")

        assert 'weightedFontFamily' in text_style
        assert text_style['weightedFontFamily']['fontFamily'] == "Courier New"
        assert 'weightedFontFamily' in fields

    def test_build_text_style_code_block_full_style(self):
        """Test building complete code block style with font and background."""
        text_style, fields = build_text_style(
            font_family="Courier New",
            background_color="#f5f5f5"
        )

        assert text_style['weightedFontFamily']['fontFamily'] == "Courier New"
        assert 'backgroundColor' in text_style
        assert 'weightedFontFamily' in fields
        assert 'backgroundColor' in fields

    def test_code_block_with_custom_background_override(self):
        """Test that custom background_color overrides code_block default."""
        # User specifies code_block=True but also background_color="#e0e0e0"
        # The custom background should be used
        custom_background = "#e0e0e0"

        text_style, fields = build_text_style(
            font_family="Courier New",
            background_color=custom_background
        )

        assert text_style['weightedFontFamily']['fontFamily'] == "Courier New"
        # The background color value will be parsed, but key should exist
        assert 'backgroundColor' in text_style

    def test_code_block_with_custom_font_override(self):
        """Test that custom font_family overrides code_block default."""
        # User specifies code_block=True but also font_family="Consolas"
        # The custom font should be used
        custom_font = "Consolas"

        text_style, fields = build_text_style(
            font_family=custom_font,
            background_color="#f5f5f5"
        )

        assert text_style['weightedFontFamily']['fontFamily'] == "Consolas"


class TestCodeBlockRequestStructure:
    """Tests for the structure of code block format requests."""

    def test_combined_insert_and_code_block_batch_structure(self):
        """Test the structure of a combined insert+code_block batch request."""
        from gdocs.docs_helpers import create_insert_text_request

        code_text = "def hello():\n    print('Hello!')"
        start_index = 100
        calculated_end_index = start_index + len(code_text)

        # Build the batch request as modify_doc_text would
        requests = []

        # First: insert text
        requests.append(create_insert_text_request(start_index, code_text))

        # Second: apply code block formatting (font + background)
        requests.append(create_format_text_request(
            start_index=start_index,
            end_index=calculated_end_index,
            font_family="Courier New",
            background_color="#f5f5f5"
        ))

        # Verify the batch structure
        assert len(requests) == 2

        # First should be insertText
        assert 'insertText' in requests[0]
        assert requests[0]['insertText']['text'] == code_text

        # Second should be updateTextStyle with code block formatting
        assert 'updateTextStyle' in requests[1]
        style_req = requests[1]['updateTextStyle']
        assert style_req['range']['startIndex'] == start_index
        assert style_req['range']['endIndex'] == calculated_end_index
        assert style_req['textStyle']['weightedFontFamily']['fontFamily'] == "Courier New"
        assert 'backgroundColor' in style_req['textStyle']

    def test_code_block_format_range_matches_text(self):
        """Ensure code block formatting is applied to exactly the inserted text range."""

        code_text = "console.log('test');"
        start_index = 75

        # Simulate the logic from modify_doc_text
        format_start = start_index
        format_end = start_index + len(code_text)

        format_req = create_format_text_request(
            format_start,
            format_end,
            font_family="Courier New",
            background_color="#f5f5f5"
        )

        range_obj = format_req['updateTextStyle']['range']
        covered_length = range_obj['endIndex'] - range_obj['startIndex']

        assert covered_length == len(code_text), (
            f"Format range covers {covered_length} chars but text is {len(code_text)} chars"
        )


class TestCodeBlockValidation:
    """Tests for code_block parameter validation."""

    def test_validate_code_block_default_font_family(self):
        """Test that code_block applies Courier New font by default."""
        from gdocs.managers.validation_manager import ValidationManager
        validator = ValidationManager()

        # The font_family that code_block sets should pass validation
        is_valid, error_msg = validator.validate_text_formatting_params(
            font_family="Courier New"
        )

        assert is_valid is True
        assert error_msg == ""

    def test_validate_code_block_default_background(self):
        """Test that code_block's default background color is valid."""
        from gdocs.managers.validation_manager import ValidationManager
        validator = ValidationManager()

        # The background_color that code_block sets should pass validation
        is_valid, error_msg = validator.validate_text_formatting_params(
            background_color="#f5f5f5"
        )

        assert is_valid is True
        assert error_msg == ""

    def test_validate_code_block_combined_styles(self):
        """Test that code_block's combined font and background are valid."""
        from gdocs.managers.validation_manager import ValidationManager
        validator = ValidationManager()

        is_valid, error_msg = validator.validate_text_formatting_params(
            font_family="Courier New",
            background_color="#f5f5f5"
        )

        assert is_valid is True
        assert error_msg == ""


class TestCodeBlockWithOtherFormatting:
    """Tests for code_block combined with other formatting options."""

    def test_code_block_with_bold(self):
        """Test code block combined with bold formatting."""
        text_style, fields = build_text_style(
            bold=True,
            font_family="Courier New",
            background_color="#f5f5f5"
        )

        assert text_style['bold'] is True
        assert text_style['weightedFontFamily']['fontFamily'] == "Courier New"
        assert 'backgroundColor' in text_style
        assert 'bold' in fields
        assert 'weightedFontFamily' in fields
        assert 'backgroundColor' in fields

    def test_code_block_with_foreground_color(self):
        """Test code block combined with foreground color (syntax highlighting)."""
        text_style, fields = build_text_style(
            font_family="Courier New",
            background_color="#f5f5f5",
            foreground_color="#0000ff"  # Blue for keywords
        )

        assert text_style['weightedFontFamily']['fontFamily'] == "Courier New"
        assert 'backgroundColor' in text_style
        assert 'foregroundColor' in text_style


class TestCodeBlockEdgeCases:
    """Edge case tests for code_block functionality."""

    def test_code_block_with_multiline_text(self):
        """Test code block with multiline code snippet."""
        code_text = """def factorial(n):
    if n <= 1:
        return 1
    return n * factorial(n - 1)"""

        start_index = 100
        end_index = start_index + len(code_text)

        request = create_format_text_request(
            start_index,
            end_index,
            font_family="Courier New",
            background_color="#f5f5f5"
        )

        assert request is not None
        range_obj = request['updateTextStyle']['range']
        assert range_obj['endIndex'] - range_obj['startIndex'] == len(code_text)

    def test_code_block_with_special_characters(self):
        """Test code block with special characters in code."""
        code_text = 'const regex = /^[a-z]+$/gi;'

        start_index = 50
        end_index = start_index + len(code_text)

        request = create_format_text_request(
            start_index,
            end_index,
            font_family="Courier New",
            background_color="#f5f5f5"
        )

        assert request is not None
        # Just verify the range is correct
        range_obj = request['updateTextStyle']['range']
        assert range_obj['startIndex'] == 50
        assert range_obj['endIndex'] == 50 + len(code_text)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
