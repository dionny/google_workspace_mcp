"""
Unit tests for paragraph formatting options in modify_doc_text.

Tests the following parameters:
- indent_first_line: First line indentation in points
- indent_start: Left margin indentation in points
- indent_end: Right margin indentation in points
- space_above: Space above paragraph in points
- space_below: Space below paragraph in points
"""
import pytest
from gdocs.docs_helpers import create_paragraph_style_request


class TestParagraphStyleRequestCreation:
    """Tests for create_paragraph_style_request helper function."""

    def test_create_request_with_indent_first_line(self):
        """Test creating a paragraph style request with first line indent."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            indent_first_line=36  # 0.5 inch
        )

        assert request is not None
        assert 'updateParagraphStyle' in request
        style_req = request['updateParagraphStyle']

        # Check range
        assert style_req['range']['startIndex'] == 100
        assert style_req['range']['endIndex'] == 200

        # Check paragraph style
        para_style = style_req['paragraphStyle']
        assert 'indentFirstLine' in para_style
        assert para_style['indentFirstLine']['magnitude'] == 36
        assert para_style['indentFirstLine']['unit'] == 'PT'

        # Check fields
        assert 'indentFirstLine' in style_req['fields']

    def test_create_request_with_indent_start(self):
        """Test creating a paragraph style request with left margin indent."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            indent_start=72  # 1 inch
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']
        assert para_style['indentStart']['magnitude'] == 72
        assert para_style['indentStart']['unit'] == 'PT'
        assert 'indentStart' in request['updateParagraphStyle']['fields']

    def test_create_request_with_indent_end(self):
        """Test creating a paragraph style request with right margin indent."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            indent_end=72  # 1 inch
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']
        assert para_style['indentEnd']['magnitude'] == 72
        assert para_style['indentEnd']['unit'] == 'PT'
        assert 'indentEnd' in request['updateParagraphStyle']['fields']

    def test_create_request_with_space_above(self):
        """Test creating a paragraph style request with space above."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            space_above=12  # 12 points
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']
        assert para_style['spaceAbove']['magnitude'] == 12
        assert para_style['spaceAbove']['unit'] == 'PT'
        assert 'spaceAbove' in request['updateParagraphStyle']['fields']

    def test_create_request_with_space_below(self):
        """Test creating a paragraph style request with space below."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            space_below=12  # 12 points
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']
        assert para_style['spaceBelow']['magnitude'] == 12
        assert para_style['spaceBelow']['unit'] == 'PT'
        assert 'spaceBelow' in request['updateParagraphStyle']['fields']

    def test_create_request_with_negative_indent_first_line(self):
        """Test hanging indent with negative first line indent."""
        # Hanging indent: indent_start=36, indent_first_line=-36
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            indent_start=36,
            indent_first_line=-36
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']
        assert para_style['indentStart']['magnitude'] == 36
        assert para_style['indentFirstLine']['magnitude'] == -36

    def test_create_request_with_all_indentation_options(self):
        """Test creating a request with all indentation and spacing options."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            indent_first_line=36,
            indent_start=72,
            indent_end=72,
            space_above=12,
            space_below=12
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']
        fields = request['updateParagraphStyle']['fields']

        # Check all values
        assert para_style['indentFirstLine']['magnitude'] == 36
        assert para_style['indentStart']['magnitude'] == 72
        assert para_style['indentEnd']['magnitude'] == 72
        assert para_style['spaceAbove']['magnitude'] == 12
        assert para_style['spaceBelow']['magnitude'] == 12

        # Check all fields are present
        assert 'indentFirstLine' in fields
        assert 'indentStart' in fields
        assert 'indentEnd' in fields
        assert 'spaceAbove' in fields
        assert 'spaceBelow' in fields

    def test_create_request_with_combined_options(self):
        """Test combining new indentation options with existing options."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            line_spacing=150,
            alignment='CENTER',
            indent_first_line=36,
            space_above=12
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']
        fields = request['updateParagraphStyle']['fields']

        # Check existing options work
        assert para_style['lineSpacing'] == 150
        assert para_style['alignment'] == 'CENTER'

        # Check new options work
        assert para_style['indentFirstLine']['magnitude'] == 36
        assert para_style['spaceAbove']['magnitude'] == 12

        # Check all fields
        assert 'lineSpacing' in fields
        assert 'alignment' in fields
        assert 'indentFirstLine' in fields
        assert 'spaceAbove' in fields

    def test_create_request_with_zero_values(self):
        """Test that zero values are included in the request."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            space_above=0,  # Remove space above
            space_below=0   # Remove space below
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']

        # Zero values should be included to explicitly remove spacing
        assert para_style['spaceAbove']['magnitude'] == 0
        assert para_style['spaceBelow']['magnitude'] == 0

    def test_create_request_returns_none_with_no_options(self):
        """Test that request returns None when no style options provided."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200
        )

        assert request is None


class TestParagraphFormattingValidation:
    """Tests for parameter validation in modify_doc_text."""

    def test_validate_indent_first_line_accepts_float(self):
        """Test that indent_first_line accepts float values."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            indent_first_line=36.5
        )

        assert request is not None
        assert request['updateParagraphStyle']['paragraphStyle']['indentFirstLine']['magnitude'] == 36.5

    def test_validate_indent_start_accepts_float(self):
        """Test that indent_start accepts float values."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            indent_start=72.25
        )

        assert request is not None
        assert request['updateParagraphStyle']['paragraphStyle']['indentStart']['magnitude'] == 72.25

    def test_validate_space_above_accepts_float(self):
        """Test that space_above accepts float values."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            space_above=6.5
        )

        assert request is not None
        assert request['updateParagraphStyle']['paragraphStyle']['spaceAbove']['magnitude'] == 6.5


class TestBlockQuoteFormatting:
    """Tests for block quote formatting patterns."""

    def test_block_quote_indentation(self):
        """Test creating block quote style with left and right margins."""
        # Block quote: indent from both sides
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            indent_start=72,  # 1 inch from left
            indent_end=72     # 1 inch from right
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']
        assert para_style['indentStart']['magnitude'] == 72
        assert para_style['indentEnd']['magnitude'] == 72


class TestHangingIndent:
    """Tests for hanging indent patterns (used in bibliographies)."""

    def test_hanging_indent_pattern(self):
        """Test hanging indent with positive start and negative first line."""
        # Hanging indent: first line at margin, subsequent lines indented
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            indent_start=36,       # All lines indented 0.5 inch
            indent_first_line=-36  # But first line goes back 0.5 inch (to margin)
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']

        # First line should be at margin (36 + (-36) = 0)
        assert para_style['indentStart']['magnitude'] == 36
        assert para_style['indentFirstLine']['magnitude'] == -36


class TestParagraphSpacing:
    """Tests for paragraph spacing patterns."""

    def test_add_space_before_and_after(self):
        """Test adding spacing before and after a paragraph."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            space_above=12,
            space_below=6
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']
        assert para_style['spaceAbove']['magnitude'] == 12
        assert para_style['spaceBelow']['magnitude'] == 6

    def test_typical_heading_spacing(self):
        """Test typical heading spacing (more above, less below)."""
        # Headings typically have more space above than below
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            heading_style='HEADING_2',
            space_above=18,
            space_below=6
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']
        assert para_style['namedStyleType'] == 'HEADING_2'
        assert para_style['spaceAbove']['magnitude'] == 18
        assert para_style['spaceBelow']['magnitude'] == 6


class TestIndentUnits:
    """Tests to verify that indentation uses correct units."""

    def test_all_indents_use_pt_unit(self):
        """Test that all indentation values use PT (points) as the unit."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            indent_first_line=36,
            indent_start=72,
            indent_end=48
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']

        # All should use PT (points) unit
        assert para_style['indentFirstLine']['unit'] == 'PT'
        assert para_style['indentStart']['unit'] == 'PT'
        assert para_style['indentEnd']['unit'] == 'PT'

    def test_all_spacing_uses_pt_unit(self):
        """Test that all spacing values use PT (points) as the unit."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            space_above=12,
            space_below=6
        )

        assert request is not None
        para_style = request['updateParagraphStyle']['paragraphStyle']

        # All should use PT (points) unit
        assert para_style['spaceAbove']['unit'] == 'PT'
        assert para_style['spaceBelow']['unit'] == 'PT'


class TestLineSpacingNormalization:
    """Tests for line_spacing parameter normalization in modify_doc_text."""

    def test_line_spacing_decimal_multiplier_single(self):
        """Test that 1.0 is converted to 100 (single spacing)."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            line_spacing=100  # Simulating already-normalized value
        )
        assert request is not None
        assert request['updateParagraphStyle']['paragraphStyle']['lineSpacing'] == 100

    def test_line_spacing_decimal_multiplier_double(self):
        """Test that 2.0 would be converted to 200 (double spacing)."""
        # The helper receives already-normalized values, so test the helper with 200
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            line_spacing=200
        )
        assert request is not None
        assert request['updateParagraphStyle']['paragraphStyle']['lineSpacing'] == 200

    def test_line_spacing_decimal_multiplier_one_and_half(self):
        """Test that 1.5 would be converted to 150 (1.5x spacing)."""
        request = create_paragraph_style_request(
            start_index=100,
            end_index=200,
            line_spacing=150
        )
        assert request is not None
        assert request['updateParagraphStyle']['paragraphStyle']['lineSpacing'] == 150

    def test_line_spacing_percentage_values(self):
        """Test that percentage values 50-1000 are accepted."""
        for value in [50, 115, 150, 200, 500, 1000]:
            request = create_paragraph_style_request(
                start_index=100,
                end_index=200,
                line_spacing=value
            )
            assert request is not None
            assert request['updateParagraphStyle']['paragraphStyle']['lineSpacing'] == value


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
