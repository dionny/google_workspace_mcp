"""
Unit tests for combined insert + format operations.

These tests verify the helper functions that support inserting text WITH formatting.
The tests focus on the request building logic.
"""
import pytest
from gdocs.docs_helpers import (
    create_insert_text_request,
    create_format_text_request,
    build_text_style,
)


class TestCombinedInsertFormatHelpers:
    """Tests for helper functions that support combined insert and format operations."""

    def test_create_insert_text_request(self):
        """Test that insert text request is properly formatted."""
        request = create_insert_text_request(100, "IMPORTANT:")

        assert 'insertText' in request
        assert request['insertText']['location']['index'] == 100
        assert request['insertText']['text'] == "IMPORTANT:"

    def test_create_format_text_request_with_all_styles(self):
        """Test formatting request with all style options."""
        request = create_format_text_request(
            start_index=100,
            end_index=110,  # 100 + len("IMPORTANT:")
            bold=True,
            italic=True,
            underline=True,
            font_size=14,
            font_family="Arial"
        )

        assert 'updateTextStyle' in request
        style_req = request['updateTextStyle']

        # Check range is correct
        assert style_req['range']['startIndex'] == 100
        assert style_req['range']['endIndex'] == 110

        # Check all style properties
        text_style = style_req['textStyle']
        assert text_style['bold'] is True
        assert text_style['italic'] is True
        assert text_style['underline'] is True
        assert text_style['fontSize']['magnitude'] == 14
        assert text_style['weightedFontFamily']['fontFamily'] == "Arial"

        # Check fields list includes all styles
        fields = style_req['fields']
        assert 'bold' in fields
        assert 'italic' in fields
        assert 'underline' in fields
        assert 'fontSize' in fields
        assert 'weightedFontFamily' in fields

    def test_build_text_style_partial(self):
        """Test building text style with only some options."""
        text_style, fields = build_text_style(bold=True, font_size=16)

        assert text_style['bold'] is True
        assert text_style['fontSize']['magnitude'] == 16
        assert 'italic' not in text_style
        assert 'underline' not in text_style

        assert 'bold' in fields
        assert 'fontSize' in fields
        assert len(fields) == 2

    def test_create_format_text_request_returns_none_without_styles(self):
        """Test that format request returns None when no styles provided."""
        request = create_format_text_request(
            start_index=0,
            end_index=10
            # No formatting options
        )

        assert request is None

    def test_combined_insert_and_format_batch_structure(self):
        """
        Test the structure of a combined insert+format batch request.
        This simulates what modify_doc_text should produce.
        """
        text = "IMPORTANT:"
        start_index = 100
        calculated_end_index = start_index + len(text)  # 110

        # Build the batch request as modify_doc_text would
        requests = []

        # First: insert text
        requests.append(create_insert_text_request(start_index, text))

        # Second: format the inserted text (range calculated from insert)
        requests.append(create_format_text_request(
            start_index=start_index,
            end_index=calculated_end_index,
            bold=True,
            italic=True
        ))

        # Verify the batch structure
        assert len(requests) == 2

        # First should be insertText
        assert 'insertText' in requests[0]
        assert requests[0]['insertText']['text'] == text

        # Second should be updateTextStyle with auto-calculated range
        assert 'updateTextStyle' in requests[1]
        assert requests[1]['updateTextStyle']['range']['startIndex'] == start_index
        assert requests[1]['updateTextStyle']['range']['endIndex'] == calculated_end_index

    def test_insert_at_index_zero_adjustment(self):
        """
        Test that index 0 gets adjusted to 1 (to avoid first section break).
        This is the behavior in modify_doc_text.
        """
        # When start_index is 0, it should be adjusted to 1
        start_index = 0
        actual_index = 1 if start_index == 0 else start_index

        text = "Header"
        format_start = actual_index
        format_end = actual_index + len(text)

        # Build requests
        insert_req = create_insert_text_request(actual_index, text)
        format_req = create_format_text_request(format_start, format_end, bold=True)

        # Verify adjusted indices
        assert insert_req['insertText']['location']['index'] == 1  # Not 0
        assert format_req['updateTextStyle']['range']['startIndex'] == 1
        assert format_req['updateTextStyle']['range']['endIndex'] == 7  # 1 + len("Header")


class TestEndIndexCalculation:
    """Tests specifically for end_index auto-calculation logic."""

    def test_end_index_calculated_from_text_length(self):
        """The key feature: end_index = start_index + len(text)."""
        test_cases = [
            {"start": 1, "text": "Hello", "expected_end": 6},
            {"start": 100, "text": "IMPORTANT:", "expected_end": 110},
            {"start": 50, "text": "x", "expected_end": 51},
            {"start": 200, "text": "Multi\nLine\nText", "expected_end": 215},
        ]

        for case in test_cases:
            calculated_end = case["start"] + len(case["text"])
            assert calculated_end == case["expected_end"], (
                f"For text '{case['text']}' at index {case['start']}, "
                f"expected end {case['expected_end']} but got {calculated_end}"
            )

    def test_format_range_matches_inserted_text(self):
        """Ensure formatting is applied to exactly the inserted text range."""
        text = "Bold Text Here"
        start_index = 75

        # Simulate the logic from modify_doc_text
        format_start = start_index
        format_end = start_index + len(text)

        format_req = create_format_text_request(format_start, format_end, bold=True)

        range_obj = format_req['updateTextStyle']['range']
        covered_length = range_obj['endIndex'] - range_obj['startIndex']

        assert covered_length == len(text), (
            f"Format range covers {covered_length} chars but text is {len(text)} chars"
        )


class TestPositionShiftCalculation:
    """Tests for position shift calculation functions."""

    def test_insert_position_shift(self):
        """Test position shift for insert operation."""
        from gdocs.docs_helpers import calculate_position_shift, OperationType

        # Insert 13 characters at position 100
        shift, affected_range = calculate_position_shift(
            OperationType.INSERT,
            start_index=100,
            end_index=None,
            text_length=13
        )

        assert shift == 13, "Insert should shift by text length"
        assert affected_range == {"start": 100, "end": 113}

    def test_delete_position_shift(self):
        """Test position shift for delete operation."""
        from gdocs.docs_helpers import calculate_position_shift, OperationType

        # Delete 50 characters from position 100 to 150
        shift, affected_range = calculate_position_shift(
            OperationType.DELETE,
            start_index=100,
            end_index=150,
            text_length=0
        )

        assert shift == -50, "Delete should shift by negative deleted length"
        assert affected_range == {"start": 100, "end": 100}

    def test_replace_position_shift_shorter(self):
        """Test position shift when replacing with shorter text."""
        from gdocs.docs_helpers import calculate_position_shift, OperationType

        # Replace 50 characters (100-150) with 8 characters
        shift, affected_range = calculate_position_shift(
            OperationType.REPLACE,
            start_index=100,
            end_index=150,
            text_length=8
        )

        assert shift == -42, "Replace should shift by (new_length - old_length)"
        assert affected_range == {"start": 100, "end": 108}

    def test_replace_position_shift_longer(self):
        """Test position shift when replacing with longer text."""
        from gdocs.docs_helpers import calculate_position_shift, OperationType

        # Replace 10 characters (100-110) with 25 characters
        shift, affected_range = calculate_position_shift(
            OperationType.REPLACE,
            start_index=100,
            end_index=110,
            text_length=25
        )

        assert shift == 15, "Replace should shift by (new_length - old_length)"
        assert affected_range == {"start": 100, "end": 125}

    def test_format_position_shift(self):
        """Test position shift for format operation (should be zero)."""
        from gdocs.docs_helpers import calculate_position_shift, OperationType

        # Format 50 characters from 100 to 150
        shift, affected_range = calculate_position_shift(
            OperationType.FORMAT,
            start_index=100,
            end_index=150,
            text_length=0
        )

        assert shift == 0, "Format should not shift positions"
        assert affected_range == {"start": 100, "end": 150}


class TestBuildOperationResult:
    """Tests for operation result building."""

    def test_build_insert_result(self):
        """Test building result for insert operation."""
        from gdocs.docs_helpers import build_operation_result, OperationType

        result = build_operation_result(
            operation_type=OperationType.INSERT,
            start_index=100,
            end_index=None,
            text="INSERTED TEXT",
            document_id="test-doc-123"
        )

        assert result.success is True
        assert result.operation == "insert"
        assert result.position_shift == 13
        assert result.affected_range == {"start": 100, "end": 113}
        assert result.inserted_length == 13
        assert "test-doc-123" in result.link

    def test_build_replace_result(self):
        """Test building result for replace operation."""
        from gdocs.docs_helpers import build_operation_result, OperationType

        result = build_operation_result(
            operation_type=OperationType.REPLACE,
            start_index=100,
            end_index=150,
            text="NEW TEXT",
            document_id="test-doc-123"
        )

        assert result.success is True
        assert result.operation == "replace"
        assert result.position_shift == -42  # 8 - 50 = -42
        assert result.affected_range == {"start": 100, "end": 108}
        assert result.original_length == 50
        assert result.new_length == 8

    def test_build_delete_result(self):
        """Test building result for delete operation."""
        from gdocs.docs_helpers import build_operation_result, OperationType

        result = build_operation_result(
            operation_type=OperationType.DELETE,
            start_index=100,
            end_index=150,
            text=None,
            document_id="test-doc-123"
        )

        assert result.success is True
        assert result.operation == "delete"
        assert result.position_shift == -50
        assert result.deleted_length == 50

    def test_build_format_result(self):
        """Test building result for format operation."""
        from gdocs.docs_helpers import build_operation_result, OperationType

        result = build_operation_result(
            operation_type=OperationType.FORMAT,
            start_index=100,
            end_index=150,
            text=None,
            document_id="test-doc-123",
            styles_applied=["bold", "italic"]
        )

        assert result.success is True
        assert result.operation == "format"
        assert result.position_shift == 0
        assert result.styles_applied == ["bold", "italic"]

    def test_result_to_dict_excludes_none(self):
        """Test that to_dict() excludes None values."""
        from gdocs.docs_helpers import build_operation_result, OperationType

        result = build_operation_result(
            operation_type=OperationType.INSERT,
            start_index=100,
            end_index=None,
            text="test",
            document_id="test-doc-123"
        )

        result_dict = result.to_dict()

        # Should NOT contain None fields
        assert "deleted_length" not in result_dict
        assert "original_length" not in result_dict
        assert "styles_applied" not in result_dict

        # Should contain populated fields
        assert "inserted_length" in result_dict
        assert result_dict["inserted_length"] == 4


class TestFollowUpEditScenario:
    """
    Tests simulating the use case from the issue:
    Making multiple edits efficiently without re-reading the document.
    """

    def test_sequential_inserts_with_shift_tracking(self):
        """
        Simulate the use case from the issue:
        Make multiple edits efficiently by tracking position shifts.
        """
        from gdocs.docs_helpers import calculate_position_shift, OperationType

        # Document is 5000 chars, first edit at position 100
        doc_length = 5000
        edit_position_1 = 100

        # First edit: insert "AAA" (3 chars)
        shift1, _ = calculate_position_shift(
            OperationType.INSERT,
            start_index=edit_position_1,
            end_index=None,
            text_length=3
        )
        assert shift1 == 3

        # Second edit was originally planned at position 200
        # But now needs to be adjusted by the shift
        original_position_2 = 200
        adjusted_position_2 = original_position_2 + shift1
        assert adjusted_position_2 == 203

        # Second edit: insert "BBB" (3 chars) at adjusted position
        shift2, _ = calculate_position_shift(
            OperationType.INSERT,
            start_index=adjusted_position_2,
            end_index=None,
            text_length=3
        )
        assert shift2 == 3

        # Third edit was originally planned at position 300
        # Needs to be adjusted by both shifts
        original_position_3 = 300
        adjusted_position_3 = original_position_3 + shift1 + shift2
        assert adjusted_position_3 == 306

    def test_mixed_operations_shift_tracking(self):
        """Test tracking shifts through mixed insert/delete/replace operations."""
        from gdocs.docs_helpers import calculate_position_shift, OperationType

        cumulative_shift = 0
        target_position = 500  # Original target position

        # Operation 1: Insert 10 chars at position 100
        shift1, _ = calculate_position_shift(
            OperationType.INSERT, 100, None, 10
        )
        cumulative_shift += shift1
        assert cumulative_shift == 10

        # Operation 2: Delete 20 chars at position 200-220
        shift2, _ = calculate_position_shift(
            OperationType.DELETE, 200, 220, 0
        )
        cumulative_shift += shift2
        assert cumulative_shift == -10  # 10 - 20 = -10

        # Operation 3: Replace 30 chars (300-330) with 50 chars
        shift3, _ = calculate_position_shift(
            OperationType.REPLACE, 300, 330, 50
        )
        cumulative_shift += shift3
        assert cumulative_shift == 10  # -10 + 20 = 10

        # Final adjusted target position
        final_position = target_position + cumulative_shift
        assert final_position == 510


class TestHyperlinkSupport:
    """Tests for hyperlink formatting support."""

    def test_build_text_style_with_link(self):
        """Test building text style with a hyperlink."""
        text_style, fields = build_text_style(link="https://example.com")

        assert 'link' in text_style
        assert text_style['link'] == {'url': 'https://example.com'}
        assert 'link' in fields
        assert len(fields) == 1

    def test_build_text_style_with_link_and_other_styles(self):
        """Test building text style with link combined with other formatting."""
        text_style, fields = build_text_style(
            bold=True,
            underline=True,
            link="https://example.com"
        )

        assert text_style['bold'] is True
        assert text_style['underline'] is True
        assert text_style['link'] == {'url': 'https://example.com'}
        assert 'bold' in fields
        assert 'underline' in fields
        assert 'link' in fields
        assert len(fields) == 3

    def test_build_text_style_remove_link(self):
        """Test that empty string link removes the link."""
        text_style, fields = build_text_style(link="")

        assert 'link' in text_style
        assert text_style['link'] is None
        assert 'link' in fields

    def test_create_format_text_request_with_link(self):
        """Test creating format request with hyperlink."""
        request = create_format_text_request(
            start_index=100,
            end_index=115,
            link="https://example.com"
        )

        assert 'updateTextStyle' in request
        style_req = request['updateTextStyle']

        assert style_req['range']['startIndex'] == 100
        assert style_req['range']['endIndex'] == 115
        assert style_req['textStyle']['link'] == {'url': 'https://example.com'}
        assert 'link' in style_req['fields']

    def test_create_format_text_request_with_link_and_bold(self):
        """Test creating format request with both link and bold."""
        request = create_format_text_request(
            start_index=50,
            end_index=70,
            bold=True,
            link="https://example.com/page"
        )

        assert 'updateTextStyle' in request
        style_req = request['updateTextStyle']

        text_style = style_req['textStyle']
        assert text_style['bold'] is True
        assert text_style['link'] == {'url': 'https://example.com/page'}

        fields = style_req['fields']
        assert 'bold' in fields
        assert 'link' in fields

    def test_combined_insert_and_link_batch_structure(self):
        """Test the structure of a combined insert+link batch request."""
        text = "Visit our website"
        start_index = 100
        calculated_end_index = start_index + len(text)

        # Build the batch request as modify_doc_text would
        requests = []

        # First: insert text
        requests.append(create_insert_text_request(start_index, text))

        # Second: add hyperlink to the inserted text
        requests.append(create_format_text_request(
            start_index=start_index,
            end_index=calculated_end_index,
            link="https://example.com"
        ))

        # Verify the batch structure
        assert len(requests) == 2

        # First should be insertText
        assert 'insertText' in requests[0]
        assert requests[0]['insertText']['text'] == text

        # Second should be updateTextStyle with link
        assert 'updateTextStyle' in requests[1]
        style_req = requests[1]['updateTextStyle']
        assert style_req['range']['startIndex'] == start_index
        assert style_req['range']['endIndex'] == calculated_end_index
        assert style_req['textStyle']['link'] == {'url': 'https://example.com'}

    def test_link_with_internal_bookmark(self):
        """Test that internal document bookmarks (#heading) are supported."""
        text_style, fields = build_text_style(link="#section-1")

        assert text_style['link'] == {'url': '#section-1'}
        assert 'link' in fields


class TestHyperlinkValidation:
    """Tests for hyperlink validation in validation_manager."""

    def test_validate_link_valid_https(self):
        """Test validation accepts valid https URLs."""
        from gdocs.managers.validation_manager import ValidationManager
        validator = ValidationManager()

        is_valid, error_msg = validator.validate_text_formatting_params(
            link="https://example.com"
        )

        assert is_valid is True
        assert error_msg == ""

    def test_validate_link_valid_http(self):
        """Test validation accepts valid http URLs."""
        from gdocs.managers.validation_manager import ValidationManager
        validator = ValidationManager()

        is_valid, error_msg = validator.validate_text_formatting_params(
            link="http://example.com"
        )

        assert is_valid is True
        assert error_msg == ""

    def test_validate_link_valid_bookmark(self):
        """Test validation accepts internal bookmarks."""
        from gdocs.managers.validation_manager import ValidationManager
        validator = ValidationManager()

        is_valid, error_msg = validator.validate_text_formatting_params(
            link="#my-heading"
        )

        assert is_valid is True
        assert error_msg == ""

    def test_validate_link_empty_string_allowed(self):
        """Test validation accepts empty string (to remove link)."""
        from gdocs.managers.validation_manager import ValidationManager
        validator = ValidationManager()

        is_valid, error_msg = validator.validate_text_formatting_params(
            link=""
        )

        assert is_valid is True
        assert error_msg == ""

    def test_validate_link_invalid_protocol(self):
        """Test validation rejects invalid URLs."""
        from gdocs.managers.validation_manager import ValidationManager
        validator = ValidationManager()

        is_valid, error_msg = validator.validate_text_formatting_params(
            link="ftp://example.com"
        )

        assert is_valid is False
        assert "link must be a valid URL" in error_msg

    def test_validate_link_non_string_rejected(self):
        """Test validation rejects non-string link values."""
        from gdocs.managers.validation_manager import ValidationManager
        validator = ValidationManager()

        is_valid, error_msg = validator.validate_text_formatting_params(
            link=123
        )

        assert is_valid is False
        assert "link must be a string" in error_msg


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
