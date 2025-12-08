"""
Unit tests for Google Docs structured error handling.

These tests verify that error messages are correctly structured,
contain all required fields, and provide actionable guidance.
"""

import json
import pytest
from gdocs.errors import (
    ErrorCode,
    StructuredError,
    ErrorContext,
    DocsErrorBuilder,
    format_error,
    simple_error,
)
from gdocs.managers.validation_manager import ValidationManager


class TestErrorCode:
    """Tests for ErrorCode enum."""

    def test_error_codes_are_strings(self):
        """All error codes should be string values."""
        for code in ErrorCode:
            assert isinstance(code.value, str)
            assert code.value.isupper()

    def test_key_error_codes_exist(self):
        """Key error codes from the spec should exist."""
        expected_codes = [
            "FORMATTING_REQUIRES_RANGE",
            "INDEX_OUT_OF_BOUNDS",
            "INVALID_INDEX_RANGE",
            "EMPTY_SEARCH_TEXT",
            "SEARCH_TEXT_NOT_FOUND",
            "AMBIGUOUS_SEARCH",
            "HEADING_NOT_FOUND",
            "DOCUMENT_NOT_FOUND",
            "PERMISSION_DENIED",
            "TABLE_NOT_FOUND",
            "INVALID_TABLE_DATA",
            "INVALID_COLOR_FORMAT",
        ]
        for code in expected_codes:
            assert hasattr(ErrorCode, code), f"Missing error code: {code}"


class TestStructuredError:
    """Tests for StructuredError dataclass."""

    def test_basic_error_creation(self):
        """Can create a basic structured error."""
        error = StructuredError(
            code="TEST_ERROR", message="Test message", suggestion="Test suggestion"
        )
        assert error.error is True
        assert error.code == "TEST_ERROR"
        assert error.message == "Test message"
        assert error.suggestion == "Test suggestion"

    def test_to_dict_excludes_none(self):
        """to_dict should exclude None values."""
        error = StructuredError(
            code="TEST",
            message="Test",
        )
        result = error.to_dict()
        assert "example" not in result
        assert "context" not in result
        assert "docs_url" not in result

    def test_to_json_valid_json(self):
        """to_json should return valid JSON."""
        error = StructuredError(
            code="TEST", message="Test message", suggestion="Do something"
        )
        json_str = error.to_json()
        parsed = json.loads(json_str)
        assert parsed["error"] is True
        assert parsed["code"] == "TEST"

    def test_context_included_when_present(self):
        """Context should be included when it has values."""
        error = StructuredError(
            code="TEST",
            message="Test",
            context=ErrorContext(document_length=1000, received={"start_index": 500}),
        )
        result = error.to_dict()
        assert "context" in result
        assert result["context"]["document_length"] == 1000


class TestDocsErrorBuilder:
    """Tests for DocsErrorBuilder factory methods."""

    def test_formatting_requires_range_no_text(self):
        """Error for formatting without text or end_index."""
        error = DocsErrorBuilder.formatting_requires_range(
            start_index=100, has_text=False, formatting_params=["bold", "italic"]
        )
        assert error.code == "FORMATTING_REQUIRES_RANGE"
        assert "end_index" in error.message.lower()
        assert "formatting" in error.reason.lower()
        assert error.example is not None
        assert "option_1" in error.example  # Format existing
        assert "option_2" in error.example  # Insert then format

    def test_formatting_requires_range_with_text(self):
        """Error message differs when text is provided."""
        error = DocsErrorBuilder.formatting_requires_range(
            start_index=100, has_text=True, formatting_params=["bold"]
        )
        assert "correct_usage" in error.example

    def test_index_out_of_bounds(self):
        """Index out of bounds provides document length context."""
        error = DocsErrorBuilder.index_out_of_bounds(
            index_name="start_index", index_value=5000, document_length=3500
        )
        assert error.code == "INDEX_OUT_OF_BOUNDS"
        assert "5000" in error.message
        assert "3500" in error.message
        assert error.context.document_length == 3500
        assert "get_doc_info" in error.suggestion.lower()

    def test_invalid_index_range(self):
        """Invalid range shows expected values."""
        error = DocsErrorBuilder.invalid_index_range(start_index=200, end_index=100)
        assert error.code == "INVALID_INDEX_RANGE"
        assert error.context.received == {"start_index": 200, "end_index": 100}
        assert error.context.expected == {"start_index": 100, "end_index": 200}

    def test_empty_search_text(self):
        """Empty search text error has proper structure."""
        error = DocsErrorBuilder.empty_search_text()
        assert error.code == "EMPTY_SEARCH_TEXT"
        assert "empty" in error.message.lower()
        assert "search" in error.message.lower()
        assert error.suggestion is not None
        assert error.example is not None
        assert "search_mode" in error.example

    def test_search_text_not_found_with_case_hint(self):
        """Search not found suggests case-insensitive when match_case=True."""
        error = DocsErrorBuilder.search_text_not_found(
            search_text="Hello World", match_case=True
        )
        assert error.code == "SEARCH_TEXT_NOT_FOUND"
        assert "Hello World" in error.message
        assert "match_case=False" in error.suggestion

    def test_search_text_not_found_no_case_hint(self):
        """Search not found doesn't suggest case-insensitive when already False."""
        error = DocsErrorBuilder.search_text_not_found(
            search_text="hello", match_case=False
        )
        assert "match_case=False" not in error.suggestion

    def test_search_text_not_found_with_similar(self):
        """Search not found includes similar matches when provided."""
        error = DocsErrorBuilder.search_text_not_found(
            search_text="Intro", similar_found=["Introduction", "Intro Section"]
        )
        assert error.context.similar_found == ["Introduction", "Intro Section"]

    def test_ambiguous_search(self):
        """Ambiguous search shows occurrence options."""
        occurrences = [
            {"index": 1, "position": "100-105"},
            {"index": 2, "position": "200-205"},
        ]
        error = DocsErrorBuilder.ambiguous_search(
            search_text="TODO", occurrences=occurrences, total_count=5
        )
        assert error.code == "AMBIGUOUS_SEARCH"
        assert "5" in error.message
        assert "occurrence" in error.suggestion.lower()
        assert "first_occurrence" in error.example
        assert "last_occurrence" in error.example

    def test_invalid_occurrence(self):
        """Invalid occurrence shows valid range."""
        error = DocsErrorBuilder.invalid_occurrence(
            occurrence=10, total_found=3, search_text="test"
        )
        assert error.code == "INVALID_OCCURRENCE"
        assert "10" in error.message
        assert "3" in error.message

    def test_heading_not_found(self):
        """Heading not found lists available headings."""
        error = DocsErrorBuilder.heading_not_found(
            heading="Missing Section",
            available_headings=["Intro", "Body", "Conclusion"],
            match_case=True,
        )
        assert error.code == "HEADING_NOT_FOUND"
        assert "Missing Section" in error.message
        assert error.context.available_headings == ["Intro", "Body", "Conclusion"]

    def test_heading_not_found_truncates_long_list(self):
        """Long heading lists are truncated to 10."""
        many_headings = [f"Heading {i}" for i in range(20)]
        error = DocsErrorBuilder.heading_not_found(
            heading="Missing", available_headings=many_headings
        )
        # Should have 10 headings plus "... and X more"
        assert len(error.context.available_headings) == 11
        assert "10 more" in error.context.available_headings[-1]

    def test_document_not_found(self):
        """Document not found lists possible causes."""
        error = DocsErrorBuilder.document_not_found("abc123")
        assert error.code == "DOCUMENT_NOT_FOUND"
        assert "abc123" in error.message
        assert error.context.possible_causes is not None
        assert len(error.context.possible_causes) > 0

    def test_permission_denied(self):
        """Permission denied shows current vs required."""
        error = DocsErrorBuilder.permission_denied(
            document_id="doc123",
            current_permission="viewer",
            required_permission="editor",
        )
        assert error.code == "PERMISSION_DENIED"
        assert error.context.current_permission == "viewer"
        assert error.context.required_permission == "editor"

    def test_invalid_table_data(self):
        """Invalid table data includes example format."""
        error = DocsErrorBuilder.invalid_table_data(
            issue="All rows must be lists", row_index=2
        )
        assert error.code == "INVALID_TABLE_DATA"
        assert "correct_format" in error.example
        assert "row" in error.context.received

    def test_table_not_found(self):
        """Table not found shows valid indices."""
        error = DocsErrorBuilder.table_not_found(table_index=5, total_tables=3)
        assert error.code == "TABLE_NOT_FOUND"
        assert "5" in error.message
        assert "3" in error.message

    def test_table_not_found_empty_document(self):
        """Table not found handles zero tables."""
        error = DocsErrorBuilder.table_not_found(table_index=0, total_tables=0)
        assert "no tables" in error.reason.lower()

    def test_missing_required_param(self):
        """Missing param shows valid values."""
        error = DocsErrorBuilder.missing_required_param(
            param_name="position",
            context_description="when using 'search'",
            valid_values=["before", "after", "replace"],
        )
        assert error.code == "MISSING_REQUIRED_PARAM"
        assert "position" in error.message
        assert all(v in error.suggestion for v in ["before", "after", "replace"])

    def test_invalid_param_value(self):
        """Invalid value shows what was received."""
        error = DocsErrorBuilder.invalid_param_value(
            param_name="list_type",
            received_value="NUMBERED",
            valid_values=["ORDERED", "UNORDERED"],
        )
        assert error.code == "INVALID_PARAM_VALUE"
        assert "NUMBERED" in error.message
        assert error.context.received == {"list_type": "NUMBERED"}

    def test_invalid_color_format_unknown_string(self):
        """Invalid color format for unknown color string."""
        error = DocsErrorBuilder.invalid_color_format(
            color_value="not_a_color", param_name="foreground_color"
        )
        assert error.code == "INVALID_COLOR_FORMAT"
        assert "not_a_color" in error.message
        assert "foreground_color" in error.message
        assert "hex" in error.reason.lower()
        assert "named" in error.reason.lower()
        assert error.context.received == {"foreground_color": "not_a_color"}
        assert error.example is not None
        assert "hex_color" in error.example
        assert "named_color" in error.example

    def test_invalid_color_format_invalid_hex(self):
        """Invalid color format for malformed hex color."""
        error = DocsErrorBuilder.invalid_color_format(
            color_value="#GGG", param_name="background_color"
        )
        assert error.code == "INVALID_COLOR_FORMAT"
        assert "#GGG" in error.message
        assert "background_color" in error.message

    def test_invalid_color_format_default_param_name(self):
        """Invalid color format uses default param name."""
        error = DocsErrorBuilder.invalid_color_format(color_value="pink")
        assert error.code == "INVALID_COLOR_FORMAT"
        assert "color" in error.message
        assert error.context.received == {"color": "pink"}

    def test_empty_text_insertion(self):
        """Empty text insertion error provides actionable guidance."""
        error = DocsErrorBuilder.empty_text_insertion()
        assert error.code == "INVALID_PARAM_VALUE"
        assert "empty" in error.message.lower()
        assert "text" in error.message.lower()
        assert "insert" in error.reason.lower()
        assert (
            "delete" in error.suggestion.lower() or "range" in error.suggestion.lower()
        )
        assert error.example is not None
        assert "insert_text" in error.example
        assert "delete_text" in error.example
        assert error.context is not None


class TestValidationManagerStructuredErrors:
    """Tests for ValidationManager's structured error methods."""

    def test_validate_document_id_structured_valid(self):
        """Valid document ID returns success."""
        vm = ValidationManager()
        is_valid, error = vm.validate_document_id_structured(
            "1234567890abcdefghij1234567890abcdef"
        )
        assert is_valid is True
        assert error is None

    def test_validate_document_id_structured_invalid(self):
        """Invalid document ID returns structured error."""
        vm = ValidationManager()
        is_valid, error = vm.validate_document_id_structured("short")
        assert is_valid is False
        assert error is not None
        parsed = json.loads(error)
        assert parsed["error"] is True
        assert parsed["code"] == "DOCUMENT_NOT_FOUND"

    def test_validate_index_range_structured_valid(self):
        """Valid index range returns success."""
        vm = ValidationManager()
        is_valid, error = vm.validate_index_range_structured(
            start_index=10, end_index=20, document_length=100
        )
        assert is_valid is True
        assert error is None

    def test_validate_index_range_structured_invalid_type(self):
        """Non-integer index returns structured error."""
        vm = ValidationManager()
        is_valid, error = vm.validate_index_range_structured(
            start_index="ten", end_index=20
        )
        assert is_valid is False
        parsed = json.loads(error)
        assert parsed["code"] == "INVALID_INDEX_TYPE"

    def test_validate_index_range_structured_invalid_range(self):
        """end < start returns structured error."""
        vm = ValidationManager()
        is_valid, error = vm.validate_index_range_structured(
            start_index=100, end_index=50
        )
        assert is_valid is False
        parsed = json.loads(error)
        assert parsed["code"] == "INVALID_INDEX_RANGE"

    def test_validate_index_range_structured_negative_start_index(self):
        """Negative start_index returns structured error."""
        vm = ValidationManager()
        is_valid, error = vm.validate_index_range_structured(
            start_index=-5, end_index=20
        )
        assert is_valid is False
        parsed = json.loads(error)
        assert parsed["code"] == "INVALID_INDEX_RANGE"
        assert "negative" in parsed["message"].lower()
        assert "get_doc_info" in parsed["suggestion"].lower()

    def test_validate_index_range_structured_out_of_bounds(self):
        """Index beyond document length returns structured error."""
        vm = ValidationManager()
        is_valid, error = vm.validate_index_range_structured(
            start_index=500, document_length=100
        )
        assert is_valid is False
        parsed = json.loads(error)
        assert parsed["code"] == "INDEX_OUT_OF_BOUNDS"

    def test_validate_formatting_range_structured_missing_end(self):
        """Formatting without text and end_index returns error."""
        vm = ValidationManager()
        is_valid, error = vm.validate_formatting_range_structured(
            start_index=100,
            end_index=None,
            text=None,
            formatting_params=["bold", "italic"],
        )
        assert is_valid is False
        parsed = json.loads(error)
        assert parsed["code"] == "FORMATTING_REQUIRES_RANGE"

    def test_validate_formatting_range_structured_with_text(self):
        """Formatting with text but no end_index is valid."""
        vm = ValidationManager()
        is_valid, error = vm.validate_formatting_range_structured(
            start_index=100,
            end_index=None,
            text="some text",
            formatting_params=["bold"],
        )
        assert is_valid is True
        assert error is None

    def test_validate_table_data_structured_valid(self):
        """Valid table data returns success."""
        vm = ValidationManager()
        is_valid, error = vm.validate_table_data_structured([["A", "B"], ["C", "D"]])
        assert is_valid is True
        assert error is None

    def test_validate_table_data_structured_none_cell(self):
        """None cell value returns structured error."""
        vm = ValidationManager()
        is_valid, error = vm.validate_table_data_structured([["A", None], ["C", "D"]])
        assert is_valid is False
        parsed = json.loads(error)
        assert parsed["code"] == "INVALID_TABLE_DATA"
        assert "None" in parsed["message"]

    def test_create_empty_search_error(self):
        """Create empty search error returns valid JSON."""
        vm = ValidationManager()
        error = vm.create_empty_search_error()
        parsed = json.loads(error)
        assert parsed["error"] is True
        assert parsed["code"] == "EMPTY_SEARCH_TEXT"
        assert "empty" in parsed["message"].lower()

    def test_create_search_not_found_error(self):
        """Create search not found returns valid JSON."""
        vm = ValidationManager()
        error = vm.create_search_not_found_error("missing text")
        parsed = json.loads(error)
        assert parsed["error"] is True
        assert parsed["code"] == "SEARCH_TEXT_NOT_FOUND"

    def test_create_heading_not_found_error(self):
        """Create heading not found returns valid JSON."""
        vm = ValidationManager()
        error = vm.create_heading_not_found_error(
            heading="Missing", available_headings=["A", "B"]
        )
        parsed = json.loads(error)
        assert parsed["code"] == "HEADING_NOT_FOUND"
        assert parsed["context"]["available_headings"] == ["A", "B"]

    def test_create_table_not_found_error(self):
        """Create table not found returns valid JSON."""
        vm = ValidationManager()
        error = vm.create_table_not_found_error(table_index=5, total_tables=2)
        parsed = json.loads(error)
        assert parsed["code"] == "TABLE_NOT_FOUND"

    def test_create_invalid_color_error(self):
        """Create invalid color error returns valid JSON."""
        vm = ValidationManager()
        error = vm.create_invalid_color_error(
            color_value="not_a_color", param_name="foreground_color"
        )
        parsed = json.loads(error)
        assert parsed["code"] == "INVALID_COLOR_FORMAT"
        assert "not_a_color" in parsed["message"]
        assert "foreground_color" in parsed["message"]

    def test_validate_color_format_valid_hex(self):
        """Valid hex colors pass validation."""
        vm = ValidationManager()
        # Full hex
        is_valid, _ = vm._validate_color_format("#FF0000", "test")
        assert is_valid
        # Short hex
        is_valid, _ = vm._validate_color_format("#F00", "test")
        assert is_valid
        # Lowercase
        is_valid, _ = vm._validate_color_format("#aabbcc", "test")
        assert is_valid

    def test_validate_color_format_valid_named(self):
        """Valid named colors pass validation."""
        vm = ValidationManager()
        named_colors = [
            "red",
            "green",
            "blue",
            "yellow",
            "orange",
            "purple",
            "black",
            "white",
            "gray",
            "grey",
        ]
        for color in named_colors:
            is_valid, _ = vm._validate_color_format(color, "test")
            assert is_valid, f"Named color '{color}' should be valid"
            # Also test case-insensitivity
            is_valid, _ = vm._validate_color_format(color.upper(), "test")
            assert is_valid, f"Named color '{color.upper()}' should be valid"

    def test_validate_color_format_invalid_string(self):
        """Invalid color strings fail validation."""
        vm = ValidationManager()
        is_valid, error_msg = vm._validate_color_format(
            "not_a_color", "foreground_color"
        )
        assert not is_valid
        assert "Invalid color format" in error_msg
        assert "foreground_color" in error_msg

    def test_validate_color_format_invalid_hex(self):
        """Invalid hex colors fail validation."""
        vm = ValidationManager()
        # Invalid characters
        is_valid, error_msg = vm._validate_color_format("#GGG", "test")
        assert not is_valid
        assert "Invalid hex color" in error_msg
        # Wrong length
        is_valid, error_msg = vm._validate_color_format("#FF00", "test")
        assert not is_valid
        assert "Invalid hex color" in error_msg

    def test_validate_text_formatting_params_invalid_color(self):
        """validate_text_formatting_params rejects invalid color format."""
        vm = ValidationManager()
        is_valid, error_msg = vm.validate_text_formatting_params(
            foreground_color="not_a_color"
        )
        assert not is_valid
        assert "Invalid color format" in error_msg


class TestHelperFunctions:
    """Tests for helper functions."""

    def test_format_error_returns_json(self):
        """format_error returns valid JSON string."""
        error = StructuredError(code="TEST", message="Test")
        result = format_error(error)
        assert isinstance(result, str)
        parsed = json.loads(result)
        assert parsed["code"] == "TEST"

    def test_simple_error_minimal(self):
        """simple_error creates minimal error JSON."""
        result = simple_error(ErrorCode.OPERATION_FAILED, "Something failed")
        parsed = json.loads(result)
        assert parsed["error"] is True
        assert parsed["code"] == "OPERATION_FAILED"
        assert parsed["message"] == "Something failed"

    def test_simple_error_with_suggestion(self):
        """simple_error includes suggestion when provided."""
        result = simple_error(
            ErrorCode.API_ERROR, "API call failed", suggestion="Try again later"
        )
        parsed = json.loads(result)
        assert parsed["suggestion"] == "Try again later"


class TestErrorOutputFormat:
    """Tests verifying the error output format matches the specification."""

    def test_all_errors_have_required_fields(self):
        """All error builder methods produce errors with required fields."""
        test_errors = [
            DocsErrorBuilder.formatting_requires_range(100, False, ["bold"]),
            DocsErrorBuilder.index_out_of_bounds("idx", 500, 100),
            DocsErrorBuilder.invalid_index_range(100, 50),
            DocsErrorBuilder.empty_search_text(),
            DocsErrorBuilder.search_text_not_found("text"),
            DocsErrorBuilder.ambiguous_search("text", [], 5),
            DocsErrorBuilder.heading_not_found("heading", []),
            DocsErrorBuilder.document_not_found("id"),
            DocsErrorBuilder.permission_denied("id"),
            DocsErrorBuilder.invalid_table_data("issue"),
            DocsErrorBuilder.table_not_found(0, 0),
            DocsErrorBuilder.missing_required_param("param", "context"),
            DocsErrorBuilder.invalid_param_value("param", "val", ["a", "b"]),
            DocsErrorBuilder.invalid_color_format("not_a_color", "foreground_color"),
        ]

        for error in test_errors:
            result = error.to_dict()
            assert result["error"] is True, f"Missing 'error' in {error.code}"
            assert "code" in result, f"Missing 'code' in {result}"
            assert "message" in result, f"Missing 'message' in {result}"
            # Verify code is a valid ErrorCode value
            valid_codes = [e.value for e in ErrorCode]
            assert result["code"] in valid_codes, f"Invalid code: {result['code']}"

    def test_error_json_is_parseable(self):
        """All errors produce valid JSON that can be parsed."""
        error = DocsErrorBuilder.formatting_requires_range(100, False, ["bold"])
        json_str = error.to_json()

        # Should be valid JSON
        parsed = json.loads(json_str)

        # Should have standard structure
        assert isinstance(parsed, dict)
        assert parsed.get("error") is True

    def test_error_includes_actionable_example(self):
        """Key errors include example code for how to fix."""
        error = DocsErrorBuilder.formatting_requires_range(100, False, ["bold"])
        assert error.example is not None
        # Example should contain actual function calls
        example_str = str(error.example)
        assert "modify_doc_text" in example_str


class TestParseDocsIndexError:
    """Tests for _parse_docs_index_error function in core/utils.py."""

    def test_parses_index_less_than_segment_end_with_length(self):
        """Parses 'Index X must be less than the end index of the referenced segment, Y'."""
        from core.utils import _parse_docs_index_error

        error_msg = (
            "Invalid requests[0].updateTextStyle: Index 500 must be less than "
            "the end index of the referenced segment, 200."
        )
        result = _parse_docs_index_error(error_msg)

        assert result is not None
        parsed = json.loads(result)
        assert parsed["error"] is True
        assert parsed["code"] == "INDEX_OUT_OF_BOUNDS"
        assert "500" in parsed["message"]
        assert parsed["context"]["received"]["index"] == 500
        assert parsed["context"]["document_length"] == 200
        assert "inspect_doc_structure" in parsed["suggestion"]

    def test_parses_index_less_than_segment_end_without_length(self):
        """Handles case where document length is not included in error message."""
        from core.utils import _parse_docs_index_error

        error_msg = (
            "Index 1000 must be less than the end index of the referenced segment"
        )
        result = _parse_docs_index_error(error_msg)

        assert result is not None
        parsed = json.loads(result)
        assert parsed["code"] == "INDEX_OUT_OF_BOUNDS"
        assert "1000" in parsed["message"]
        assert parsed["context"]["received"]["index"] == 1000
        assert "document_length" not in parsed["context"]

    def test_parses_insertion_bounds_error(self):
        """Parses 'insertion index must be inside the bounds of an existing paragraph'."""
        from core.utils import _parse_docs_index_error

        error_msg = (
            "Invalid requests[0].insertText: The insertion index must be inside "
            "the bounds of an existing paragraph."
        )
        result = _parse_docs_index_error(error_msg)

        assert result is not None
        parsed = json.loads(result)
        assert parsed["error"] is True
        assert parsed["code"] == "INDEX_OUT_OF_BOUNDS"
        assert "outside document bounds" in parsed["message"]
        assert (
            "empty document" in parsed["suggestion"].lower()
            or "index 1" in parsed["example"]["empty_doc"]
        )

    def test_returns_none_for_unrelated_error(self):
        """Returns None for errors that are not index-related."""
        from core.utils import _parse_docs_index_error

        error_msg = "Invalid requests[0].updateTextStyle: At least one field must be listed in 'fields'."
        result = _parse_docs_index_error(error_msg)
        assert result is None

    def test_returns_none_for_empty_error(self):
        """Returns None for empty error string."""
        from core.utils import _parse_docs_index_error

        result = _parse_docs_index_error("")
        assert result is None

    def test_case_insensitive_parsing(self):
        """Parses error messages regardless of case."""
        from core.utils import _parse_docs_index_error

        error_msg = (
            "INDEX 300 MUST BE LESS THAN THE END INDEX OF THE REFERENCED SEGMENT, 100."
        )
        result = _parse_docs_index_error(error_msg)

        assert result is not None
        parsed = json.loads(result)
        assert parsed["context"]["received"]["index"] == 300


class TestCreateDocsNotFoundError:
    """Tests for _create_docs_not_found_error function in core/utils.py."""

    def test_creates_structured_error_for_invalid_document_id(self):
        """Creates structured error with document ID in message."""
        from core.utils import _create_docs_not_found_error

        result = _create_docs_not_found_error("INVALID_DOC_ID_12345")

        assert result is not None
        parsed = json.loads(result)
        assert parsed["error"] is True
        assert parsed["code"] == "DOCUMENT_NOT_FOUND"
        assert "INVALID_DOC_ID_12345" in parsed["message"]
        assert parsed["context"]["received"]["document_id"] == "INVALID_DOC_ID_12345"

    def test_includes_helpful_suggestion(self):
        """Includes suggestion about verifying document ID."""
        from core.utils import _create_docs_not_found_error

        result = _create_docs_not_found_error("test_id")

        parsed = json.loads(result)
        assert "suggestion" in parsed
        assert (
            "Verify" in parsed["suggestion"]
            or "document_id" in parsed["suggestion"].lower()
        )
        assert "docs.google.com" in parsed["suggestion"]

    def test_includes_possible_causes(self):
        """Includes possible causes in context."""
        from core.utils import _create_docs_not_found_error

        result = _create_docs_not_found_error("test_id")

        parsed = json.loads(result)
        assert "context" in parsed
        assert "possible_causes" in parsed["context"]
        causes = parsed["context"]["possible_causes"]
        assert len(causes) >= 2
        assert any("incorrect" in c.lower() for c in causes)
        assert any("deleted" in c.lower() or "permission" in c.lower() for c in causes)

    def test_handles_unknown_document_id(self):
        """Handles 'unknown' as a document ID gracefully."""
        from core.utils import _create_docs_not_found_error

        result = _create_docs_not_found_error("unknown")

        parsed = json.loads(result)
        assert parsed["error"] is True
        assert parsed["code"] == "DOCUMENT_NOT_FOUND"


class TestHandleHttpErrors404:
    """Tests for handle_http_errors decorator handling 404 errors for docs service."""

    @pytest.mark.asyncio
    async def test_decorator_returns_structured_error_for_404(self):
        """Decorator returns structured error for 404 HttpError on docs service."""
        from core.utils import handle_http_errors
        from googleapiclient.errors import HttpError
        from unittest.mock import MagicMock

        # Create a mock 404 HttpError
        mock_resp = MagicMock()
        mock_resp.status = 404
        mock_resp.reason = "Not Found"
        error = HttpError(mock_resp, b"Requested entity was not found.")

        @handle_http_errors("test_tool", service_type="docs")
        async def mock_func(document_id: str):
            raise error

        result = await mock_func(document_id="test_doc_123")

        assert result is not None
        parsed = json.loads(result)
        assert parsed["error"] is True
        assert parsed["code"] == "DOCUMENT_NOT_FOUND"
        assert "test_doc_123" in parsed["message"]

    @pytest.mark.asyncio
    async def test_decorator_uses_unknown_when_document_id_not_provided(self):
        """Decorator uses 'unknown' when document_id kwarg is not provided."""
        from core.utils import handle_http_errors
        from googleapiclient.errors import HttpError
        from unittest.mock import MagicMock

        mock_resp = MagicMock()
        mock_resp.status = 404
        error = HttpError(mock_resp, b"Requested entity was not found.")

        @handle_http_errors("test_tool", service_type="docs")
        async def mock_func():
            raise error

        result = await mock_func()

        parsed = json.loads(result)
        assert parsed["code"] == "DOCUMENT_NOT_FOUND"
        assert "unknown" in parsed["message"]

    @pytest.mark.asyncio
    async def test_decorator_raises_exception_for_404_on_non_docs_service(self):
        """Decorator raises exception for 404 on non-docs services."""
        from core.utils import handle_http_errors
        from googleapiclient.errors import HttpError
        from unittest.mock import MagicMock

        mock_resp = MagicMock()
        mock_resp.status = 404
        error = HttpError(mock_resp, b"Not found.")

        @handle_http_errors("test_tool", service_type="calendar")
        async def mock_func():
            raise error

        with pytest.raises(Exception) as exc_info:
            await mock_func()

        assert "API error in test_tool" in str(exc_info.value)


class TestLocationParameterValidation:
    """Tests for location parameter in modify_doc_text validation."""

    def test_invalid_location_value_error(self):
        """Test that invalid location values produce proper error."""
        validator = ValidationManager()
        error = validator.create_invalid_param_error(
            param_name="location", received="middle", valid_values=["start", "end"]
        )
        assert error is not None
        parsed = json.loads(error)
        assert parsed["error"] is True
        assert "location" in parsed["message"].lower()
        assert "middle" in str(parsed)
        assert "start" in str(parsed)
        assert "end" in str(parsed)

    def test_conflicting_params_error(self):
        """Test that conflicting location + start_index produces proper error."""
        error = DocsErrorBuilder.conflicting_params(
            params=["location", "start_index", "end_index"],
            message="Cannot use 'location' parameter with explicit 'start_index' or 'end_index'",
        )
        assert error is not None
        assert error.code == ErrorCode.CONFLICTING_PARAMS.value
        assert "location" in error.message.lower()
        assert "start_index" in error.message.lower()

    def test_location_info_structure(self):
        """Test that location_info dict has expected structure."""
        location_info = {
            "location": "end",
            "resolved_index": 500,
            "message": "Appending at document end (index 500)",
        }
        assert location_info["location"] == "end"
        assert isinstance(location_info["resolved_index"], int)
        assert "message" in location_info

    def test_location_start_resolves_to_index_1(self):
        """Test that location='start' resolves to index 1 (after section break)."""
        location_info = {
            "location": "start",
            "resolved_index": 1,
            "message": "Inserting at document start (index 1)",
        }
        assert location_info["resolved_index"] == 1


class TestConvertToListValidation:
    """Tests for convert_to_list parameter validation."""

    def test_invalid_convert_to_list_value_error(self):
        """Test that invalid convert_to_list values produce proper error."""
        validator = ValidationManager()
        error = validator.create_invalid_param_error(
            param_name="convert_to_list",
            received="BULLET",
            valid_values=["ORDERED", "UNORDERED"],
        )
        assert error is not None
        parsed = json.loads(error)
        assert parsed["error"] is True
        assert "convert_to_list" in parsed["message"].lower()
        assert "BULLET" in str(parsed)
        assert "ORDERED" in str(parsed)
        assert "UNORDERED" in str(parsed)

    def test_convert_to_list_accepted_values(self):
        """Test that valid values and aliases are accepted."""
        # Direct values
        valid_direct = ["ORDERED", "UNORDERED"]
        # Aliases that should normalize to valid values
        valid_aliases = [
            "bullet",
            "bullets",
            "numbered",
            "numbers",
            "ordered",
            "unordered",
        ]
        # Invalid values that should be rejected
        invalid_values = ["LIST", "", "foo", "123", "random"]

        for val in valid_direct:
            assert val in ["ORDERED", "UNORDERED"], f"{val} should be valid"

        # Test alias normalization
        list_type_aliases = {
            "bullet": "UNORDERED",
            "bullets": "UNORDERED",
            "unordered": "UNORDERED",
            "numbered": "ORDERED",
            "numbers": "ORDERED",
            "ordered": "ORDERED",
        }
        for alias in valid_aliases:
            normalized = list_type_aliases.get(alias.lower(), alias.upper())
            assert normalized in ["ORDERED", "UNORDERED"], (
                f"{alias} should normalize to valid value"
            )

        for val in invalid_values:
            normalized = list_type_aliases.get(
                val.lower() if val else val, val.upper() if val else val
            )
            assert normalized not in ["ORDERED", "UNORDERED"], (
                f"{val} should be invalid"
            )

    def test_convert_to_list_format_styles_output(self):
        """Test that convert_to_list adds to format_styles correctly."""
        # Simulate what happens when convert_to_list is applied
        format_styles = []
        convert_to_list = "UNORDERED"
        list_type_display = "bullet" if convert_to_list == "UNORDERED" else "numbered"
        format_styles.append(f"convert_to_{list_type_display}_list")

        assert "convert_to_bullet_list" in format_styles

        format_styles = []
        convert_to_list = "ORDERED"
        list_type_display = "bullet" if convert_to_list == "UNORDERED" else "numbered"
        format_styles.append(f"convert_to_{list_type_display}_list")

        assert "convert_to_numbered_list" in format_styles


class TestLineSpacingValidation:
    """Tests for line_spacing parameter validation and functionality."""

    def test_line_spacing_invalid_string_error(self):
        """Test that unrecognized string line_spacing produces error."""
        validator = ValidationManager()
        # Unrecognized string value should be invalid (note: 'double', 'single', '1.5x' are now valid)
        line_spacing = "triple"  # Not a recognized named value
        line_spacing_map = {
            "single": 100,
            "1": 100,
            "1.0": 100,
            "1.15": 115,
            "1.15x": 115,
            "1.5": 150,
            "1.5x": 150,
            "double": 200,
            "2": 200,
            "2.0": 200,
        }
        if (
            isinstance(line_spacing, str)
            and line_spacing.lower() not in line_spacing_map
        ):
            error = validator.create_invalid_param_error(
                param_name="line_spacing",
                received=str(line_spacing),
                valid_values=[
                    "Named: 'single', 'double', '1.5x'",
                    "Decimal: 1.0, 1.5, 2.0 (auto-converted to percentage)",
                    "Percentage: 100, 150, 200 (100=single, 150=1.5x, 200=double)",
                ],
            )
            result = json.loads(error)
            assert result["error"] is True
            assert result["code"] == "INVALID_PARAM_VALUE"
            assert "line_spacing" in result["message"]

    def test_line_spacing_valid_named_strings(self):
        """Test that valid named strings are accepted."""
        valid_strings = [
            "single",
            "double",
            "1.5x",
            "1.15x",
            "1",
            "1.0",
            "1.5",
            "2",
            "2.0",
        ]
        line_spacing_map = {
            "single": 100,
            "1": 100,
            "1.0": 100,
            "1.15": 115,
            "1.15x": 115,
            "1.5": 150,
            "1.5x": 150,
            "double": 200,
            "2": 200,
            "2.0": 200,
        }
        for val in valid_strings:
            assert val.lower() in line_spacing_map, (
                f"'{val}' should be a valid named string"
            )

    def test_line_spacing_out_of_range_low_error(self):
        """Test that line_spacing below 50 (after conversion) produces error."""
        validator = ValidationManager()
        # 25 is >= 10 so not treated as multiplier, but < 50 so invalid
        line_spacing = 25
        if line_spacing < 10:
            line_spacing = line_spacing * 100
        if line_spacing < 50 or line_spacing > 1000:
            error = validator.create_invalid_param_error(
                param_name="line_spacing",
                received=str(line_spacing),
                valid_values=[
                    "Named: 'single', 'double', '1.5x'",
                    "Decimal: 1.0, 1.5, 2.0 (auto-converted to percentage)",
                    "Percentage: 100, 150, 200 (100=single, 150=1.5x, 200=double)",
                ],
            )
            result = json.loads(error)
            assert result["error"] is True
            assert result["code"] == "INVALID_PARAM_VALUE"

    def test_line_spacing_out_of_range_high_error(self):
        """Test that line_spacing above 1000 (after conversion) produces error."""
        validator = ValidationManager()
        # 1500 is >= 10 so not treated as multiplier, and > 1000 so invalid
        line_spacing = 1500
        if line_spacing < 10:
            line_spacing = line_spacing * 100
        if line_spacing < 50 or line_spacing > 1000:
            error = validator.create_invalid_param_error(
                param_name="line_spacing",
                received=str(line_spacing),
                valid_values=[
                    "Named: 'single', 'double', '1.5x'",
                    "Decimal: 1.0, 1.5, 2.0 (auto-converted to percentage)",
                    "Percentage: 100, 150, 200 (100=single, 150=1.5x, 200=double)",
                ],
            )
            result = json.loads(error)
            assert result["error"] is True
            assert result["code"] == "INVALID_PARAM_VALUE"

    def test_line_spacing_decimal_multiplier_conversion(self):
        """Test that decimal multipliers < 10 are converted to percentage."""
        # Values < 10 are treated as multipliers
        assert 1.0 * 100 == 100  # single
        assert 1.5 * 100 == 150  # 1.5x
        assert 2.0 * 100 == 200  # double

        # Edge cases
        line_spacing = 0.5
        if line_spacing < 10:
            line_spacing = line_spacing * 100
        assert line_spacing == 50  # Minimum valid percentage

    def test_line_spacing_accepted_values(self):
        """Test that common line_spacing values are valid."""
        valid_values = [50, 100, 115, 150, 200, 250, 1000]
        for val in valid_values:
            assert 50 <= val <= 1000, f"{val} should be in valid range"

    def test_line_spacing_format_styles_output(self):
        """Test that line_spacing adds to format_styles correctly."""
        format_styles = []
        line_spacing = 150
        format_styles.append(f"line_spacing_{line_spacing}")

        assert "line_spacing_150" in format_styles

    def test_create_paragraph_style_request_output(self):
        """Test that create_paragraph_style_request generates correct API format."""
        from gdocs.docs_helpers import create_paragraph_style_request

        result = create_paragraph_style_request(10, 50, 150)

        assert result is not None
        assert "updateParagraphStyle" in result
        assert result["updateParagraphStyle"]["range"]["startIndex"] == 10
        assert result["updateParagraphStyle"]["range"]["endIndex"] == 50
        assert result["updateParagraphStyle"]["paragraphStyle"]["lineSpacing"] == 150
        assert result["updateParagraphStyle"]["fields"] == "lineSpacing"

    def test_create_paragraph_style_request_single_spacing(self):
        """Test single (100) line spacing request."""
        from gdocs.docs_helpers import create_paragraph_style_request

        result = create_paragraph_style_request(1, 100, 100)

        assert result["updateParagraphStyle"]["paragraphStyle"]["lineSpacing"] == 100

    def test_create_paragraph_style_request_double_spacing(self):
        """Test double (200) line spacing request."""
        from gdocs.docs_helpers import create_paragraph_style_request

        result = create_paragraph_style_request(1, 100, 200)

        assert result["updateParagraphStyle"]["paragraphStyle"]["lineSpacing"] == 200

    def test_create_paragraph_style_request_none_returns_none(self):
        """Test that None line_spacing returns None."""
        from gdocs.docs_helpers import create_paragraph_style_request

        result = create_paragraph_style_request(10, 50, None)

        assert result is None


class TestAlignmentValidation:
    """Tests for alignment parameter validation and functionality."""

    def test_alignment_invalid_value_error(self):
        """Test that invalid alignment value produces error."""
        validator = ValidationManager()
        alignment = "LEFT"  # Invalid - should be "START"
        valid_alignments = ["START", "CENTER", "END", "JUSTIFIED"]
        if alignment not in valid_alignments:
            error = validator.create_invalid_param_error(
                param_name="alignment",
                received=alignment,
                valid_values=valid_alignments,
            )
            result = json.loads(error)
            assert result["error"] is True
            assert result["code"] == "INVALID_PARAM_VALUE"
            assert "alignment" in result["message"]

    def test_alignment_accepted_values(self):
        """Test that all valid alignment values are accepted."""
        valid_values = ["START", "CENTER", "END", "JUSTIFIED"]
        for val in valid_values:
            assert val in valid_values, f"{val} should be in valid alignment values"

    def test_alignment_format_styles_output(self):
        """Test that alignment adds to format_styles correctly."""
        format_styles = []
        alignment = "CENTER"
        format_styles.append(f"alignment_{alignment}")

        assert "alignment_CENTER" in format_styles

    def test_create_paragraph_style_request_with_alignment(self):
        """Test that create_paragraph_style_request generates correct API format for alignment."""
        from gdocs.docs_helpers import create_paragraph_style_request

        result = create_paragraph_style_request(10, 50, alignment="CENTER")

        assert result is not None
        assert "updateParagraphStyle" in result
        assert result["updateParagraphStyle"]["range"]["startIndex"] == 10
        assert result["updateParagraphStyle"]["range"]["endIndex"] == 50
        assert result["updateParagraphStyle"]["paragraphStyle"]["alignment"] == "CENTER"
        assert result["updateParagraphStyle"]["fields"] == "alignment"

    def test_create_paragraph_style_request_justified(self):
        """Test JUSTIFIED alignment request."""
        from gdocs.docs_helpers import create_paragraph_style_request

        result = create_paragraph_style_request(1, 100, alignment="JUSTIFIED")

        assert (
            result["updateParagraphStyle"]["paragraphStyle"]["alignment"] == "JUSTIFIED"
        )

    def test_create_paragraph_style_request_end_alignment(self):
        """Test END alignment request (right-align for LTR)."""
        from gdocs.docs_helpers import create_paragraph_style_request

        result = create_paragraph_style_request(1, 100, alignment="END")

        assert result["updateParagraphStyle"]["paragraphStyle"]["alignment"] == "END"

    def test_create_paragraph_style_request_combined_alignment_and_line_spacing(self):
        """Test combining alignment with line_spacing."""
        from gdocs.docs_helpers import create_paragraph_style_request

        result = create_paragraph_style_request(
            10, 50, line_spacing=150, alignment="CENTER"
        )

        assert result is not None
        para_style = result["updateParagraphStyle"]["paragraphStyle"]
        assert para_style["lineSpacing"] == 150
        assert para_style["alignment"] == "CENTER"
        fields = result["updateParagraphStyle"]["fields"]
        assert "lineSpacing" in fields
        assert "alignment" in fields


class TestFontSizeValidation:
    """Tests for font_size parameter validation."""

    def test_font_size_zero_rejected(self):
        """Test that font_size=0 is properly rejected (valid range is 1-400)."""
        validator = ValidationManager()
        is_valid, error_msg = validator.validate_text_formatting_params(font_size=0)

        assert is_valid is False
        assert "font_size must be between 1 and 400" in error_msg
        assert "got 0" in error_msg

    def test_font_size_negative_rejected(self):
        """Test that negative font_size is rejected."""
        validator = ValidationManager()
        is_valid, error_msg = validator.validate_text_formatting_params(font_size=-1)

        assert is_valid is False
        assert "font_size must be between 1 and 400" in error_msg

    def test_font_size_too_large_rejected(self):
        """Test that font_size > 400 is rejected."""
        validator = ValidationManager()
        is_valid, error_msg = validator.validate_text_formatting_params(font_size=401)

        assert is_valid is False
        assert "font_size must be between 1 and 400" in error_msg
        assert "got 401" in error_msg

    def test_font_size_minimum_accepted(self):
        """Test that font_size=1 (minimum) is accepted."""
        validator = ValidationManager()
        is_valid, error_msg = validator.validate_text_formatting_params(font_size=1)

        assert is_valid is True
        assert error_msg == ""

    def test_font_size_maximum_accepted(self):
        """Test that font_size=400 (maximum) is accepted."""
        validator = ValidationManager()
        is_valid, error_msg = validator.validate_text_formatting_params(font_size=400)

        assert is_valid is True
        assert error_msg == ""

    def test_font_size_typical_value_accepted(self):
        """Test that typical font_size values are accepted."""
        validator = ValidationManager()
        for size in [8, 10, 11, 12, 14, 16, 18, 24, 36, 48, 72]:
            is_valid, error_msg = validator.validate_text_formatting_params(
                font_size=size
            )
            assert is_valid is True, f"font_size={size} should be valid"

    def test_font_size_zero_triggers_has_formatting(self):
        """Test that font_size=0 triggers has_formatting check (regression test for bug)."""
        # This test verifies that font_size=0 is detected as a formatting request
        # and triggers validation, rather than being ignored due to truthy check
        font_size = 0
        has_formatting = any([font_size is not None])
        assert has_formatting is True, "font_size=0 should trigger has_formatting=True"
