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
            "SEARCH_TEXT_NOT_FOUND",
            "AMBIGUOUS_SEARCH",
            "HEADING_NOT_FOUND",
            "DOCUMENT_NOT_FOUND",
            "PERMISSION_DENIED",
            "TABLE_NOT_FOUND",
            "INVALID_TABLE_DATA",
        ]
        for code in expected_codes:
            assert hasattr(ErrorCode, code), f"Missing error code: {code}"


class TestStructuredError:
    """Tests for StructuredError dataclass."""

    def test_basic_error_creation(self):
        """Can create a basic structured error."""
        error = StructuredError(
            code="TEST_ERROR",
            message="Test message",
            suggestion="Test suggestion"
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
            code="TEST",
            message="Test message",
            suggestion="Do something"
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
            context=ErrorContext(
                document_length=1000,
                received={"start_index": 500}
            )
        )
        result = error.to_dict()
        assert "context" in result
        assert result["context"]["document_length"] == 1000


class TestDocsErrorBuilder:
    """Tests for DocsErrorBuilder factory methods."""

    def test_formatting_requires_range_no_text(self):
        """Error for formatting without text or end_index."""
        error = DocsErrorBuilder.formatting_requires_range(
            start_index=100,
            has_text=False,
            formatting_params=["bold", "italic"]
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
            start_index=100,
            has_text=True,
            formatting_params=["bold"]
        )
        assert "correct_usage" in error.example

    def test_index_out_of_bounds(self):
        """Index out of bounds provides document length context."""
        error = DocsErrorBuilder.index_out_of_bounds(
            index_name="start_index",
            index_value=5000,
            document_length=3500
        )
        assert error.code == "INDEX_OUT_OF_BOUNDS"
        assert "5000" in error.message
        assert "3500" in error.message
        assert error.context.document_length == 3500
        assert "inspect_doc_structure" in error.suggestion.lower()

    def test_invalid_index_range(self):
        """Invalid range shows expected values."""
        error = DocsErrorBuilder.invalid_index_range(
            start_index=200,
            end_index=100
        )
        assert error.code == "INVALID_INDEX_RANGE"
        assert error.context.received == {"start_index": 200, "end_index": 100}
        assert error.context.expected == {"start_index": 100, "end_index": 200}

    def test_search_text_not_found_with_case_hint(self):
        """Search not found suggests case-insensitive when match_case=True."""
        error = DocsErrorBuilder.search_text_not_found(
            search_text="Hello World",
            match_case=True
        )
        assert error.code == "SEARCH_TEXT_NOT_FOUND"
        assert "Hello World" in error.message
        assert "match_case=False" in error.suggestion

    def test_search_text_not_found_no_case_hint(self):
        """Search not found doesn't suggest case-insensitive when already False."""
        error = DocsErrorBuilder.search_text_not_found(
            search_text="hello",
            match_case=False
        )
        assert "match_case=False" not in error.suggestion

    def test_search_text_not_found_with_similar(self):
        """Search not found includes similar matches when provided."""
        error = DocsErrorBuilder.search_text_not_found(
            search_text="Intro",
            similar_found=["Introduction", "Intro Section"]
        )
        assert error.context.similar_found == ["Introduction", "Intro Section"]

    def test_ambiguous_search(self):
        """Ambiguous search shows occurrence options."""
        occurrences = [
            {"index": 1, "position": "100-105"},
            {"index": 2, "position": "200-205"},
        ]
        error = DocsErrorBuilder.ambiguous_search(
            search_text="TODO",
            occurrences=occurrences,
            total_count=5
        )
        assert error.code == "AMBIGUOUS_SEARCH"
        assert "5" in error.message
        assert "occurrence" in error.suggestion.lower()
        assert "first_occurrence" in error.example
        assert "last_occurrence" in error.example

    def test_invalid_occurrence(self):
        """Invalid occurrence shows valid range."""
        error = DocsErrorBuilder.invalid_occurrence(
            occurrence=10,
            total_found=3,
            search_text="test"
        )
        assert error.code == "INVALID_OCCURRENCE"
        assert "10" in error.message
        assert "3" in error.message

    def test_heading_not_found(self):
        """Heading not found lists available headings."""
        error = DocsErrorBuilder.heading_not_found(
            heading="Missing Section",
            available_headings=["Intro", "Body", "Conclusion"],
            match_case=True
        )
        assert error.code == "HEADING_NOT_FOUND"
        assert "Missing Section" in error.message
        assert error.context.available_headings == ["Intro", "Body", "Conclusion"]

    def test_heading_not_found_truncates_long_list(self):
        """Long heading lists are truncated to 10."""
        many_headings = [f"Heading {i}" for i in range(20)]
        error = DocsErrorBuilder.heading_not_found(
            heading="Missing",
            available_headings=many_headings
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
            required_permission="editor"
        )
        assert error.code == "PERMISSION_DENIED"
        assert error.context.current_permission == "viewer"
        assert error.context.required_permission == "editor"

    def test_invalid_table_data(self):
        """Invalid table data includes example format."""
        error = DocsErrorBuilder.invalid_table_data(
            issue="All rows must be lists",
            row_index=2
        )
        assert error.code == "INVALID_TABLE_DATA"
        assert "correct_format" in error.example
        assert "row" in error.context.received

    def test_table_not_found(self):
        """Table not found shows valid indices."""
        error = DocsErrorBuilder.table_not_found(
            table_index=5,
            total_tables=3
        )
        assert error.code == "TABLE_NOT_FOUND"
        assert "5" in error.message
        assert "3" in error.message

    def test_table_not_found_empty_document(self):
        """Table not found handles zero tables."""
        error = DocsErrorBuilder.table_not_found(
            table_index=0,
            total_tables=0
        )
        assert "no tables" in error.reason.lower()

    def test_missing_required_param(self):
        """Missing param shows valid values."""
        error = DocsErrorBuilder.missing_required_param(
            param_name="position",
            context_description="when using 'search'",
            valid_values=["before", "after", "replace"]
        )
        assert error.code == "MISSING_REQUIRED_PARAM"
        assert "position" in error.message
        assert all(v in error.suggestion for v in ["before", "after", "replace"])

    def test_invalid_param_value(self):
        """Invalid value shows what was received."""
        error = DocsErrorBuilder.invalid_param_value(
            param_name="list_type",
            received_value="NUMBERED",
            valid_values=["ORDERED", "UNORDERED"]
        )
        assert error.code == "INVALID_PARAM_VALUE"
        assert "NUMBERED" in error.message
        assert error.context.received == {"list_type": "NUMBERED"}


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
            start_index=10,
            end_index=20,
            document_length=100
        )
        assert is_valid is True
        assert error is None

    def test_validate_index_range_structured_invalid_type(self):
        """Non-integer index returns structured error."""
        vm = ValidationManager()
        is_valid, error = vm.validate_index_range_structured(
            start_index="ten",
            end_index=20
        )
        assert is_valid is False
        parsed = json.loads(error)
        assert parsed["code"] == "INVALID_INDEX_TYPE"

    def test_validate_index_range_structured_invalid_range(self):
        """end < start returns structured error."""
        vm = ValidationManager()
        is_valid, error = vm.validate_index_range_structured(
            start_index=100,
            end_index=50
        )
        assert is_valid is False
        parsed = json.loads(error)
        assert parsed["code"] == "INVALID_INDEX_RANGE"

    def test_validate_index_range_structured_out_of_bounds(self):
        """Index beyond document length returns structured error."""
        vm = ValidationManager()
        is_valid, error = vm.validate_index_range_structured(
            start_index=500,
            document_length=100
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
            formatting_params=["bold", "italic"]
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
            formatting_params=["bold"]
        )
        assert is_valid is True
        assert error is None

    def test_validate_table_data_structured_valid(self):
        """Valid table data returns success."""
        vm = ValidationManager()
        is_valid, error = vm.validate_table_data_structured([
            ["A", "B"],
            ["C", "D"]
        ])
        assert is_valid is True
        assert error is None

    def test_validate_table_data_structured_none_cell(self):
        """None cell value returns structured error."""
        vm = ValidationManager()
        is_valid, error = vm.validate_table_data_structured([
            ["A", None],
            ["C", "D"]
        ])
        assert is_valid is False
        parsed = json.loads(error)
        assert parsed["code"] == "INVALID_TABLE_DATA"
        assert "None" in parsed["message"]

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
            heading="Missing",
            available_headings=["A", "B"]
        )
        parsed = json.loads(error)
        assert parsed["code"] == "HEADING_NOT_FOUND"
        assert parsed["context"]["available_headings"] == ["A", "B"]

    def test_create_table_not_found_error(self):
        """Create table not found returns valid JSON."""
        vm = ValidationManager()
        error = vm.create_table_not_found_error(
            table_index=5,
            total_tables=2
        )
        parsed = json.loads(error)
        assert parsed["code"] == "TABLE_NOT_FOUND"


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
        result = simple_error(
            ErrorCode.OPERATION_FAILED,
            "Something failed"
        )
        parsed = json.loads(result)
        assert parsed["error"] is True
        assert parsed["code"] == "OPERATION_FAILED"
        assert parsed["message"] == "Something failed"

    def test_simple_error_with_suggestion(self):
        """simple_error includes suggestion when provided."""
        result = simple_error(
            ErrorCode.API_ERROR,
            "API call failed",
            suggestion="Try again later"
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
            DocsErrorBuilder.search_text_not_found("text"),
            DocsErrorBuilder.ambiguous_search("text", [], 5),
            DocsErrorBuilder.heading_not_found("heading", []),
            DocsErrorBuilder.document_not_found("id"),
            DocsErrorBuilder.permission_denied("id"),
            DocsErrorBuilder.invalid_table_data("issue"),
            DocsErrorBuilder.table_not_found(0, 0),
            DocsErrorBuilder.missing_required_param("param", "context"),
            DocsErrorBuilder.invalid_param_value("param", "val", ["a", "b"]),
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
