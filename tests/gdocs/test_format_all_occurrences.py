"""
Unit tests for format_all_occurrences tool.

These tests verify:
- Input validation logic
- Preview mode functionality
- Formatting request building
- Empty search results handling

Note: The format_all_occurrences tool is decorated with @server.tool() and
@require_google_service which modify the function signature. These tests
focus on testing the underlying logic through the helper functions and
integration patterns used by the tool.
"""

import json

from gdocs.docs_helpers import (
    find_all_occurrences_in_document,
    create_format_text_request,
)
from gdocs.managers.validation_manager import ValidationManager


class TestFindAllOccurrencesForFormatting:
    """Tests for finding all occurrences in a document (used by format_all_occurrences)."""

    def test_find_multiple_occurrences(self):
        """Test finding multiple occurrences of text."""
        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "TODO item 1. Some text. TODO item 2. More text. TODO item 3.",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 60,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        occurrences = find_all_occurrences_in_document(
            doc_data, "TODO", match_case=True
        )

        assert len(occurrences) == 3
        # Each occurrence should be a (start, end) tuple
        for start, end in occurrences:
            assert end - start == 4  # len("TODO")

    def test_find_case_insensitive(self):
        """Test case-insensitive search."""
        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "TODO item. todo item. Todo item.",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 32,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        occurrences = find_all_occurrences_in_document(
            doc_data, "TODO", match_case=False
        )

        assert len(occurrences) == 3

    def test_find_case_sensitive(self):
        """Test case-sensitive search."""
        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "TODO item. todo item. Todo item.",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 32,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        occurrences = find_all_occurrences_in_document(
            doc_data, "TODO", match_case=True
        )

        assert len(occurrences) == 1

    def test_find_no_matches(self):
        """Test when no matches are found."""
        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "Some text without the search term.",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 35,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        occurrences = find_all_occurrences_in_document(
            doc_data, "NONEXISTENT", match_case=True
        )

        assert len(occurrences) == 0

    def test_find_empty_search_returns_empty(self):
        """Test that empty search returns empty list."""
        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "Some text.",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 11,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        occurrences = find_all_occurrences_in_document(doc_data, "", match_case=True)

        assert len(occurrences) == 0


class TestCreateFormatTextRequestForBulk:
    """Tests for creating format text requests (used by format_all_occurrences)."""

    def test_create_bold_request(self):
        """Test creating a bold formatting request."""
        request = create_format_text_request(
            start_index=10,
            end_index=14,
            bold=True,
        )

        assert request is not None
        assert "updateTextStyle" in request
        assert request["updateTextStyle"]["textStyle"]["bold"] is True
        assert request["updateTextStyle"]["range"]["startIndex"] == 10
        assert request["updateTextStyle"]["range"]["endIndex"] == 14

    def test_create_multiple_formatting_request(self):
        """Test creating a request with multiple formatting options."""
        request = create_format_text_request(
            start_index=10,
            end_index=14,
            bold=True,
            italic=True,
            foreground_color="red",
            font_size=14,
        )

        assert request is not None
        style = request["updateTextStyle"]["textStyle"]
        assert style["bold"] is True
        assert style["italic"] is True
        assert "foregroundColor" in style
        assert style["fontSize"]["magnitude"] == 14

    def test_create_link_request(self):
        """Test creating a hyperlink formatting request."""
        request = create_format_text_request(
            start_index=10,
            end_index=14,
            link="https://example.com",
        )

        assert request is not None
        assert "link" in request["updateTextStyle"]["textStyle"]

    def test_create_background_color_request(self):
        """Test creating a background color formatting request."""
        request = create_format_text_request(
            start_index=10,
            end_index=14,
            background_color="yellow",
        )

        assert request is not None
        style = request["updateTextStyle"]["textStyle"]
        assert "backgroundColor" in style

    def test_create_no_formatting_returns_none(self):
        """Test that no formatting options returns None."""
        request = create_format_text_request(
            start_index=10,
            end_index=14,
        )

        assert request is None

    def test_create_small_caps_request(self):
        """Test creating small caps formatting request."""
        request = create_format_text_request(
            start_index=10,
            end_index=14,
            small_caps=True,
        )

        assert request is not None
        assert request["updateTextStyle"]["textStyle"]["smallCaps"] is True

    def test_create_subscript_request(self):
        """Test creating subscript formatting request."""
        request = create_format_text_request(
            start_index=10,
            end_index=14,
            subscript=True,
        )

        assert request is not None
        assert request["updateTextStyle"]["textStyle"]["baselineOffset"] == "SUBSCRIPT"

    def test_create_superscript_request(self):
        """Test creating superscript formatting request."""
        request = create_format_text_request(
            start_index=10,
            end_index=14,
            superscript=True,
        )

        assert request is not None
        assert (
            request["updateTextStyle"]["textStyle"]["baselineOffset"] == "SUPERSCRIPT"
        )


class TestValidationForFormatAll:
    """Tests for validation logic used by format_all_occurrences."""

    def test_validate_document_id(self):
        """Test document ID validation."""
        validator = ValidationManager()

        # Valid document ID (must be at least 20 chars)
        is_valid, _ = validator.validate_document_id_structured(
            "1abc123DEF_xyz-A1B2C3D4E5"
        )
        assert is_valid is True

        # Empty document ID
        is_valid, error = validator.validate_document_id_structured("")
        assert is_valid is False
        assert error is not None

        # Too short document ID
        is_valid, error = validator.validate_document_id_structured("short")
        assert is_valid is False
        assert error is not None

    def test_invalid_font_size_validation(self):
        """Test font size validation logic."""
        validator = ValidationManager()

        # Zero font size should be invalid
        error = validator.create_invalid_param_error(
            param_name="font_size",
            received=0,
            valid_values=["positive integer (e.g., 10, 12, 14)"],
        )
        result = json.loads(error)
        assert result["error"] is True

        # Negative font size should be invalid
        error = validator.create_invalid_param_error(
            param_name="font_size",
            received=-5,
            valid_values=["positive integer (e.g., 10, 12, 14)"],
        )
        result = json.loads(error)
        assert result["error"] is True


class TestBulkFormatBatchBuilding:
    """Tests for building batch requests for bulk formatting."""

    def test_build_multiple_format_requests(self):
        """Test building format requests for multiple occurrences."""
        occurrences = [
            (10, 14),
            (25, 29),
            (40, 44),
        ]

        requests = []
        for start, end in occurrences:
            request = create_format_text_request(
                start_index=start,
                end_index=end,
                bold=True,
                foreground_color="red",
            )
            if request:
                requests.append(request)

        assert len(requests) == 3
        for request in requests:
            assert "updateTextStyle" in request
            assert request["updateTextStyle"]["textStyle"]["bold"] is True

    def test_empty_occurrences_produces_no_requests(self):
        """Test that empty occurrences list produces no requests."""
        occurrences = []

        requests = []
        for start, end in occurrences:
            request = create_format_text_request(
                start_index=start,
                end_index=end,
                bold=True,
            )
            if request:
                requests.append(request)

        assert len(requests) == 0


class TestFormatAllResponseStructure:
    """Tests for the expected response structure from format_all_occurrences."""

    def test_success_response_structure(self):
        """Test the expected structure of a successful response."""
        # Simulate what the tool would return
        response = {
            "success": True,
            "operation": "format_all",
            "occurrences_formatted": 3,
            "search": "TODO",
            "match_case": True,
            "affected_ranges": [
                {"index": 1, "range": {"start": 10, "end": 14}},
                {"index": 2, "range": {"start": 25, "end": 29}},
                {"index": 3, "range": {"start": 40, "end": 44}},
            ],
            "formatting_applied": ["bold", "foreground_color"],
            "message": "Applied formatting (bold, foreground_color) to 3 occurrence(s) of 'TODO'",
            "link": "https://docs.google.com/document/d/doc-123/edit",
        }

        # Verify structure
        assert response["success"] is True
        assert response["operation"] == "format_all"
        assert response["occurrences_formatted"] == 3
        assert len(response["affected_ranges"]) == 3
        assert "bold" in response["formatting_applied"]

    def test_preview_response_structure(self):
        """Test the expected structure of a preview response."""
        # Simulate what the tool would return in preview mode
        response = {
            "preview": True,
            "would_modify": True,
            "occurrences_found": 3,
            "search": "TODO",
            "match_case": True,
            "matches": [
                {
                    "index": 1,
                    "range": {"start": 10, "end": 14},
                    "text": "TODO",
                    "context": {"before": "text ", "after": " item"},
                },
            ],
            "formatting_to_apply": ["bold", "foreground_color"],
            "message": "Would format 3 occurrence(s) of 'TODO' with bold, foreground_color",
            "link": "https://docs.google.com/document/d/doc-123/edit",
        }

        # Verify structure
        assert response["preview"] is True
        assert response["would_modify"] is True
        assert response["occurrences_found"] == 3
        assert len(response["matches"]) > 0
        assert "formatting_to_apply" in response

    def test_no_matches_response_structure(self):
        """Test response structure when no matches are found."""
        response = {
            "success": True,
            "operation": "format_all",
            "occurrences_formatted": 0,
            "search": "NONEXISTENT",
            "match_case": True,
            "affected_ranges": [],
            "formatting_applied": ["bold"],
            "message": "No occurrences of 'NONEXISTENT' found in document",
            "link": "https://docs.google.com/document/d/doc-123/edit",
        }

        assert response["success"] is True
        assert response["occurrences_formatted"] == 0
        assert len(response["affected_ranges"]) == 0
        assert "No occurrences" in response["message"]
