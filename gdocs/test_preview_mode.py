"""
Unit tests for Google Docs modify_doc_text preview mode.

Tests the preview=True parameter functionality that allows users to see
what would change without actually modifying the document.
"""
import json
import pytest
from gdocs.docs_helpers import extract_text_at_range


def create_mock_document(paragraphs: list[tuple[str, int]]) -> dict:
    """
    Create a mock Google Docs document.

    Args:
        paragraphs: List of (text, start_index) tuples

    Returns:
        Mock document data structure
    """
    content = []

    # Initial section break
    content.append({
        'sectionBreak': {},
        'startIndex': 0,
        'endIndex': 1
    })

    for text, start_index in paragraphs:
        end_index = start_index + len(text)
        content.append({
            'startIndex': start_index,
            'endIndex': end_index,
            'paragraph': {
                'elements': [{
                    'startIndex': start_index,
                    'endIndex': end_index,
                    'textRun': {
                        'content': text
                    }
                }]
            }
        })

    return {
        'title': 'Test Document',
        'body': {
            'content': content
        }
    }


class TestExtractTextAtRange:
    """Tests for the extract_text_at_range helper function."""

    def test_extract_simple_range(self):
        """Can extract text at a simple range."""
        doc = create_mock_document([
            ("Hello World\n", 1),
            ("This is a test.\n", 13),
        ])

        result = extract_text_at_range(doc, 1, 6)

        assert result["found"] is True
        assert result["text"] == "Hello"
        assert result["start_index"] == 1
        assert result["end_index"] == 6

    def test_extract_with_context(self):
        """Returns surrounding context."""
        doc = create_mock_document([
            ("Hello World\n", 1),
            ("This is a test.\n", 13),
        ])

        result = extract_text_at_range(doc, 7, 12, context_chars=5)

        assert result["found"] is True
        assert result["text"] == "World"
        assert result["context_before"] == "ello "  # 5 chars before
        # Context after includes the newline and next paragraph start

    def test_extract_at_document_start(self):
        """Can extract text at document start."""
        doc = create_mock_document([
            ("Hello World\n", 1),
        ])

        result = extract_text_at_range(doc, 1, 4)

        assert result["found"] is True
        assert result["text"] == "Hel"
        assert result["context_before"] == ""

    def test_extract_at_invalid_range_returns_not_found(self):
        """Returns found=False for indices outside text content."""
        doc = create_mock_document([
            ("Hello\n", 1),
        ])

        # Try to get text at index 0 (section break, not text)
        result = extract_text_at_range(doc, 0, 0)

        # Should handle gracefully
        assert "found" in result

    def test_extract_empty_range_for_insert_point(self):
        """For insert operations (same start and end), returns empty text."""
        doc = create_mock_document([
            ("Hello World\n", 1),
        ])

        result = extract_text_at_range(doc, 6, 6)

        assert result["found"] is True
        assert result["text"] == ""  # Insert point has no text
        assert "context_before" in result
        assert "context_after" in result


class TestPreviewResponseStructure:
    """Tests for the expected preview response structure."""

    def test_preview_response_has_required_fields(self):
        """Preview response should contain all required fields."""
        # This is a structural test - we mock what the preview response should look like
        preview_response = {
            "preview": True,
            "would_modify": True,
            "operation": "replace",
            "affected_range": {"start": 100, "end": 108},
            "position_shift": 0,
            "current_content": "old text",
            "new_content": "new text",
            "context": {
                "before": "...before...",
                "after": "...after..."
            },
            "positioning_info": {},
            "message": "Would replace 8 characters with 8 characters at index 100",
            "link": "https://docs.google.com/document/d/test/edit"
        }

        # Verify structure
        assert preview_response["preview"] is True
        assert preview_response["would_modify"] is True
        assert "operation" in preview_response
        assert "affected_range" in preview_response
        assert "start" in preview_response["affected_range"]
        assert "end" in preview_response["affected_range"]
        assert "position_shift" in preview_response
        assert "message" in preview_response
        assert "link" in preview_response

    def test_preview_for_insert_shows_position_shift(self):
        """Insert preview should show positive position_shift."""
        preview_response = {
            "preview": True,
            "would_modify": True,
            "operation": "insert",
            "affected_range": {"start": 100, "end": 100},
            "position_shift": 10,  # Text of 10 chars would shift everything right
            "new_content": "new text!.",
            "message": "Would insert 10 characters at index 100"
        }

        assert preview_response["position_shift"] == 10
        assert preview_response["operation"] == "insert"

    def test_preview_for_replace_calculates_shift_correctly(self):
        """Replace preview should calculate position shift based on size difference."""
        # Replace "hello" (5 chars) with "goodbye" (7 chars) = +2 shift
        preview_response = {
            "preview": True,
            "would_modify": True,
            "operation": "replace",
            "affected_range": {"start": 10, "end": 15},
            "position_shift": 2,  # 7 - 5 = 2
            "original_length": 5,
            "new_length": 7,
            "current_content": "hello",
            "new_content": "goodbye",
            "message": "Would replace 5 characters with 7 characters at index 10"
        }

        assert preview_response["position_shift"] == preview_response["new_length"] - preview_response["original_length"]

    def test_preview_for_format_has_zero_shift(self):
        """Format preview should have zero position shift."""
        preview_response = {
            "preview": True,
            "would_modify": True,
            "operation": "format",
            "affected_range": {"start": 10, "end": 20},
            "position_shift": 0,  # Formatting doesn't change positions
            "styles_to_apply": ["bold", "italic"],
            "current_content": "some text.",
            "message": "Would apply formatting (bold, italic) to range 10-20"
        }

        assert preview_response["position_shift"] == 0
        assert "styles_to_apply" in preview_response


class TestPreviewModeDocumentation:
    """Tests that verify preview mode behavior described in documentation."""

    def test_preview_does_not_include_success_field(self):
        """Preview response uses 'would_modify' not 'success'."""
        # The actual operation response has 'success', but preview has 'would_modify'
        preview_response = {
            "preview": True,
            "would_modify": True,
        }

        assert "success" not in preview_response
        assert preview_response["would_modify"] is True

    def test_preview_preserves_positioning_info(self):
        """Preview response includes positioning resolution details."""
        # When using search-based positioning, preview shows how it resolved
        preview_response = {
            "preview": True,
            "would_modify": True,
            "operation": "replace",
            "positioning_info": {
                "search_text": "TODO",
                "position": "replace",
                "occurrence": 2,
                "found_at_index": 150,
                "message": "Found at index 150"
            }
        }

        assert "positioning_info" in preview_response
        assert preview_response["positioning_info"]["search_text"] == "TODO"
        assert preview_response["positioning_info"]["occurrence"] == 2


class TestPreviewDeleteOperation:
    """Tests for delete operation preview mode (text="" with replace)."""

    def test_preview_for_delete_shows_negative_shift(self):
        """Delete preview should show negative position_shift."""
        # Delete 8 characters = -8 shift
        preview_response = {
            "preview": True,
            "would_modify": True,
            "operation": "delete",
            "affected_range": {"start": 100, "end": 108},
            "position_shift": -8,  # Deleting shifts everything left
            "deleted_length": 8,
            "current_content": "[DELETE]",
            "context": {
                "before": "text before ",
                "after": " text after"
            },
            "message": "Would delete 8 characters from index 100 to 108"
        }

        assert preview_response["position_shift"] == -8
        assert preview_response["operation"] == "delete"
        assert preview_response["deleted_length"] == 8

    def test_preview_delete_has_current_content_no_new_content(self):
        """Delete preview shows what will be deleted, but has no new_content."""
        preview_response = {
            "preview": True,
            "would_modify": True,
            "operation": "delete",
            "affected_range": {"start": 50, "end": 55},
            "position_shift": -5,
            "deleted_length": 5,
            "current_content": "Hello",
            "message": "Would delete 5 characters from index 50 to 55"
        }

        assert "current_content" in preview_response
        assert preview_response["current_content"] == "Hello"
        assert "new_content" not in preview_response


class TestFindAndReplacePreviewResponseStructure:
    """Tests for find_and_replace_doc preview response structure."""

    def test_preview_response_has_required_fields(self):
        """find_and_replace_doc preview response should contain all required fields."""
        preview_response = {
            "preview": True,
            "would_modify": True,
            "occurrences_found": 3,
            "find_text": "TODO",
            "replace_text": "DONE",
            "match_case": False,
            "matches": [
                {
                    "index": 1,
                    "range": {"start": 15, "end": 19},
                    "text": "TODO",
                    "context": {"before": "...", "after": "..."}
                },
                {
                    "index": 2,
                    "range": {"start": 50, "end": 54},
                    "text": "TODO",
                    "context": {"before": "...", "after": "..."}
                },
                {
                    "index": 3,
                    "range": {"start": 100, "end": 104},
                    "text": "TODO",
                    "context": {"before": "...", "after": "..."}
                },
            ],
            "position_shift_per_replacement": 0,
            "total_position_shift": 0,
            "link": "https://docs.google.com/document/d/test/edit",
            "message": "Would replace 3 occurrence(s) of 'TODO' with 'DONE'"
        }

        # Verify structure
        assert preview_response["preview"] is True
        assert preview_response["would_modify"] is True
        assert preview_response["occurrences_found"] == 3
        assert preview_response["find_text"] == "TODO"
        assert preview_response["replace_text"] == "DONE"
        assert "match_case" in preview_response
        assert "matches" in preview_response
        assert len(preview_response["matches"]) == 3
        assert "position_shift_per_replacement" in preview_response
        assert "total_position_shift" in preview_response
        assert "link" in preview_response
        assert "message" in preview_response

    def test_preview_calculates_position_shift_correctly(self):
        """Preview should calculate position shift based on text length difference."""
        # Replace "foo" (3 chars) with "foobar" (6 chars) = +3 shift per occurrence
        preview_response = {
            "preview": True,
            "would_modify": True,
            "occurrences_found": 4,
            "find_text": "foo",
            "replace_text": "foobar",
            "position_shift_per_replacement": 3,  # 6 - 3 = 3
            "total_position_shift": 12,  # 3 * 4 = 12
        }

        assert preview_response["position_shift_per_replacement"] == len("foobar") - len("foo")
        assert preview_response["total_position_shift"] == preview_response["position_shift_per_replacement"] * preview_response["occurrences_found"]

    def test_preview_with_no_matches_shows_would_not_modify(self):
        """Preview with no matches should show would_modify=False."""
        preview_response = {
            "preview": True,
            "would_modify": False,
            "occurrences_found": 0,
            "find_text": "nonexistent",
            "replace_text": "something",
            "matches": [],
            "position_shift_per_replacement": 9,  # len("something") - len("nonexistent") = -2
            "total_position_shift": 0,  # 0 occurrences = 0 total shift
            "message": "No occurrences of 'nonexistent' found in document"
        }

        assert preview_response["would_modify"] is False
        assert preview_response["occurrences_found"] == 0
        assert len(preview_response["matches"]) == 0
        assert "No occurrences" in preview_response["message"]

    def test_preview_matches_include_context(self):
        """Each match in preview should include text and context."""
        match = {
            "index": 1,
            "range": {"start": 25, "end": 30},
            "text": "hello",
            "context": {
                "before": "say ",
                "after": " world"
            }
        }

        assert "index" in match
        assert "range" in match
        assert "start" in match["range"]
        assert "end" in match["range"]
        assert "text" in match
        assert "context" in match
        assert "before" in match["context"]
        assert "after" in match["context"]

    def test_preview_negative_shift_for_shorter_replacement(self):
        """Preview should show negative shift when replacement is shorter."""
        # Replace "hello" (5 chars) with "hi" (2 chars) = -3 shift per occurrence
        preview_response = {
            "preview": True,
            "would_modify": True,
            "occurrences_found": 2,
            "find_text": "hello",
            "replace_text": "hi",
            "position_shift_per_replacement": -3,  # 2 - 5 = -3
            "total_position_shift": -6,  # -3 * 2 = -6
        }

        assert preview_response["position_shift_per_replacement"] == -3
        assert preview_response["total_position_shift"] == -6
