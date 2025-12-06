"""
Unit tests for insert_doc_image functionality.

Tests verify proper handling of optional width/height parameters.
"""
import pytest
from gdocs.docs_helpers import create_insert_image_request


class TestCreateInsertImageRequest:
    """Tests for create_insert_image_request helper function."""

    def test_request_without_dimensions(self):
        """Image request without dimensions should NOT include objectSize."""
        request = create_insert_image_request(
            index=1,
            image_uri="https://example.com/image.png",
        )

        assert "insertInlineImage" in request
        assert request["insertInlineImage"]["location"]["index"] == 1
        assert request["insertInlineImage"]["uri"] == "https://example.com/image.png"
        # Key assertion: no objectSize when dimensions not specified
        assert "objectSize" not in request["insertInlineImage"]

    def test_request_with_width_only(self):
        """Image request with only width should include width in objectSize."""
        request = create_insert_image_request(
            index=1,
            image_uri="https://example.com/image.png",
            width=200,
        )

        assert "objectSize" in request["insertInlineImage"]
        assert "width" in request["insertInlineImage"]["objectSize"]
        assert request["insertInlineImage"]["objectSize"]["width"]["magnitude"] == 200
        assert "height" not in request["insertInlineImage"]["objectSize"]

    def test_request_with_height_only(self):
        """Image request with only height should include height in objectSize."""
        request = create_insert_image_request(
            index=1,
            image_uri="https://example.com/image.png",
            height=150,
        )

        assert "objectSize" in request["insertInlineImage"]
        assert "height" in request["insertInlineImage"]["objectSize"]
        assert request["insertInlineImage"]["objectSize"]["height"]["magnitude"] == 150
        assert "width" not in request["insertInlineImage"]["objectSize"]

    def test_request_with_both_dimensions(self):
        """Image request with both dimensions should include full objectSize."""
        request = create_insert_image_request(
            index=1,
            image_uri="https://example.com/image.png",
            width=200,
            height=150,
        )

        object_size = request["insertInlineImage"]["objectSize"]
        assert object_size["width"]["magnitude"] == 200
        assert object_size["width"]["unit"] == "PT"
        assert object_size["height"]["magnitude"] == 150
        assert object_size["height"]["unit"] == "PT"

    def test_request_with_explicit_none_dimensions(self):
        """Image request with explicit None dimensions should NOT include objectSize."""
        request = create_insert_image_request(
            index=1,
            image_uri="https://example.com/image.png",
            width=None,
            height=None,
        )

        # Should behave same as not passing dimensions at all
        assert "objectSize" not in request["insertInlineImage"]

    def test_request_with_zero_dimensions_are_included(self):
        """Verify that 0 is NOT a valid dimension value that gets included.

        Note: The Google Docs API rejects width=0 or height=0, so callers
        should use None instead of 0 to omit dimensions. This test documents
        that 0 values DO get included in the request (unlike None).
        """
        # This test documents current behavior - 0 values ARE included
        # The fix was in docs_tools.py to default to None instead of 0
        request = create_insert_image_request(
            index=1,
            image_uri="https://example.com/image.png",
            width=0,
            height=0,
        )

        # 0 is not None, so it gets included - this would fail at API level
        assert "objectSize" in request["insertInlineImage"]
        assert request["insertInlineImage"]["objectSize"]["width"]["magnitude"] == 0
