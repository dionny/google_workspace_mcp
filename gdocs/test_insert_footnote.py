"""
Unit tests for insert_doc_footnote functionality.

Tests verify proper handling of footnote creation requests and text insertion.
"""
from gdocs.docs_helpers import (
    create_insert_footnote_request,
    create_insert_text_in_footnote_request,
)


class TestCreateInsertFootnoteRequest:
    """Tests for create_insert_footnote_request helper function."""

    def test_basic_footnote_request(self):
        """Footnote request should have correct structure."""
        request = create_insert_footnote_request(index=10)

        assert "createFootnote" in request
        assert "location" in request["createFootnote"]
        assert request["createFootnote"]["location"]["index"] == 10

    def test_footnote_request_at_start(self):
        """Footnote request at document start (index 1)."""
        request = create_insert_footnote_request(index=1)

        assert request["createFootnote"]["location"]["index"] == 1

    def test_footnote_request_structure(self):
        """Verify full request structure matches Google Docs API spec."""
        request = create_insert_footnote_request(index=42)

        # Should have exactly one key at top level
        assert len(request) == 1
        assert "createFootnote" in request

        # Location should have exactly index key
        location = request["createFootnote"]["location"]
        assert len(location) == 1
        assert "index" in location


class TestCreateInsertTextInFootnoteRequest:
    """Tests for create_insert_text_in_footnote_request helper function."""

    def test_basic_text_insertion(self):
        """Text insertion request for footnote should have correct structure."""
        request = create_insert_text_in_footnote_request(
            footnote_id="kix.abc123",
            index=1,
            text="This is footnote content."
        )

        assert "insertText" in request
        assert "location" in request["insertText"]
        assert request["insertText"]["location"]["segmentId"] == "kix.abc123"
        assert request["insertText"]["location"]["index"] == 1
        assert request["insertText"]["text"] == "This is footnote content."

    def test_text_insertion_with_special_characters(self):
        """Text with special characters should be preserved."""
        text = "Citation: Author (2024). \"Title\" p.42"
        request = create_insert_text_in_footnote_request(
            footnote_id="kix.def456",
            index=1,
            text=text
        )

        assert request["insertText"]["text"] == text

    def test_text_insertion_with_unicode(self):
        """Unicode characters should be preserved."""
        text = "Reference: \u00a9 2024 \u2013 M\u00fcller et al."
        request = create_insert_text_in_footnote_request(
            footnote_id="kix.unicode",
            index=1,
            text=text
        )

        assert request["insertText"]["text"] == text

    def test_multiline_text_insertion(self):
        """Multiline text should be preserved."""
        text = "Line 1\nLine 2\nLine 3"
        request = create_insert_text_in_footnote_request(
            footnote_id="kix.multiline",
            index=1,
            text=text
        )

        assert request["insertText"]["text"] == text
        assert "\n" in request["insertText"]["text"]

    def test_segment_id_is_included(self):
        """Verify segment ID is properly set for footnote targeting."""
        request = create_insert_text_in_footnote_request(
            footnote_id="kix.footnote123",
            index=0,
            text="test"
        )

        # The segmentId tells the API to insert into the footnote, not the body
        assert request["insertText"]["location"]["segmentId"] == "kix.footnote123"
