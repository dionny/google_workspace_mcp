"""
Unit tests for section_break insertion functionality.

Tests verify proper creation of insertSectionBreak requests.
"""
from gdocs.docs_helpers import create_insert_section_break_request


class TestCreateInsertSectionBreakRequest:
    """Tests for create_insert_section_break_request helper function."""

    def test_request_structure_default_type(self):
        """Section break request should have correct structure with default NEXT_PAGE type."""
        request = create_insert_section_break_request(index=1)

        assert "insertSectionBreak" in request
        assert "location" in request["insertSectionBreak"]
        assert "sectionType" in request["insertSectionBreak"]
        assert "index" in request["insertSectionBreak"]["location"]
        assert request["insertSectionBreak"]["location"]["index"] == 1
        assert request["insertSectionBreak"]["sectionType"] == "NEXT_PAGE"

    def test_request_with_next_page_type(self):
        """Section break request should use NEXT_PAGE type when specified."""
        request = create_insert_section_break_request(index=50, section_type="NEXT_PAGE")

        assert request["insertSectionBreak"]["sectionType"] == "NEXT_PAGE"
        assert request["insertSectionBreak"]["location"]["index"] == 50

    def test_request_with_continuous_type(self):
        """Section break request should use CONTINUOUS type when specified."""
        request = create_insert_section_break_request(index=25, section_type="CONTINUOUS")

        assert request["insertSectionBreak"]["sectionType"] == "CONTINUOUS"
        assert request["insertSectionBreak"]["location"]["index"] == 25

    def test_request_with_different_index(self):
        """Section break request should use provided index."""
        request = create_insert_section_break_request(index=100)

        assert request["insertSectionBreak"]["location"]["index"] == 100

    def test_request_at_document_start(self):
        """Section break request at index 1 (document start)."""
        request = create_insert_section_break_request(index=1)

        assert request["insertSectionBreak"]["location"]["index"] == 1

    def test_request_has_correct_fields(self):
        """Section break request should have location and sectionType fields."""
        request = create_insert_section_break_request(index=50)

        # Should have 'location' and 'sectionType' keys within 'insertSectionBreak'
        assert len(request["insertSectionBreak"]) == 2
        assert "location" in request["insertSectionBreak"]
        assert "sectionType" in request["insertSectionBreak"]
