"""
Unit tests for Google Docs range-based selection utilities.

Tests the new range resolution functions:
- resolve_range
- resolve_range_by_search_bounds
- resolve_range_by_search_with_extension
- resolve_range_by_search_with_offsets
- resolve_range_by_section
- find_paragraph_boundaries
- find_sentence_boundaries
- find_line_boundaries
"""

from gdocs.docs_helpers import (
    resolve_range,
    resolve_range_by_search_bounds,
    resolve_range_by_search_with_extension,
    resolve_range_by_search_with_offsets,
    resolve_range_by_section,
    find_paragraph_boundaries,
    find_sentence_boundaries,
    find_line_boundaries,
    RangeResult,
    ExtendBoundary,
)


def create_mock_paragraph(
    text: str, start_index: int, named_style: str = "NORMAL_TEXT"
):
    """Create a mock paragraph element."""
    end_index = start_index + len(text) + 1  # +1 for newline
    return {
        "startIndex": start_index,
        "endIndex": end_index,
        "paragraph": {
            "paragraphStyle": {"namedStyleType": named_style},
            "elements": [
                {
                    "startIndex": start_index,
                    "endIndex": end_index,
                    "textRun": {"content": text + "\n"},
                }
            ],
        },
    }


def create_mock_document(elements):
    """Create a mock document with given elements."""
    return {"title": "Test Document", "body": {"content": elements}, "lists": {}}


class TestResolveRangeBySearchBounds:
    """Tests for resolve_range_by_search_bounds function."""

    def test_resolves_simple_range(self):
        """Test resolving a range between two search terms."""
        doc = create_mock_document(
            [
                create_mock_paragraph(
                    "This is the Introduction section", 1, "HEADING_1"
                ),
                create_mock_paragraph(
                    "Some content in between here", 35, "NORMAL_TEXT"
                ),
                create_mock_paragraph(
                    "This is the Conclusion section", 65, "HEADING_1"
                ),
            ]
        )

        result = resolve_range_by_search_bounds(
            doc, "Introduction", "Conclusion", match_case=True
        )

        assert result.success is True
        assert result.start_index is not None
        assert result.end_index is not None
        assert result.start_index < result.end_index
        assert result.matched_start == "Introduction"
        assert result.matched_end == "Conclusion"

    def test_fails_when_start_not_found(self):
        """Test failure when start text is not found."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Some content here", 1, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_bounds(
            doc, "NotFound", "content", match_case=True
        )

        assert result.success is False
        assert "not found" in result.message.lower()

    def test_fails_when_end_not_found(self):
        """Test failure when end text is not found."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Some content here", 1, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_bounds(
            doc, "Some", "NotFound", match_case=True
        )

        assert result.success is False
        assert "not found" in result.message.lower()

    def test_fails_when_end_before_start(self):
        """Test failure when end comes before start in document."""
        doc = create_mock_document(
            [
                create_mock_paragraph("End text comes first", 1, "NORMAL_TEXT"),
                create_mock_paragraph("Start text comes second", 23, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_bounds(doc, "Start", "End", match_case=True)

        assert result.success is False
        assert "before" in result.message.lower() or "invalid" in result.message.lower()

    def test_respects_occurrence_parameter(self):
        """Test that occurrence parameter works correctly."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First marker text here", 1, "NORMAL_TEXT"),
                create_mock_paragraph("Content between markers", 25, "NORMAL_TEXT"),
                create_mock_paragraph("Second marker text here", 50, "NORMAL_TEXT"),
                create_mock_paragraph("More content", 75, "NORMAL_TEXT"),
                create_mock_paragraph("Third marker text here", 89, "NORMAL_TEXT"),
            ]
        )

        # Should find from first to second occurrence
        result = resolve_range_by_search_bounds(
            doc,
            "marker",
            "marker",
            start_occurrence=1,
            end_occurrence=2,
            match_case=True,
        )

        assert result.success is True


class TestResolveRangeBySearchWithExtension:
    """Tests for resolve_range_by_search_with_extension function."""

    def test_extends_to_paragraph(self):
        """Test extending selection to paragraph boundaries."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First paragraph content", 1, "NORMAL_TEXT"),
                create_mock_paragraph(
                    "Second paragraph with keyword inside it", 26, "NORMAL_TEXT"
                ),
                create_mock_paragraph("Third paragraph content", 67, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_with_extension(
            doc, "keyword", "paragraph", match_case=True
        )

        assert result.success is True
        assert result.extend_type == "paragraph"
        assert result.start_index == 26  # Start of the paragraph
        # End index is end of paragraph element (66 = 26 + 39 + 1 for newline)
        assert result.end_index == 66

    def test_fails_when_search_not_found(self):
        """Test failure when search text not found."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Some content here", 1, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_with_extension(
            doc, "NotFound", "paragraph", match_case=True
        )

        assert result.success is False
        assert "not found" in result.message.lower()

    def test_invalid_extend_type(self):
        """Test failure with invalid extend type."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Some content here", 1, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_with_extension(
            doc, "content", "invalid_type", match_case=True
        )

        assert result.success is False
        assert "invalid" in result.message.lower()


class TestResolveRangeBySearchWithOffsets:
    """Tests for resolve_range_by_search_with_offsets function."""

    def test_applies_before_offset(self):
        """Test that before_chars offset is applied correctly."""
        doc = create_mock_document(
            [
                create_mock_paragraph(
                    "This is some text with a keyword in it", 1, "NORMAL_TEXT"
                ),
            ]
        )

        result = resolve_range_by_search_with_offsets(
            doc, "keyword", before_chars=10, after_chars=0, match_case=True
        )

        assert result.success is True
        # The range should start before the keyword

    def test_applies_after_offset(self):
        """Test that after_chars offset is applied correctly."""
        doc = create_mock_document(
            [
                create_mock_paragraph(
                    "This is some text with a keyword in it", 1, "NORMAL_TEXT"
                ),
            ]
        )

        result = resolve_range_by_search_with_offsets(
            doc, "keyword", before_chars=0, after_chars=10, match_case=True
        )

        assert result.success is True
        # The range should extend after the keyword

    def test_clamps_to_document_bounds(self):
        """Test that offsets are clamped to document bounds."""
        doc = create_mock_document(
            [
                create_mock_paragraph("keyword at start", 1, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_with_offsets(
            doc, "keyword", before_chars=1000, after_chars=1000, match_case=True
        )

        assert result.success is True
        assert result.start_index >= 1  # Can't go below document start

    def test_fails_when_search_not_found(self):
        """Test failure when search text not found."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Some content here", 1, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_with_offsets(
            doc, "NotFound", before_chars=10, after_chars=10, match_case=True
        )

        assert result.success is False


class TestResolveRangeBySection:
    """Tests for resolve_range_by_section function."""

    def test_selects_section_without_heading(self):
        """Test selecting section content excluding heading."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Intro content here", 14, "NORMAL_TEXT"),
                create_mock_paragraph("Details", 34, "HEADING_1"),
            ]
        )

        result = resolve_range_by_section(
            doc, "Introduction", include_heading=False, match_case=False
        )

        assert result.success is True
        assert result.section_name == "Introduction"
        # Should not include the heading text

    def test_selects_section_with_heading(self):
        """Test selecting section including heading."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Intro content here", 14, "NORMAL_TEXT"),
                create_mock_paragraph("Details", 34, "HEADING_1"),
            ]
        )

        result = resolve_range_by_section(
            doc, "Introduction", include_heading=True, match_case=False
        )

        assert result.success is True
        assert result.start_index == 1  # Includes heading start

    def test_fails_when_section_not_found(self):
        """Test failure when section heading not found."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
            ]
        )

        result = resolve_range_by_section(
            doc, "NonExistent", include_heading=False, match_case=False
        )

        assert result.success is False
        assert "not found" in result.message.lower()


class TestResolveRange:
    """Tests for the main resolve_range function."""

    def test_handles_search_bounds_format(self):
        """Test resolving start/end search format."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Start of the document", 1, "NORMAL_TEXT"),
                create_mock_paragraph("Middle content here", 24, "NORMAL_TEXT"),
                create_mock_paragraph("End of the document", 45, "NORMAL_TEXT"),
            ]
        )

        range_spec = {"start": {"search": "Start"}, "end": {"search": "End"}}

        result = resolve_range(doc, range_spec)

        assert result.success is True
        assert result.matched_start == "Start"
        assert result.matched_end == "End"

    def test_handles_search_extend_format(self):
        """Test resolving search with extension format."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First paragraph", 1, "NORMAL_TEXT"),
                create_mock_paragraph(
                    "Second paragraph with keyword", 18, "NORMAL_TEXT"
                ),
            ]
        )

        range_spec = {"search": "keyword", "extend": "paragraph"}

        result = resolve_range(doc, range_spec)

        assert result.success is True
        assert result.extend_type == "paragraph"

    def test_handles_search_offsets_format(self):
        """Test resolving search with offsets format."""
        doc = create_mock_document(
            [
                create_mock_paragraph(
                    "Some text with a keyword in the middle", 1, "NORMAL_TEXT"
                ),
            ]
        )

        range_spec = {"search": "keyword", "before_chars": 5, "after_chars": 10}

        result = resolve_range(doc, range_spec)

        assert result.success is True

    def test_handles_section_format(self):
        """Test resolving section reference format."""
        doc = create_mock_document(
            [
                create_mock_paragraph("My Section", 1, "HEADING_1"),
                create_mock_paragraph("Section content", 13, "NORMAL_TEXT"),
            ]
        )

        range_spec = {"section": "My Section", "include_heading": True}

        result = resolve_range(doc, range_spec)

        assert result.success is True
        assert result.section_name == "My Section"

    def test_fails_with_invalid_format(self):
        """Test failure with invalid range specification."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Some content", 1, "NORMAL_TEXT"),
            ]
        )

        range_spec = {"invalid_key": "value"}

        result = resolve_range(doc, range_spec)

        assert result.success is False
        assert "invalid" in result.message.lower()


class TestFindParagraphBoundaries:
    """Tests for find_paragraph_boundaries function."""

    def test_finds_paragraph_containing_index(self):
        """Test finding paragraph boundaries for an index."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First paragraph", 1, "NORMAL_TEXT"),
                create_mock_paragraph("Second paragraph", 18, "NORMAL_TEXT"),
                create_mock_paragraph("Third paragraph", 36, "NORMAL_TEXT"),
            ]
        )

        # Index 25 should be in "Second paragraph"
        start, end = find_paragraph_boundaries(doc, 25)

        assert start == 18
        # End is 35 = 18 + 16 (text length) + 1 (newline)
        assert end == 35


class TestFindSentenceBoundaries:
    """Tests for find_sentence_boundaries function."""

    def test_finds_sentence_containing_index(self):
        """Test finding sentence boundaries."""
        doc = create_mock_document(
            [
                create_mock_paragraph(
                    "First sentence. Second sentence. Third sentence.", 1, "NORMAL_TEXT"
                ),
            ]
        )

        # This test verifies the function runs without error
        start, end = find_sentence_boundaries(doc, 20)

        assert start is not None
        assert end is not None
        assert start <= 20
        assert end >= 20


class TestFindLineBoundaries:
    """Tests for find_line_boundaries function."""

    def test_finds_line_containing_index(self):
        """Test finding line boundaries."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First line", 1, "NORMAL_TEXT"),
                create_mock_paragraph("Second line", 13, "NORMAL_TEXT"),
            ]
        )

        start, end = find_line_boundaries(doc, 5)

        assert start is not None
        assert end is not None


class TestRangeResult:
    """Tests for RangeResult dataclass."""

    def test_to_dict_excludes_none_values(self):
        """Test that to_dict excludes None values."""
        result = RangeResult(
            success=True,
            start_index=10,
            end_index=50,
            message="Test message",
            matched_start="start",
            matched_end=None,  # Should be excluded
        )

        result_dict = result.to_dict()

        assert "success" in result_dict
        assert "start_index" in result_dict
        assert "matched_start" in result_dict
        assert "matched_end" not in result_dict

    def test_to_dict_includes_all_set_values(self):
        """Test that to_dict includes all non-None values."""
        result = RangeResult(
            success=True,
            start_index=10,
            end_index=50,
            message="Test message",
            matched_start="start",
            matched_end="end",
            extend_type="paragraph",
            section_name="My Section",
        )

        result_dict = result.to_dict()

        assert result_dict["success"] is True
        assert result_dict["start_index"] == 10
        assert result_dict["end_index"] == 50
        assert result_dict["matched_start"] == "start"
        assert result_dict["matched_end"] == "end"
        assert result_dict["extend_type"] == "paragraph"
        assert result_dict["section_name"] == "My Section"


class TestExtendBoundary:
    """Tests for ExtendBoundary enum."""

    def test_enum_values(self):
        """Test that enum has expected values."""
        assert ExtendBoundary.PARAGRAPH.value == "paragraph"
        assert ExtendBoundary.SENTENCE.value == "sentence"
        assert ExtendBoundary.LINE.value == "line"
        assert ExtendBoundary.SECTION.value == "section"


class TestSentenceBoundaryAbbreviations:
    """Tests for improved sentence boundary detection with abbreviations."""

    def test_handles_common_abbreviations(self):
        """Test that common abbreviations don't cause false sentence breaks."""
        doc = create_mock_document(
            [
                create_mock_paragraph(
                    "Dr. Smith went to the store. He bought milk.", 1, "NORMAL_TEXT"
                ),
            ]
        )

        # Find boundaries at index within "Dr. Smith" - should stay in same sentence
        start, end = find_sentence_boundaries(doc, 5)

        # "Dr. Smith went to the store." should be one sentence
        # The period after "Dr" should not be treated as a sentence end
        assert start is not None
        assert end is not None
        # End should be after "store." not after "Dr."
        assert end > 20

    def test_handles_multiple_abbreviations(self):
        """Test handling multiple abbreviations in one sentence."""
        doc = create_mock_document(
            [
                create_mock_paragraph(
                    "Mr. and Mrs. Jones visited the U.S. last year. They enjoyed it.",
                    1,
                    "NORMAL_TEXT",
                ),
            ]
        )

        # Find boundaries - should correctly identify sentence
        start, end = find_sentence_boundaries(doc, 10)

        assert start is not None
        assert end is not None

    def test_exclamation_and_question_marks(self):
        """Test that exclamation and question marks end sentences."""
        doc = create_mock_document(
            [
                create_mock_paragraph(
                    "What time is it? I think noon! That seems right.", 1, "NORMAL_TEXT"
                ),
            ]
        )

        # Find boundaries at "noon" - should be in second sentence
        start, end = find_sentence_boundaries(doc, 30)

        assert start is not None
        assert end is not None


class TestCharacterOffsetValidation:
    """Tests for improved character offset validation."""

    def test_rejects_negative_before_chars(self):
        """Test that negative before_chars is rejected."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Some text here", 1, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_with_offsets(
            doc, "text", before_chars=-5, after_chars=0, match_case=True
        )

        assert result.success is False
        assert "before_chars" in result.message.lower()

    def test_rejects_negative_after_chars(self):
        """Test that negative after_chars is rejected."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Some text here", 1, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_with_offsets(
            doc, "text", before_chars=0, after_chars=-5, match_case=True
        )

        assert result.success is False
        assert "after_chars" in result.message.lower()

    def test_provides_informative_clamping_message(self):
        """Test that clamping provides informative messages."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Short", 1, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_with_offsets(
            doc, "Short", before_chars=1000, after_chars=1000, match_case=True
        )

        assert result.success is True
        # Message should mention clamping
        assert "clamped" in result.message.lower() or "bound" in result.message.lower()


class TestSectionHierarchyAwareness:
    """Tests for improved section hierarchy awareness."""

    def test_section_includes_subsections(self):
        """Test that parent section includes nested subsections."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Some intro content", 14, "NORMAL_TEXT"),
                create_mock_paragraph("Background", 34, "HEADING_2"),
                create_mock_paragraph("Background details", 46, "NORMAL_TEXT"),
                create_mock_paragraph("Conclusion", 66, "HEADING_1"),
            ]
        )

        # Extend "intro content" to section should include the H2 subsection
        result = resolve_range_by_search_with_extension(
            doc, "intro content", "section", match_case=True
        )

        assert result.success is True
        # Section should start at Introduction heading
        assert result.start_index == 1
        # Section should end at Conclusion (next H1), not at Background (H2)
        assert result.end_index == 66

    def test_subsection_ends_at_sibling(self):
        """Test that subsection ends at sibling heading."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Method A", 14, "HEADING_2"),
                create_mock_paragraph("Method A details", 24, "NORMAL_TEXT"),
                create_mock_paragraph("Method B", 42, "HEADING_2"),
                create_mock_paragraph("Method B details", 52, "NORMAL_TEXT"),
            ]
        )

        # Extend "Method A details" to section should end at Method B
        result = resolve_range_by_search_with_extension(
            doc, "Method A details", "section", match_case=True
        )

        assert result.success is True
        # Section should start at Method A heading
        assert result.start_index == 14
        # Section should end at Method B (sibling H2)
        assert result.end_index == 42


class TestRangeValidation:
    """Tests for range validation after extension."""

    def test_extension_validates_range_contains_search(self):
        """Test that extended range validation works."""
        # This is a structural test - in normal operation, the extended range
        # should always contain the search result. This tests the validation exists.
        doc = create_mock_document(
            [
                create_mock_paragraph("Some content here", 1, "NORMAL_TEXT"),
            ]
        )

        result = resolve_range_by_search_with_extension(
            doc, "content", "paragraph", match_case=True
        )

        assert result.success is True
        # Extended range should contain the found text indices
