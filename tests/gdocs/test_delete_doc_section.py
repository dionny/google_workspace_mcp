"""
Unit tests for Google Docs delete_doc_section tool.

Tests the delete_doc_section functionality that allows users to delete
an entire section (heading + content) or just the content by heading name.
"""
from gdocs.docs_structure import (
    find_section_by_heading,
    extract_structural_elements,
)


def create_mock_paragraph(text: str, start_index: int, named_style: str = 'NORMAL_TEXT'):
    """Create a mock paragraph element."""
    end_index = start_index + len(text) + 1  # +1 for newline
    return {
        'startIndex': start_index,
        'endIndex': end_index,
        'paragraph': {
            'paragraphStyle': {
                'namedStyleType': named_style
            },
            'elements': [{
                'startIndex': start_index,
                'endIndex': end_index,
                'textRun': {
                    'content': text + '\n'
                }
            }]
        }
    }


def create_mock_document(elements):
    """Create a mock document with given elements."""
    return {
        'title': 'Test Document',
        'body': {
            'content': elements
        },
        'lists': {}
    }


class TestDeleteDocSectionRangeCalculation:
    """Tests for delete_doc_section range calculation logic."""

    def test_calculates_delete_range_including_heading(self):
        """Test that delete range includes heading when include_heading=True."""
        doc = create_mock_document([
            create_mock_paragraph('Introduction', 1, 'HEADING_1'),
            create_mock_paragraph('Intro content here', 14, 'NORMAL_TEXT'),
            create_mock_paragraph('Details', 34, 'HEADING_1'),
            create_mock_paragraph('Details content', 43, 'NORMAL_TEXT'),
        ])

        section = find_section_by_heading(doc, 'Introduction')

        # When include_heading=True, delete_start should be section start
        assert section is not None
        delete_start = section['start_index']
        delete_end = section['end_index']

        assert delete_start == 1  # Starts at heading
        assert delete_end == 34  # Ends at next section

    def test_calculates_delete_range_excluding_heading(self):
        """Test that delete range excludes heading when include_heading=False."""
        doc = create_mock_document([
            create_mock_paragraph('Introduction', 1, 'HEADING_1'),
            create_mock_paragraph('Intro content here', 14, 'NORMAL_TEXT'),
            create_mock_paragraph('Details', 34, 'HEADING_1'),
        ])

        section = find_section_by_heading(doc, 'Introduction')
        elements = extract_structural_elements(doc)

        # Find heading end for include_heading=False case
        heading_end = section['start_index']
        for elem in elements:
            if elem['type'].startswith('heading') and elem['text'].strip() == 'Introduction':
                heading_end = elem['end_index']
                break

        delete_start = heading_end
        delete_end = section['end_index']

        assert delete_start == 14  # Starts after heading
        assert delete_end == 34  # Ends at next section

    def test_calculates_range_for_last_section(self):
        """Test range calculation for the last section in document."""
        doc = create_mock_document([
            create_mock_paragraph('First', 1, 'HEADING_1'),
            create_mock_paragraph('First content', 8, 'NORMAL_TEXT'),
            create_mock_paragraph('Last Section', 23, 'HEADING_1'),
            create_mock_paragraph('Final content', 37, 'NORMAL_TEXT'),
        ])

        section = find_section_by_heading(doc, 'Last Section')

        assert section is not None
        # Last section should extend to document end
        assert section['start_index'] == 23
        assert section['end_index'] >= 37

    def test_handles_section_with_subsections(self):
        """Test range includes subsections."""
        doc = create_mock_document([
            create_mock_paragraph('Chapter 1', 1, 'HEADING_1'),
            create_mock_paragraph('Chapter content', 12, 'NORMAL_TEXT'),
            create_mock_paragraph('Section 1.1', 29, 'HEADING_2'),
            create_mock_paragraph('Section content', 42, 'NORMAL_TEXT'),
            create_mock_paragraph('Chapter 2', 59, 'HEADING_1'),
        ])

        section = find_section_by_heading(doc, 'Chapter 1')

        assert section is not None
        assert section['start_index'] == 1
        assert section['end_index'] == 59  # Includes subsections
        assert len(section['subsections']) == 1
        assert section['subsections'][0]['heading'] == 'Section 1.1'


class TestDeleteDocSectionValidation:
    """Tests for delete_doc_section input validation."""

    def test_section_not_found_returns_none(self):
        """Test that missing heading returns None from find_section_by_heading."""
        doc = create_mock_document([
            create_mock_paragraph('Existing Section', 1, 'HEADING_1'),
            create_mock_paragraph('Content', 18, 'NORMAL_TEXT'),
        ])

        section = find_section_by_heading(doc, 'NonExistent')
        assert section is None

    def test_case_sensitive_search(self):
        """Test case-sensitive heading search."""
        doc = create_mock_document([
            create_mock_paragraph('Introduction', 1, 'HEADING_1'),
            create_mock_paragraph('Content', 14, 'NORMAL_TEXT'),
        ])

        # match_case=True (default) should not find case mismatch
        section = find_section_by_heading(doc, 'INTRODUCTION', match_case=True)
        assert section is None

    def test_case_insensitive_search(self):
        """Test case-insensitive heading search."""
        doc = create_mock_document([
            create_mock_paragraph('Introduction', 1, 'HEADING_1'),
            create_mock_paragraph('Content', 14, 'NORMAL_TEXT'),
        ])

        # match_case=False should find case mismatch
        section = find_section_by_heading(doc, 'INTRODUCTION', match_case=False)
        assert section is not None
        assert section['heading'] == 'Introduction'


class TestDeleteDocSectionEdgeCases:
    """Tests for edge cases in delete_doc_section."""

    def test_minimal_section_content(self):
        """Test section with minimal content between headings.

        Note: The docs_structure module treats consecutive headings without
        any content between them as potential "style bleed" artifacts and
        may combine them into a single section. To properly test section
        boundaries, there must be at least some content between headings.
        """
        doc = create_mock_document([
            create_mock_paragraph('Short Section', 1, 'HEADING_1'),
            create_mock_paragraph('x', 16, 'NORMAL_TEXT'),  # Minimal content
            create_mock_paragraph('Next Section', 19, 'HEADING_1'),
        ])

        section = find_section_by_heading(doc, 'Short Section')

        assert section is not None
        assert section['start_index'] == 1
        assert section['end_index'] == 19

        # Characters to delete includes heading and minimal content
        characters_to_delete = section['end_index'] - section['start_index']
        assert characters_to_delete == 18

    def test_section_at_document_end(self):
        """Test section that extends to end of document."""
        doc = create_mock_document([
            create_mock_paragraph('Only Section', 1, 'HEADING_1'),
            create_mock_paragraph('Some content here', 14, 'NORMAL_TEXT'),
        ])

        section = find_section_by_heading(doc, 'Only Section')

        assert section is not None
        assert section['start_index'] == 1
        # Should extend to end of document
        assert section['end_index'] >= 14

    def test_nested_heading_levels(self):
        """Test that lower-level headings don't terminate section."""
        doc = create_mock_document([
            create_mock_paragraph('Main Section', 1, 'HEADING_1'),
            create_mock_paragraph('Content', 14, 'NORMAL_TEXT'),
            create_mock_paragraph('Subsection', 23, 'HEADING_2'),
            create_mock_paragraph('Sub content', 35, 'NORMAL_TEXT'),
            create_mock_paragraph('Sub-subsection', 48, 'HEADING_3'),
            create_mock_paragraph('Deep content', 64, 'NORMAL_TEXT'),
            create_mock_paragraph('Next Main', 78, 'HEADING_1'),
        ])

        section = find_section_by_heading(doc, 'Main Section')

        assert section is not None
        # Section should include all nested headings
        assert section['end_index'] == 78
        assert len(section['subsections']) == 2

    def test_whitespace_in_heading(self):
        """Test heading with leading/trailing whitespace."""
        doc = create_mock_document([
            create_mock_paragraph('  Spaced Heading  ', 1, 'HEADING_1'),
            create_mock_paragraph('Content', 21, 'NORMAL_TEXT'),
        ])

        # Search should match despite whitespace
        section = find_section_by_heading(doc, 'Spaced Heading')
        assert section is not None


class TestDeleteDocSectionEndOfDocumentAdjustment:
    """Tests for the end-of-document newline exclusion logic in delete_doc_section."""

    def test_delete_end_adjusted_when_section_at_document_end(self):
        """
        Test that delete_end is adjusted by -1 when section extends to document end.

        The Google Docs API doesn't allow deleting the final newline character that
        terminates the document body segment. When delete_end equals doc_end_index,
        we must subtract 1 to avoid the API error:
        "Invalid requests[0].deleteContentRange: The range cannot include the
        newline character at the end of the segment."
        """
        # Create a document where the section extends to the end
        doc = create_mock_document([
            create_mock_paragraph('Last Section', 1, 'HEADING_1'),
            create_mock_paragraph('Final content here', 15, 'NORMAL_TEXT'),
        ])

        # Get the document body end index (what the API would return)
        body_content = doc['body']['content']
        doc_end_index = body_content[-1]['endIndex']

        # Get the section
        section = find_section_by_heading(doc, 'Last Section')
        assert section is not None

        # The section end_index should match the document end
        assert section['end_index'] == doc_end_index

        # Simulate the adjustment logic from delete_doc_section
        delete_end = section['end_index']
        if delete_end == doc_end_index:
            delete_end = delete_end - 1

        # After adjustment, delete_end should be 1 less than doc end
        assert delete_end == doc_end_index - 1

    def test_delete_end_not_adjusted_when_section_not_at_document_end(self):
        """Test that delete_end is NOT adjusted when section is not at document end."""
        doc = create_mock_document([
            create_mock_paragraph('First Section', 1, 'HEADING_1'),
            create_mock_paragraph('Content', 16, 'NORMAL_TEXT'),
            create_mock_paragraph('Second Section', 25, 'HEADING_1'),
            create_mock_paragraph('More content', 41, 'NORMAL_TEXT'),
        ])

        # Get the document body end index
        body_content = doc['body']['content']
        doc_end_index = body_content[-1]['endIndex']

        # Get the first section (not at end)
        section = find_section_by_heading(doc, 'First Section')
        assert section is not None

        # Section end should NOT equal document end
        assert section['end_index'] != doc_end_index
        assert section['end_index'] == 25  # Ends at next section start

        # Simulate the adjustment logic - should NOT be adjusted
        delete_end = section['end_index']
        original_delete_end = delete_end
        if delete_end == doc_end_index:
            delete_end = delete_end - 1

        # delete_end should remain unchanged
        assert delete_end == original_delete_end
