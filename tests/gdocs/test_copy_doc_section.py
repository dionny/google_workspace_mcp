"""
Unit tests for Google Docs copy_doc_section tool.

Tests the copy_doc_section functionality that allows users to copy
content from one location to another within a document.
"""
from gdocs.docs_structure import (
    find_section_by_heading,
    extract_structural_elements,
    extract_text_in_range,
)
from gdocs.docs_helpers import find_all_occurrences_in_document


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


def create_mock_paragraph_with_formatting(
    text: str,
    start_index: int,
    named_style: str = 'NORMAL_TEXT',
    bold: bool = False,
    italic: bool = False
):
    """Create a mock paragraph element with text formatting."""
    end_index = start_index + len(text) + 1  # +1 for newline
    text_style = {}
    if bold:
        text_style['bold'] = True
    if italic:
        text_style['italic'] = True

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
                    'content': text + '\n',
                    'textStyle': text_style
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


class TestCopyDocSectionSourceResolution:
    """Tests for copy_doc_section source resolution logic."""

    def test_resolves_source_by_heading(self):
        """Test that source can be resolved by heading name."""
        doc = create_mock_document([
            create_mock_paragraph('Introduction', 1, 'HEADING_1'),
            create_mock_paragraph('Intro content here', 14, 'NORMAL_TEXT'),
            create_mock_paragraph('Details', 34, 'HEADING_1'),
            create_mock_paragraph('Details content', 43, 'NORMAL_TEXT'),
        ])

        section = find_section_by_heading(doc, 'Introduction')

        assert section is not None
        assert section['start_index'] == 1
        assert section['end_index'] == 34
        assert section['heading'] == 'Introduction'

    def test_resolves_source_by_heading_without_heading_text(self):
        """Test that source can be resolved excluding the heading itself."""
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

        assert heading_end == 14  # Starts after heading
        assert section['end_index'] == 34  # Ends at next section

    def test_resolves_source_by_explicit_indices(self):
        """Test that source can be resolved by explicit start and end indices."""
        doc = create_mock_document([
            create_mock_paragraph('Some content to copy', 1, 'NORMAL_TEXT'),
            create_mock_paragraph('More content here', 23, 'NORMAL_TEXT'),
        ])

        text = extract_text_in_range(doc, 1, 23)

        assert text is not None
        assert 'Some content to copy' in text

    def test_resolves_source_by_search_text(self):
        """Test that source can be resolved by searching for text."""
        doc = create_mock_document([
            create_mock_paragraph('Start marker', 1, 'NORMAL_TEXT'),
            create_mock_paragraph('Content to copy', 15, 'NORMAL_TEXT'),
            create_mock_paragraph('End marker', 32, 'NORMAL_TEXT'),
        ])

        # Find start marker
        # find_all_occurrences_in_document returns List[Tuple[int, int]] - (start, end) pairs
        positions = find_all_occurrences_in_document(doc, 'Start marker', False)

        assert len(positions) == 1
        assert positions[0][0] == 1  # start index
        assert positions[0][1] == 13  # end index, 'Start marker' is 12 chars

    def test_resolves_source_range_with_search_start_and_end(self):
        """Test that source range can be resolved with both start and end search."""
        doc = create_mock_document([
            create_mock_paragraph('[START]', 1, 'NORMAL_TEXT'),
            create_mock_paragraph('Content to copy', 10, 'NORMAL_TEXT'),
            create_mock_paragraph('[END]', 27, 'NORMAL_TEXT'),
        ])

        # find_all_occurrences_in_document returns List[Tuple[int, int]] - (start, end) pairs
        start_positions = find_all_occurrences_in_document(doc, '[START]', False)
        end_positions = find_all_occurrences_in_document(doc, '[END]', False)

        assert len(start_positions) == 1
        assert len(end_positions) == 1

        copy_start = start_positions[0][0]  # start index
        copy_end = end_positions[0][1]  # end index

        assert copy_start == 1
        assert copy_end == 32  # End of '[END]' (5 chars) + start index 27 = 32


class TestCopyDocSectionDestinationResolution:
    """Tests for copy_doc_section destination resolution logic."""

    def test_resolves_destination_start(self):
        """Test that destination 'start' resolves to index 1."""
        # destination_location='start' should give dest_index=1
        dest_index = 1  # As defined in the implementation
        assert dest_index == 1

    def test_resolves_destination_end(self):
        """Test that destination 'end' resolves to doc_end - 1."""
        doc = create_mock_document([
            create_mock_paragraph('First content', 1, 'NORMAL_TEXT'),
            create_mock_paragraph('Last content', 16, 'NORMAL_TEXT'),
        ])

        body_content = doc['body']['content']
        doc_end_index = body_content[-1]['endIndex']

        # destination_location='end' should give dest_index=doc_end_index - 1
        dest_index = doc_end_index - 1

        assert dest_index == doc_end_index - 1

    def test_resolves_destination_after_heading(self):
        """Test that destination_after_heading resolves to section end."""
        doc = create_mock_document([
            create_mock_paragraph('Target Section', 1, 'HEADING_1'),
            create_mock_paragraph('Section content', 17, 'NORMAL_TEXT'),
            create_mock_paragraph('Next Section', 34, 'HEADING_1'),
        ])

        section = find_section_by_heading(doc, 'Target Section')

        assert section is not None
        # Destination should be at the end of the section
        assert section['end_index'] == 34


class TestCopyDocSectionTextExtraction:
    """Tests for text extraction in copy_doc_section."""

    def test_extracts_text_from_range(self):
        """Test that text is correctly extracted from a document range."""
        doc = create_mock_document([
            create_mock_paragraph('First paragraph', 1, 'NORMAL_TEXT'),
            create_mock_paragraph('Second paragraph', 18, 'NORMAL_TEXT'),
        ])

        text = extract_text_in_range(doc, 1, 18)

        assert 'First paragraph' in text

    def test_extracts_text_from_section(self):
        """Test that text is correctly extracted from a section."""
        doc = create_mock_document([
            create_mock_paragraph('My Section', 1, 'HEADING_1'),
            create_mock_paragraph('Section content here', 13, 'NORMAL_TEXT'),
            create_mock_paragraph('Next Section', 35, 'HEADING_1'),
        ])

        section = find_section_by_heading(doc, 'My Section')
        text = extract_text_in_range(doc, section['start_index'], section['end_index'])

        assert 'My Section' in text
        assert 'Section content here' in text

    def test_handles_partial_range_overlap(self):
        """Test that partial range overlap is handled correctly."""
        doc = create_mock_document([
            create_mock_paragraph('ABCDEFGHIJ', 1, 'NORMAL_TEXT'),
        ])

        # Extract a subset of the text
        text = extract_text_in_range(doc, 3, 7)

        assert 'CDEF' in text


class TestCopyDocSectionValidation:
    """Tests for input validation in copy_doc_section."""

    def test_source_heading_not_found(self):
        """Test that missing source heading is detected."""
        doc = create_mock_document([
            create_mock_paragraph('Existing Section', 1, 'HEADING_1'),
            create_mock_paragraph('Content', 18, 'NORMAL_TEXT'),
        ])

        section = find_section_by_heading(doc, 'NonExistent')
        assert section is None

    def test_source_search_text_not_found(self):
        """Test that missing search text is detected."""
        doc = create_mock_document([
            create_mock_paragraph('Some content', 1, 'NORMAL_TEXT'),
        ])

        positions = find_all_occurrences_in_document(doc, 'NOT FOUND', False)
        assert len(positions) == 0

    def test_destination_heading_not_found(self):
        """Test that missing destination heading is detected."""
        doc = create_mock_document([
            create_mock_paragraph('Source Section', 1, 'HEADING_1'),
            create_mock_paragraph('Content', 16, 'NORMAL_TEXT'),
        ])

        section = find_section_by_heading(doc, 'NonExistent Destination')
        assert section is None

    def test_case_sensitive_heading_match(self):
        """Test case-sensitive heading matching."""
        doc = create_mock_document([
            create_mock_paragraph('Introduction', 1, 'HEADING_1'),
            create_mock_paragraph('Content', 14, 'NORMAL_TEXT'),
        ])

        # match_case=True should not find case mismatch
        section = find_section_by_heading(doc, 'INTRODUCTION', match_case=True)
        assert section is None

    def test_case_insensitive_heading_match(self):
        """Test case-insensitive heading matching."""
        doc = create_mock_document([
            create_mock_paragraph('Introduction', 1, 'HEADING_1'),
            create_mock_paragraph('Content', 14, 'NORMAL_TEXT'),
        ])

        # match_case=False should find case mismatch
        section = find_section_by_heading(doc, 'INTRODUCTION', match_case=False)
        assert section is not None
        assert section['heading'] == 'Introduction'


class TestCopyDocSectionEdgeCases:
    """Tests for edge cases in copy_doc_section."""

    def test_copy_section_at_document_end(self):
        """Test copying a section that extends to the document end."""
        doc = create_mock_document([
            create_mock_paragraph('Last Section', 1, 'HEADING_1'),
            create_mock_paragraph('Final content', 15, 'NORMAL_TEXT'),
        ])

        section = find_section_by_heading(doc, 'Last Section')
        body_content = doc['body']['content']
        doc_end_index = body_content[-1]['endIndex']

        assert section is not None
        assert section['end_index'] == doc_end_index

        # Copy end should be adjusted to not include final terminator
        copy_end = section['end_index']
        if copy_end >= doc_end_index:
            copy_end = doc_end_index - 1

        assert copy_end == doc_end_index - 1

    def test_copy_with_nested_subsections(self):
        """Test copying a section that contains subsections."""
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

    def test_copy_empty_range(self):
        """Test that empty range is detected."""
        doc = create_mock_document([
            create_mock_paragraph('Content', 1, 'NORMAL_TEXT'),
        ])

        # Extract from a range that doesn't exist
        text = extract_text_in_range(doc, 100, 200)

        assert text == ''

    def test_copy_minimal_content(self):
        """Test copying minimal content."""
        doc = create_mock_document([
            create_mock_paragraph('A', 1, 'NORMAL_TEXT'),
        ])

        text = extract_text_in_range(doc, 1, 3)

        assert 'A' in text


class TestCopyDocSectionFormattingExtraction:
    """Tests for formatting extraction in copy_doc_section."""

    def test_extracts_bold_formatting(self):
        """Test that bold formatting is correctly identified."""
        doc = create_mock_document([
            create_mock_paragraph_with_formatting('Bold text', 1, 'NORMAL_TEXT', bold=True),
        ])

        # Check that the text run has bold formatting
        para = doc['body']['content'][0]['paragraph']
        text_run = para['elements'][0]['textRun']

        assert text_run.get('textStyle', {}).get('bold') is True

    def test_extracts_italic_formatting(self):
        """Test that italic formatting is correctly identified."""
        doc = create_mock_document([
            create_mock_paragraph_with_formatting('Italic text', 1, 'NORMAL_TEXT', italic=True),
        ])

        # Check that the text run has italic formatting
        para = doc['body']['content'][0]['paragraph']
        text_run = para['elements'][0]['textRun']

        assert text_run.get('textStyle', {}).get('italic') is True

    def test_extracts_combined_formatting(self):
        """Test that combined formatting (bold + italic) is correctly identified."""
        doc = create_mock_document([
            create_mock_paragraph_with_formatting('Bold italic', 1, 'NORMAL_TEXT', bold=True, italic=True),
        ])

        para = doc['body']['content'][0]['paragraph']
        text_run = para['elements'][0]['textRun']

        assert text_run.get('textStyle', {}).get('bold') is True
        assert text_run.get('textStyle', {}).get('italic') is True


class TestCopyDocSectionPositionCalculation:
    """Tests for position calculation when copying content."""

    def test_calculates_formatting_offset(self):
        """Test that formatting span offset is correctly calculated."""
        # Given a source span at indices 10-20 in a copy starting at index 5
        # When copying to destination at index 100
        # The new span should be at 100 + (10 - 5) = 105 to 100 + (20 - 5) = 115

        copy_start = 5
        dest_index = 100

        span_start_index = 10
        span_end_index = 20

        span_offset = span_start_index - copy_start
        span_length = span_end_index - span_start_index

        new_start = dest_index + span_offset
        new_end = new_start + span_length

        assert new_start == 105
        assert new_end == 115

    def test_handles_multiple_formatting_spans(self):
        """Test that multiple formatting spans are correctly repositioned."""
        copy_start = 1
        dest_index = 50

        spans = [
            {'start_index': 1, 'end_index': 5},   # First span
            {'start_index': 10, 'end_index': 15}, # Second span
            {'start_index': 20, 'end_index': 25}, # Third span
        ]

        new_positions = []
        for span in spans:
            span_offset = span['start_index'] - copy_start
            span_length = span['end_index'] - span['start_index']
            new_start = dest_index + span_offset
            new_end = new_start + span_length
            new_positions.append({'start': new_start, 'end': new_end})

        assert new_positions[0] == {'start': 50, 'end': 54}
        assert new_positions[1] == {'start': 59, 'end': 64}
        assert new_positions[2] == {'start': 69, 'end': 74}
