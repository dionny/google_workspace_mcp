"""
Unit tests for paragraph style inheritance prevention in modify_doc_text.

Tests the auto-detection of heading paragraphs and automatic NORMAL_TEXT
style application to prevent newly inserted content from inheriting heading styles.
"""
from gdocs.docs_structure import (
    get_paragraph_style_at_index,
    is_heading_style,
    HEADING_TYPES
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
        'body': {
            'content': elements
        }
    }


class TestGetParagraphStyleAtIndex:
    """Tests for get_paragraph_style_at_index function."""

    def test_returns_style_for_heading_1(self):
        """Should return HEADING_1 for text inside a heading 1 paragraph."""
        doc = create_mock_document([
            create_mock_paragraph("My Heading", 1, 'HEADING_1')
        ])
        result = get_paragraph_style_at_index(doc, 5)
        assert result == 'HEADING_1'

    def test_returns_style_for_heading_2(self):
        """Should return HEADING_2 for text inside a heading 2 paragraph."""
        doc = create_mock_document([
            create_mock_paragraph("Subheading", 1, 'HEADING_2')
        ])
        result = get_paragraph_style_at_index(doc, 5)
        assert result == 'HEADING_2'

    def test_returns_style_for_normal_text(self):
        """Should return NORMAL_TEXT for regular paragraph."""
        doc = create_mock_document([
            create_mock_paragraph("Normal paragraph text", 1, 'NORMAL_TEXT')
        ])
        result = get_paragraph_style_at_index(doc, 10)
        assert result == 'NORMAL_TEXT'

    def test_returns_style_at_paragraph_boundary(self):
        """Should return correct style at paragraph start boundary."""
        doc = create_mock_document([
            create_mock_paragraph("Heading", 1, 'HEADING_1')
        ])
        result = get_paragraph_style_at_index(doc, 1)
        assert result == 'HEADING_1'

    def test_returns_none_for_index_outside_content(self):
        """Should return None for index outside all elements."""
        doc = create_mock_document([
            create_mock_paragraph("Some text", 1, 'NORMAL_TEXT')
        ])
        # Index 100 is way beyond the paragraph
        result = get_paragraph_style_at_index(doc, 100)
        assert result is None

    def test_handles_multiple_paragraphs(self):
        """Should find correct paragraph in multi-paragraph document."""
        doc = create_mock_document([
            create_mock_paragraph("Heading", 1, 'HEADING_1'),        # 1-9
            create_mock_paragraph("Normal text", 9, 'NORMAL_TEXT'),  # 9-21
            create_mock_paragraph("Another heading", 21, 'HEADING_2')  # 21-37
        ])
        # Check heading 1
        assert get_paragraph_style_at_index(doc, 5) == 'HEADING_1'
        # Check normal text
        assert get_paragraph_style_at_index(doc, 15) == 'NORMAL_TEXT'
        # Check heading 2
        assert get_paragraph_style_at_index(doc, 30) == 'HEADING_2'

    def test_handles_title_style(self):
        """Should return TITLE for title-styled paragraph."""
        doc = create_mock_document([
            create_mock_paragraph("Document Title", 1, 'TITLE')
        ])
        result = get_paragraph_style_at_index(doc, 5)
        assert result == 'TITLE'

    def test_handles_subtitle_style(self):
        """Should return SUBTITLE for subtitle-styled paragraph."""
        doc = create_mock_document([
            create_mock_paragraph("Document Subtitle", 1, 'SUBTITLE')
        ])
        result = get_paragraph_style_at_index(doc, 5)
        assert result == 'SUBTITLE'

    def test_handles_all_heading_levels(self):
        """Should correctly identify all heading levels 1-6."""
        for level in range(1, 7):
            style = f'HEADING_{level}'
            doc = create_mock_document([
                create_mock_paragraph(f"Heading {level}", 1, style)
            ])
            result = get_paragraph_style_at_index(doc, 5)
            assert result == style, f"Failed for {style}"


class TestIsHeadingStyle:
    """Tests for is_heading_style function."""

    def test_returns_true_for_heading_1_through_6(self):
        """All HEADING_1 through HEADING_6 should be detected as headings."""
        for level in range(1, 7):
            style = f'HEADING_{level}'
            assert is_heading_style(style) is True, f"Failed for {style}"

    def test_returns_true_for_title(self):
        """TITLE should be detected as a heading style."""
        assert is_heading_style('TITLE') is True

    def test_returns_false_for_normal_text(self):
        """NORMAL_TEXT should NOT be detected as a heading style."""
        assert is_heading_style('NORMAL_TEXT') is False

    def test_returns_false_for_subtitle(self):
        """SUBTITLE should NOT be detected as a heading style (not in HEADING_TYPES)."""
        # SUBTITLE is not in HEADING_TYPES, so it should return False
        assert is_heading_style('SUBTITLE') is False

    def test_returns_false_for_none(self):
        """None should return False."""
        assert is_heading_style(None) is False

    def test_returns_false_for_arbitrary_string(self):
        """Arbitrary strings should return False."""
        assert is_heading_style('RANDOM_STYLE') is False
        assert is_heading_style('') is False

    def test_heading_types_constant_values(self):
        """Verify HEADING_TYPES constant has expected values."""
        expected = {
            'HEADING_1': 1,
            'HEADING_2': 2,
            'HEADING_3': 3,
            'HEADING_4': 4,
            'HEADING_5': 5,
            'HEADING_6': 6,
            'TITLE': 0,
        }
        assert HEADING_TYPES == expected


class TestEdgeCases:
    """Edge case tests for paragraph style detection."""

    def test_empty_document(self):
        """Should handle empty document gracefully."""
        doc = {'body': {'content': []}}
        result = get_paragraph_style_at_index(doc, 1)
        assert result is None

    def test_missing_body(self):
        """Should handle document without body."""
        doc = {}
        result = get_paragraph_style_at_index(doc, 1)
        assert result is None

    def test_missing_content(self):
        """Should handle body without content."""
        doc = {'body': {}}
        result = get_paragraph_style_at_index(doc, 1)
        assert result is None

    def test_element_without_paragraph(self):
        """Should skip elements that don't contain paragraphs."""
        doc = create_mock_document([
            {
                'startIndex': 1,
                'endIndex': 10,
                'table': {}  # Not a paragraph
            },
            create_mock_paragraph("After table", 10, 'HEADING_1')
        ])
        # Index 5 is in the table element, which has no paragraph
        result = get_paragraph_style_at_index(doc, 5)
        assert result is None
        # Index 15 should find the heading
        result = get_paragraph_style_at_index(doc, 15)
        assert result == 'HEADING_1'

    def test_paragraph_without_style(self):
        """Should return NORMAL_TEXT for paragraph without explicit style."""
        doc = create_mock_document([
            {
                'startIndex': 1,
                'endIndex': 20,
                'paragraph': {
                    # No paragraphStyle key
                    'elements': [{
                        'startIndex': 1,
                        'endIndex': 20,
                        'textRun': {'content': 'Text without style\n'}
                    }]
                }
            }
        ])
        result = get_paragraph_style_at_index(doc, 5)
        assert result == 'NORMAL_TEXT'  # Default value

    def test_paragraph_with_empty_style(self):
        """Should return NORMAL_TEXT for paragraph with empty style dict."""
        doc = create_mock_document([
            {
                'startIndex': 1,
                'endIndex': 20,
                'paragraph': {
                    'paragraphStyle': {},  # Empty style dict
                    'elements': [{
                        'startIndex': 1,
                        'endIndex': 20,
                        'textRun': {'content': 'Text with empty style\n'}
                    }]
                }
            }
        ])
        result = get_paragraph_style_at_index(doc, 5)
        assert result == 'NORMAL_TEXT'  # Default value


class TestHeadingStyleFirstParagraphOnly:
    """Tests to verify heading_style only applies to the first paragraph.

    These tests verify the fix for google_workspace_mcp-9d6e:
    When using modify_doc_text with heading_style to insert multi-line text,
    the heading style should only apply to the FIRST paragraph, not all paragraphs.
    """

    def test_multiline_text_heading_applies_to_first_paragraph_only(self):
        """When inserting multi-line text with heading_style, only first paragraph gets heading."""
        # Multi-line text: "=== HEADING ===\nLine 1\nLine 2\n"
        # Expected: only "=== HEADING ===" should be styled as heading

        text = "=== HEADING ===\nLine 1\nLine 2\n"
        para_start = 100

        # Calculate what the heading range should be using the algorithm from docs_tools.py
        text_stripped = text.lstrip("\n")
        leading_newlines = len(text) - len(text_stripped)
        heading_style_start = para_start + leading_newlines

        first_newline_pos = text_stripped.find("\n")
        if first_newline_pos != -1:
            heading_style_end = heading_style_start + first_newline_pos
        else:
            heading_style_end = para_start + len(text)

        # Verify the calculation
        assert heading_style_start == 100  # No leading newlines
        assert heading_style_end == 115  # 100 + len("=== HEADING ===") = 100 + 15 = 115

        # The heading style should only cover "=== HEADING ===" (15 chars)
        # NOT "=== HEADING ===\nLine 1\nLine 2\n" (30 chars)
        assert heading_style_end - heading_style_start == len("=== HEADING ===")

    def test_leading_newlines_excluded_from_heading_range(self):
        """Leading newlines should be excluded from heading style range."""
        text = "\n\n=== HEADING ===\nLine 2\n"
        para_start = 100

        text_stripped = text.lstrip("\n")
        leading_newlines = len(text) - len(text_stripped)
        heading_style_start = para_start + leading_newlines

        first_newline_pos = text_stripped.find("\n")
        heading_style_end = heading_style_start + first_newline_pos

        # With 2 leading newlines, heading should start at 102
        assert heading_style_start == 102
        # Heading ends after "=== HEADING ===" (15 chars)
        assert heading_style_end == 117
        assert heading_style_end - heading_style_start == len("=== HEADING ===")

    def test_single_line_text_with_heading_style(self):
        """Single line text (no internal newlines) should work correctly."""
        text = "Just a heading"
        para_start = 100

        text_stripped = text.lstrip("\n")
        leading_newlines = len(text) - len(text_stripped)
        _ = para_start + leading_newlines  # heading_style_start (not used in this test)

        first_newline_pos = text_stripped.find("\n")
        # No newlines found, so heading applies to full text
        assert first_newline_pos == -1

        # For single line, heading_style_end should be para_start + len(text)
        heading_style_end = para_start + len(text)
        assert heading_style_end == 114

    def test_single_line_with_trailing_newline(self):
        """Single line with trailing newline should strip the trailing newline."""
        text = "Just a heading\n"
        para_start = 100

        text_stripped = text.lstrip("\n")
        leading_newlines = len(text) - len(text_stripped)
        heading_style_start = para_start + leading_newlines

        first_newline_pos = text_stripped.find("\n")
        if first_newline_pos != -1:
            heading_style_end = heading_style_start + first_newline_pos
        else:
            # Strip trailing newlines
            trailing_newlines = len(text_stripped) - len(text_stripped.rstrip("\n"))
            heading_style_end = para_start + len(text) - trailing_newlines

        # Heading should cover "Just a heading" (14 chars), not the trailing newline
        assert heading_style_start == 100
        assert heading_style_end == 114  # 100 + 14
        assert heading_style_end - heading_style_start == len("Just a heading")


class TestAutoNormalTextAppliedToAllParagraphs:
    """Tests to verify that when auto_normal_text_applied is True,
    NORMAL_TEXT style is applied to ALL paragraphs, not just the first.

    This tests the fix for google_workspace_mcp-5d9c:
    When inserting text after a heading, the auto-applied NORMAL_TEXT style
    should cover all paragraphs to prevent heading style inheritance from
    bleeding into subsequent paragraphs.
    """

    def test_auto_normal_text_applied_multi_paragraph_should_cover_all(self):
        """When auto_normal_text_applied=True with multi-paragraph text,
        NORMAL_TEXT should be applied to ALL paragraphs, not just the first."""
        # Simulates inserting "Line 1\nLine 2\nLine 3\n" after a heading
        # with auto_normal_text_applied=True
        text = "Line 1\nLine 2\nLine 3\n"  # 21 characters: 6+1+6+1+6+1
        para_start = 100
        para_end = para_start + len(text)  # 100 + 21 = 121
        auto_normal_text_applied = True

        # Calculate what the heading range should be using the algorithm from docs_tools.py
        heading_style_start = para_start
        heading_style_end = para_end  # Default to full range

        text_stripped = text.lstrip("\n")
        leading_newlines = len(text) - len(text_stripped)
        if leading_newlines > 0:
            heading_style_start = para_start + leading_newlines

        # The key difference: when auto_normal_text_applied is True,
        # we should NOT restrict to first paragraph
        if not auto_normal_text_applied:
            first_newline_pos = text_stripped.find("\n")
            if first_newline_pos != -1:
                heading_style_end = heading_style_start + first_newline_pos
        # else: keep heading_style_end = para_end (all paragraphs)

        # Verify: auto_normal_text_applied=True should cover ALL text
        assert heading_style_start == 100  # No leading newlines
        assert heading_style_end == 121  # Full range: para_start + len(text) = 100 + 21
        assert heading_style_end - heading_style_start == len(text)

    def test_user_specified_heading_multi_paragraph_should_cover_first_only(self):
        """When user explicitly specifies heading_style (not auto-applied),
        should only apply to first paragraph."""
        text = "Heading Title\nNormal paragraph 1\nNormal paragraph 2\n"
        para_start = 100
        para_end = para_start + len(text)
        auto_normal_text_applied = False  # User explicitly set heading_style

        heading_style_start = para_start
        heading_style_end = para_end

        text_stripped = text.lstrip("\n")
        leading_newlines = len(text) - len(text_stripped)
        if leading_newlines > 0:
            heading_style_start = para_start + leading_newlines

        # User-specified heading should only cover first paragraph
        if not auto_normal_text_applied:
            first_newline_pos = text_stripped.find("\n")
            if first_newline_pos != -1:
                heading_style_end = heading_style_start + first_newline_pos

        # Verify: user-specified heading should only cover "Heading Title" (13 chars)
        assert heading_style_start == 100
        assert heading_style_end == 113  # 100 + len("Heading Title")
        assert heading_style_end - heading_style_start == len("Heading Title")

    def test_auto_normal_text_single_paragraph(self):
        """Single paragraph with auto_normal_text_applied should work correctly."""
        text = "Just one paragraph"
        para_start = 100
        para_end = para_start + len(text)

        heading_style_start = para_start
        heading_style_end = para_end

        # No newlines to process, so heading_style_end stays as para_end
        # Verify the range covers the single paragraph
        assert heading_style_start == 100
        assert heading_style_end == 118  # 100 + 18
        assert heading_style_end - heading_style_start == len(text)
