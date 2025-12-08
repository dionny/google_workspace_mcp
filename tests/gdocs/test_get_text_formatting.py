"""
Unit tests for get_text_formatting tool.

Tests the text formatting extraction functionality including:
- Basic formatting (bold, italic, underline, etc.)
- Font properties (size, family)
- Colors (foreground, background)
- Links
- Mixed formatting detection
- Range-based extraction
"""

from gdocs.docs_tools import (
    _extract_text_formatting_from_range,
    _build_formatting_info,
    _color_to_hex,
    _has_mixed_formatting,
)


def create_mock_paragraph_with_style(
    text: str,
    start_index: int,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    strikethrough: bool = None,
    small_caps: bool = None,
    baseline_offset: str = None,
    font_size: int = None,
    font_family: str = None,
    foreground_color: dict = None,
    background_color: dict = None,
    link_url: str = None,
):
    """Create a mock paragraph element with specified text style."""
    end_index = start_index + len(text) + 1  # +1 for newline

    text_style = {}
    if bold is not None:
        text_style["bold"] = bold
    if italic is not None:
        text_style["italic"] = italic
    if underline is not None:
        text_style["underline"] = underline
    if strikethrough is not None:
        text_style["strikethrough"] = strikethrough
    if small_caps is not None:
        text_style["smallCaps"] = small_caps
    if baseline_offset is not None:
        text_style["baselineOffset"] = baseline_offset
    if font_size is not None:
        text_style["fontSize"] = {"magnitude": font_size, "unit": "PT"}
    if font_family is not None:
        text_style["weightedFontFamily"] = {"fontFamily": font_family}
    if foreground_color is not None:
        text_style["foregroundColor"] = foreground_color
    if background_color is not None:
        text_style["backgroundColor"] = background_color
    if link_url is not None:
        text_style["link"] = {"url": link_url}

    return {
        "startIndex": start_index,
        "endIndex": end_index,
        "paragraph": {
            "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
            "elements": [
                {
                    "startIndex": start_index,
                    "endIndex": end_index,
                    "textRun": {"content": text + "\n", "textStyle": text_style},
                }
            ],
        },
    }


def create_mock_doc(content_elements):
    """Create a mock document with given content elements."""
    return {"body": {"content": content_elements}}


class TestBuildFormattingInfo:
    """Tests for the _build_formatting_info helper function."""

    def test_empty_style_returns_defaults(self):
        """Test that empty style returns default values."""
        result = _build_formatting_info({}, 10, 20, "test text")

        assert result["start_index"] == 10
        assert result["end_index"] == 20
        assert result["text"] == "test text"
        assert result["bold"] is False
        assert result["italic"] is False
        assert result["underline"] is False
        assert result["strikethrough"] is False
        assert result["small_caps"] is False
        assert result["baseline_offset"] == "NONE"
        assert "font_size" not in result
        assert "font_family" not in result
        assert "foreground_color" not in result
        assert "background_color" not in result
        assert "link_url" not in result

    def test_bold_style(self):
        """Test bold style extraction."""
        result = _build_formatting_info({"bold": True}, 0, 10, "bold text")
        assert result["bold"] is True

    def test_italic_style(self):
        """Test italic style extraction."""
        result = _build_formatting_info({"italic": True}, 0, 10, "italic text")
        assert result["italic"] is True

    def test_underline_style(self):
        """Test underline style extraction."""
        result = _build_formatting_info({"underline": True}, 0, 10, "underline")
        assert result["underline"] is True

    def test_strikethrough_style(self):
        """Test strikethrough style extraction."""
        result = _build_formatting_info({"strikethrough": True}, 0, 10, "struck")
        assert result["strikethrough"] is True

    def test_small_caps_style(self):
        """Test small caps style extraction."""
        result = _build_formatting_info({"smallCaps": True}, 0, 10, "small")
        assert result["small_caps"] is True

    def test_superscript_style(self):
        """Test superscript baseline offset extraction."""
        result = _build_formatting_info(
            {"baselineOffset": "SUPERSCRIPT"}, 0, 5, "super"
        )
        assert result["baseline_offset"] == "SUPERSCRIPT"

    def test_subscript_style(self):
        """Test subscript baseline offset extraction."""
        result = _build_formatting_info({"baselineOffset": "SUBSCRIPT"}, 0, 5, "sub")
        assert result["baseline_offset"] == "SUBSCRIPT"

    def test_baseline_offset_unspecified_normalized_to_none(self):
        """Test that BASELINE_OFFSET_UNSPECIFIED is normalized to NONE."""
        result = _build_formatting_info(
            {"baselineOffset": "BASELINE_OFFSET_UNSPECIFIED"}, 0, 10, "normal text"
        )
        assert result["baseline_offset"] == "NONE"

    def test_font_size(self):
        """Test font size extraction."""
        result = _build_formatting_info(
            {"fontSize": {"magnitude": 14, "unit": "PT"}}, 0, 10, "sized"
        )
        assert result["font_size"] == 14

    def test_font_family(self):
        """Test font family extraction."""
        result = _build_formatting_info(
            {"weightedFontFamily": {"fontFamily": "Arial"}}, 0, 10, "arial"
        )
        assert result["font_family"] == "Arial"

    def test_foreground_color(self):
        """Test foreground color extraction."""
        color = {"color": {"rgbColor": {"red": 1.0, "green": 0.0, "blue": 0.0}}}
        result = _build_formatting_info({"foregroundColor": color}, 0, 10, "red")
        assert result["foreground_color"] == "#FF0000"

    def test_background_color(self):
        """Test background color extraction."""
        color = {"color": {"rgbColor": {"red": 1.0, "green": 1.0, "blue": 0.0}}}
        result = _build_formatting_info({"backgroundColor": color}, 0, 10, "yellow")
        assert result["background_color"] == "#FFFF00"

    def test_link_url(self):
        """Test link URL extraction."""
        result = _build_formatting_info(
            {"link": {"url": "https://example.com"}}, 0, 10, "link"
        )
        assert result["link_url"] == "https://example.com"

    def test_multiple_styles_combined(self):
        """Test extraction of multiple styles at once."""
        text_style = {
            "bold": True,
            "italic": True,
            "fontSize": {"magnitude": 18, "unit": "PT"},
            "weightedFontFamily": {"fontFamily": "Georgia"},
            "foregroundColor": {
                "color": {"rgbColor": {"red": 0.0, "green": 0.0, "blue": 1.0}}
            },
        }
        result = _build_formatting_info(text_style, 100, 150, "styled text")

        assert result["start_index"] == 100
        assert result["end_index"] == 150
        assert result["text"] == "styled text"
        assert result["bold"] is True
        assert result["italic"] is True
        assert result["font_size"] == 18
        assert result["font_family"] == "Georgia"
        assert result["foreground_color"] == "#0000FF"


class TestColorToHex:
    """Tests for the _color_to_hex helper function."""

    def test_empty_color_object(self):
        """Test empty color object returns empty string."""
        assert _color_to_hex({}) == ""
        assert _color_to_hex(None) == ""

    def test_missing_rgb_color(self):
        """Test color object without rgbColor returns empty string."""
        assert _color_to_hex({"color": {}}) == ""

    def test_red_color(self):
        """Test red color conversion."""
        color = {"color": {"rgbColor": {"red": 1.0, "green": 0.0, "blue": 0.0}}}
        assert _color_to_hex(color) == "#FF0000"

    def test_green_color(self):
        """Test green color conversion."""
        color = {"color": {"rgbColor": {"red": 0.0, "green": 1.0, "blue": 0.0}}}
        assert _color_to_hex(color) == "#00FF00"

    def test_blue_color(self):
        """Test blue color conversion."""
        color = {"color": {"rgbColor": {"red": 0.0, "green": 0.0, "blue": 1.0}}}
        assert _color_to_hex(color) == "#0000FF"

    def test_black_color(self):
        """Test black color conversion (explicit zero values)."""
        color = {"color": {"rgbColor": {"red": 0.0, "green": 0.0, "blue": 0.0}}}
        assert _color_to_hex(color) == "#000000"

    def test_empty_rgb_color(self):
        """Test empty rgbColor returns empty string (no explicit color)."""
        color = {"color": {"rgbColor": {}}}
        assert _color_to_hex(color) == ""

    def test_white_color(self):
        """Test white color conversion."""
        color = {"color": {"rgbColor": {"red": 1.0, "green": 1.0, "blue": 1.0}}}
        assert _color_to_hex(color) == "#FFFFFF"

    def test_gray_color(self):
        """Test gray color conversion."""
        color = {"color": {"rgbColor": {"red": 0.5, "green": 0.5, "blue": 0.5}}}
        assert _color_to_hex(color) == "#7F7F7F"


class TestHasMixedFormatting:
    """Tests for the _has_mixed_formatting helper function."""

    def test_empty_list(self):
        """Test empty list returns False."""
        assert _has_mixed_formatting([]) is False

    def test_single_span(self):
        """Test single span returns False."""
        spans = [{"bold": True, "italic": False, "font_size": 12}]
        assert _has_mixed_formatting(spans) is False

    def test_identical_spans(self):
        """Test identical spans return False."""
        spans = [
            {"bold": True, "italic": False, "font_size": 12},
            {"bold": True, "italic": False, "font_size": 12},
        ]
        assert _has_mixed_formatting(spans) is False

    def test_different_bold(self):
        """Test different bold values return True."""
        spans = [
            {"bold": True, "italic": False},
            {"bold": False, "italic": False},
        ]
        assert _has_mixed_formatting(spans) is True

    def test_different_font_size(self):
        """Test different font sizes return True."""
        spans = [
            {"bold": False, "font_size": 12},
            {"bold": False, "font_size": 14},
        ]
        assert _has_mixed_formatting(spans) is True

    def test_different_colors(self):
        """Test different colors return True."""
        spans = [
            {"bold": False, "foreground_color": "#FF0000"},
            {"bold": False, "foreground_color": "#0000FF"},
        ]
        assert _has_mixed_formatting(spans) is True

    def test_one_with_link_one_without(self):
        """Test one span with link, one without returns True."""
        spans = [
            {"bold": False, "link_url": "https://example.com"},
            {"bold": False},
        ]
        assert _has_mixed_formatting(spans) is True


class TestExtractTextFormattingFromRange:
    """Tests for the _extract_text_formatting_from_range function."""

    def test_empty_document(self):
        """Test empty document returns empty list."""
        doc_data = create_mock_doc([])
        result = _extract_text_formatting_from_range(doc_data, 0, 100)
        assert result == []

    def test_single_paragraph_no_style(self):
        """Test extracting from paragraph with no explicit style."""
        para = create_mock_paragraph_with_style("Hello World", 1)
        doc_data = create_mock_doc([para])

        result = _extract_text_formatting_from_range(doc_data, 1, 12)

        assert len(result) == 1
        assert result[0]["text"] == "Hello World"
        assert result[0]["bold"] is False
        assert result[0]["italic"] is False

    def test_single_paragraph_with_bold(self):
        """Test extracting from bold paragraph."""
        para = create_mock_paragraph_with_style("Bold Text", 1, bold=True)
        doc_data = create_mock_doc([para])

        result = _extract_text_formatting_from_range(doc_data, 1, 10)

        assert len(result) == 1
        assert result[0]["bold"] is True

    def test_single_paragraph_with_multiple_styles(self):
        """Test extracting paragraph with multiple style attributes."""
        para = create_mock_paragraph_with_style(
            "Styled", 1, bold=True, italic=True, font_size=16, font_family="Arial"
        )
        doc_data = create_mock_doc([para])

        result = _extract_text_formatting_from_range(doc_data, 1, 7)

        assert len(result) == 1
        assert result[0]["bold"] is True
        assert result[0]["italic"] is True
        assert result[0]["font_size"] == 16
        assert result[0]["font_family"] == "Arial"

    def test_partial_range_extraction(self):
        """Test extracting a partial range from a text run."""
        # Text "Hello World\n" at indices 1-13
        para = create_mock_paragraph_with_style("Hello World", 1, bold=True)
        doc_data = create_mock_doc([para])

        # Only get "World" (indices 7-12)
        result = _extract_text_formatting_from_range(doc_data, 7, 12)

        assert len(result) == 1
        assert result[0]["text"] == "World"
        assert result[0]["start_index"] == 7
        assert result[0]["end_index"] == 12

    def test_range_outside_content(self):
        """Test range completely outside content returns empty list."""
        para = create_mock_paragraph_with_style("Hello", 1)  # Indices 1-7
        doc_data = create_mock_doc([para])

        result = _extract_text_formatting_from_range(doc_data, 100, 200)
        assert result == []

    def test_with_link(self):
        """Test extracting text with link."""
        para = create_mock_paragraph_with_style(
            "Click here", 1, link_url="https://example.com"
        )
        doc_data = create_mock_doc([para])

        result = _extract_text_formatting_from_range(doc_data, 1, 11)

        assert len(result) == 1
        assert result[0]["link_url"] == "https://example.com"

    def test_with_colors(self):
        """Test extracting text with colors."""
        fg_color = {"color": {"rgbColor": {"red": 1.0, "green": 0.0, "blue": 0.0}}}
        bg_color = {"color": {"rgbColor": {"red": 1.0, "green": 1.0, "blue": 0.0}}}
        para = create_mock_paragraph_with_style(
            "Colored", 1, foreground_color=fg_color, background_color=bg_color
        )
        doc_data = create_mock_doc([para])

        result = _extract_text_formatting_from_range(doc_data, 1, 8)

        assert len(result) == 1
        assert result[0]["foreground_color"] == "#FF0000"
        assert result[0]["background_color"] == "#FFFF00"


class TestExtractTextFormattingFromRangeWithMixedContent:
    """Tests for extracting formatting from documents with mixed content."""

    def test_multiple_paragraphs(self):
        """Test extracting from multiple paragraphs."""
        para1 = create_mock_paragraph_with_style("First", 1, bold=True)
        # para1 ends at index 7 (1 + 5 + 1 for newline)
        para2 = create_mock_paragraph_with_style("Second", 7, italic=True)
        doc_data = create_mock_doc([para1, para2])

        # Get both paragraphs (indices 1-14 covers both)
        result = _extract_text_formatting_from_range(doc_data, 1, 14)

        assert len(result) == 2
        assert result[0]["text"] == "First\n"
        assert result[0]["bold"] is True
        # Second paragraph includes newline in the textRun content
        assert result[1]["text"] == "Second\n"
        assert result[1]["italic"] is True

    def test_multiple_text_runs_in_paragraph(self):
        """Test extracting from paragraph with multiple text runs (mixed formatting)."""
        # Create a paragraph with two text runs: "Bold" (bold) and " Normal" (not bold)
        doc_data = {
            "body": {
                "content": [
                    {
                        "startIndex": 1,
                        "endIndex": 13,
                        "paragraph": {
                            "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                            "elements": [
                                {
                                    "startIndex": 1,
                                    "endIndex": 5,
                                    "textRun": {
                                        "content": "Bold",
                                        "textStyle": {"bold": True},
                                    },
                                },
                                {
                                    "startIndex": 5,
                                    "endIndex": 13,
                                    "textRun": {
                                        "content": " Normal\n",
                                        "textStyle": {},
                                    },
                                },
                            ],
                        },
                    }
                ]
            }
        }

        result = _extract_text_formatting_from_range(doc_data, 1, 13)

        assert len(result) == 2
        assert result[0]["text"] == "Bold"
        assert result[0]["bold"] is True
        assert result[1]["text"] == " Normal\n"
        assert result[1]["bold"] is False


class TestExtractTextFormattingFromTable:
    """Tests for extracting formatting from table cells."""

    def test_table_cell_formatting(self):
        """Test extracting formatting from table cell content."""
        doc_data = {
            "body": {
                "content": [
                    {
                        "startIndex": 1,
                        "endIndex": 50,
                        "table": {
                            "tableRows": [
                                {
                                    "tableCells": [
                                        {
                                            "startIndex": 2,
                                            "endIndex": 15,
                                            "content": [
                                                {
                                                    "paragraph": {
                                                        "elements": [
                                                            {
                                                                "startIndex": 3,
                                                                "endIndex": 12,
                                                                "textRun": {
                                                                    "content": "Cell text",
                                                                    "textStyle": {
                                                                        "bold": True
                                                                    },
                                                                },
                                                            }
                                                        ]
                                                    }
                                                }
                                            ],
                                        }
                                    ]
                                }
                            ]
                        },
                    }
                ]
            }
        }

        result = _extract_text_formatting_from_range(doc_data, 3, 12)

        assert len(result) == 1
        assert result[0]["text"] == "Cell text"
        assert result[0]["bold"] is True


class TestValidationScenarios:
    """Tests for validation error scenarios - testing the helper functions directly."""

    def test_negative_indices_handled(self):
        """Test that helper functions handle edge cases gracefully."""
        doc_data = create_mock_doc([])
        # This should not crash, just return empty
        result = _extract_text_formatting_from_range(doc_data, -10, -5)
        assert result == []

    def test_zero_range_returns_empty_text(self):
        """Test zero-width range returns span with empty text."""
        para = create_mock_paragraph_with_style("Hello", 1)
        doc_data = create_mock_doc([para])

        # Zero-width range at index 5 (within "Hello" text)
        # The function returns a span but with empty text since start==end
        result = _extract_text_formatting_from_range(doc_data, 5, 5)
        # When start_index == end_index, text substring is empty
        if result:
            assert result[0]["text"] == ""
            assert result[0]["start_index"] == 5
            assert result[0]["end_index"] == 5


class TestSearchBasedPositioning:
    """Tests for search-based positioning used in get_text_formatting."""

    def test_find_text_basic(self):
        """Test basic text search returns correct indices."""
        from gdocs.docs_helpers import find_text_in_document

        para = create_mock_paragraph_with_style("Hello World Test", 1, bold=True)
        doc_data = create_mock_doc([para])

        result = find_text_in_document(doc_data, "World")
        assert result is not None
        start, end = result
        assert start == 7  # "Hello " is 6 chars, then "World" starts at index 7
        assert end == 12   # "World" is 5 chars long

    def test_find_text_not_found(self):
        """Test search for non-existent text returns None."""
        from gdocs.docs_helpers import find_text_in_document

        para = create_mock_paragraph_with_style("Hello World", 1)
        doc_data = create_mock_doc([para])

        result = find_text_in_document(doc_data, "Missing")
        assert result is None

    def test_find_text_case_sensitive(self):
        """Test case-sensitive search."""
        from gdocs.docs_helpers import find_text_in_document

        para = create_mock_paragraph_with_style("Hello World", 1)
        doc_data = create_mock_doc([para])

        # Should not find lowercase "world"
        result = find_text_in_document(doc_data, "world", match_case=True)
        assert result is None

        # Should find with case-insensitive search
        result = find_text_in_document(doc_data, "world", match_case=False)
        assert result is not None

    def test_find_text_occurrence(self):
        """Test finding specific occurrence of repeated text."""
        from gdocs.docs_helpers import find_text_in_document

        # Create document with "test" appearing multiple times
        doc_data = {
            "body": {
                "content": [
                    {
                        "startIndex": 1,
                        "endIndex": 20,
                        "paragraph": {
                            "elements": [
                                {
                                    "startIndex": 1,
                                    "endIndex": 20,
                                    "textRun": {
                                        "content": "test one test two\n",
                                        "textStyle": {},
                                    },
                                }
                            ]
                        },
                    }
                ]
            }
        }

        # First occurrence
        result = find_text_in_document(doc_data, "test", occurrence=1)
        assert result is not None
        assert result[0] == 1  # First "test" starts at index 1

        # Second occurrence
        result = find_text_in_document(doc_data, "test", occurrence=2)
        assert result is not None
        assert result[0] == 10  # Second "test" starts at index 10

        # Last occurrence
        result = find_text_in_document(doc_data, "test", occurrence=-1)
        assert result is not None
        assert result[0] == 10

    def test_find_text_empty_search(self):
        """Test empty search string returns None."""
        from gdocs.docs_helpers import find_text_in_document

        para = create_mock_paragraph_with_style("Hello", 1)
        doc_data = create_mock_doc([para])

        result = find_text_in_document(doc_data, "")
        assert result is None

    def test_search_formatting_extraction_integration(self):
        """Test that search-found indices work with formatting extraction."""
        from gdocs.docs_helpers import find_text_in_document

        # Create a document with "Bold Text" that is bold
        para = create_mock_paragraph_with_style("Some Bold Text Here", 1, bold=True)
        doc_data = create_mock_doc([para])

        # Find "Bold Text"
        result = find_text_in_document(doc_data, "Bold Text")
        assert result is not None
        start, end = result

        # Extract formatting at found indices
        formatting = _extract_text_formatting_from_range(doc_data, start, end)
        assert len(formatting) == 1
        assert formatting[0]["text"] == "Bold Text"
        assert formatting[0]["bold"] is True
