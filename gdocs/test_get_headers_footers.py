"""
Unit tests for get_doc_headers_footers tool.

Tests the header/footer content extraction functionality including:
- Extracting content from headers and footers
- Filtering by section type (header only, footer only, both)
- Handling empty headers/footers
- Handling documents without headers/footers
- Type inference from section IDs
"""

from gdocs.docs_tools import (
    _extract_section_text,
    _infer_header_footer_type,
)


class TestExtractSectionText:
    """Tests for the _extract_section_text helper function."""

    def test_empty_section_data(self):
        """Test empty section data returns empty string."""
        result = _extract_section_text({})
        assert result == ""

    def test_empty_content_list(self):
        """Test empty content list returns empty string."""
        section_data = {"content": []}
        result = _extract_section_text(section_data)
        assert result == ""

    def test_simple_text_content(self):
        """Test extracting simple text from header/footer."""
        section_data = {
            "content": [
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "My Header\n"}}
                        ]
                    }
                }
            ]
        }
        result = _extract_section_text(section_data)
        assert result == "My Header"

    def test_multiple_paragraphs(self):
        """Test extracting text from multiple paragraphs."""
        section_data = {
            "content": [
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "Line 1\n"}}
                        ]
                    }
                },
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "Line 2\n"}}
                        ]
                    }
                }
            ]
        }
        result = _extract_section_text(section_data)
        assert result == "Line 1\nLine 2"

    def test_multiple_text_runs_in_paragraph(self):
        """Test extracting from paragraph with multiple text runs."""
        section_data = {
            "content": [
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "Bold "}},
                            {"textRun": {"content": "Text\n"}}
                        ]
                    }
                }
            ]
        }
        result = _extract_section_text(section_data)
        assert result == "Bold Text"

    def test_only_newline_returns_empty(self):
        """Test that section with only newline is treated as empty after strip."""
        section_data = {
            "content": [
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "\n"}}
                        ]
                    }
                }
            ]
        }
        result = _extract_section_text(section_data)
        assert result == ""

    def test_non_paragraph_elements_ignored(self):
        """Test that non-paragraph elements are ignored."""
        section_data = {
            "content": [
                {"sectionBreak": {}},
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "Header text\n"}}
                        ]
                    }
                },
                {"table": {}}
            ]
        }
        result = _extract_section_text(section_data)
        assert result == "Header text"

    def test_elements_without_text_run(self):
        """Test paragraph elements without textRun are handled."""
        section_data = {
            "content": [
                {
                    "paragraph": {
                        "elements": [
                            {"inlineObjectElement": {"inlineObjectId": "obj1"}},
                            {"textRun": {"content": "Text after image\n"}}
                        ]
                    }
                }
            ]
        }
        result = _extract_section_text(section_data)
        assert result == "Text after image"


class TestInferHeaderFooterType:
    """Tests for the _infer_header_footer_type helper function."""

    def test_default_type_with_kix_prefix(self):
        """Test standard section ID returns DEFAULT."""
        assert _infer_header_footer_type("kix.abc123xyz") == "DEFAULT"

    def test_first_page_type(self):
        """Test first page detection from section ID."""
        assert _infer_header_footer_type("kix.firstpage123") == "FIRST_PAGE"
        assert _infer_header_footer_type("kix.first_header") == "FIRST_PAGE"
        assert _infer_header_footer_type("FIRST_PAGE_abc") == "FIRST_PAGE"

    def test_even_page_type(self):
        """Test even page detection from section ID."""
        assert _infer_header_footer_type("kix.evenpage123") == "EVEN_PAGE"
        assert _infer_header_footer_type("kix.even_header") == "EVEN_PAGE"
        assert _infer_header_footer_type("EVEN_PAGE_xyz") == "EVEN_PAGE"

    def test_case_insensitive(self):
        """Test that type inference is case insensitive."""
        assert _infer_header_footer_type("kix.FIRSTPAGE") == "FIRST_PAGE"
        assert _infer_header_footer_type("kix.FirstPage") == "FIRST_PAGE"
        assert _infer_header_footer_type("kix.EVENPAGE") == "EVEN_PAGE"

    def test_empty_section_id(self):
        """Test empty section ID returns DEFAULT."""
        assert _infer_header_footer_type("") == "DEFAULT"

    def test_arbitrary_section_id(self):
        """Test arbitrary section ID returns DEFAULT."""
        assert _infer_header_footer_type("random_string_123") == "DEFAULT"


class TestGetDocHeadersFootersResponseStructure:
    """Tests for verifying the response structure of get_doc_headers_footers."""

    def test_headers_dict_structure(self):
        """Test that header extraction produces expected structure."""
        header_data = {
            "content": [
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "Company Name\n"}}
                        ]
                    }
                }
            ]
        }

        content = _extract_section_text(header_data)
        header_id = "kix.abc123"

        result = {
            "type": _infer_header_footer_type(header_id),
            "content": content,
            "is_empty": not content.strip(),
        }

        assert result["type"] == "DEFAULT"
        assert result["content"] == "Company Name"
        assert result["is_empty"] is False

    def test_empty_header_is_flagged(self):
        """Test that empty header is correctly flagged."""
        header_data = {
            "content": [
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "\n"}}
                        ]
                    }
                }
            ]
        }

        content = _extract_section_text(header_data)

        result = {
            "type": _infer_header_footer_type("kix.xyz"),
            "content": content,
            "is_empty": not content.strip(),
        }

        assert result["content"] == ""
        assert result["is_empty"] is True


class TestEdgeCases:
    """Tests for edge cases in header/footer extraction."""

    def test_content_with_special_characters(self):
        """Test extracting content with special characters."""
        section_data = {
            "content": [
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "Page © 2024 - Confidential™\n"}}
                        ]
                    }
                }
            ]
        }
        result = _extract_section_text(section_data)
        assert result == "Page © 2024 - Confidential™"

    def test_content_with_unicode(self):
        """Test extracting content with unicode characters."""
        section_data = {
            "content": [
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "日本語ヘッダー 🔒\n"}}
                        ]
                    }
                }
            ]
        }
        result = _extract_section_text(section_data)
        assert result == "日本語ヘッダー 🔒"

    def test_multiline_content(self):
        """Test extracting multi-line header/footer content."""
        section_data = {
            "content": [
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "Line 1\n"}}
                        ]
                    }
                },
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "Line 2\n"}}
                        ]
                    }
                },
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "Line 3\n"}}
                        ]
                    }
                }
            ]
        }
        result = _extract_section_text(section_data)
        assert result == "Line 1\nLine 2\nLine 3"

    def test_whitespace_only_content(self):
        """Test that whitespace-only content is stripped properly."""
        section_data = {
            "content": [
                {
                    "paragraph": {
                        "elements": [
                            {"textRun": {"content": "   \t  \n"}}
                        ]
                    }
                }
            ]
        }
        result = _extract_section_text(section_data)
        # rstrip(\n) removes the newline, but not other whitespace
        assert result == "   \t  "
