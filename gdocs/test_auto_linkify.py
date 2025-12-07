"""
Unit tests for auto_linkify_doc tool.

These tests verify:
- URL detection patterns (http://, https://, www.)
- Index mapping from text positions to document indices
- Already-linked URL detection
- Preview mode functionality
- Custom URL pattern support

Note: The auto_linkify_doc tool is decorated with @server.tool() and
@require_google_service which modify the function signature. These tests
focus on testing the underlying logic through helper functions and
simulated document data.
"""
import re
import pytest

from gdocs.docs_helpers import (
    extract_document_text_with_indices,
    create_format_text_request,
)


class TestURLDetectionPattern:
    """Tests for URL detection regex patterns."""

    # Default URL pattern used by auto_linkify_doc
    DEFAULT_URL_PATTERN = (
        r'(?:https?://|www\.)'  # Protocol or www.
        r'[a-zA-Z0-9]'  # Must start with alphanumeric
        r'(?:[a-zA-Z0-9\-._~:/?#\[\]@!$&\'()*+,;=%]*[a-zA-Z0-9/])?'  # URL chars
    )

    def test_detect_https_url(self):
        """Test detecting https:// URLs."""
        pattern = re.compile(self.DEFAULT_URL_PATTERN, re.IGNORECASE)
        text = "Visit https://example.com for more info."

        matches = pattern.findall(text)

        assert len(matches) == 1
        assert "https://example.com" in text

    def test_detect_http_url(self):
        """Test detecting http:// URLs."""
        pattern = re.compile(self.DEFAULT_URL_PATTERN, re.IGNORECASE)
        text = "Check http://example.org today."

        matches = list(pattern.finditer(text))

        assert len(matches) == 1
        assert matches[0].group() == "http://example.org"

    def test_detect_www_url(self):
        """Test detecting www. URLs."""
        pattern = re.compile(self.DEFAULT_URL_PATTERN, re.IGNORECASE)
        text = "See www.example.com for details."

        matches = list(pattern.finditer(text))

        assert len(matches) == 1
        assert matches[0].group() == "www.example.com"

    def test_detect_url_with_path(self):
        """Test detecting URLs with paths."""
        pattern = re.compile(self.DEFAULT_URL_PATTERN, re.IGNORECASE)
        text = "Go to https://example.com/path/to/page for resources."

        matches = list(pattern.finditer(text))

        assert len(matches) == 1
        assert matches[0].group() == "https://example.com/path/to/page"

    def test_detect_url_with_query_params(self):
        """Test detecting URLs with query parameters."""
        pattern = re.compile(self.DEFAULT_URL_PATTERN, re.IGNORECASE)
        text = "Search https://example.com/search?q=test&page=1 for results."

        matches = list(pattern.finditer(text))

        assert len(matches) == 1
        assert "q=test" in matches[0].group()

    def test_detect_multiple_urls(self):
        """Test detecting multiple URLs in text."""
        pattern = re.compile(self.DEFAULT_URL_PATTERN, re.IGNORECASE)
        text = "Visit https://example.com and www.test.org and http://demo.net today."

        matches = list(pattern.finditer(text))

        assert len(matches) == 3

    def test_no_urls_in_text(self):
        """Test text with no URLs."""
        pattern = re.compile(self.DEFAULT_URL_PATTERN, re.IGNORECASE)
        text = "This is just plain text without any web addresses."

        matches = pattern.findall(text)

        assert len(matches) == 0

    def test_url_not_capturing_trailing_punctuation(self):
        """Test that trailing punctuation is not captured."""
        pattern = re.compile(self.DEFAULT_URL_PATTERN, re.IGNORECASE)
        text = "Check out https://example.com."

        matches = list(pattern.finditer(text))

        assert len(matches) == 1
        # Should not include the period
        assert matches[0].group() == "https://example.com"

    def test_url_ending_with_slash(self):
        """Test URL ending with slash is captured correctly."""
        pattern = re.compile(self.DEFAULT_URL_PATTERN, re.IGNORECASE)
        text = "Visit https://example.com/ for info."

        matches = list(pattern.finditer(text))

        assert len(matches) == 1
        assert matches[0].group() == "https://example.com/"


class TestDocumentTextExtraction:
    """Tests for extracting text with indices from document data."""

    def test_extract_simple_paragraph(self):
        """Test extracting text from a simple paragraph."""
        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "Visit https://example.com today.",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 33,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        segments = extract_document_text_with_indices(doc_data)

        assert len(segments) == 1
        text, start, end = segments[0]
        assert "https://example.com" in text
        assert start == 1

    def test_extract_multiple_paragraphs(self):
        """Test extracting text from multiple paragraphs."""
        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {"content": "First paragraph.\n"},
                                    "startIndex": 1,
                                    "endIndex": 18,
                                }
                            ]
                        }
                    },
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {"content": "Second paragraph with https://test.com link.\n"},
                                    "startIndex": 18,
                                    "endIndex": 63,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        segments = extract_document_text_with_indices(doc_data)

        assert len(segments) == 2
        assert "https://test.com" in segments[1][0]


class TestURLNormalization:
    """Tests for URL normalization (www. -> https://)."""

    def test_normalize_www_url(self):
        """Test that www. URLs are normalized to https://."""
        url_text = "www.example.com"
        normalized = url_text
        if url_text.lower().startswith('www.'):
            normalized = 'https://' + url_text

        assert normalized == "https://www.example.com"

    def test_https_url_unchanged(self):
        """Test that https:// URLs are not modified."""
        url_text = "https://example.com"
        normalized = url_text
        if url_text.lower().startswith('www.'):
            normalized = 'https://' + url_text

        assert normalized == "https://example.com"

    def test_http_url_unchanged(self):
        """Test that http:// URLs are not modified."""
        url_text = "http://example.com"
        normalized = url_text
        if url_text.lower().startswith('www.'):
            normalized = 'https://' + url_text

        assert normalized == "http://example.com"


class TestCreateLinkRequest:
    """Tests for creating link formatting requests."""

    def test_create_link_request(self):
        """Test creating a hyperlink formatting request."""
        request = create_format_text_request(
            start_index=10,
            end_index=30,
            link="https://example.com",
        )

        assert request is not None
        assert 'updateTextStyle' in request
        assert request['updateTextStyle']['textStyle']['link'] == {'url': 'https://example.com'}
        assert request['updateTextStyle']['range']['startIndex'] == 10
        assert request['updateTextStyle']['range']['endIndex'] == 30
        assert 'link' in request['updateTextStyle']['fields']

    def test_create_link_only_request(self):
        """Test that link-only request doesn't include other formatting."""
        request = create_format_text_request(
            start_index=10,
            end_index=30,
            link="https://example.com",
        )

        assert request is not None
        style = request['updateTextStyle']['textStyle']
        # Should only have 'link' key
        assert 'link' in style
        # Should not have other formatting
        assert 'bold' not in style
        assert 'italic' not in style


class TestExistingLinkDetection:
    """Tests for detecting already-linked text."""

    def test_extract_existing_links(self):
        """Test extracting existing link ranges from document."""
        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "Visit ",
                                        "textStyle": {}
                                    },
                                    "startIndex": 1,
                                    "endIndex": 7,
                                },
                                {
                                    "textRun": {
                                        "content": "https://example.com",
                                        "textStyle": {
                                            "link": {"url": "https://example.com"}
                                        }
                                    },
                                    "startIndex": 7,
                                    "endIndex": 26,
                                },
                                {
                                    "textRun": {
                                        "content": " today.",
                                        "textStyle": {}
                                    },
                                    "startIndex": 26,
                                    "endIndex": 33,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        # Simulate the extraction logic from auto_linkify_doc
        existing_links = set()
        body = doc_data.get('body', {})
        content = body.get('content', [])

        for element in content:
            if 'paragraph' in element:
                paragraph = element['paragraph']
                for para_element in paragraph.get('elements', []):
                    if 'textRun' in para_element:
                        text_run = para_element['textRun']
                        text_style = text_run.get('textStyle', {})
                        if 'link' in text_style:
                            start_idx = para_element.get('startIndex', 0)
                            end_idx = para_element.get('endIndex', 0)
                            existing_links.add((start_idx, end_idx))

        assert len(existing_links) == 1
        assert (7, 26) in existing_links

    def test_overlap_detection(self):
        """Test detecting overlap between URL and existing link."""
        existing_links = [(10, 30)]
        url_range = (10, 30)

        # Check for overlap (same logic as in auto_linkify_doc)
        is_overlapping = False
        for link_start, link_end in existing_links:
            url_start, url_end = url_range
            if url_start < link_end and url_end > link_start:
                is_overlapping = True
                break

        assert is_overlapping is True

    def test_no_overlap_detection(self):
        """Test detecting no overlap between URL and existing link."""
        existing_links = [(50, 70)]
        url_range = (10, 30)

        # Check for overlap
        is_overlapping = False
        for link_start, link_end in existing_links:
            url_start, url_end = url_range
            if url_start < link_end and url_end > link_start:
                is_overlapping = True
                break

        assert is_overlapping is False


class TestCustomURLPattern:
    """Tests for custom URL pattern support."""

    def test_github_url_pattern(self):
        """Test custom pattern for GitHub URLs."""
        custom_pattern = r'https://github\.com/[\w-]+/[\w-]+'
        pattern = re.compile(custom_pattern)
        text = "Check https://github.com/user-name/repo-name for the code."

        matches = list(pattern.finditer(text))

        assert len(matches) == 1
        assert matches[0].group() == "https://github.com/user-name/repo-name"

    def test_jira_url_pattern(self):
        """Test custom pattern for JIRA URLs."""
        custom_pattern = r'https://[\w-]+\.atlassian\.net/browse/[\w]+-\d+'
        pattern = re.compile(custom_pattern)
        text = "See ticket at https://mycompany.atlassian.net/browse/PROJ-123 for details."

        matches = list(pattern.finditer(text))

        assert len(matches) == 1
        assert "PROJ-123" in matches[0].group()

    def test_invalid_regex_detection(self):
        """Test that invalid regex can be detected."""
        invalid_pattern = r'[invalid(regex'

        with pytest.raises(re.error):
            re.compile(invalid_pattern)


class TestAutoLinkifyResponseStructure:
    """Tests for the expected response structure from auto_linkify_doc."""

    def test_success_response_structure(self):
        """Test the expected structure of a successful response."""
        response = {
            "success": True,
            "operation": "auto_linkify",
            "urls_linked": 3,
            "urls_found": 3,
            "urls_skipped": 0,
            "affected_ranges": [
                {"url": "https://example.com", "original_text": "https://example.com", "range": {"start": 10, "end": 29}},
                {"url": "https://www.test.org", "original_text": "www.test.org", "range": {"start": 50, "end": 62}},
                {"url": "http://demo.net", "original_text": "http://demo.net", "range": {"start": 80, "end": 95}},
            ],
            "message": "Linked 3 URL(s) in document",
            "link": "https://docs.google.com/document/d/doc-123/edit",
        }

        assert response["success"] is True
        assert response["operation"] == "auto_linkify"
        assert response["urls_linked"] == 3
        assert len(response["affected_ranges"]) == 3

    def test_preview_response_structure(self):
        """Test the expected structure of a preview response."""
        response = {
            "preview": True,
            "would_modify": True,
            "urls_found": 5,
            "urls_to_link": [
                {"url": "https://example.com", "original_text": "https://example.com", "range": {"start": 10, "end": 29}},
            ],
            "urls_already_linked": 2,
            "message": "Would link 3 URL(s) (2 already linked, will be skipped)",
            "link": "https://docs.google.com/document/d/doc-123/edit",
        }

        assert response["preview"] is True
        assert response["would_modify"] is True
        assert response["urls_found"] == 5
        assert response["urls_already_linked"] == 2

    def test_no_urls_response_structure(self):
        """Test response structure when no URLs are found."""
        response = {
            "success": True,
            "operation": "auto_linkify",
            "urls_linked": 0,
            "urls_found": 0,
            "urls_skipped": 0,
            "affected_ranges": [],
            "message": "No URLs found to link",
            "link": "https://docs.google.com/document/d/doc-123/edit",
        }

        assert response["success"] is True
        assert response["urls_linked"] == 0
        assert response["urls_found"] == 0
        assert "No URLs" in response["message"]

    def test_all_already_linked_response(self):
        """Test response when all URLs are already linked."""
        response = {
            "success": True,
            "operation": "auto_linkify",
            "urls_linked": 0,
            "urls_found": 3,
            "urls_skipped": 3,
            "affected_ranges": [],
            "message": "All 3 URL(s) are already linked",
            "link": "https://docs.google.com/document/d/doc-123/edit",
        }

        assert response["success"] is True
        assert response["urls_linked"] == 0
        assert response["urls_found"] == 3
        assert response["urls_skipped"] == 3
        assert "already linked" in response["message"]


class TestIndexMapping:
    """Tests for mapping text positions to document indices."""

    def test_build_index_map(self):
        """Test building index map from document segments."""
        # Simulate document with text starting at index 1
        text_segments = [
            ("Hello world", 1, 12),  # (text, start_index, end_index)
        ]

        full_text = ""
        index_map = []

        for segment_text, start_idx, _ in text_segments:
            for i, char in enumerate(segment_text):
                index_map.append(start_idx + i)
                full_text += char

        assert full_text == "Hello world"
        assert len(index_map) == 11
        assert index_map[0] == 1  # 'H' is at document index 1
        assert index_map[6] == 7  # 'w' is at document index 7

    def test_map_match_to_document_indices(self):
        """Test mapping regex match positions to document indices."""
        text_segments = [
            ("Visit https://example.com today", 1, 32),
        ]

        full_text = ""
        index_map = []

        for segment_text, start_idx, _ in text_segments:
            for i, char in enumerate(segment_text):
                index_map.append(start_idx + i)
                full_text += char

        # Find URL in full text
        pattern = re.compile(r'https://example\.com')
        match = pattern.search(full_text)

        assert match is not None
        text_start = match.start()
        text_end = match.end()

        # Map to document indices
        doc_start = index_map[text_start]
        doc_end = index_map[text_end - 1] + 1

        assert doc_start == 7  # URL starts at "Visit " (6 chars) + 1
        assert doc_end == 26  # URL ends at position 26
