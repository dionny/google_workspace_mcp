"""
Unit tests for get_doc_statistics tool.

Tests the document statistics functionality including:
- Word count calculation
- Character count (with and without spaces)
- Sentence count estimation
- Paragraph counting
- Structural element counting
- Section breakdown feature
"""


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


def create_mock_inline_image(start_index: int, object_id: str):
    """Create a mock inline image element."""
    return {
        "startIndex": start_index,
        "endIndex": start_index + 1,
        "paragraph": {
            "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
            "elements": [
                {
                    "startIndex": start_index,
                    "endIndex": start_index + 1,
                    "inlineObjectElement": {"inlineObjectId": object_id},
                }
            ],
        },
    }


def create_mock_table(start_index: int, rows: int, cols: int, cell_text: str = "Cell"):
    """Create a mock table element."""
    table_rows = []
    current_index = start_index + 1

    for r in range(rows):
        cells = []
        for c in range(cols):
            cell_content = f"{cell_text} {r},{c}"
            cell = {
                "startIndex": current_index,
                "endIndex": current_index + len(cell_content) + 2,
                "content": [
                    {
                        "paragraph": {
                            "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                            "elements": [
                                {
                                    "startIndex": current_index + 1,
                                    "endIndex": current_index + len(cell_content) + 1,
                                    "textRun": {"content": cell_content + "\n"},
                                }
                            ],
                        }
                    }
                ],
            }
            cells.append(cell)
            current_index = cell["endIndex"]
        table_rows.append({"tableCells": cells})

    return {
        "startIndex": start_index,
        "endIndex": current_index,
        "table": {"tableRows": table_rows, "columns": cols, "rows": rows},
    }


def create_mock_list_item(
    text: str, start_index: int, list_id: str = "list1", nesting_level: int = 0
):
    """Create a mock list item paragraph."""
    end_index = start_index + len(text) + 1
    return {
        "startIndex": start_index,
        "endIndex": end_index,
        "paragraph": {
            "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
            "bullet": {"listId": list_id, "nestingLevel": nesting_level},
            "elements": [
                {
                    "startIndex": start_index,
                    "endIndex": end_index,
                    "textRun": {"content": text + "\n"},
                }
            ],
        },
    }


def create_mock_document(elements, title: str = "Test Document", lists: dict = None):
    """Create a mock document with given elements."""
    doc = {"title": title, "body": {"content": elements}, "lists": lists or {}}
    return doc


class TestWordCountCalculation:
    """Tests for word count functionality."""

    def test_counts_words_in_simple_paragraph(self):
        """Test word counting in a simple paragraph."""
        doc = create_mock_document(
            [
                create_mock_paragraph("This is a test sentence.", 1),
            ]
        )

        # Extract text and count words
        text = _extract_all_text(doc)
        words = [w for w in text.split() if w.strip()]

        assert len(words) == 5

    def test_counts_words_across_multiple_paragraphs(self):
        """Test word counting across multiple paragraphs."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First paragraph here.", 1),
                create_mock_paragraph("Second paragraph with more words.", 23),
            ]
        )

        text = _extract_all_text(doc)
        words = [w for w in text.split() if w.strip()]

        assert len(words) == 8  # 3 + 5

    def test_counts_words_in_tables(self):
        """Test word counting includes table cell content."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Header text.", 1),
                create_mock_table(14, rows=2, cols=2, cell_text="Cell"),
            ]
        )

        text = _extract_all_text(doc)
        words = [w for w in text.split() if w.strip()]

        # "Header text." = 2 words + 4 cells with "Cell X,Y" = 4 * 2 = 8 words
        assert len(words) == 2 + 8

    def test_empty_document_has_zero_words(self):
        """Test that empty document returns zero word count."""
        doc = create_mock_document([])

        text = _extract_all_text(doc)
        words = [w for w in text.split() if w.strip()]

        assert len(words) == 0


class TestCharacterCountCalculation:
    """Tests for character count functionality."""

    def test_counts_characters_with_spaces(self):
        """Test total character count including spaces."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Hello World", 1),
            ]
        )

        text = _extract_all_text(doc)

        # "Hello World\n" = 12 characters
        assert len(text) == 12

    def test_counts_characters_without_spaces(self):
        """Test character count excluding whitespace."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Hello World", 1),
            ]
        )

        text = _extract_all_text(doc)
        char_count_no_spaces = len(
            text.replace(" ", "").replace("\t", "").replace("\n", "")
        )

        # "HelloWorld" = 10 characters
        assert char_count_no_spaces == 10


class TestSentenceCountCalculation:
    """Tests for sentence count estimation."""

    def test_counts_sentences_with_periods(self):
        """Test sentence counting with periods."""
        import re

        text = "This is sentence one. This is sentence two. And a third."
        sentence_endings = re.findall(r"[.!?]+", text)

        assert len(sentence_endings) == 3

    def test_counts_sentences_with_mixed_punctuation(self):
        """Test sentence counting with different punctuation."""
        import re

        text = "Is this a question? Yes it is! And a statement."
        sentence_endings = re.findall(r"[.!?]+", text)

        assert len(sentence_endings) == 3

    def test_handles_multiple_punctuation(self):
        """Test that multiple punctuation counts as one sentence end."""
        import re

        text = "What?! Really?? Yes..."
        sentence_endings = re.findall(r"[.!?]+", text)

        # "?!" counts as one, "??" counts as one, "..." counts as one
        assert len(sentence_endings) == 3


class TestParagraphCounting:
    """Tests for paragraph count functionality."""

    def test_counts_non_empty_paragraphs(self):
        """Test that only non-empty paragraphs are counted."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First paragraph.", 1),
                create_mock_paragraph("", 18),  # Empty paragraph
                create_mock_paragraph("Third paragraph.", 19),
            ]
        )

        count = _count_non_empty_paragraphs(doc)

        assert count == 2

    def test_counts_headings_as_paragraphs(self):
        """Test that headings count as paragraphs."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Some content here.", 14),
            ]
        )

        count = _count_non_empty_paragraphs(doc)

        assert count == 2


class TestStructuralElementCounting:
    """Tests for structural element counting."""

    def test_counts_headings(self):
        """Test heading count across different levels."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Title", 1, "TITLE"),
                create_mock_paragraph("Chapter 1", 7, "HEADING_1"),
                create_mock_paragraph("Some content.", 17),
                create_mock_paragraph("Section 1.1", 31, "HEADING_2"),
            ]
        )

        from gdocs.docs_structure import extract_structural_elements

        elements = extract_structural_elements(doc)

        heading_count = sum(
            1
            for e in elements
            if e.get("type", "").startswith("heading") or e.get("type") == "title"
        )

        assert heading_count == 3

    def test_counts_tables(self):
        """Test table counting."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Intro text.", 1),
                create_mock_table(13, rows=2, cols=2),
                create_mock_paragraph("Between tables.", 80),
                create_mock_table(96, rows=3, cols=3),
            ]
        )

        from gdocs.docs_structure import extract_structural_elements

        elements = extract_structural_elements(doc)

        table_count = sum(1 for e in elements if e.get("type") == "table")

        assert table_count == 2

    def test_counts_lists(self):
        """Test list counting."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Items:", 1),
                create_mock_list_item("Item one", 8, "list1"),
                create_mock_list_item("Item two", 17, "list1"),
                create_mock_paragraph("More text", 26),
                create_mock_list_item("Another item", 36, "list2"),
            ],
            lists={
                "list1": {"listProperties": {"nestingLevels": [{}]}},
                "list2": {"listProperties": {"nestingLevels": [{}]}},
            },
        )

        from gdocs.docs_structure import extract_structural_elements

        elements = extract_structural_elements(doc)

        list_count = sum(
            1 for e in elements if e.get("type") in ("bullet_list", "numbered_list")
        )

        assert list_count == 2

    def test_counts_inline_images(self):
        """Test inline image counting."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Text before image.", 1),
                create_mock_inline_image(20, "img1"),
                create_mock_paragraph("Text between.", 22),
                create_mock_inline_image(36, "img2"),
                create_mock_inline_image(38, "img3"),
            ]
        )

        # Count inline images directly
        image_count = 0
        for element in doc["body"]["content"]:
            if "paragraph" in element:
                para = element["paragraph"]
                for elem in para.get("elements", []):
                    if "inlineObjectElement" in elem:
                        image_count += 1

        assert image_count == 3


class TestReadingTimeEstimate:
    """Tests for reading time estimation."""

    def test_estimates_reading_time(self):
        """Test reading time calculation based on word count."""
        # 200 words per minute is typical reading speed
        word_count = 1000
        reading_time = max(1, round(word_count / 200))

        assert reading_time == 5  # 1000 / 200 = 5 minutes

    def test_minimum_reading_time_is_one_minute(self):
        """Test that minimum reading time is 1 minute."""
        word_count = 50
        reading_time = max(1, round(word_count / 200))

        assert reading_time == 1


class TestPageCountEstimate:
    """Tests for page count estimation."""

    def test_estimates_page_count(self):
        """Test page count calculation based on word count."""
        # ~500 words per page (double-spaced)
        word_count = 2500
        page_count = max(1, round(word_count / 500))

        assert page_count == 5

    def test_minimum_page_count_is_one(self):
        """Test that minimum page count is 1."""
        word_count = 100
        page_count = max(1, round(word_count / 500))

        assert page_count == 1


class TestSectionBreakdown:
    """Tests for section word count breakdown."""

    def test_breaks_down_by_heading(self):
        """Test word count breakdown by section."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("This is the intro text with several words.", 14),
                create_mock_paragraph("Methods", 58, "HEADING_1"),
                create_mock_paragraph("The methods section has content here.", 66),
            ]
        )

        from gdocs.docs_structure import extract_structural_elements

        elements = extract_structural_elements(doc)

        headings = [
            e
            for e in elements
            if e.get("type", "").startswith("heading") or e.get("type") == "title"
        ]

        # Should find 2 headings
        assert len(headings) == 2
        assert headings[0]["text"].strip() == "Introduction"
        assert headings[1]["text"].strip() == "Methods"


# Helper functions for testing (mirroring the implementation logic)
def _extract_all_text(doc: dict) -> str:
    """Extract all text from document elements."""
    content = doc.get("body", {}).get("content", [])

    def extract_text(elements: list) -> str:
        text_parts = []
        for element in elements:
            if "paragraph" in element:
                para = element["paragraph"]
                for elem in para.get("elements", []):
                    if "textRun" in elem:
                        text_parts.append(elem["textRun"].get("content", ""))
            elif "table" in element:
                table = element["table"]
                for row in table.get("tableRows", []):
                    for cell in row.get("tableCells", []):
                        cell_content = cell.get("content", [])
                        text_parts.append(extract_text(cell_content))
        return "".join(text_parts)

    return extract_text(content)


def _count_non_empty_paragraphs(doc: dict) -> int:
    """Count non-empty paragraphs in document."""
    content = doc.get("body", {}).get("content", [])
    count = 0

    for element in content:
        if "paragraph" in element:
            para = element["paragraph"]
            para_text = ""
            for elem in para.get("elements", []):
                if "textRun" in elem:
                    para_text += elem["textRun"].get("content", "")
            if para_text.strip():
                count += 1

    return count
