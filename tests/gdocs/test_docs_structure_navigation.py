"""
Unit tests for Google Docs document structure navigation.

Tests the new structural navigation functions:
- extract_structural_elements
- build_headings_outline
- find_section_by_heading
- get_all_headings
- find_section_insertion_point
- find_elements_by_type
- get_element_ancestors
- get_heading_siblings
"""

from gdocs.docs_structure import (
    extract_structural_elements,
    build_headings_outline,
    find_section_by_heading,
    get_all_headings,
    find_section_insertion_point,
    find_elements_by_type,
    get_element_ancestors,
    get_heading_siblings,
    extract_text_in_range,
    HEADING_TYPES,
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


def create_mock_table(start_index: int, rows: int = 2, columns: int = 2):
    """Create a mock table element."""
    table_rows = []
    for _ in range(rows):
        cells = []
        for _ in range(columns):
            cells.append(
                {"startIndex": start_index, "endIndex": start_index + 10, "content": []}
            )
        table_rows.append({"tableCells": cells})

    return {
        "startIndex": start_index,
        "endIndex": start_index + 100,
        "table": {"tableRows": table_rows},
    }


def create_mock_list_item(
    text: str, start_index: int, list_id: str, nesting_level: int = 0
):
    """Create a mock list item paragraph element."""
    end_index = start_index + len(text) + 1  # +1 for newline
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


def create_mock_document(elements, lists=None):
    """Create a mock document with given elements.

    Args:
        elements: List of content elements (paragraphs, tables, etc.)
        lists: Optional dict of list metadata. Keys are list IDs, values are list properties.
               Example: {'kix.abc123': {'listProperties': {'nestingLevels': [{'glyphType': 'DECIMAL'}]}}}
    """
    return {
        "title": "Test Document",
        "body": {"content": elements},
        "lists": lists or {},
    }


class TestExtractStructuralElements:
    """Tests for extract_structural_elements function."""

    def test_extracts_headings(self):
        """Test that headings are correctly identified and extracted."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Some content here", 14, "NORMAL_TEXT"),
                create_mock_paragraph("Details", 33, "HEADING_2"),
            ]
        )

        elements = extract_structural_elements(doc)

        assert len(elements) == 3
        assert elements[0]["type"] == "heading1"
        assert elements[0]["text"] == "Introduction"
        assert elements[0]["level"] == 1
        assert elements[1]["type"] == "paragraph"
        assert elements[1]["text"] == "Some content here"
        assert elements[2]["type"] == "heading2"
        assert elements[2]["level"] == 2

    def test_extracts_tables(self):
        """Test that tables are correctly identified."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Before table", 1, "NORMAL_TEXT"),
                create_mock_table(20, rows=3, columns=4),
            ]
        )

        elements = extract_structural_elements(doc)

        assert len(elements) == 2
        assert elements[0]["type"] == "paragraph"
        assert elements[1]["type"] == "table"
        assert elements[1]["rows"] == 3
        assert elements[1]["columns"] == 4

    def test_extracts_all_heading_levels(self):
        """Test extraction of all heading levels 1-6."""
        content = []
        idx = 1
        for level in range(1, 7):
            style = f"HEADING_{level}"
            elem = create_mock_paragraph(f"Heading {level}", idx, style)
            content.append(elem)
            idx = elem["endIndex"]

        doc = create_mock_document(content)
        elements = extract_structural_elements(doc)

        assert len(elements) == 6
        for i, elem in enumerate(elements):
            assert elem["type"] == f"heading{i + 1}"
            assert elem["level"] == i + 1

    def test_handles_empty_document(self):
        """Test handling of document with no content."""
        doc = create_mock_document([])
        elements = extract_structural_elements(doc)
        assert elements == []

    def test_skips_empty_paragraphs(self):
        """Test that empty paragraphs are not included."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Real content", 1, "NORMAL_TEXT"),
                {
                    "startIndex": 15,
                    "endIndex": 16,
                    "paragraph": {
                        "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                        "elements": [{"textRun": {"content": "\n"}}],
                    },
                },
                create_mock_paragraph("More content", 20, "NORMAL_TEXT"),
            ]
        )

        elements = extract_structural_elements(doc)

        # Empty paragraph should be skipped
        assert len(elements) == 2
        assert elements[0]["text"] == "Real content"
        assert elements[1]["text"] == "More content"


class TestBuildHeadingsOutline:
    """Tests for build_headings_outline function."""

    def test_builds_flat_outline(self):
        """Test outline with same-level headings."""
        elements = [
            {
                "type": "heading1",
                "text": "Chapter 1",
                "level": 1,
                "start_index": 1,
                "end_index": 10,
            },
            {
                "type": "heading1",
                "text": "Chapter 2",
                "level": 1,
                "start_index": 50,
                "end_index": 60,
            },
            {
                "type": "heading1",
                "text": "Chapter 3",
                "level": 1,
                "start_index": 100,
                "end_index": 110,
            },
        ]

        outline = build_headings_outline(elements)

        assert len(outline) == 3
        assert outline[0]["text"] == "Chapter 1"
        assert outline[0]["children"] == []
        assert outline[1]["text"] == "Chapter 2"
        assert outline[2]["text"] == "Chapter 3"

    def test_builds_nested_outline(self):
        """Test outline with nested headings."""
        elements = [
            {
                "type": "heading1",
                "text": "Chapter 1",
                "level": 1,
                "start_index": 1,
                "end_index": 10,
            },
            {
                "type": "heading2",
                "text": "Section 1.1",
                "level": 2,
                "start_index": 20,
                "end_index": 30,
            },
            {
                "type": "heading2",
                "text": "Section 1.2",
                "level": 2,
                "start_index": 40,
                "end_index": 50,
            },
            {
                "type": "heading1",
                "text": "Chapter 2",
                "level": 1,
                "start_index": 60,
                "end_index": 70,
            },
        ]

        outline = build_headings_outline(elements)

        assert len(outline) == 2
        assert outline[0]["text"] == "Chapter 1"
        assert len(outline[0]["children"]) == 2
        assert outline[0]["children"][0]["text"] == "Section 1.1"
        assert outline[0]["children"][1]["text"] == "Section 1.2"
        assert outline[1]["text"] == "Chapter 2"
        assert outline[1]["children"] == []

    def test_builds_deeply_nested_outline(self):
        """Test outline with multiple levels of nesting."""
        elements = [
            {
                "type": "heading1",
                "text": "H1",
                "level": 1,
                "start_index": 1,
                "end_index": 10,
            },
            {
                "type": "heading2",
                "text": "H2",
                "level": 2,
                "start_index": 20,
                "end_index": 30,
            },
            {
                "type": "heading3",
                "text": "H3",
                "level": 3,
                "start_index": 40,
                "end_index": 50,
            },
            {
                "type": "heading4",
                "text": "H4",
                "level": 4,
                "start_index": 60,
                "end_index": 70,
            },
        ]

        outline = build_headings_outline(elements)

        assert len(outline) == 1
        assert outline[0]["text"] == "H1"
        assert len(outline[0]["children"]) == 1
        assert outline[0]["children"][0]["text"] == "H2"
        assert len(outline[0]["children"][0]["children"]) == 1
        assert outline[0]["children"][0]["children"][0]["text"] == "H3"
        assert len(outline[0]["children"][0]["children"][0]["children"]) == 1

    def test_handles_no_headings(self):
        """Test with no headings in elements."""
        elements = [
            {
                "type": "paragraph",
                "text": "Just text",
                "start_index": 1,
                "end_index": 10,
            },
        ]

        outline = build_headings_outline(elements)
        assert outline == []


class TestFindSectionByHeading:
    """Tests for find_section_by_heading function."""

    def test_finds_section(self):
        """Test finding a section by heading."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Intro content here", 14, "NORMAL_TEXT"),
                create_mock_paragraph("Details", 34, "HEADING_1"),
                create_mock_paragraph("Details content", 43, "NORMAL_TEXT"),
            ]
        )

        section = find_section_by_heading(doc, "Introduction")

        assert section is not None
        assert section["heading"] == "Introduction"
        assert section["level"] == 1
        assert section["start_index"] == 1
        # Section ends at next same-level heading
        assert section["end_index"] == 34

    def test_finds_section_case_insensitive(self):
        """Test case-insensitive heading search."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Content", 14, "NORMAL_TEXT"),
            ]
        )

        section = find_section_by_heading(doc, "INTRODUCTION", match_case=False)
        assert section is not None
        assert section["heading"] == "Introduction"

    def test_case_sensitive_fails_when_case_differs(self):
        """Test case-sensitive search fails with different case."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Content", 14, "NORMAL_TEXT"),
            ]
        )

        section = find_section_by_heading(doc, "INTRODUCTION", match_case=True)
        assert section is None

    def test_returns_none_for_missing_heading(self):
        """Test returns None when heading not found."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
            ]
        )

        section = find_section_by_heading(doc, "NonExistent")
        assert section is None

    def test_section_includes_subsections(self):
        """Test that section includes subsection headings."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Chapter 1", 1, "HEADING_1"),
                create_mock_paragraph("Content", 12, "NORMAL_TEXT"),
                create_mock_paragraph("Section 1.1", 21, "HEADING_2"),
                create_mock_paragraph("Subsection content", 34, "NORMAL_TEXT"),
                create_mock_paragraph("Chapter 2", 54, "HEADING_1"),
            ]
        )

        section = find_section_by_heading(doc, "Chapter 1")

        assert section is not None
        assert len(section["subsections"]) == 1
        assert section["subsections"][0]["heading"] == "Section 1.1"
        assert section["subsections"][0]["level"] == 2

    def test_last_section_extends_to_document_end(self):
        """Test that the last section extends to document end."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First", 1, "HEADING_1"),
                create_mock_paragraph("Content", 8, "NORMAL_TEXT"),
                create_mock_paragraph("Last Section", 17, "HEADING_1"),
                create_mock_paragraph("Final content", 31, "NORMAL_TEXT"),
            ]
        )

        section = find_section_by_heading(doc, "Last Section")

        assert section is not None
        # Should extend to end of document
        assert section["end_index"] >= 31

    def test_section_returns_content_text(self):
        """Test that section content is extracted correctly (fixes bug 8937)."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph(
                    "This is the introduction content.", 14, "NORMAL_TEXT"
                ),
                create_mock_paragraph("Next Section", 49, "HEADING_1"),
            ]
        )

        section = find_section_by_heading(doc, "Introduction")

        assert section is not None
        assert section["content"] != ""
        assert "introduction content" in section["content"].lower()

    def test_section_content_with_multiple_paragraphs(self):
        """Test section content includes multiple paragraphs."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Chapter", 1, "HEADING_1"),
                create_mock_paragraph("First paragraph here.", 10, "NORMAL_TEXT"),
                create_mock_paragraph("Second paragraph here.", 33, "NORMAL_TEXT"),
                create_mock_paragraph("Next Chapter", 57, "HEADING_1"),
            ]
        )

        section = find_section_by_heading(doc, "Chapter")

        assert section is not None
        assert "First paragraph" in section["content"]
        assert "Second paragraph" in section["content"]

    def test_section_ends_at_same_level_heading_after_false_heading(self):
        """Test that section correctly ends at real same-level heading even after false headings.

        This tests the bug where style-bleed creates empty headings, and the next
        real heading gets incorrectly marked as false because it immediately follows
        the empty (false) heading.

        Bug: google_workspace_mcp-4fcb
        """
        # Create a document where style bleed creates an empty heading between sections
        # Structure:
        # 1. 'Section A' (H2) - target section
        # 2. Content paragraph
        # 3. Empty heading (H2) - style bleed artifact (immediately after content)
        # 4. 'Section B' (H2) - real next section (immediately after empty heading)
        # 5. Section B content
        doc = create_mock_document(
            [
                create_mock_paragraph("Section A", 1, "HEADING_2"),
                create_mock_paragraph("Content in section A", 12, "NORMAL_TEXT"),
                # Empty heading created by style bleed - starts right after content
                {
                    "startIndex": 34,
                    "endIndex": 35,
                    "paragraph": {
                        "paragraphStyle": {"namedStyleType": "HEADING_2"},
                        "elements": [{"textRun": {"content": "\n"}}],
                    },
                },
                # Real next section - starts right after empty heading
                create_mock_paragraph("Section B", 35, "HEADING_2"),
                create_mock_paragraph("Content in section B", 46, "NORMAL_TEXT"),
            ]
        )

        section = find_section_by_heading(doc, "Section A")

        assert section is not None
        # Section A should end at Section B (index 35), NOT include Section B content
        assert section["end_index"] == 35
        assert "Content in section A" in section["content"]
        assert "Section B" not in section["content"]
        assert "Content in section B" not in section["content"]
        # Subsections should not include Section B
        for sub in section.get("subsections", []):
            assert sub["heading"] != "Section B"


class TestExtractTextInRange:
    """Tests for extract_text_in_range function."""

    def test_extracts_text_from_single_paragraph(self):
        """Test extracting text from a single paragraph."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Hello world", 1, "NORMAL_TEXT"),
            ]
        )

        text = extract_text_in_range(doc, 1, 13)
        assert "Hello world" in text

    def test_extracts_text_from_range(self):
        """Test extracting text from a specific range."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First paragraph", 1, "NORMAL_TEXT"),
                create_mock_paragraph("Second paragraph", 18, "NORMAL_TEXT"),
            ]
        )

        # Extract just the second paragraph
        text = extract_text_in_range(doc, 18, 36)
        assert "Second" in text
        assert "First" not in text

    def test_returns_empty_for_out_of_range(self):
        """Test returns empty string for indices outside document."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Only content", 1, "NORMAL_TEXT"),
            ]
        )

        text = extract_text_in_range(doc, 100, 200)
        assert text == ""

    def test_extracts_partial_paragraph(self):
        """Test extracting partial text when range partially overlaps."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Hello world", 1, "NORMAL_TEXT"),
            ]
        )

        # Try to get just the middle portion
        text = extract_text_in_range(doc, 1, 6)
        assert text == "Hello"


class TestGetAllHeadings:
    """Tests for get_all_headings function."""

    def test_returns_all_headings(self):
        """Test that all headings are returned."""
        doc = create_mock_document(
            [
                create_mock_paragraph("H1 Title", 1, "HEADING_1"),
                create_mock_paragraph("Normal text", 11, "NORMAL_TEXT"),
                create_mock_paragraph("H2 Subtitle", 24, "HEADING_2"),
                create_mock_paragraph("More text", 37, "NORMAL_TEXT"),
            ]
        )

        headings = get_all_headings(doc)

        assert len(headings) == 2
        assert headings[0]["text"] == "H1 Title"
        assert headings[0]["level"] == 1
        assert headings[1]["text"] == "H2 Subtitle"
        assert headings[1]["level"] == 2

    def test_returns_empty_for_no_headings(self):
        """Test returns empty list when no headings."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Just text", 1, "NORMAL_TEXT"),
            ]
        )

        headings = get_all_headings(doc)
        assert headings == []


class TestFindSectionInsertionPoint:
    """Tests for find_section_insertion_point function."""

    def test_finds_start_position(self):
        """Test finding insertion point at section start."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Content here", 14, "NORMAL_TEXT"),
            ]
        )

        index = find_section_insertion_point(doc, "Introduction", "start")

        # Should be right after the heading text
        assert index is not None
        assert index == 14  # End of heading element

    def test_finds_end_position(self):
        """Test finding insertion point at section end."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Content here", 14, "NORMAL_TEXT"),
                create_mock_paragraph("Next Section", 28, "HEADING_1"),
            ]
        )

        index = find_section_insertion_point(doc, "Introduction", "end")

        # Should be at start of next same-level heading
        assert index is not None
        assert index == 28

    def test_returns_none_for_missing_heading(self):
        """Test returns None when heading not found."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
            ]
        )

        index = find_section_insertion_point(doc, "NonExistent", "start")
        assert index is None

    def test_finds_start_position_case_insensitive(self):
        """Test finding insertion point with case-insensitive search.

        Bug: google_workspace_mcp-e0fc
        When using heading parameter with different case, the search should
        succeed when match_case=False.
        """
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Content here", 14, "NORMAL_TEXT"),
            ]
        )

        # Search with different case, case-insensitive
        index = find_section_insertion_point(
            doc, "INTRODUCTION", "start", match_case=False
        )

        assert index is not None
        assert index == 14  # End of heading element

    def test_finds_end_position_case_insensitive(self):
        """Test finding end insertion point with case-insensitive search."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Content here", 14, "NORMAL_TEXT"),
                create_mock_paragraph("Next Section", 28, "HEADING_1"),
            ]
        )

        # Search with different case, case-insensitive
        index = find_section_insertion_point(
            doc, "introduction", "end", match_case=False
        )

        assert index is not None
        assert index == 28

    def test_case_sensitive_fails_when_case_differs(self):
        """Test case-sensitive search fails with different case."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Content here", 14, "NORMAL_TEXT"),
            ]
        )

        # Search with different case, case-sensitive
        index = find_section_insertion_point(
            doc, "INTRODUCTION", "start", match_case=True
        )

        assert index is None


class TestHeadingTypes:
    """Tests for HEADING_TYPES constant."""

    def test_heading_types_defined(self):
        """Test that all heading types are defined."""
        assert "HEADING_1" in HEADING_TYPES
        assert "HEADING_2" in HEADING_TYPES
        assert "HEADING_3" in HEADING_TYPES
        assert "HEADING_4" in HEADING_TYPES
        assert "HEADING_5" in HEADING_TYPES
        assert "HEADING_6" in HEADING_TYPES
        assert "TITLE" in HEADING_TYPES

    def test_heading_levels_correct(self):
        """Test that heading levels are correct."""
        assert HEADING_TYPES["HEADING_1"] == 1
        assert HEADING_TYPES["HEADING_6"] == 6
        assert HEADING_TYPES["TITLE"] == 0  # Title is level 0


class TestFindElementsByType:
    """Tests for find_elements_by_type function."""

    def test_finds_tables(self):
        """Test finding all tables in a document."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Before table", 1, "NORMAL_TEXT"),
                create_mock_table(20, rows=3, columns=4),
                create_mock_paragraph("Between tables", 130, "NORMAL_TEXT"),
                create_mock_table(150, rows=2, columns=2),
            ]
        )

        elements = find_elements_by_type(doc, "table")

        assert len(elements) == 2
        assert elements[0]["type"] == "table"
        assert elements[0]["rows"] == 3
        assert elements[0]["columns"] == 4
        assert elements[1]["rows"] == 2

    def test_finds_all_headings(self):
        """Test finding all headings with alias 'heading'."""
        doc = create_mock_document(
            [
                create_mock_paragraph("H1 Title", 1, "HEADING_1"),
                create_mock_paragraph("Normal text", 11, "NORMAL_TEXT"),
                create_mock_paragraph("H2 Subtitle", 24, "HEADING_2"),
                create_mock_paragraph("H3 Detail", 37, "HEADING_3"),
            ]
        )

        elements = find_elements_by_type(doc, "heading")

        assert len(elements) == 3
        assert elements[0]["level"] == 1
        assert elements[1]["level"] == 2
        assert elements[2]["level"] == 3

    def test_finds_specific_heading_level(self):
        """Test finding headings of a specific level."""
        doc = create_mock_document(
            [
                create_mock_paragraph("H1 First", 1, "HEADING_1"),
                create_mock_paragraph("H2 Sub1", 11, "HEADING_2"),
                create_mock_paragraph("H1 Second", 21, "HEADING_1"),
                create_mock_paragraph("H2 Sub2", 32, "HEADING_2"),
            ]
        )

        elements = find_elements_by_type(doc, "heading2")

        assert len(elements) == 2
        assert all(e["level"] == 2 for e in elements)
        assert elements[0]["text"] == "H2 Sub1"
        assert elements[1]["text"] == "H2 Sub2"

    def test_finds_paragraphs(self):
        """Test finding paragraphs (non-heading, non-list)."""
        doc = create_mock_document(
            [
                create_mock_paragraph("H1 Title", 1, "HEADING_1"),
                create_mock_paragraph("Para 1", 11, "NORMAL_TEXT"),
                create_mock_paragraph("Para 2", 19, "NORMAL_TEXT"),
            ]
        )

        elements = find_elements_by_type(doc, "paragraph")

        assert len(elements) == 2
        assert elements[0]["text"] == "Para 1"
        assert elements[1]["text"] == "Para 2"

    def test_returns_empty_when_type_not_found(self):
        """Test returns empty list when element type not found."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Just text", 1, "NORMAL_TEXT"),
            ]
        )

        elements = find_elements_by_type(doc, "table")
        assert elements == []

    def test_case_insensitive_type_matching(self):
        """Test that element type matching is case insensitive."""
        doc = create_mock_document(
            [
                create_mock_table(1, rows=2, columns=2),
            ]
        )

        elements1 = find_elements_by_type(doc, "TABLE")
        elements2 = find_elements_by_type(doc, "Table")
        elements3 = find_elements_by_type(doc, "table")

        assert len(elements1) == 1
        assert len(elements2) == 1
        assert len(elements3) == 1

    def test_finds_bullet_lists(self):
        """Test finding bullet lists in a document."""
        list_id = "kix.abc123"
        doc = create_mock_document(
            [
                create_mock_paragraph("Before list", 1, "NORMAL_TEXT"),
                create_mock_list_item("Item 1", 14, list_id),
                create_mock_list_item("Item 2", 22, list_id),
                create_mock_paragraph("After list", 30, "NORMAL_TEXT"),
            ],
            lists={
                list_id: {
                    "listProperties": {
                        "nestingLevels": [{}]  # No glyphType = bullet list
                    }
                }
            },
        )

        elements = find_elements_by_type(doc, "bullet_list")

        assert len(elements) == 1
        assert elements[0]["type"] == "bullet_list"
        assert len(elements[0]["items"]) == 2
        assert elements[0]["items"][0]["text"] == "Item 1"
        assert elements[0]["items"][1]["text"] == "Item 2"

    def test_finds_numbered_lists(self):
        """Test finding numbered lists in a document."""
        list_id = "kix.def456"
        doc = create_mock_document(
            [
                create_mock_paragraph("Before list", 1, "NORMAL_TEXT"),
                create_mock_list_item("First", 14, list_id),
                create_mock_list_item("Second", 21, list_id),
                create_mock_list_item("Third", 29, list_id),
            ],
            lists={
                list_id: {
                    "listProperties": {
                        "nestingLevels": [{"glyphType": "DECIMAL"}]  # numbered list
                    }
                }
            },
        )

        elements = find_elements_by_type(doc, "numbered_list")

        assert len(elements) == 1
        assert elements[0]["type"] == "numbered_list"
        assert len(elements[0]["items"]) == 3

    def test_finds_all_lists_with_alias(self):
        """Test finding all lists (both bullet and numbered) with 'list' alias."""
        bullet_list_id = "kix.bullets"
        numbered_list_id = "kix.numbers"
        doc = create_mock_document(
            [
                create_mock_list_item("Bullet 1", 1, bullet_list_id),
                create_mock_list_item("Bullet 2", 11, bullet_list_id),
                create_mock_paragraph("Between lists", 21, "NORMAL_TEXT"),
                create_mock_list_item("Number 1", 36, numbered_list_id),
                create_mock_list_item("Number 2", 46, numbered_list_id),
            ],
            lists={
                bullet_list_id: {
                    "listProperties": {
                        "nestingLevels": [{}]  # bullet list
                    }
                },
                numbered_list_id: {
                    "listProperties": {
                        "nestingLevels": [{"glyphType": "DECIMAL"}]  # numbered list
                    }
                },
            },
        )

        elements = find_elements_by_type(doc, "list")

        assert len(elements) == 2
        types = {e["type"] for e in elements}
        assert "bullet_list" in types
        assert "numbered_list" in types

    def test_plural_aliases_work(self):
        """Test that plural forms of element types work (fixes bug cb80)."""
        bullet_list_id = "kix.bullets"
        doc = create_mock_document(
            [
                create_mock_paragraph("H1 Title", 1, "HEADING_1"),
                create_mock_paragraph("Normal text", 11, "NORMAL_TEXT"),
                create_mock_table(24, rows=2, columns=3),
                create_mock_list_item("Item 1", 130, bullet_list_id),
            ],
            lists={
                bullet_list_id: {
                    "listProperties": {
                        "nestingLevels": [{}]  # bullet list
                    }
                }
            },
        )

        # Test 'tables' works like 'table'
        tables_result = find_elements_by_type(doc, "tables")
        table_result = find_elements_by_type(doc, "table")
        assert len(tables_result) == len(table_result) == 1
        assert tables_result[0]["type"] == "table"

        # Test 'headings' works like 'heading'
        headings_result = find_elements_by_type(doc, "headings")
        heading_result = find_elements_by_type(doc, "heading")
        assert len(headings_result) == len(heading_result) == 1

        # Test 'paragraphs' works like 'paragraph'
        paragraphs_result = find_elements_by_type(doc, "paragraphs")
        paragraph_result = find_elements_by_type(doc, "paragraph")
        assert len(paragraphs_result) == len(paragraph_result) == 1

        # Test 'lists' works like 'list'
        lists_result = find_elements_by_type(doc, "lists")
        list_result = find_elements_by_type(doc, "list")
        assert len(lists_result) == len(list_result) == 1

    def test_toc_alias_works(self):
        """Test that 'toc' alias works for 'table_of_contents'."""
        # Create a simple document (no actual TOC, just testing the alias mechanism)
        doc = create_mock_document(
            [
                create_mock_paragraph("Content", 1, "NORMAL_TEXT"),
            ]
        )

        # Both should return empty (no TOC in doc), but neither should error
        toc_result = find_elements_by_type(doc, "toc")
        full_result = find_elements_by_type(doc, "table_of_contents")
        assert toc_result == full_result == []


class TestGetElementAncestors:
    """Tests for get_element_ancestors function."""

    def test_finds_single_ancestor(self):
        """Test finding ancestor when under one heading."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Introduction", 1, "HEADING_1"),
                create_mock_paragraph("Content here", 14, "NORMAL_TEXT"),
            ]
        )

        ancestors = get_element_ancestors(doc, 20)

        assert len(ancestors) == 1
        assert ancestors[0]["text"] == "Introduction"
        assert ancestors[0]["level"] == 1

    def test_finds_nested_ancestors(self):
        """Test finding multiple ancestors with nested headings."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Chapter", 1, "HEADING_1"),
                create_mock_paragraph("Section", 10, "HEADING_2"),
                create_mock_paragraph("Subsection", 20, "HEADING_3"),
                create_mock_paragraph("Content paragraph here", 32, "NORMAL_TEXT"),
            ]
        )

        # Query a position within the nested content (35 is within Content paragraph)
        ancestors = get_element_ancestors(doc, 35)

        assert len(ancestors) == 3
        assert ancestors[0]["text"] == "Chapter"
        assert ancestors[0]["level"] == 1
        assert ancestors[1]["text"] == "Section"
        assert ancestors[1]["level"] == 2
        assert ancestors[2]["text"] == "Subsection"
        assert ancestors[2]["level"] == 3

    def test_returns_empty_before_first_heading(self):
        """Test returns empty when index is before any heading."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Normal text first", 1, "NORMAL_TEXT"),
                create_mock_paragraph("Heading later", 20, "HEADING_1"),
            ]
        )

        ancestors = get_element_ancestors(doc, 5)
        assert ancestors == []

    def test_correctly_handles_section_boundaries(self):
        """Test that section boundaries are correctly identified."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First H1", 1, "HEADING_1"),
                create_mock_paragraph("Content", 11, "NORMAL_TEXT"),
                create_mock_paragraph("Second H1", 20, "HEADING_1"),
                create_mock_paragraph("More content", 31, "NORMAL_TEXT"),
            ]
        )

        # Position 15 is in first section
        ancestors1 = get_element_ancestors(doc, 15)
        assert len(ancestors1) == 1
        assert ancestors1[0]["text"] == "First H1"

        # Position 35 is in second section
        ancestors2 = get_element_ancestors(doc, 35)
        assert len(ancestors2) == 1
        assert ancestors2[0]["text"] == "Second H1"


class TestGetHeadingSiblings:
    """Tests for get_heading_siblings function."""

    def test_finds_siblings(self):
        """Test finding previous and next siblings."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Section 1", 1, "HEADING_2"),
                create_mock_paragraph("Content 1", 12, "NORMAL_TEXT"),
                create_mock_paragraph("Section 2", 24, "HEADING_2"),
                create_mock_paragraph("Content 2", 35, "NORMAL_TEXT"),
                create_mock_paragraph("Section 3", 47, "HEADING_2"),
            ]
        )

        result = get_heading_siblings(doc, "Section 2")

        assert result["found"] is True
        assert result["heading"]["text"] == "Section 2"
        assert result["level"] == 2
        assert result["previous"]["text"] == "Section 1"
        assert result["next"]["text"] == "Section 3"
        assert result["siblings_count"] == 3
        assert result["position_in_siblings"] == 2

    def test_first_sibling_has_no_previous(self):
        """Test that first sibling has no previous."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First", 1, "HEADING_2"),
                create_mock_paragraph("Second", 8, "HEADING_2"),
            ]
        )

        result = get_heading_siblings(doc, "First")

        assert result["found"] is True
        assert result["previous"] is None
        assert result["next"]["text"] == "Second"
        assert result["position_in_siblings"] == 1

    def test_last_sibling_has_no_next(self):
        """Test that last sibling has no next."""
        doc = create_mock_document(
            [
                create_mock_paragraph("First", 1, "HEADING_2"),
                create_mock_paragraph("Last", 8, "HEADING_2"),
            ]
        )

        result = get_heading_siblings(doc, "Last")

        assert result["found"] is True
        assert result["previous"]["text"] == "First"
        assert result["next"] is None
        assert result["position_in_siblings"] == 2

    def test_returns_not_found_for_missing_heading(self):
        """Test returns found=False when heading not found."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Existing", 1, "HEADING_2"),
            ]
        )

        result = get_heading_siblings(doc, "NonExistent")

        assert result["found"] is False
        assert "heading" not in result

    def test_case_insensitive_search(self):
        """Test case-insensitive heading search."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Section One", 1, "HEADING_2"),
                create_mock_paragraph("Section Two", 14, "HEADING_2"),
            ]
        )

        result = get_heading_siblings(doc, "SECTION ONE", match_case=False)

        assert result["found"] is True
        assert result["heading"]["text"] == "Section One"

    def test_case_sensitive_fails_when_case_differs(self):
        """Test case-sensitive search fails with different case."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Section One", 1, "HEADING_2"),
            ]
        )

        result = get_heading_siblings(doc, "SECTION ONE", match_case=True)

        assert result["found"] is False

    def test_only_finds_siblings_at_same_level(self):
        """Test that siblings are only at the same heading level."""
        doc = create_mock_document(
            [
                create_mock_paragraph("Chapter", 1, "HEADING_1"),
                create_mock_paragraph("Section 1", 10, "HEADING_2"),
                create_mock_paragraph("Section 2", 22, "HEADING_2"),
            ]
        )

        result = get_heading_siblings(doc, "Section 1")

        assert result["found"] is True
        assert result["siblings_count"] == 2  # Only H2s
        assert result["previous"] is None  # No previous H2
        assert result["next"]["text"] == "Section 2"


class TestMultiTabSupport:
    """Tests for multi-tab document support in parse_document_structure."""

    def test_get_body_for_tab_with_single_tab_document(self):
        """Test get_body_for_tab returns first tab's body for documents with tabs."""
        from gdocs.docs_structure import get_body_for_tab

        doc_data = {
            "title": "Test Document",
            "tabs": [
                {
                    "tabProperties": {"tabId": "t.0"},
                    "documentTab": {
                        "body": {
                            "content": [
                                {"startIndex": 0, "endIndex": 1, "sectionBreak": {}},
                                {
                                    "startIndex": 1,
                                    "endIndex": 50,
                                    "paragraph": {"elements": []},
                                },
                            ]
                        }
                    },
                }
            ],
        }

        body = get_body_for_tab(doc_data)
        assert body is not None
        assert "content" in body
        assert len(body["content"]) == 2

    def test_get_body_for_tab_with_specific_tab_id(self):
        """Test get_body_for_tab returns correct tab's body when tab_id is specified."""
        from gdocs.docs_structure import get_body_for_tab

        doc_data = {
            "title": "Multi-Tab Document",
            "tabs": [
                {
                    "tabProperties": {"tabId": "t.0"},
                    "documentTab": {
                        "body": {
                            "content": [
                                {"startIndex": 0, "endIndex": 1, "sectionBreak": {}},
                                {
                                    "startIndex": 1,
                                    "endIndex": 25000,
                                    "paragraph": {"elements": []},
                                },
                            ]
                        }
                    },
                },
                {
                    "tabProperties": {"tabId": "t.second"},
                    "documentTab": {
                        "body": {
                            "content": [
                                {"startIndex": 0, "endIndex": 1, "sectionBreak": {}},
                                {
                                    "startIndex": 1,
                                    "endIndex": 7000,
                                    "paragraph": {"elements": []},
                                },
                            ]
                        }
                    },
                },
            ],
        }

        # Request second tab specifically
        body = get_body_for_tab(doc_data, "t.second")
        assert body is not None
        assert "content" in body
        # Second tab has content ending at 7000, not 25000
        assert body["content"][1]["endIndex"] == 7000

    def test_get_body_for_tab_with_nested_child_tabs(self):
        """Test get_body_for_tab finds tabs in nested childTabs structure."""
        from gdocs.docs_structure import get_body_for_tab

        doc_data = {
            "title": "Nested Tabs Document",
            "tabs": [
                {
                    "tabProperties": {"tabId": "t.parent"},
                    "documentTab": {
                        "body": {"content": [{"startIndex": 0, "endIndex": 100}]}
                    },
                    "childTabs": [
                        {
                            "tabProperties": {"tabId": "t.child"},
                            "documentTab": {
                                "body": {
                                    "content": [{"startIndex": 0, "endIndex": 200}]
                                }
                            },
                        }
                    ],
                }
            ],
        }

        # Request nested child tab
        body = get_body_for_tab(doc_data, "t.child")
        assert body is not None
        assert body["content"][0]["endIndex"] == 200

    def test_get_body_for_tab_fallback_to_root_body(self):
        """Test get_body_for_tab falls back to root body when no tabs array."""
        from gdocs.docs_structure import get_body_for_tab

        # Legacy document format without tabs array
        doc_data = {
            "title": "Legacy Document",
            "body": {"content": [{"startIndex": 0, "endIndex": 500}]},
        }

        body = get_body_for_tab(doc_data)
        assert body is not None
        assert body["content"][0]["endIndex"] == 500

    def test_parse_document_structure_with_tab_id(self):
        """Test parse_document_structure correctly parses specific tab."""
        from gdocs.docs_structure import parse_document_structure

        doc_data = {
            "title": "Multi-Tab Document",
            "tabs": [
                {
                    "tabProperties": {"tabId": "t.0"},
                    "documentTab": {
                        "body": {
                            "content": [
                                {"startIndex": 0, "endIndex": 1, "sectionBreak": {}},
                                {
                                    "startIndex": 1,
                                    "endIndex": 25000,
                                    "paragraph": {
                                        "elements": [
                                            {"textRun": {"content": "Tab 1 content"}}
                                        ],
                                        "paragraphStyle": {},
                                    },
                                },
                            ]
                        }
                    },
                },
                {
                    "tabProperties": {"tabId": "t.second"},
                    "documentTab": {
                        "body": {
                            "content": [
                                {"startIndex": 0, "endIndex": 1, "sectionBreak": {}},
                                {
                                    "startIndex": 1,
                                    "endIndex": 7768,
                                    "paragraph": {
                                        "elements": [
                                            {"textRun": {"content": "Tab 2 content"}}
                                        ],
                                        "paragraphStyle": {},
                                    },
                                },
                            ]
                        }
                    },
                },
            ],
        }

        # Parse first tab (default)
        structure1 = parse_document_structure(doc_data)
        assert structure1["total_length"] == 25000

        # Parse second tab specifically
        structure2 = parse_document_structure(doc_data, "t.second")
        assert structure2["total_length"] == 7768

    def test_parse_document_structure_with_invalid_tab_id_falls_back(self):
        """Test parse_document_structure falls back to default when tab_id not found."""
        from gdocs.docs_structure import parse_document_structure

        doc_data = {
            "title": "Test Document",
            "tabs": [
                {
                    "tabProperties": {"tabId": "t.0"},
                    "documentTab": {
                        "body": {
                            "content": [
                                {"startIndex": 0, "endIndex": 1, "sectionBreak": {}},
                                {
                                    "startIndex": 1,
                                    "endIndex": 100,
                                    "paragraph": {"elements": [], "paragraphStyle": {}},
                                },
                            ]
                        }
                    },
                }
            ],
        }

        # Request non-existent tab - should fall back to default
        structure = parse_document_structure(doc_data, "t.nonexistent")
        assert structure["total_length"] == 100
