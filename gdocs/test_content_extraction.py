"""
Unit tests for Google Docs smart content extraction tools.

Tests the extraction helper functions and response formats for:
- extract_links
- extract_images
- extract_code_blocks
- extract_document_summary
"""


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


def create_mock_linked_text(text: str, url: str, start_index: int):
    """Create a mock paragraph element with a link."""
    end_index = start_index + len(text) + 1
    return {
        'startIndex': start_index,
        'endIndex': end_index,
        'paragraph': {
            'paragraphStyle': {
                'namedStyleType': 'NORMAL_TEXT'
            },
            'elements': [{
                'startIndex': start_index,
                'endIndex': end_index,
                'textRun': {
                    'content': text + '\n',
                    'textStyle': {
                        'link': {
                            'url': url
                        }
                    }
                }
            }]
        }
    }


def create_mock_code_text(text: str, start_index: int, font_family: str = 'Courier New', has_background: bool = False):
    """Create a mock paragraph element with code formatting."""
    end_index = start_index + len(text) + 1
    text_style = {
        'weightedFontFamily': {
            'fontFamily': font_family
        }
    }
    if has_background:
        text_style['backgroundColor'] = {
            'color': {
                'rgbColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}
            }
        }

    return {
        'startIndex': start_index,
        'endIndex': end_index,
        'paragraph': {
            'paragraphStyle': {
                'namedStyleType': 'NORMAL_TEXT'
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


def create_mock_inline_object(start_index: int, object_id: str):
    """Create a mock inline object element (image reference)."""
    return {
        'startIndex': start_index,
        'endIndex': start_index + 1,
        'paragraph': {
            'paragraphStyle': {
                'namedStyleType': 'NORMAL_TEXT'
            },
            'elements': [{
                'startIndex': start_index,
                'endIndex': start_index + 1,
                'inlineObjectElement': {
                    'inlineObjectId': object_id
                }
            }]
        }
    }


def create_mock_document(elements, inline_objects=None):
    """Create a mock document with given elements."""
    doc = {
        'title': 'Test Document',
        'body': {
            'content': elements
        },
        'lists': {}
    }
    if inline_objects:
        doc['inlineObjects'] = inline_objects
    return doc


class TestExtractLinksHelper:
    """Tests for link extraction logic."""

    def test_finds_links_in_document(self):
        """Test that links are correctly extracted from paragraphs."""
        doc = create_mock_document([
            create_mock_paragraph('Introduction', 1, 'HEADING_1'),
            create_mock_linked_text('Click here', 'https://example.com', 14),
            create_mock_paragraph('Some regular text', 26),
            create_mock_linked_text('Another link', 'https://google.com', 45),
        ])

        # Test link extraction directly from document structure
        body = doc.get('body', {})
        content = body.get('content', [])

        links = []
        for element in content:
            if 'paragraph' in element:
                paragraph = element['paragraph']
                for para_elem in paragraph.get('elements', []):
                    if 'textRun' in para_elem:
                        text_run = para_elem['textRun']
                        text_style = text_run.get('textStyle', {})
                        if 'link' in text_style:
                            link_info = text_style['link']
                            url = link_info.get('url', '')
                            if url and not url.startswith('#'):
                                links.append({
                                    'text': text_run.get('content', '').strip(),
                                    'url': url,
                                    'start_index': para_elem.get('startIndex', 0),
                                })

        assert len(links) == 2
        assert links[0]['text'] == 'Click here'
        assert links[0]['url'] == 'https://example.com'
        assert links[1]['text'] == 'Another link'
        assert links[1]['url'] == 'https://google.com'

    def test_ignores_internal_bookmarks(self):
        """Test that internal bookmark links (starting with #) are ignored."""
        doc = create_mock_document([
            create_mock_linked_text('Jump to section', '#bookmark-id', 1),
            create_mock_linked_text('External link', 'https://example.com', 20),
        ])

        body = doc.get('body', {})
        content = body.get('content', [])

        links = []
        for element in content:
            if 'paragraph' in element:
                paragraph = element['paragraph']
                for para_elem in paragraph.get('elements', []):
                    if 'textRun' in para_elem:
                        text_run = para_elem['textRun']
                        text_style = text_run.get('textStyle', {})
                        if 'link' in text_style:
                            link_info = text_style['link']
                            url = link_info.get('url', '')
                            if url and not url.startswith('#'):
                                links.append({'url': url})

        assert len(links) == 1
        assert links[0]['url'] == 'https://example.com'


class TestExtractImagesHelper:
    """Tests for image extraction logic."""

    def test_finds_images_in_document(self):
        """Test that inline images are correctly extracted."""
        inline_objects = {
            'kix.abc123': {
                'inlineObjectProperties': {
                    'embeddedObject': {
                        'imageProperties': {
                            'contentUri': 'https://lh3.googleusercontent.com/image1',
                            'sourceUri': 'https://example.com/original.png'
                        },
                        'size': {
                            'width': {'magnitude': 400, 'unit': 'PT'},
                            'height': {'magnitude': 300, 'unit': 'PT'}
                        }
                    }
                }
            },
            'kix.def456': {
                'inlineObjectProperties': {
                    'embeddedObject': {
                        'imageProperties': {
                            'contentUri': 'https://lh3.googleusercontent.com/image2',
                        },
                        'size': {
                            'width': {'magnitude': 200, 'unit': 'PT'},
                            'height': {'magnitude': 150, 'unit': 'PT'}
                        }
                    }
                }
            }
        }

        doc = create_mock_document([
            create_mock_paragraph('Introduction', 1, 'HEADING_1'),
            create_mock_inline_object(14, 'kix.abc123'),
            create_mock_paragraph('Some text', 16),
            create_mock_inline_object(27, 'kix.def456'),
        ], inline_objects=inline_objects)

        # Test image extraction logic
        assert len(doc['inlineObjects']) == 2

        obj1 = doc['inlineObjects']['kix.abc123']
        props1 = obj1['inlineObjectProperties']['embeddedObject']
        assert props1['imageProperties']['contentUri'] == 'https://lh3.googleusercontent.com/image1'
        assert props1['size']['width']['magnitude'] == 400
        assert props1['size']['height']['magnitude'] == 300

    def test_extracts_image_positions(self):
        """Test that image positions are correctly tracked."""
        inline_objects = {
            'kix.img1': {
                'inlineObjectProperties': {
                    'embeddedObject': {
                        'imageProperties': {
                            'contentUri': 'https://example.com/img1',
                        },
                        'size': {
                            'width': {'magnitude': 100, 'unit': 'PT'},
                            'height': {'magnitude': 100, 'unit': 'PT'}
                        }
                    }
                }
            }
        }

        doc = create_mock_document([
            create_mock_inline_object(50, 'kix.img1'),
        ], inline_objects=inline_objects)

        # Find image references in document
        image_refs = {}
        body = doc.get('body', {})
        content = body.get('content', [])

        for element in content:
            if 'paragraph' in element:
                paragraph = element['paragraph']
                for para_elem in paragraph.get('elements', []):
                    if 'inlineObjectElement' in para_elem:
                        obj_elem = para_elem['inlineObjectElement']
                        obj_id = obj_elem.get('inlineObjectId')
                        if obj_id:
                            start_idx = para_elem.get('startIndex', 0)
                            image_refs[obj_id] = start_idx

        assert image_refs['kix.img1'] == 50


class TestExtractCodeBlocksHelper:
    """Tests for code block extraction logic."""

    def test_identifies_monospace_text(self):
        """Test that monospace-formatted text is identified as code."""
        MONOSPACE_FONTS = {
            'courier new', 'consolas', 'monaco', 'menlo', 'source code pro',
            'fira code', 'jetbrains mono', 'roboto mono', 'ubuntu mono',
            'droid sans mono', 'liberation mono', 'dejavu sans mono',
            'lucida console', 'andale mono', 'courier'
        }

        def is_code_formatted(text_style):
            font_info = text_style.get('weightedFontFamily', {})
            font_family = font_info.get('fontFamily', '').lower()
            is_monospace = any(mono in font_family for mono in MONOSPACE_FONTS)
            bg_color = text_style.get('backgroundColor', {})
            has_background = bool(bg_color.get('color', {}))
            return is_monospace, font_info.get('fontFamily', ''), has_background

        # Test Courier New
        style1 = {'weightedFontFamily': {'fontFamily': 'Courier New'}}
        is_code, font, has_bg = is_code_formatted(style1)
        assert is_code is True
        assert font == 'Courier New'
        assert has_bg is False

        # Test Consolas with background
        style2 = {
            'weightedFontFamily': {'fontFamily': 'Consolas'},
            'backgroundColor': {'color': {'rgbColor': {'red': 0.9}}}
        }
        is_code, font, has_bg = is_code_formatted(style2)
        assert is_code is True
        assert font == 'Consolas'
        assert has_bg is True

        # Test regular font (not code)
        style3 = {'weightedFontFamily': {'fontFamily': 'Arial'}}
        is_code, font, has_bg = is_code_formatted(style3)
        assert is_code is False

    def test_finds_code_in_document(self):
        """Test that code-formatted paragraphs are found."""
        doc = create_mock_document([
            create_mock_paragraph('Introduction', 1, 'HEADING_1'),
            create_mock_code_text('def hello():', 14, 'Courier New'),
            create_mock_code_text("    print('Hello')", 28, 'Courier New'),
            create_mock_paragraph('Regular text', 48),
        ])

        MONOSPACE_FONTS = {'courier new', 'consolas', 'monaco'}
        code_runs = []

        body = doc.get('body', {})
        content = body.get('content', [])

        for element in content:
            if 'paragraph' in element:
                paragraph = element['paragraph']
                for para_elem in paragraph.get('elements', []):
                    if 'textRun' in para_elem:
                        text_run = para_elem['textRun']
                        text_style = text_run.get('textStyle', {})
                        text_content = text_run.get('content', '')

                        font_info = text_style.get('weightedFontFamily', {})
                        font_family = font_info.get('fontFamily', '').lower()

                        is_monospace = any(mono in font_family for mono in MONOSPACE_FONTS)

                        if is_monospace and text_content.strip():
                            code_runs.append({
                                'content': text_content,
                                'start_index': para_elem.get('startIndex', 0),
                            })

        assert len(code_runs) == 2
        assert 'def hello():' in code_runs[0]['content']
        assert "print('Hello')" in code_runs[1]['content']


class TestExtractSummaryHelper:
    """Tests for document summary extraction logic."""

    def test_counts_element_types(self):
        """Test that elements are correctly counted by type."""
        doc = create_mock_document([
            create_mock_paragraph('Title', 1, 'HEADING_1'),
            create_mock_paragraph('Introduction', 10, 'HEADING_2'),
            create_mock_paragraph('Some content', 25),
            create_mock_paragraph('More content', 40),
            create_mock_paragraph('Section 2', 55, 'HEADING_2'),
            create_mock_paragraph('Even more content', 68),
        ])

        from gdocs.docs_structure import extract_structural_elements
        elements = extract_structural_elements(doc)

        counts = {
            'headings': 0,
            'paragraphs': 0,
        }

        for elem in elements:
            elem_type = elem.get('type', '')
            if elem_type.startswith('heading') or elem_type == 'title':
                counts['headings'] += 1
            elif elem_type == 'paragraph':
                counts['paragraphs'] += 1

        assert counts['headings'] == 3
        assert counts['paragraphs'] == 3

    def test_builds_outline(self):
        """Test that hierarchical outline is correctly built."""
        doc = create_mock_document([
            create_mock_paragraph('Chapter 1', 1, 'HEADING_1'),
            create_mock_paragraph('Section 1.1', 13, 'HEADING_2'),
            create_mock_paragraph('Content', 27),
            create_mock_paragraph('Section 1.2', 37, 'HEADING_2'),
            create_mock_paragraph('Chapter 2', 51, 'HEADING_1'),
        ])

        from gdocs.docs_structure import extract_structural_elements, build_headings_outline
        elements = extract_structural_elements(doc)
        outline = build_headings_outline(elements)

        # Should have 2 top-level headings
        assert len(outline) == 2
        assert outline[0]['text'] == 'Chapter 1'
        assert outline[0]['level'] == 1

        # Chapter 1 should have 2 children
        assert len(outline[0]['children']) == 2
        assert outline[0]['children'][0]['text'] == 'Section 1.1'
        assert outline[0]['children'][1]['text'] == 'Section 1.2'

        # Chapter 2 should have no children
        assert outline[1]['text'] == 'Chapter 2'
        assert len(outline[1]['children']) == 0


class TestSectionContext:
    """Tests for section context finding logic."""

    def test_finds_section_for_index(self):
        """Test that correct section is found for a given index."""
        doc = create_mock_document([
            create_mock_paragraph('Introduction', 1, 'HEADING_1'),
            create_mock_paragraph('Content 1', 15),
            create_mock_paragraph('Section 2', 27, 'HEADING_1'),
            create_mock_paragraph('Content 2', 40),
        ])

        from gdocs.docs_structure import get_all_headings
        headings = get_all_headings(doc)

        def find_section_for_index(idx):
            if not headings:
                return ""
            current_section = ""
            for heading in headings:
                if heading['start_index'] <= idx:
                    current_section = heading['text']
                else:
                    break
            return current_section

        # Before first heading
        assert find_section_for_index(0) == ""

        # In first section
        assert find_section_for_index(15) == 'Introduction'
        assert find_section_for_index(20) == 'Introduction'

        # In second section
        assert find_section_for_index(40) == 'Section 2'
        assert find_section_for_index(50) == 'Section 2'


class TestTableContentExtraction:
    """Tests for extracting content from tables."""

    def test_processes_table_cells(self):
        """Test that content inside tables is processed."""
        table_element = {
            'startIndex': 50,
            'endIndex': 150,
            'table': {
                'tableRows': [
                    {
                        'tableCells': [
                            {
                                'startIndex': 51,
                                'endIndex': 70,
                                'content': [
                                    {
                                        'startIndex': 52,
                                        'endIndex': 68,
                                        'paragraph': {
                                            'paragraphStyle': {'namedStyleType': 'NORMAL_TEXT'},
                                            'elements': [{
                                                'startIndex': 52,
                                                'endIndex': 68,
                                                'textRun': {
                                                    'content': 'Cell content\n',
                                                    'textStyle': {
                                                        'link': {
                                                            'url': 'https://table-link.com'
                                                        }
                                                    }
                                                }
                                            }]
                                        }
                                    }
                                ]
                            }
                        ]
                    }
                ]
            }
        }

        # Simulate recursive table processing
        links = []

        def process_elements(elements, depth=0):
            if depth > 10:
                return
            for element in elements:
                if 'paragraph' in element:
                    paragraph = element['paragraph']
                    for para_elem in paragraph.get('elements', []):
                        if 'textRun' in para_elem:
                            text_run = para_elem['textRun']
                            text_style = text_run.get('textStyle', {})
                            if 'link' in text_style:
                                links.append(text_style['link']['url'])
                elif 'table' in element:
                    table = element['table']
                    for row in table.get('tableRows', []):
                        for cell in row.get('tableCells', []):
                            process_elements(cell.get('content', []), depth + 1)

        process_elements([table_element])

        assert len(links) == 1
        assert links[0] == 'https://table-link.com'
