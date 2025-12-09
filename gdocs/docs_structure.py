"""
Google Docs Document Structure Parsing and Analysis

This module provides utilities for parsing and analyzing the structure
of Google Docs documents, including finding tables, cells, headings, and other elements.
"""

import logging
from typing import Any, Optional

logger = logging.getLogger(__name__)


# Google Docs heading types mapped to levels
HEADING_TYPES = {
    "HEADING_1": 1,
    "HEADING_2": 2,
    "HEADING_3": 3,
    "HEADING_4": 4,
    "HEADING_5": 5,
    "HEADING_6": 6,
    "TITLE": 0,  # Document title style
}

# Element type names for structural navigation
ELEMENT_TYPE_NAMES = {
    "heading1": "HEADING_1",
    "heading2": "HEADING_2",
    "heading3": "HEADING_3",
    "heading4": "HEADING_4",
    "heading5": "HEADING_5",
    "heading6": "HEADING_6",
    "title": "TITLE",
    "paragraph": "NORMAL_TEXT",
    "table": "table",
    "bullet_list": "bullet_list",
    "numbered_list": "numbered_list",
}


def get_body_for_tab(doc_data: dict[str, Any], tab_id: str = None) -> dict[str, Any]:
    """
    Get the body content for a specific tab in a document.

    For multi-tab documents fetched with includeTabsContent=True, this extracts
    the body content for the specified tab. If tab_id is None, returns the default
    tab's body (either root body for legacy format or first tab's body).

    Args:
        doc_data: Raw document data from Google Docs API (fetched with includeTabsContent=True)
        tab_id: Optional tab ID to get body for. If None, uses default/first tab.

    Returns:
        Body dictionary with 'content' array, or empty dict if not found
    """

    def find_tab_recursive(tabs, target_id):
        """Recursively search tabs and child tabs for the target tab ID."""
        for tab in tabs:
            tab_props = tab.get("tabProperties", {})
            current_id = tab_props.get("tabId")
            if current_id == target_id:
                return tab.get("documentTab", {}).get("body", {})
            # Check child tabs
            child_tabs = tab.get("childTabs", [])
            if child_tabs:
                result = find_tab_recursive(child_tabs, target_id)
                if result:
                    return result
        return None

    tabs = doc_data.get("tabs", [])

    if tab_id is not None and tabs:
        # Look for specific tab
        body = find_tab_recursive(tabs, tab_id)
        if body is not None:
            return body
        # Tab not found, log warning and fall through to default
        logger.warning(f"Tab '{tab_id}' not found in document, using default body")

    # Return default body:
    # 1. If tabs exist, use first tab's documentTab.body
    # 2. Fall back to root body (legacy/single-tab format)
    if tabs:
        first_tab = tabs[0]
        return first_tab.get("documentTab", {}).get("body", {})

    return doc_data.get("body", {})


def parse_document_structure(
    doc_data: dict[str, Any], tab_id: str = None
) -> dict[str, Any]:
    """
    Parse the full document structure into a navigable format.

    Args:
        doc_data: Raw document data from Google Docs API (fetched with includeTabsContent=True
                  for multi-tab support)
        tab_id: Optional tab ID for multi-tab documents. If provided, parses that specific
                tab's content. If None, uses the default/first tab.

    Returns:
        Dictionary containing parsed structure with elements and their positions
    """
    structure = {
        "title": doc_data.get("title", ""),
        "body": [],
        "tables": [],
        "headers": {},
        "footers": {},
        "total_length": 0,
    }

    # Get body content for the specified tab (or default tab if not specified)
    body = get_body_for_tab(doc_data, tab_id)
    content = body.get("content", [])

    for element in content:
        element_info = _parse_element(element)
        if element_info:
            structure["body"].append(element_info)
            if element_info["type"] == "table":
                structure["tables"].append(element_info)

    # Calculate total document length
    if structure["body"]:
        last_element = structure["body"][-1]
        structure["total_length"] = last_element.get("end_index", 0)

    # Parse headers and footers
    for header_id, header_data in doc_data.get("headers", {}).items():
        structure["headers"][header_id] = _parse_segment(header_data)

    for footer_id, footer_data in doc_data.get("footers", {}).items():
        structure["footers"][footer_id] = _parse_segment(footer_data)

    return structure


def _parse_element(element: dict[str, Any]) -> Optional[dict[str, Any]]:
    """
    Parse a single document element.

    Args:
        element: Element data from document

    Returns:
        Parsed element information or None
    """
    element_info = {
        "start_index": element.get("startIndex", 0),
        "end_index": element.get("endIndex", 0),
    }

    if "paragraph" in element:
        paragraph = element["paragraph"]
        element_info["type"] = "paragraph"
        element_info["text"] = _extract_paragraph_text(paragraph)
        element_info["style"] = paragraph.get("paragraphStyle", {})

    elif "table" in element:
        table = element["table"]
        element_info["type"] = "table"
        element_info["rows"] = len(table.get("tableRows", []))
        element_info["columns"] = len(
            table.get("tableRows", [{}])[0].get("tableCells", [])
        )
        element_info["cells"] = _parse_table_cells(table)
        element_info["table_style"] = table.get("tableStyle", {})

    elif "sectionBreak" in element:
        element_info["type"] = "section_break"
        element_info["section_style"] = element["sectionBreak"].get("sectionStyle", {})

    elif "tableOfContents" in element:
        element_info["type"] = "table_of_contents"

    else:
        return None

    return element_info


def _parse_table_cells(table: dict[str, Any]) -> list[list[dict[str, Any]]]:
    """
    Parse table cells with their positions and content.

    Args:
        table: Table element data

    Returns:
        2D list of cell information
    """
    cells = []
    for row_idx, row in enumerate(table.get("tableRows", [])):
        row_cells = []
        for col_idx, cell in enumerate(row.get("tableCells", [])):
            # Find the first paragraph in the cell for insertion
            insertion_index = cell.get("startIndex", 0) + 1  # Default fallback

            # Look for the first paragraph in cell content
            content_elements = cell.get("content", [])
            for element in content_elements:
                if "paragraph" in element:
                    paragraph = element["paragraph"]
                    # Get the first element in the paragraph
                    para_elements = paragraph.get("elements", [])
                    if para_elements:
                        first_element = para_elements[0]
                        if "startIndex" in first_element:
                            insertion_index = first_element["startIndex"]
                            break

            # Extract cell span information from tableCellStyle
            cell_style = cell.get("tableCellStyle", {})
            row_span = cell_style.get("rowSpan", 1)
            column_span = cell_style.get("columnSpan", 1)

            cell_info = {
                "row": row_idx,
                "column": col_idx,
                "start_index": cell.get("startIndex", 0),
                "end_index": cell.get("endIndex", 0),
                "insertion_index": insertion_index,  # Where to insert text in this cell
                "content": _extract_cell_text(cell),
                "content_elements": content_elements,
                "row_span": row_span,
                "column_span": column_span,
            }
            row_cells.append(cell_info)
        cells.append(row_cells)
    return cells


def _extract_paragraph_text(paragraph: dict[str, Any]) -> str:
    """Extract text from a paragraph element."""
    text_parts = []
    for element in paragraph.get("elements", []):
        if "textRun" in element:
            text_parts.append(element["textRun"].get("content", ""))
    return "".join(text_parts)


def _extract_cell_text(cell: dict[str, Any]) -> str:
    """Extract text content from a table cell."""
    text_parts = []
    for element in cell.get("content", []):
        if "paragraph" in element:
            text_parts.append(_extract_paragraph_text(element["paragraph"]))
    return "".join(text_parts)


def extract_text_in_range(
    doc_data: dict[str, Any], start_index: int, end_index: int
) -> str:
    """
    Extract all text content from a document between given indices.

    This extracts text directly from the raw document body, not just
    from structural elements. This ensures we capture all text even
    if it wasn't identified as a structural element.

    Args:
        doc_data: Raw document data from Google Docs API
        start_index: Starting character position (inclusive)
        end_index: Ending character position (exclusive)

    Returns:
        All text content in the specified range
    """
    text_parts = []
    body = doc_data.get("body", {})
    content = body.get("content", [])

    for element in content:
        elem_start = element.get("startIndex", 0)
        elem_end = element.get("endIndex", 0)

        # Skip elements completely outside our range
        if elem_end <= start_index or elem_start >= end_index:
            continue

        if "paragraph" in element:
            paragraph = element["paragraph"]
            for para_elem in paragraph.get("elements", []):
                pe_start = para_elem.get("startIndex", 0)
                pe_end = para_elem.get("endIndex", 0)

                # Skip elements completely outside our range
                if pe_end <= start_index or pe_start >= end_index:
                    continue

                if "textRun" in para_elem:
                    content_text = para_elem["textRun"].get("content", "")

                    # Handle partial overlap
                    # Calculate which portion of this text is in our range
                    text_start = max(0, start_index - pe_start)
                    text_end = min(len(content_text), end_index - pe_start)

                    if text_start < text_end:
                        text_parts.append(content_text[text_start:text_end])

        elif "table" in element:
            # For tables, extract cell content
            table = element["table"]
            for row in table.get("tableRows", []):
                for cell in row.get("tableCells", []):
                    cell_start = cell.get("startIndex", 0)
                    cell_end = cell.get("endIndex", 0)

                    # Skip cells completely outside our range
                    if cell_end <= start_index or cell_start >= end_index:
                        continue

                    cell_text = _extract_cell_text(cell)
                    if cell_text:
                        text_parts.append(cell_text)

    return "".join(text_parts)


def _parse_segment(segment_data: dict[str, Any]) -> dict[str, Any]:
    """Parse a document segment (header/footer)."""
    return {
        "content": segment_data.get("content", []),
        "start_index": segment_data.get("content", [{}])[0].get("startIndex", 0)
        if segment_data.get("content")
        else 0,
        "end_index": segment_data.get("content", [{}])[-1].get("endIndex", 0)
        if segment_data.get("content")
        else 0,
    }


def find_tables(doc_data: dict[str, Any]) -> list[dict[str, Any]]:
    """
    Find all tables in the document with their positions and dimensions.

    Args:
        doc_data: Raw document data from Google Docs API

    Returns:
        List of table information dictionaries
    """
    tables = []
    structure = parse_document_structure(doc_data)

    for idx, table_info in enumerate(structure["tables"]):
        tables.append(
            {
                "index": idx,
                "start_index": table_info["start_index"],
                "end_index": table_info["end_index"],
                "rows": table_info["rows"],
                "columns": table_info["columns"],
                "cells": table_info["cells"],
            }
        )

    return tables


def get_table_cell_indices(
    doc_data: dict[str, Any], table_index: int = 0
) -> Optional[list[list[tuple[int, int]]]]:
    """
    Get content indices for all cells in a specific table.

    Args:
        doc_data: Raw document data from Google Docs API
        table_index: Index of the table (0-based)

    Returns:
        2D list of (start_index, end_index) tuples for each cell, or None if table not found
    """
    tables = find_tables(doc_data)

    if table_index >= len(tables):
        logger.warning(
            f"Table index {table_index} not found. Document has {len(tables)} tables."
        )
        return None

    table = tables[table_index]
    cell_indices = []

    for row in table["cells"]:
        row_indices = []
        for cell in row:
            # Each cell contains at least one paragraph
            # Find the first paragraph in the cell for content insertion
            cell_content = cell.get("content_elements", [])
            if cell_content:
                # Look for the first paragraph in cell content
                first_para = None
                for element in cell_content:
                    if "paragraph" in element:
                        first_para = element["paragraph"]
                        break

                if first_para and "elements" in first_para and first_para["elements"]:
                    # Insert at the start of the first text run in the paragraph
                    first_text_element = first_para["elements"][0]
                    if "textRun" in first_text_element:
                        start_idx = first_text_element.get(
                            "startIndex", cell["start_index"] + 1
                        )
                        end_idx = first_text_element.get("endIndex", start_idx + 1)
                        row_indices.append((start_idx, end_idx))
                        continue

            # Fallback: use cell boundaries with safe margins
            content_start = cell["start_index"] + 1
            content_end = cell["end_index"] - 1
            row_indices.append((content_start, content_end))
        cell_indices.append(row_indices)

    return cell_indices


def find_element_at_index(
    doc_data: dict[str, Any], index: int
) -> Optional[dict[str, Any]]:
    """
    Find what element exists at a given index in the document.

    Args:
        doc_data: Raw document data from Google Docs API
        index: Position in the document

    Returns:
        Information about the element at that position, or None
    """
    structure = parse_document_structure(doc_data)

    for element in structure["body"]:
        if element["start_index"] <= index < element["end_index"]:
            element_copy = element.copy()

            # If it's a table, find which cell contains the index
            if element["type"] == "table" and "cells" in element:
                for row_idx, row in enumerate(element["cells"]):
                    for col_idx, cell in enumerate(row):
                        if cell["start_index"] <= index < cell["end_index"]:
                            element_copy["containing_cell"] = {
                                "row": row_idx,
                                "column": col_idx,
                                "cell_start": cell["start_index"],
                                "cell_end": cell["end_index"],
                            }
                            break

            return element_copy

    return None


def get_next_paragraph_index(doc_data: dict[str, Any], after_index: int = 0) -> int:
    """
    Find the next safe position to insert content after a given index.

    Args:
        doc_data: Raw document data from Google Docs API
        after_index: Index after which to find insertion point

    Returns:
        Safe index for insertion
    """
    structure = parse_document_structure(doc_data)

    # Find the first paragraph element after the given index
    for element in structure["body"]:
        if element["type"] == "paragraph" and element["start_index"] > after_index:
            # Insert at the end of the previous element or start of this paragraph
            return element["start_index"]

    # If no paragraph found, return the end of document
    return structure["total_length"] - 1 if structure["total_length"] > 0 else 1


def analyze_document_complexity(doc_data: dict[str, Any]) -> dict[str, Any]:
    """
    Analyze document complexity and provide statistics.

    Args:
        doc_data: Raw document data from Google Docs API

    Returns:
        Dictionary with document statistics
    """
    structure = parse_document_structure(doc_data)

    stats = {
        "total_elements": len(structure["body"]),
        "tables": len(structure["tables"]),
        "paragraphs": sum(1 for e in structure["body"] if e.get("type") == "paragraph"),
        "section_breaks": sum(
            1 for e in structure["body"] if e.get("type") == "section_break"
        ),
        "total_length": structure["total_length"],
        "has_headers": bool(structure["headers"]),
        "has_footers": bool(structure["footers"]),
    }

    # Add table statistics
    if structure["tables"]:
        total_cells = sum(
            table["rows"] * table["columns"] for table in structure["tables"]
        )
        stats["total_table_cells"] = total_cells
        stats["largest_table"] = max(
            (t["rows"] * t["columns"] for t in structure["tables"]), default=0
        )

    return stats


def _get_paragraph_style_type(paragraph: dict[str, Any]) -> str:
    """
    Get the style type of a paragraph (heading level, normal text, etc.).

    Args:
        paragraph: Paragraph element from document

    Returns:
        Style type string (e.g., 'HEADING_1', 'NORMAL_TEXT')
    """
    style = paragraph.get("paragraphStyle", {})
    named_style = style.get("namedStyleType", "NORMAL_TEXT")
    return named_style


def _is_list_paragraph(paragraph: dict[str, Any]) -> Optional[str]:
    """
    Check if a paragraph is part of a list.

    Args:
        paragraph: Paragraph element from document

    Returns:
        'bullet_list' or 'numbered_list' if list, None otherwise
    """
    bullet = paragraph.get("bullet")
    if not bullet:
        return None

    # Check the nesting level - all list items have bullets
    list_id = bullet.get("listId")
    if list_id:
        # Google Docs uses different glyph types for different list styles
        # We determine type based on the list properties in the document
        return "bullet_list"  # Default; actual type requires checking document lists

    return None


def extract_structural_elements(doc_data: dict[str, Any]) -> list[dict[str, Any]]:
    """
    Extract all structural elements from a document with detailed type information.

    This function identifies headings, paragraphs, lists, tables, and other
    structural elements with their positions and content.

    Args:
        doc_data: Raw document data from Google Docs API

    Returns:
        List of structural element dictionaries with type, text, positions, and metadata
    """
    elements = []
    body = doc_data.get("body", {})
    content = body.get("content", [])
    lists_info = doc_data.get("lists", {})

    # Track list context for grouping list items
    current_list = None

    for element in content:
        start_idx = element.get("startIndex", 0)
        end_idx = element.get("endIndex", 0)

        if "paragraph" in element:
            para = element["paragraph"]
            text = _extract_paragraph_text(para).strip()
            style_type = _get_paragraph_style_type(para)

            # Check if this is a heading
            if style_type in HEADING_TYPES:
                level = HEADING_TYPES[style_type]
                elem_type = f"heading{level}" if level > 0 else "title"

                elements.append(
                    {
                        "type": elem_type,
                        "text": text,
                        "start_index": start_idx,
                        "end_index": end_idx,
                        "level": level,
                        "style_type": style_type,
                    }
                )
                current_list = None

            elif _is_list_paragraph(para):
                # This is a list item
                bullet = para.get("bullet", {})
                list_id = bullet.get("listId")
                nesting_level = bullet.get("nestingLevel", 0)

                # Determine list type from document lists
                list_type = "bullet_list"
                if list_id and list_id in lists_info:
                    list_props = lists_info[list_id].get("listProperties", {})
                    nesting_props = list_props.get("nestingLevels", [])
                    if nesting_props and nesting_level < len(nesting_props):
                        glyph_type = nesting_props[nesting_level].get("glyphType")
                        if glyph_type and glyph_type != "GLYPH_TYPE_UNSPECIFIED":
                            list_type = "numbered_list"

                # Group consecutive list items
                if current_list and current_list["list_id"] == list_id:
                    # Add to existing list
                    current_list["items"].append(
                        {
                            "text": text,
                            "start_index": start_idx,
                            "end_index": end_idx,
                            "nesting_level": nesting_level,
                        }
                    )
                    current_list["end_index"] = end_idx
                else:
                    # Start new list
                    if current_list:
                        elements.append(current_list)

                    current_list = {
                        "type": list_type,
                        "list_type": list_type.replace("_list", ""),
                        "start_index": start_idx,
                        "end_index": end_idx,
                        "list_id": list_id,
                        "items": [
                            {
                                "text": text,
                                "start_index": start_idx,
                                "end_index": end_idx,
                                "nesting_level": nesting_level,
                            }
                        ],
                    }
            else:
                # Regular paragraph
                if current_list:
                    elements.append(current_list)
                    current_list = None

                # Only add non-empty paragraphs
                if text:
                    elements.append(
                        {
                            "type": "paragraph",
                            "text": text,
                            "start_index": start_idx,
                            "end_index": end_idx,
                        }
                    )

        elif "table" in element:
            if current_list:
                elements.append(current_list)
                current_list = None

            table = element["table"]
            rows = len(table.get("tableRows", []))
            cols = (
                len(table.get("tableRows", [{}])[0].get("tableCells", []))
                if rows > 0
                else 0
            )

            elements.append(
                {
                    "type": "table",
                    "start_index": start_idx,
                    "end_index": end_idx,
                    "rows": rows,
                    "columns": cols,
                }
            )

        elif "sectionBreak" in element:
            if current_list:
                elements.append(current_list)
                current_list = None

        elif "tableOfContents" in element:
            if current_list:
                elements.append(current_list)
                current_list = None

            elements.append(
                {
                    "type": "table_of_contents",
                    "start_index": start_idx,
                    "end_index": end_idx,
                }
            )

    # Don't forget any trailing list
    if current_list:
        elements.append(current_list)

    return elements


def build_headings_outline(elements: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Build a hierarchical outline from document headings.

    Args:
        elements: List of structural elements from extract_structural_elements

    Returns:
        Hierarchical list of headings with nested children
    """
    headings = [
        e for e in elements if e["type"].startswith("heading") or e["type"] == "title"
    ]

    if not headings:
        return []

    # Build hierarchy
    outline = []
    stack = []  # Stack of (level, heading_dict) tuples

    for heading in headings:
        level = heading.get("level", 0)
        heading_item = {
            "level": level,
            "text": heading["text"],
            "start_index": heading["start_index"],
            "end_index": heading["end_index"],
            "children": [],
        }

        # Pop items from stack until we find a parent (lower level number)
        while stack and stack[-1][0] >= level:
            stack.pop()

        if stack:
            # Add as child of the top of stack
            stack[-1][1]["children"].append(heading_item)
        else:
            # Top-level heading
            outline.append(heading_item)

        stack.append((level, heading_item))

    return outline


def find_section_by_heading(
    doc_data: dict[str, Any], heading_text: str, match_case: bool = False
) -> Optional[dict[str, Any]]:
    """
    Find a section by its heading text.

    A section includes all content from the heading until the next heading
    of the same or higher level (smaller number = higher level).

    Args:
        doc_data: Raw document data from Google Docs API
        heading_text: Text of the heading to find
        match_case: Whether to match case exactly

    Returns:
        Dictionary with section info including start_index, end_index, content, and subsections
        Returns None if heading not found
    """
    elements = extract_structural_elements(doc_data)

    # Find the target heading
    target_heading = None
    target_idx = -1

    search_text = heading_text if match_case else heading_text.lower()

    for i, elem in enumerate(elements):
        if elem["type"].startswith("heading") or elem["type"] == "title":
            elem_text = elem["text"] if match_case else elem["text"].lower()
            if elem_text.strip() == search_text.strip():
                target_heading = elem
                target_idx = i
                break

    if target_heading is None:
        return None

    target_level = target_heading.get("level", 0)
    section_start = target_heading["start_index"]
    section_end = None

    # Find where this section ends
    # It ends when we hit a heading of same or higher level (smaller number)
    # NOTE: We skip "false" headings that appear to be content paragraphs incorrectly styled.
    # This can happen when heading styles bleed into subsequent paragraphs.
    # Heuristics for a "false" heading:
    # 1. Empty text (just whitespace)
    # 2. Immediately follows the target heading (within ~2 chars - just newlines)
    # 3. Very long text (>100 chars) that looks like body content
    #
    # IMPORTANT: We track which elements are "false" so we don't use them as evidence
    # for marking subsequent headings as false. A real heading following a false heading
    # should still be considered real.
    false_heading_indices = set()
    for i in range(target_idx + 1, len(elements)):
        elem = elements[i]
        if elem["type"].startswith("heading") or elem["type"] == "title":
            elem_level = elem.get("level", 0)
            if elem_level <= target_level:
                elem_text = elem.get("text", "").strip()
                elem_start = elem.get("start_index", 0)

                # Skip "false" headings that are likely style bleed artifacts
                # This happens when heading styles bleed into subsequent paragraphs due to
                # Google Docs API paragraph style inheritance.
                is_false_heading = False

                # Empty heading - definitely false (style bleed creates empty-text headings)
                if not elem_text:
                    is_false_heading = True
                # Very long text (>60 chars) is unlikely to be a real heading
                # Real headings are typically short (chapter titles, section names)
                elif len(elem_text) > 60:
                    is_false_heading = True
                # Immediately follows another heading (likely style bleed)
                # Check if there's no content between the previous element and this one
                # A real heading typically has at least some content before it
                elif i > 0:
                    prev_elem = elements[i - 1]
                    prev_end = prev_elem.get("end_index", 0)
                    # If this heading starts right after the previous element ends
                    # (allowing for just newlines/whitespace), it's likely style bleed
                    if elem_start - prev_end <= 2:
                        # Check if the previous element is also a heading
                        # BUT only if that heading wasn't already marked as false
                        prev_is_heading = (
                            prev_elem["type"].startswith("heading")
                            or prev_elem["type"] == "title"
                        )
                        prev_is_false = (i - 1) in false_heading_indices
                        if prev_is_heading and not prev_is_false:
                            is_false_heading = True

                if is_false_heading:
                    false_heading_indices.add(i)
                else:
                    section_end = elem["start_index"]
                    break

    # If no ending heading found, section goes to end of document
    if section_end is None:
        # Get document end
        body = doc_data.get("body", {})
        content = body.get("content", [])
        if content:
            section_end = content[-1].get("endIndex", section_start + 1)
        else:
            section_end = section_start + 1

    # Extract section content
    section_elements = []
    subsections = []

    for i in range(target_idx + 1, len(elements)):
        elem = elements[i]
        if elem["start_index"] >= section_end:
            break

        if elem["type"].startswith("heading") or elem["type"] == "title":
            # This is a subsection
            subsections.append(
                {
                    "heading": elem["text"],
                    "level": elem.get("level", 0),
                    "start_index": elem["start_index"],
                    "end_index": elem["end_index"],
                }
            )

        section_elements.append(elem)

    # Extract content directly from the raw document
    # Start after the heading text ends, up to section end
    content_start = target_heading["end_index"]
    raw_content = extract_text_in_range(doc_data, content_start, section_end)

    # Clean up the content (remove leading/trailing whitespace and newlines)
    content = raw_content.strip()

    return {
        "heading": target_heading["text"],
        "level": target_level,
        "start_index": section_start,
        "end_index": section_end,
        "content": content,
        "elements": section_elements,
        "subsections": subsections,
    }


def get_all_headings(doc_data: dict[str, Any]) -> list[dict[str, Any]]:
    """
    Get all headings in the document with their positions and levels.

    Args:
        doc_data: Raw document data from Google Docs API

    Returns:
        List of heading dictionaries with text, level, and position info
    """
    elements = extract_structural_elements(doc_data)
    return [
        {
            "text": e["text"],
            "level": e.get("level", 0),
            "type": e["type"],
            "start_index": e["start_index"],
            "end_index": e["end_index"],
        }
        for e in elements
        if e["type"].startswith("heading") or e["type"] == "title"
    ]


def find_section_insertion_point(
    doc_data: dict[str, Any],
    heading_text: str,
    position: str = "end",
    match_case: bool = False,
) -> Optional[int]:
    """
    Find the insertion point for content within a section.

    Args:
        doc_data: Raw document data from Google Docs API
        heading_text: Text of the heading identifying the section
        position: Where to insert - 'start' (after heading), 'end' (end of section)
        match_case: Whether to match case exactly

    Returns:
        Index for insertion, or None if section not found
    """
    section = find_section_by_heading(doc_data, heading_text, match_case)
    if not section:
        return None

    if position == "start":
        # Insert right after the heading
        elements = extract_structural_elements(doc_data)
        for elem in elements:
            if (elem["type"].startswith("heading") or elem["type"] == "title") and elem[
                "text"
            ].strip() == section["heading"].strip():
                return elem["end_index"]
        return section["start_index"] + len(section["heading"]) + 1
    else:
        # Insert at end of section
        return section["end_index"]


def find_elements_by_type(
    doc_data: dict[str, Any], element_type: str
) -> list[dict[str, Any]]:
    """
    Find all elements of a specific type in the document.

    This function identifies all occurrences of a given element type
    (tables, lists, headings, paragraphs) and returns their positions
    and metadata.

    Args:
        doc_data: Raw document data from Google Docs API
        element_type: Type of element to find. Supported values:
            - 'table' (or 'tables'): All tables in the document
            - 'heading' (or 'headings'): All headings (any level)
            - 'heading1' through 'heading6': Specific heading levels
            - 'title': Document title style elements
            - 'paragraph' (or 'paragraphs'): Regular paragraphs (non-heading, non-list)
            - 'bullet_list' (or 'bullet_lists'): Bulleted lists
            - 'numbered_list' (or 'numbered_lists'): Numbered/ordered lists
            - 'list' (or 'lists'): Both bullet and numbered lists
            - 'table_of_contents' (or 'toc'): Table of contents elements

    Returns:
        List of element dictionaries, each containing:
            - type: Element type string
            - start_index: Starting character position
            - end_index: Ending character position
            - Additional fields depending on element type:
                - For tables: rows, columns
                - For headings: text, level
                - For paragraphs: text
                - For lists: items (list of item dicts)

    Example:
        # Find all tables
        tables = find_elements_by_type(doc_data, 'table')
        for table in tables:
            print(f"Table at {table['start_index']}: {table['rows']}x{table['columns']}")

        # Find all H2 headings
        h2s = find_elements_by_type(doc_data, 'heading2')
        for h in h2s:
            print(f"H2: {h['text']} at position {h['start_index']}")
    """
    elements = extract_structural_elements(doc_data)

    # Normalize element type for matching
    search_type = element_type.lower().strip()

    # Map common aliases (including plural forms)
    type_aliases = {
        "heading": [
            "heading1",
            "heading2",
            "heading3",
            "heading4",
            "heading5",
            "heading6",
            "title",
        ],
        "headings": [
            "heading1",
            "heading2",
            "heading3",
            "heading4",
            "heading5",
            "heading6",
            "title",
        ],
        "list": ["bullet_list", "numbered_list"],
        "lists": ["bullet_list", "numbered_list"],
        "toc": ["table_of_contents"],
        # Plural -> singular mappings
        "tables": ["table"],
        "paragraphs": ["paragraph"],
        "bullet_lists": ["bullet_list"],
        "numbered_lists": ["numbered_list"],
    }

    # Determine which types to match
    if search_type in type_aliases:
        match_types = type_aliases[search_type]
    else:
        match_types = [search_type]

    # Filter elements by type
    matched = []
    for elem in elements:
        elem_type = elem.get("type", "").lower()
        if elem_type in match_types:
            matched.append(elem)

    return matched


def get_element_ancestors(doc_data: dict[str, Any], index: int) -> list[dict[str, Any]]:
    """
    Find the parent section/heading hierarchy for an element at a given position.

    This function traces the hierarchical path from the document root down to
    the element at the specified index, showing which sections contain it.

    Args:
        doc_data: Raw document data from Google Docs API
        index: Character position in the document

    Returns:
        List of ancestor headings from root to most specific, each containing:
            - type: Heading type (e.g., 'heading1', 'heading2')
            - text: Heading text
            - level: Heading level (0=title, 1-6 for headings)
            - start_index: Start position of the heading
            - end_index: End position of the heading
            - section_end: Where this section ends (next same/higher level heading)

        Returns empty list if the index is before any heading or document is empty.

    Example:
        # Find what section contains position 500
        ancestors = get_element_ancestors(doc_data, 500)
        for a in ancestors:
            print(f"{'  ' * a['level']}{a['text']}")

        # Might output:
        # Introduction (level 1)
        #   Background (level 2)
        #     Technical Details (level 3)
    """
    elements = extract_structural_elements(doc_data)

    # Extract all headings with their section ranges
    headings = []
    for i, elem in enumerate(elements):
        if elem["type"].startswith("heading") or elem["type"] == "title":
            level = elem.get("level", 0)

            # Determine section end: next heading of same or higher level
            section_end = None
            for j in range(i + 1, len(elements)):
                next_elem = elements[j]
                if (
                    next_elem["type"].startswith("heading")
                    or next_elem["type"] == "title"
                ):
                    next_level = next_elem.get("level", 0)
                    if next_level <= level:
                        section_end = next_elem["start_index"]
                        break

            # If no ending heading found, section goes to document end
            if section_end is None:
                body = doc_data.get("body", {})
                content = body.get("content", [])
                if content:
                    section_end = content[-1].get("endIndex", elem["end_index"])
                else:
                    section_end = elem["end_index"]

            headings.append(
                {
                    "type": elem["type"],
                    "text": elem.get("text", ""),
                    "level": level,
                    "start_index": elem["start_index"],
                    "end_index": elem["end_index"],
                    "section_end": section_end,
                }
            )

    # Find all ancestors that contain this index
    ancestors = []
    for heading in headings:
        # Check if this heading's section contains the index
        # Section starts at heading start and ends at section_end
        if heading["start_index"] <= index < heading["section_end"]:
            ancestors.append(heading)

    # Sort by level to get hierarchy from root to leaf
    ancestors.sort(key=lambda x: x["level"])

    return ancestors


def get_heading_siblings(
    doc_data: dict[str, Any], heading_text: str, match_case: bool = False
) -> dict[str, Any]:
    """
    Find the previous and next headings at the same level as the specified heading.

    This function helps navigate between sibling sections in a document,
    allowing movement to the previous or next section at the same hierarchy level.

    Args:
        doc_data: Raw document data from Google Docs API
        heading_text: Text of the heading to find siblings for
        match_case: Whether to match case exactly when finding the heading

    Returns:
        Dictionary containing:
            - found: Boolean indicating if the heading was found
            - heading: The matched heading info (if found)
            - level: The heading level
            - previous: Previous sibling heading info, or None if first at this level
            - next: Next sibling heading info, or None if last at this level
            - siblings_count: Total count of headings at this level
            - position_in_siblings: 1-based position among siblings (e.g., 2 of 5)

        Returns {"found": False} if heading not found.

    Example:
        # Navigate between H2 sections
        result = get_heading_siblings(doc_data, "Methods")
        if result['found']:
            if result['previous']:
                print(f"Previous: {result['previous']['text']}")
            if result['next']:
                print(f"Next: {result['next']['text']}")
            print(f"Position: {result['position_in_siblings']} of {result['siblings_count']}")
    """
    elements = extract_structural_elements(doc_data)

    # Find the target heading
    target_heading = None
    search_text = heading_text if match_case else heading_text.lower()

    headings_at_level = []

    for i, elem in enumerate(elements):
        if elem["type"].startswith("heading") or elem["type"] == "title":
            elem_text = elem.get("text", "")
            compare_text = elem_text if match_case else elem_text.lower()

            if compare_text.strip() == search_text.strip():
                target_heading = elem

    if target_heading is None:
        return {"found": False}

    target_level = target_heading.get("level", 0)

    # Find all headings at the same level
    for elem in elements:
        if elem["type"].startswith("heading") or elem["type"] == "title":
            if elem.get("level", 0) == target_level:
                headings_at_level.append(
                    {
                        "type": elem["type"],
                        "text": elem.get("text", ""),
                        "level": elem.get("level", 0),
                        "start_index": elem["start_index"],
                        "end_index": elem["end_index"],
                    }
                )

    # Find position in siblings list
    position = -1
    for i, h in enumerate(headings_at_level):
        if h["start_index"] == target_heading["start_index"]:
            position = i
            break

    previous_sibling = headings_at_level[position - 1] if position > 0 else None
    next_sibling = (
        headings_at_level[position + 1]
        if position < len(headings_at_level) - 1
        else None
    )

    return {
        "found": True,
        "heading": {
            "type": target_heading["type"],
            "text": target_heading.get("text", ""),
            "level": target_level,
            "start_index": target_heading["start_index"],
            "end_index": target_heading["end_index"],
        },
        "level": target_level,
        "previous": previous_sibling,
        "next": next_sibling,
        "siblings_count": len(headings_at_level),
        "position_in_siblings": position + 1,  # 1-based
    }


def get_paragraph_style_at_index(doc_data: dict[str, Any], index: int) -> Optional[str]:
    """
    Get the paragraph style type at a specific index in the document.

    This is useful for detecting if text is being inserted after a heading,
    which would cause it to inherit the heading's style.

    Args:
        doc_data: Raw document data from Google Docs API
        index: The document index position to check

    Returns:
        The paragraph style type (e.g., 'HEADING_1', 'NORMAL_TEXT') or None if not found
    """
    body = doc_data.get("body", {})
    content = body.get("content", [])

    for element in content:
        start_idx = element.get("startIndex", 0)
        end_idx = element.get("endIndex", 0)

        # Check if the index falls within this element's range
        if start_idx <= index < end_idx:
            if "paragraph" in element:
                return _get_paragraph_style_type(element["paragraph"])

    return None


def is_heading_style(style_type: Optional[str]) -> bool:
    """
    Check if a paragraph style type is a heading style.

    Args:
        style_type: The paragraph style type (e.g., 'HEADING_1', 'NORMAL_TEXT')

    Returns:
        True if the style is a heading (HEADING_1-6 or TITLE), False otherwise
    """
    if style_type is None:
        return False
    return style_type in HEADING_TYPES
