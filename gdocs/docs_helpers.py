"""
Google Docs Helper Functions

This module provides utility functions for common Google Docs operations
to simplify the implementation of document editing tools.
"""
import logging
from typing import Dict, Any, Optional, Tuple, List
from enum import Enum
from dataclasses import dataclass, asdict

logger = logging.getLogger(__name__)


class OperationType(str, Enum):
    """Type of document modification operation."""
    INSERT = "insert"
    REPLACE = "replace"
    DELETE = "delete"
    FORMAT = "format"


@dataclass
class OperationResult:
    """
    Result of a document modification operation with position shift information.

    This enables efficient follow-up edits without re-reading the document
    by providing the exact position shift caused by the operation.
    """
    success: bool
    operation: str  # OperationType value
    position_shift: int  # Positive = positions shifted right, negative = shifted left
    affected_range: Dict[str, int]  # {"start": x, "end": y}
    message: str
    link: str

    # Optional fields depending on operation type
    inserted_length: Optional[int] = None
    deleted_length: Optional[int] = None
    original_length: Optional[int] = None
    new_length: Optional[int] = None
    styles_applied: Optional[List[str]] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary, excluding None values."""
        result = asdict(self)
        # Remove None values for cleaner output
        return {k: v for k, v in result.items() if v is not None}


def calculate_position_shift(
    operation_type: OperationType,
    start_index: int,
    end_index: Optional[int],
    text_length: int
) -> Tuple[int, Dict[str, int]]:
    """
    Calculate the position shift caused by a document operation.

    Args:
        operation_type: Type of operation performed
        start_index: Start position of the operation
        end_index: End position (for replace/delete operations)
        text_length: Length of inserted/new text (0 for delete, format)

    Returns:
        Tuple of (position_shift, affected_range)
        - position_shift: How much positions after the operation shifted
        - affected_range: {"start": x, "end": y} of the affected area
    """
    if operation_type == OperationType.INSERT:
        # Insert: all positions >= start_index shift by text_length
        shift = text_length
        affected_range = {"start": start_index, "end": start_index + text_length}

    elif operation_type == OperationType.DELETE:
        # Delete: all positions >= end_index shift by -(end_index - start_index)
        deleted_length = (end_index or start_index) - start_index
        shift = -deleted_length
        affected_range = {"start": start_index, "end": start_index}

    elif operation_type == OperationType.REPLACE:
        # Replace: shift = new_length - old_length
        old_length = (end_index or start_index) - start_index
        shift = text_length - old_length
        affected_range = {"start": start_index, "end": start_index + text_length}

    elif operation_type == OperationType.FORMAT:
        # Format: no position shift
        shift = 0
        affected_range = {"start": start_index, "end": end_index or start_index}

    else:
        shift = 0
        affected_range = {"start": start_index, "end": end_index or start_index}

    return shift, affected_range


def build_operation_result(
    operation_type: OperationType,
    start_index: int,
    end_index: Optional[int],
    text: Optional[str],
    document_id: str,
    extra_info: Optional[Dict[str, Any]] = None,
    styles_applied: Optional[List[str]] = None
) -> OperationResult:
    """
    Build an OperationResult with calculated position shift.

    Args:
        operation_type: Type of operation performed
        start_index: Start position of the operation
        end_index: End position (for replace/delete operations)
        text: Text that was inserted/replaced (None for delete/format)
        document_id: ID of the document
        extra_info: Additional info to include in message
        styles_applied: List of style names applied (for format operations)

    Returns:
        OperationResult with all fields populated
    """
    text_length = len(text) if text else 0
    shift, affected_range = calculate_position_shift(
        operation_type, start_index, end_index, text_length
    )

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    # Build message based on operation type
    if operation_type == OperationType.INSERT:
        message = f"Inserted {text_length} characters at index {start_index}"
        result = OperationResult(
            success=True,
            operation=operation_type.value,
            position_shift=shift,
            affected_range=affected_range,
            message=message,
            link=link,
            inserted_length=text_length
        )

    elif operation_type == OperationType.DELETE:
        deleted_length = (end_index or start_index) - start_index
        message = f"Deleted {deleted_length} characters from index {start_index} to {end_index}"
        result = OperationResult(
            success=True,
            operation=operation_type.value,
            position_shift=shift,
            affected_range=affected_range,
            message=message,
            link=link,
            deleted_length=deleted_length
        )

    elif operation_type == OperationType.REPLACE:
        original_length = (end_index or start_index) - start_index
        message = f"Replaced {original_length} characters with {text_length} characters at index {start_index}"
        result = OperationResult(
            success=True,
            operation=operation_type.value,
            position_shift=shift,
            affected_range=affected_range,
            message=message,
            link=link,
            original_length=original_length,
            new_length=text_length
        )

    elif operation_type == OperationType.FORMAT:
        format_length = (end_index or start_index) - start_index
        styles_str = ", ".join(styles_applied) if styles_applied else "styles"
        message = f"Applied {styles_str} to {format_length} characters at index {start_index}-{end_index}"
        result = OperationResult(
            success=True,
            operation=operation_type.value,
            position_shift=shift,
            affected_range=affected_range,
            message=message,
            link=link,
            styles_applied=styles_applied
        )

    else:
        message = "Operation completed"
        result = OperationResult(
            success=True,
            operation=operation_type.value,
            position_shift=shift,
            affected_range=affected_range,
            message=message,
            link=link
        )

    # Add extra info to message if provided
    if extra_info:
        extra_parts = []
        if 'search_text' in extra_info:
            extra_parts.append(f"Search: '{extra_info['search_text']}'")
        if 'heading' in extra_info:
            extra_parts.append(f"Section: '{extra_info['heading']}'")
        if extra_parts:
            result.message += f" ({', '.join(extra_parts)})"

    return result


class SearchPosition(str, Enum):
    """Position relative to search result for text operations."""
    BEFORE = "before"
    AFTER = "after"
    REPLACE = "replace"


def extract_document_text_with_indices(doc_data: Dict[str, Any]) -> List[Tuple[str, int, int]]:
    """
    Extract all text content from a document with their start and end indices.

    Args:
        doc_data: Raw document data from Google Docs API

    Returns:
        List of tuples (text_content, start_index, end_index)
    """
    text_segments = []
    body = doc_data.get('body', {})
    content = body.get('content', [])

    def extract_from_elements(elements: List[Dict[str, Any]]) -> None:
        """Recursively extract text from document elements."""
        for element in elements:
            if 'paragraph' in element:
                paragraph = element['paragraph']
                for para_element in paragraph.get('elements', []):
                    if 'textRun' in para_element:
                        text_run = para_element['textRun']
                        text = text_run.get('content', '')
                        start_idx = para_element.get('startIndex', 0)
                        end_idx = para_element.get('endIndex', start_idx + len(text))
                        if text:
                            text_segments.append((text, start_idx, end_idx))
            elif 'table' in element:
                # Extract text from table cells
                table = element['table']
                for row in table.get('tableRows', []):
                    for cell in row.get('tableCells', []):
                        cell_content = cell.get('content', [])
                        extract_from_elements(cell_content)

    extract_from_elements(content)
    return text_segments


def extract_text_at_range(
    doc_data: Dict[str, Any],
    start_index: int,
    end_index: int,
    context_chars: int = 50
) -> Dict[str, Any]:
    """
    Extract text at a specific index range from a document.

    Args:
        doc_data: Raw document data from Google Docs API
        start_index: Start of the range
        end_index: End of the range
        context_chars: Number of characters of context to include before/after

    Returns:
        Dictionary with:
        - text: The text at the specified range
        - context_before: Text before the range
        - context_after: Text after the range
        - found: Whether text was found at the range
    """
    text_segments = extract_document_text_with_indices(doc_data)

    # Build full document text with index mapping
    full_text = ""
    index_map = []  # Maps position in full_text to document index
    reverse_map = {}  # Maps document index to position in full_text

    for segment_text, start_idx, _ in text_segments:
        for i, char in enumerate(segment_text):
            doc_idx = start_idx + i
            reverse_map[doc_idx] = len(full_text)
            index_map.append(doc_idx)
            full_text += char

    # Find positions in full_text
    text_start_pos = reverse_map.get(start_index)
    text_end_pos = reverse_map.get(end_index - 1)  # end_index is exclusive

    if text_start_pos is None:
        # Start index not in text (might be at end or in non-text element)
        # Try to find the closest position
        for idx in range(start_index, end_index):
            if idx in reverse_map:
                text_start_pos = reverse_map[idx]
                break

    if text_end_pos is None and end_index > start_index:
        # Find closest end position
        for idx in range(end_index - 1, start_index - 1, -1):
            if idx in reverse_map:
                text_end_pos = reverse_map[idx]
                break

    result = {
        "text": "",
        "context_before": "",
        "context_after": "",
        "found": False,
        "start_index": start_index,
        "end_index": end_index
    }

    if text_start_pos is not None:
        result["found"] = True

        if text_end_pos is not None and text_end_pos >= text_start_pos:
            result["text"] = full_text[text_start_pos:text_end_pos + 1]
        else:
            # Just the character at start (for insert point)
            result["text"] = ""

        # Extract context
        context_start = max(0, text_start_pos - context_chars)
        context_end = min(len(full_text), (text_end_pos or text_start_pos) + 1 + context_chars)

        result["context_before"] = full_text[context_start:text_start_pos]
        if text_end_pos is not None:
            result["context_after"] = full_text[text_end_pos + 1:context_end]
        else:
            result["context_after"] = full_text[text_start_pos:context_end]

    return result


def find_text_in_document(
    doc_data: Dict[str, Any],
    search_text: str,
    occurrence: int = 1,
    match_case: bool = True
) -> Optional[Tuple[int, int]]:
    """
    Find text in document and return its start and end indices.

    Args:
        doc_data: Raw document data from Google Docs API
        search_text: Text to search for
        occurrence: Which occurrence to find (1=first, 2=second, -1=last)
        match_case: Whether to match case exactly

    Returns:
        Tuple of (start_index, end_index) or None if not found
    """
    if not search_text:
        return None

    # Build full document text with index mapping
    text_segments = extract_document_text_with_indices(doc_data)

    # Reconstruct document text and build index mapping
    full_text = ""
    index_map = []  # Maps position in full_text to document index

    for segment_text, start_idx, _ in text_segments:
        full_text += segment_text
        # Map each character position in the segment
        for i, _ in enumerate(segment_text):
            index_map.append(start_idx + i)

    # Search for text
    search_target = search_text if match_case else search_text.lower()
    search_in = full_text if match_case else full_text.lower()

    # Find all occurrences
    occurrences = []
    pos = 0
    while True:
        found = search_in.find(search_target, pos)
        if found == -1:
            break
        occurrences.append(found)
        pos = found + 1

    if not occurrences:
        return None

    # Select the appropriate occurrence
    if occurrence == -1:
        # Last occurrence
        target_idx = occurrences[-1]
    elif occurrence > 0 and occurrence <= len(occurrences):
        target_idx = occurrences[occurrence - 1]
    else:
        return None

    # Map back to document indices
    if target_idx >= len(index_map) or target_idx + len(search_text) > len(index_map):
        return None

    doc_start = index_map[target_idx]
    doc_end = index_map[target_idx + len(search_text) - 1] + 1

    return (doc_start, doc_end)


def find_all_occurrences_in_document(
    doc_data: Dict[str, Any],
    search_text: str,
    match_case: bool = True
) -> List[Tuple[int, int]]:
    """
    Find all occurrences of text in document.

    Args:
        doc_data: Raw document data from Google Docs API
        search_text: Text to search for
        match_case: Whether to match case exactly

    Returns:
        List of (start_index, end_index) tuples for all occurrences
    """
    if not search_text:
        return []

    # Build full document text with index mapping
    text_segments = extract_document_text_with_indices(doc_data)

    # Reconstruct document text and build index mapping
    full_text = ""
    index_map = []

    for segment_text, start_idx, end_idx in text_segments:
        full_text += segment_text
        for i, _ in enumerate(segment_text):
            index_map.append(start_idx + i)

    # Search for text
    search_target = search_text if match_case else search_text.lower()
    search_in = full_text if match_case else full_text.lower()

    # Find all occurrences
    results = []
    pos = 0
    while True:
        found = search_in.find(search_target, pos)
        if found == -1:
            break

        if found < len(index_map) and found + len(search_text) <= len(index_map):
            doc_start = index_map[found]
            doc_end = index_map[found + len(search_text) - 1] + 1
            results.append((doc_start, doc_end))

        pos = found + 1

    return results


def calculate_search_based_indices(
    doc_data: Dict[str, Any],
    search_text: str,
    position: str,
    occurrence: int = 1,
    match_case: bool = True
) -> Tuple[bool, Optional[int], Optional[int], str]:
    """
    Calculate start and end indices based on search text and position.

    Args:
        doc_data: Raw document data from Google Docs API
        search_text: Text to search for
        position: Where to operate relative to found text ("before", "after", "replace")
        occurrence: Which occurrence to target (1=first, 2=second, -1=last)
        match_case: Whether to match case exactly

    Returns:
        Tuple of (success, start_index, end_index, message)
        - For 'before': Returns position just before found text
        - For 'after': Returns position just after found text
        - For 'replace': Returns the range of the found text
    """
    # Find the text
    found = find_text_in_document(doc_data, search_text, occurrence, match_case)

    if found is None:
        # Get occurrence info for error message
        all_occurrences = find_all_occurrences_in_document(doc_data, search_text, match_case)
        if not all_occurrences:
            return (False, None, None, f"Text '{search_text}' not found in document")
        else:
            return (False, None, None,
                    f"Occurrence {occurrence} of '{search_text}' not found. "
                    f"Document contains {len(all_occurrences)} occurrence(s).")

    found_start, found_end = found

    # Calculate indices based on position
    if position == SearchPosition.BEFORE.value:
        # Insert point is right before the found text
        return (True, found_start, found_start, f"Found at index {found_start}")
    elif position == SearchPosition.AFTER.value:
        # Insert point is right after the found text
        return (True, found_end, found_end, f"Found at index {found_start}, inserting after index {found_end}")
    elif position == SearchPosition.REPLACE.value:
        # Return the full range to replace
        return (True, found_start, found_end, f"Found at index range {found_start}-{found_end}")
    else:
        return (False, None, None, f"Invalid position '{position}'. Use 'before', 'after', or 'replace'.")


def _parse_color(color_str: str) -> Dict[str, Any]:
    """
    Parse a color string (hex or named) to Google Docs API color format.
    
    Args:
        color_str: Color as hex (#FF0000, #F00) or CSS named color
        
    Returns:
        Dictionary with rgbColor format for Google Docs API
    """
    # Handle hex colors
    if color_str.startswith('#'):
        hex_color = color_str.lstrip('#')
        # Handle short hex (#F00 -> #FF0000)
        if len(hex_color) == 3:
            hex_color = ''.join(c*2 for c in hex_color)
        if len(hex_color) != 6:
            raise ValueError(f"Invalid hex color: {color_str}")
        r = int(hex_color[0:2], 16) / 255.0
        g = int(hex_color[2:4], 16) / 255.0
        b = int(hex_color[4:6], 16) / 255.0
        return {'color': {'rgbColor': {'red': r, 'green': g, 'blue': b}}}
    
    # Handle common named colors
    named_colors = {
        'red': (1.0, 0.0, 0.0),
        'green': (0.0, 1.0, 0.0),
        'blue': (0.0, 0.0, 1.0),
        'yellow': (1.0, 1.0, 0.0),
        'orange': (1.0, 0.65, 0.0),
        'purple': (0.5, 0.0, 0.5),
        'black': (0.0, 0.0, 0.0),
        'white': (1.0, 1.0, 1.0),
        'gray': (0.5, 0.5, 0.5),
        'grey': (0.5, 0.5, 0.5),
    }
    color_lower = color_str.lower()
    if color_lower in named_colors:
        r, g, b = named_colors[color_lower]
        return {'color': {'rgbColor': {'red': r, 'green': g, 'blue': b}}}
    
    raise ValueError(f"Unknown color format: {color_str}. Use hex (#FF0000) or named colors.")


def build_text_style(
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    strikethrough: bool = None,
    font_size: int = None,
    font_family: str = None,
    link: str = None,
    foreground_color: str = None,
    background_color: str = None,
) -> tuple[Dict[str, Any], list[str]]:
    """
    Build text style object for Google Docs API requests.

    Args:
        bold: Whether text should be bold
        italic: Whether text should be italic
        underline: Whether text should be underlined
        strikethrough: Whether text should have strikethrough
        font_size: Font size in points
        font_family: Font family name
        link: URL to create a hyperlink (use empty string "" to remove existing link)
        foreground_color: Text color as hex (#FF0000) or named color (red, blue, etc.)
        background_color: Background/highlight color as hex or named color

    Returns:
        Tuple of (text_style_dict, list_of_field_names)
    """
    text_style = {}
    fields = []

    if bold is not None:
        text_style['bold'] = bold
        fields.append('bold')

    if italic is not None:
        text_style['italic'] = italic
        fields.append('italic')

    if underline is not None:
        text_style['underline'] = underline
        fields.append('underline')

    if strikethrough is not None:
        text_style['strikethrough'] = strikethrough
        fields.append('strikethrough')

    if font_size is not None:
        text_style['fontSize'] = {'magnitude': font_size, 'unit': 'PT'}
        fields.append('fontSize')

    if font_family is not None:
        text_style['weightedFontFamily'] = {'fontFamily': font_family}
        fields.append('weightedFontFamily')

    if link is not None:
        if link == "":
            # Empty string removes the link
            text_style['link'] = None
        else:
            text_style['link'] = {'url': link}
        fields.append('link')

    if foreground_color is not None:
        text_style['foregroundColor'] = _parse_color(foreground_color)
        fields.append('foregroundColor')

    if background_color is not None:
        text_style['backgroundColor'] = _parse_color(background_color)
        fields.append('backgroundColor')

    return text_style, fields

def create_insert_text_request(index: int, text: str) -> Dict[str, Any]:
    """
    Create an insertText request for Google Docs API.
    
    Args:
        index: Position to insert text
        text: Text to insert
    
    Returns:
        Dictionary representing the insertText request
    """
    return {
        'insertText': {
            'location': {'index': index},
            'text': text
        }
    }

def create_insert_text_segment_request(index: int, text: str, segment_id: str) -> Dict[str, Any]:
    """
    Create an insertText request for Google Docs API with segmentId (for headers/footers).
    
    Args:
        index: Position to insert text
        text: Text to insert
        segment_id: Segment ID (for targeting headers/footers)
    
    Returns:
        Dictionary representing the insertText request with segmentId
    """
    return {
        'insertText': {
            'location': {
                'segmentId': segment_id,
                'index': index
            },
            'text': text
        }
    }

def create_delete_range_request(start_index: int, end_index: int) -> Dict[str, Any]:
    """
    Create a deleteContentRange request for Google Docs API.
    
    Args:
        start_index: Start position of content to delete
        end_index: End position of content to delete
    
    Returns:
        Dictionary representing the deleteContentRange request
    """
    return {
        'deleteContentRange': {
            'range': {
                'startIndex': start_index,
                'endIndex': end_index
            }
        }
    }

def create_format_text_request(
    start_index: int,
    end_index: int,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    strikethrough: bool = None,
    font_size: int = None,
    font_family: str = None,
    link: str = None,
    foreground_color: str = None,
    background_color: str = None,
) -> Optional[Dict[str, Any]]:
    """
    Create an updateTextStyle request for Google Docs API.

    Args:
        start_index: Start position of text to format
        end_index: End position of text to format
        bold: Whether text should be bold
        italic: Whether text should be italic
        underline: Whether text should be underlined
        strikethrough: Whether text should have strikethrough
        font_size: Font size in points
        font_family: Font family name
        link: URL to create a hyperlink (use empty string "" to remove existing link)
        foreground_color: Text color as hex (#FF0000) or named color
        background_color: Background/highlight color as hex or named color

    Returns:
        Dictionary representing the updateTextStyle request, or None if no styles provided
    """
    text_style, fields = build_text_style(
        bold, italic, underline, strikethrough, font_size, font_family, link,
        foreground_color, background_color
    )
    
    if not text_style:
        return None
    
    return {
        'updateTextStyle': {
            'range': {
                'startIndex': start_index,
                'endIndex': end_index
            },
            'textStyle': text_style,
            'fields': ','.join(fields)
        }
    }

def create_find_replace_request(
    find_text: str, 
    replace_text: str, 
    match_case: bool = False
) -> Dict[str, Any]:
    """
    Create a replaceAllText request for Google Docs API.
    
    Args:
        find_text: Text to find
        replace_text: Text to replace with
        match_case: Whether to match case exactly
    
    Returns:
        Dictionary representing the replaceAllText request
    """
    return {
        'replaceAllText': {
            'containsText': {
                'text': find_text,
                'matchCase': match_case
            },
            'replaceText': replace_text
        }
    }

def create_insert_table_request(index: int, rows: int, columns: int) -> Dict[str, Any]:
    """
    Create an insertTable request for Google Docs API.
    
    Args:
        index: Position to insert table
        rows: Number of rows
        columns: Number of columns
    
    Returns:
        Dictionary representing the insertTable request
    """
    return {
        'insertTable': {
            'location': {'index': index},
            'rows': rows,
            'columns': columns
        }
    }

def create_insert_page_break_request(index: int) -> Dict[str, Any]:
    """
    Create an insertPageBreak request for Google Docs API.
    
    Args:
        index: Position to insert page break
    
    Returns:
        Dictionary representing the insertPageBreak request
    """
    return {
        'insertPageBreak': {
            'location': {'index': index}
        }
    }

def create_insert_image_request(
    index: int, 
    image_uri: str,
    width: int = None,
    height: int = None
) -> Dict[str, Any]:
    """
    Create an insertInlineImage request for Google Docs API.
    
    Args:
        index: Position to insert image
        image_uri: URI of the image (Drive URL or public URL)
        width: Image width in points
        height: Image height in points
    
    Returns:
        Dictionary representing the insertInlineImage request
    """
    request = {
        'insertInlineImage': {
            'location': {'index': index},
            'uri': image_uri
        }
    }
    
    # Add size properties if specified
    object_size = {}
    if width is not None:
        object_size['width'] = {'magnitude': width, 'unit': 'PT'}
    if height is not None:
        object_size['height'] = {'magnitude': height, 'unit': 'PT'}
    
    if object_size:
        request['insertInlineImage']['objectSize'] = object_size
    
    return request

def create_bullet_list_request(
    start_index: int, 
    end_index: int,
    list_type: str = "UNORDERED"
) -> Dict[str, Any]:
    """
    Create a createParagraphBullets request for Google Docs API.
    
    Args:
        start_index: Start of text range to convert to list
        end_index: End of text range to convert to list
        list_type: Type of list ("UNORDERED" or "ORDERED")
    
    Returns:
        Dictionary representing the createParagraphBullets request
    """
    bullet_preset = (
        'BULLET_DISC_CIRCLE_SQUARE' 
        if list_type == "UNORDERED" 
        else 'NUMBERED_DECIMAL_ALPHA_ROMAN'
    )
    
    return {
        'createParagraphBullets': {
            'range': {
                'startIndex': start_index,
                'endIndex': end_index
            },
            'bulletPreset': bullet_preset
        }
    }

# =============================================================================
# Range-Based Selection for Semantic Editing
# =============================================================================


class ExtendBoundary(str, Enum):
    """Boundary type for extending search results."""
    PARAGRAPH = "paragraph"
    SENTENCE = "sentence"
    LINE = "line"
    SECTION = "section"


@dataclass
class RangeResult:
    """
    Result of range resolution with detailed information.

    This enables semantic text selection with clear feedback about
    what was matched and how the range was resolved.
    """
    success: bool
    start_index: Optional[int]
    end_index: Optional[int]
    message: str

    # Detailed match information
    matched_start: Optional[str] = None  # What text matched the start
    matched_end: Optional[str] = None    # What text matched the end
    extend_type: Optional[str] = None    # If boundary extension was used
    section_name: Optional[str] = None   # If section-based selection

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary, excluding None values."""
        result = asdict(self)
        return {k: v for k, v in result.items() if v is not None}


def find_paragraph_boundaries(
    doc_data: Dict[str, Any],
    index: int
) -> Tuple[int, int]:
    """
    Find the paragraph boundaries containing a given index.

    Args:
        doc_data: Raw document data from Google Docs API
        index: Position within the document

    Returns:
        Tuple of (paragraph_start, paragraph_end)
    """
    body = doc_data.get('body', {})
    content = body.get('content', [])

    for element in content:
        start_idx = element.get('startIndex', 0)
        end_idx = element.get('endIndex', 0)

        if start_idx <= index < end_idx:
            if 'paragraph' in element:
                return (start_idx, end_idx)
            elif 'table' in element:
                # For tables, find the cell containing the index
                table = element['table']
                for row in table.get('tableRows', []):
                    for cell in row.get('tableCells', []):
                        cell_start = cell.get('startIndex', 0)
                        cell_end = cell.get('endIndex', 0)
                        if cell_start <= index < cell_end:
                            # Find paragraph within cell
                            for cell_elem in cell.get('content', []):
                                if 'paragraph' in cell_elem:
                                    p_start = cell_elem.get('startIndex', 0)
                                    p_end = cell_elem.get('endIndex', 0)
                                    if p_start <= index < p_end:
                                        return (p_start, p_end)

    # Fallback: return a small range around the index
    return (index, index + 1)


def find_sentence_boundaries(
    doc_data: Dict[str, Any],
    index: int
) -> Tuple[int, int]:
    """
    Find the sentence boundaries containing a given index.

    Sentences are detected by common sentence-ending punctuation
    followed by whitespace or end of text. Handles common abbreviations
    (Mr., Mrs., Dr., etc.) to avoid false sentence breaks.

    Args:
        doc_data: Raw document data from Google Docs API
        index: Position within the document

    Returns:
        Tuple of (sentence_start, sentence_end)
    """
    import re

    # Get text segments with indices
    text_segments = extract_document_text_with_indices(doc_data)

    # Build full text and index mapping
    full_text = ""
    index_map = []

    for segment_text, start_idx, _ in text_segments:
        full_text += segment_text
        for i, _ in enumerate(segment_text):
            index_map.append(start_idx + i)

    if not index_map:
        return (index, index + 1)

    # Find position in full_text that corresponds to our index
    text_pos = None
    for i, doc_idx in enumerate(index_map):
        if doc_idx >= index:
            text_pos = i
            break

    if text_pos is None:
        text_pos = len(full_text) - 1

    # Common abbreviations that should not end sentences
    # These have periods but are not sentence endings
    abbreviations = {
        'mr', 'mrs', 'ms', 'dr', 'prof', 'sr', 'jr', 'vs', 'etc', 'inc', 'ltd',
        'corp', 'co', 'st', 'ave', 'blvd', 'rd', 'apt', 'no', 'vol', 'pg', 'pp',
        'jan', 'feb', 'mar', 'apr', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec',
        'fig', 'eg', 'ie', 'cf', 'al', 'ed', 'rev', 'gen', 'gov', 'sen', 'rep',
        'hon', 'col', 'maj', 'capt', 'lt', 'sgt', 'pvt', 'est', 'approx'
    }

    def is_sentence_end(text: str, pos: int) -> bool:
        """
        Check if a period/punctuation at position pos is a real sentence end.

        Returns True if this appears to be a sentence-ending punctuation.
        """
        if pos < 0 or pos >= len(text):
            return False

        char = text[pos]
        if char not in '.!?':
            return False

        # Exclamation and question marks are almost always sentence ends
        if char in '!?':
            # Check if followed by space/newline or end of text
            if pos + 1 >= len(text):
                return True
            next_char = text[pos + 1]
            return next_char in ' \t\n\r'

        # For periods, check if this might be an abbreviation
        if char == '.':
            # Check if followed by space/newline or end of text
            if pos + 1 >= len(text):
                return True
            next_char = text[pos + 1]

            # If not followed by whitespace, probably not a sentence end
            if next_char not in ' \t\n\r':
                return False

            # If followed by a lowercase letter after whitespace, probably not a sentence end
            # (e.g., "Dr. smith" is unlikely, but "Dr. Smith" is valid)
            if pos + 2 < len(text) and text[pos + 2].islower():
                # Might still be a real sentence end if the word is a common starter
                pass  # Fall through to abbreviation check

            # Check for known abbreviations
            # Look backward to find the word before the period
            word_start = pos
            while word_start > 0 and text[word_start - 1].isalpha():
                word_start -= 1

            word_before = text[word_start:pos].lower()
            if word_before in abbreviations:
                return False

            # Check for single-letter abbreviations (e.g., "A. Smith", "J.D.")
            if len(word_before) == 1:
                return False

            # Check for numeric patterns (e.g., "1.", "2.5")
            if word_start > 0 and text[word_start - 1].isdigit():
                return False

            # Check for ellipsis (...)
            if pos >= 2 and text[pos - 2:pos + 1] == '...':
                # Check if this ellipsis ends a sentence
                if pos + 1 >= len(text):
                    return True
                return text[pos + 1] in ' \t\n\r'

            return True

        return False

    def find_sentence_end_positions(text: str) -> List[int]:
        """Find all sentence end positions in text."""
        positions = []
        for i, char in enumerate(text):
            if char in '.!?' and is_sentence_end(text, i):
                # Include trailing whitespace as part of the sentence end
                end_pos = i + 1
                while end_pos < len(text) and text[end_pos] in ' \t':
                    end_pos += 1
                positions.append(end_pos)
        return positions

    # Find all sentence ends in the text
    sentence_ends = find_sentence_end_positions(full_text)

    # Find sentence start (the end of the previous sentence or start of text)
    sentence_start_pos = 0
    for end_pos in sentence_ends:
        if end_pos <= text_pos:
            sentence_start_pos = end_pos
        else:
            break

    # Find sentence end (the next sentence end or end of text)
    sentence_end_pos = len(full_text)
    for end_pos in sentence_ends:
        if end_pos > text_pos:
            sentence_end_pos = end_pos
            break

    # Map back to document indices
    if sentence_start_pos >= len(index_map):
        sentence_start_pos = len(index_map) - 1
    if sentence_end_pos > len(index_map):
        sentence_end_pos = len(index_map)

    doc_start = index_map[sentence_start_pos] if sentence_start_pos < len(index_map) else index
    doc_end = index_map[sentence_end_pos - 1] + 1 if sentence_end_pos > 0 and sentence_end_pos <= len(index_map) else index + 1

    return (doc_start, doc_end)


def find_line_boundaries(
    doc_data: Dict[str, Any],
    index: int
) -> Tuple[int, int]:
    """
    Find the line boundaries containing a given index.

    Lines are detected by newline characters.

    Args:
        doc_data: Raw document data from Google Docs API
        index: Position within the document

    Returns:
        Tuple of (line_start, line_end)
    """
    # Get text segments with indices
    text_segments = extract_document_text_with_indices(doc_data)

    # Build full text and index mapping
    full_text = ""
    index_map = []

    for segment_text, start_idx, _ in text_segments:
        full_text += segment_text
        for i, _ in enumerate(segment_text):
            index_map.append(start_idx + i)

    if not index_map:
        return (index, index + 1)

    # Find position in full_text that corresponds to our index
    text_pos = None
    for i, doc_idx in enumerate(index_map):
        if doc_idx >= index:
            text_pos = i
            break

    if text_pos is None:
        text_pos = len(full_text) - 1

    # Find line start (search backward for newline or start of text)
    line_start_pos = full_text.rfind('\n', 0, text_pos)
    if line_start_pos == -1:
        line_start_pos = 0
    else:
        line_start_pos += 1  # Move past the newline

    # Find line end (search forward for newline or end of text)
    line_end_pos = full_text.find('\n', text_pos)
    if line_end_pos == -1:
        line_end_pos = len(full_text)
    else:
        line_end_pos += 1  # Include the newline

    # Map back to document indices
    if line_start_pos >= len(index_map):
        line_start_pos = len(index_map) - 1
    if line_end_pos > len(index_map):
        line_end_pos = len(index_map)

    doc_start = index_map[line_start_pos] if line_start_pos < len(index_map) else index
    doc_end = index_map[line_end_pos - 1] + 1 if line_end_pos > 0 and line_end_pos <= len(index_map) else index + 1

    return (doc_start, doc_end)


def resolve_range_by_search_bounds(
    doc_data: Dict[str, Any],
    start_search: str,
    end_search: str,
    start_occurrence: int = 1,
    end_occurrence: int = 1,
    match_case: bool = True
) -> RangeResult:
    """
    Resolve a range defined by start and end search terms.

    The range includes from the start of the start match to
    the end of the end match.

    Args:
        doc_data: Raw document data from Google Docs API
        start_search: Text to search for range start
        end_search: Text to search for range end
        start_occurrence: Which occurrence of start_search (1=first, -1=last)
        end_occurrence: Which occurrence of end_search (1=first, -1=last)
        match_case: Whether to match case exactly

    Returns:
        RangeResult with resolved indices or error information
    """
    # Find start position
    start_result = find_text_in_document(doc_data, start_search, start_occurrence, match_case)
    if start_result is None:
        all_start = find_all_occurrences_in_document(doc_data, start_search, match_case)
        if not all_start:
            return RangeResult(
                success=False,
                start_index=None,
                end_index=None,
                message=f"Start text '{start_search}' not found in document"
            )
        return RangeResult(
            success=False,
            start_index=None,
            end_index=None,
            message=f"Occurrence {start_occurrence} of start text '{start_search}' not found. "
                    f"Document contains {len(all_start)} occurrence(s)."
        )

    start_idx, _ = start_result

    # Find end position
    end_result = find_text_in_document(doc_data, end_search, end_occurrence, match_case)
    if end_result is None:
        all_end = find_all_occurrences_in_document(doc_data, end_search, match_case)
        if not all_end:
            return RangeResult(
                success=False,
                start_index=None,
                end_index=None,
                message=f"End text '{end_search}' not found in document"
            )
        return RangeResult(
            success=False,
            start_index=None,
            end_index=None,
            message=f"Occurrence {end_occurrence} of end text '{end_search}' not found. "
                    f"Document contains {len(all_end)} occurrence(s)."
        )

    _, end_idx = end_result

    # Validate range makes sense
    if end_idx <= start_idx:
        return RangeResult(
            success=False,
            start_index=None,
            end_index=None,
            message=f"Invalid range: end text '{end_search}' (at {end_idx}) "
                    f"comes before or at start text '{start_search}' (at {start_idx})"
        )

    return RangeResult(
        success=True,
        start_index=start_idx,
        end_index=end_idx,
        message=f"Range resolved: {start_idx}-{end_idx}",
        matched_start=start_search,
        matched_end=end_search
    )


def resolve_range_by_search_with_extension(
    doc_data: Dict[str, Any],
    search: str,
    extend_to: str,
    occurrence: int = 1,
    match_case: bool = True
) -> RangeResult:
    """
    Find text and extend the selection to boundary (paragraph/sentence/line).

    Args:
        doc_data: Raw document data from Google Docs API
        search: Text to search for
        extend_to: Boundary type ("paragraph", "sentence", "line")
        occurrence: Which occurrence of search text (1=first, -1=last)
        match_case: Whether to match case exactly

    Returns:
        RangeResult with resolved indices extended to boundaries
    """
    # Find the search text
    result = find_text_in_document(doc_data, search, occurrence, match_case)
    if result is None:
        all_occurrences = find_all_occurrences_in_document(doc_data, search, match_case)
        if not all_occurrences:
            return RangeResult(
                success=False,
                start_index=None,
                end_index=None,
                message=f"Text '{search}' not found in document"
            )
        return RangeResult(
            success=False,
            start_index=None,
            end_index=None,
            message=f"Occurrence {occurrence} of '{search}' not found. "
                    f"Document contains {len(all_occurrences)} occurrence(s)."
        )

    found_start, found_end = result

    # Extend to boundary
    extend_to_lower = extend_to.lower()
    if extend_to_lower == ExtendBoundary.PARAGRAPH.value:
        start_idx, end_idx = find_paragraph_boundaries(doc_data, found_start)
    elif extend_to_lower == ExtendBoundary.SENTENCE.value:
        start_idx, end_idx = find_sentence_boundaries(doc_data, found_start)
    elif extend_to_lower == ExtendBoundary.LINE.value:
        start_idx, end_idx = find_line_boundaries(doc_data, found_start)
    elif extend_to_lower == ExtendBoundary.SECTION.value:
        # For section, we need structural navigation
        # Import here to avoid circular dependency
        from gdocs.docs_structure import extract_structural_elements

        elements = extract_structural_elements(doc_data)

        # Find which section contains this text using proper hierarchy awareness
        # A section is bounded by the next heading of SAME OR HIGHER level (lower number)
        section_start = None
        section_end = None
        current_heading = None
        current_heading_level = float('inf')  # Default to highest level (lowest priority)

        # First pass: find the closest heading before the found text
        for elem in elements:
            if elem['type'].startswith('heading') or elem['type'] == 'title':
                elem_level = elem.get('level', 0)
                if elem['end_index'] <= found_start:
                    # This heading is before our text - track it
                    current_heading = elem
                    current_heading_level = elem_level
                    section_start = elem['start_index']

        # Second pass: find section end (next heading of same or higher level)
        if current_heading is not None:
            found_current = False
            for elem in elements:
                if elem['type'].startswith('heading') or elem['type'] == 'title':
                    if elem == current_heading:
                        found_current = True
                        continue

                    if found_current:
                        elem_level = elem.get('level', 0)
                        # Section ends at headings of same or higher level (lower number)
                        # This properly handles nested subsections - they don't end the parent section
                        if elem_level <= current_heading_level:
                            section_end = elem['start_index']
                            break

        if section_start is None:
            # Text is before any heading - use document start
            section_start = 1

        if section_end is None:
            # Section goes to end of document
            body = doc_data.get('body', {})
            content = body.get('content', [])
            if content:
                section_end = content[-1].get('endIndex', found_end)
            else:
                section_end = found_end

        start_idx, end_idx = section_start, section_end
    else:
        return RangeResult(
            success=False,
            start_index=None,
            end_index=None,
            message=f"Invalid extend_to value '{extend_to}'. "
                    f"Use: paragraph, sentence, line, or section"
        )

    # Validate that extended range doesn't go backward (start > end)
    if start_idx > end_idx:
        return RangeResult(
            success=False,
            start_index=None,
            end_index=None,
            message=f"Range extension error: extending to {extend_to} resulted in invalid range "
                    f"where start ({start_idx}) > end ({end_idx}). "
                    f"The search text '{search}' was found at index {found_start}, "
                    f"but boundary detection failed. Try a different extend type or search term."
        )

    # Validate that the extended range still contains the original search result
    if start_idx > found_start or end_idx < found_end:
        return RangeResult(
            success=False,
            start_index=None,
            end_index=None,
            message=f"Range extension error: the {extend_to} boundary "
                    f"({start_idx}-{end_idx}) does not contain the search result "
                    f"({found_start}-{found_end}). This may indicate a structural "
                    f"issue with the document or an edge case in boundary detection."
        )

    return RangeResult(
        success=True,
        start_index=start_idx,
        end_index=end_idx,
        message=f"Range extended to {extend_to}: {start_idx}-{end_idx}",
        matched_start=search,
        extend_type=extend_to
    )


def resolve_range_by_search_with_offsets(
    doc_data: Dict[str, Any],
    search: str,
    before_chars: int = 0,
    after_chars: int = 0,
    occurrence: int = 1,
    match_case: bool = True
) -> RangeResult:
    """
    Find text and expand the selection by character offsets.

    Args:
        doc_data: Raw document data from Google Docs API
        search: Text to search for
        before_chars: Characters to include before the match (must be non-negative)
        after_chars: Characters to include after the match (must be non-negative)
        occurrence: Which occurrence of search text (1=first, -1=last)
        match_case: Whether to match case exactly

    Returns:
        RangeResult with resolved indices including offsets
    """
    # Validate offset parameters
    if before_chars < 0:
        return RangeResult(
            success=False,
            start_index=None,
            end_index=None,
            message=f"Invalid before_chars value ({before_chars}): must be non-negative. "
                    f"Use a positive value to include characters before the match."
        )

    if after_chars < 0:
        return RangeResult(
            success=False,
            start_index=None,
            end_index=None,
            message=f"Invalid after_chars value ({after_chars}): must be non-negative. "
                    f"Use a positive value to include characters after the match."
        )

    # Find the search text
    result = find_text_in_document(doc_data, search, occurrence, match_case)
    if result is None:
        all_occurrences = find_all_occurrences_in_document(doc_data, search, match_case)
        if not all_occurrences:
            return RangeResult(
                success=False,
                start_index=None,
                end_index=None,
                message=f"Text '{search}' not found in document"
            )
        return RangeResult(
            success=False,
            start_index=None,
            end_index=None,
            message=f"Occurrence {occurrence} of '{search}' not found. "
                    f"Document contains {len(all_occurrences)} occurrence(s)."
        )

    found_start, found_end = result

    # Get document bounds
    body = doc_data.get('body', {})
    content = body.get('content', [])
    doc_end = content[-1].get('endIndex', found_end) if content else found_end
    doc_start = 1  # Google Docs starts at index 1 (0 is reserved)

    # Apply offsets with safe clamping to document bounds
    # Use safe arithmetic to prevent integer overflow with very large values
    try:
        requested_start = found_start - before_chars
    except (OverflowError, ValueError):
        requested_start = doc_start

    try:
        requested_end = found_end + after_chars
    except (OverflowError, ValueError):
        requested_end = doc_end

    # Clamp to document bounds
    start_idx = max(doc_start, requested_start)
    end_idx = min(doc_end, requested_end)

    # Track if clamping occurred for informative message
    start_clamped = start_idx > requested_start
    end_clamped = end_idx < requested_end

    # Build informative message
    message_parts = [f"Range with offsets: {start_idx}-{end_idx}"]
    message_parts.append(f"(match at {found_start}-{found_end})")

    if start_clamped or end_clamped:
        clamp_notes = []
        if start_clamped:
            clamp_notes.append(f"start clamped from {requested_start} to {start_idx}")
        if end_clamped:
            clamp_notes.append(f"end clamped from {requested_end} to {end_idx}")
        message_parts.append(f"Note: {', '.join(clamp_notes)} (document bounds: {doc_start}-{doc_end})")

    return RangeResult(
        success=True,
        start_index=start_idx,
        end_index=end_idx,
        message=" ".join(message_parts),
        matched_start=search
    )


def resolve_range_by_section(
    doc_data: Dict[str, Any],
    section_heading: str,
    include_heading: bool = False,
    include_subsections: bool = True,
    match_case: bool = False
) -> RangeResult:
    """
    Select a complete section by its heading.

    Args:
        doc_data: Raw document data from Google Docs API
        section_heading: Text of the heading to select
        include_heading: Whether to include the heading text in the selection
        include_subsections: Whether to include nested subsections
        match_case: Whether to match case exactly

    Returns:
        RangeResult with section boundaries
    """
    # Import here to avoid circular dependency
    from gdocs.docs_structure import find_section_by_heading, get_all_headings

    section = find_section_by_heading(doc_data, section_heading, match_case)

    if section is None:
        all_headings = get_all_headings(doc_data)
        heading_list = [h['text'] for h in all_headings[:10]]
        return RangeResult(
            success=False,
            start_index=None,
            end_index=None,
            message=f"Section '{section_heading}' not found. "
                    f"Available headings: {heading_list}" +
                    ("..." if len(all_headings) > 10 else "")
        )

    start_idx = section['start_index']
    end_idx = section['end_index']

    # Adjust start if we don't want the heading
    if not include_heading:
        # Find the end of the heading (first newline after heading start)
        heading_end = section['start_index'] + len(section['heading']) + 1
        # Look for actual end in elements
        if section.get('elements'):
            first_elem = section['elements'][0] if section['elements'] else None
            if first_elem:
                start_idx = first_elem['start_index']
            else:
                start_idx = heading_end
        else:
            start_idx = heading_end

    # Adjust end if we don't want subsections
    if not include_subsections and section.get('subsections'):
        # End at the first subsection
        first_sub = section['subsections'][0]
        end_idx = first_sub['start_index']

    return RangeResult(
        success=True,
        start_index=start_idx,
        end_index=end_idx,
        message=f"Section '{section_heading}' selected: {start_idx}-{end_idx}",
        section_name=section['heading'],
        matched_start=section['heading']
    )


def resolve_range(
    doc_data: Dict[str, Any],
    range_spec: Dict[str, Any]
) -> RangeResult:
    """
    Resolve a range specification to start/end indices.

    This is the main entry point for range resolution, supporting multiple
    specification formats:

    1. Search-based bounds:
       {"start": {"search": "text", "occurrence": 1},
        "end": {"search": "text", "occurrence": 1}}

    2. Search with extension:
       {"search": "keyword", "extend": "paragraph"}

    3. Search with offsets:
       {"search": "keyword", "before_chars": 50, "after_chars": 100}

    4. Section reference:
       {"section": "Heading Name", "include_heading": False}

    Args:
        doc_data: Raw document data from Google Docs API
        range_spec: Range specification dictionary

    Returns:
        RangeResult with resolved indices or error information
    """
    match_case = range_spec.get('match_case', True)

    # Option 1: Search-based bounds (start/end)
    if 'start' in range_spec and 'end' in range_spec:
        start_spec = range_spec['start']
        end_spec = range_spec['end']

        if isinstance(start_spec, dict) and 'search' in start_spec:
            start_search = start_spec['search']
            start_occurrence = start_spec.get('occurrence', 1)
        else:
            return RangeResult(
                success=False,
                start_index=None,
                end_index=None,
                message="Invalid range start specification. "
                        "Expected {'search': 'text', 'occurrence': N}"
            )

        if isinstance(end_spec, dict) and 'search' in end_spec:
            end_search = end_spec['search']
            end_occurrence = end_spec.get('occurrence', 1)
        else:
            return RangeResult(
                success=False,
                start_index=None,
                end_index=None,
                message="Invalid range end specification. "
                        "Expected {'search': 'text', 'occurrence': N}"
            )

        return resolve_range_by_search_bounds(
            doc_data, start_search, end_search,
            start_occurrence, end_occurrence, match_case
        )

    # Option 2: Search with extension to boundary
    if 'search' in range_spec and 'extend' in range_spec:
        return resolve_range_by_search_with_extension(
            doc_data,
            range_spec['search'],
            range_spec['extend'],
            range_spec.get('occurrence', 1),
            match_case
        )

    # Option 3: Search with character offsets
    if 'search' in range_spec and ('before_chars' in range_spec or 'after_chars' in range_spec):
        return resolve_range_by_search_with_offsets(
            doc_data,
            range_spec['search'],
            range_spec.get('before_chars', 0),
            range_spec.get('after_chars', 0),
            range_spec.get('occurrence', 1),
            match_case
        )

    # Option 4: Section reference
    if 'section' in range_spec:
        return resolve_range_by_section(
            doc_data,
            range_spec['section'],
            range_spec.get('include_heading', False),
            range_spec.get('include_subsections', True),
            match_case
        )

    return RangeResult(
        success=False,
        start_index=None,
        end_index=None,
        message="Invalid range specification. Supported formats:\n"
                "1. {start: {search: 'text'}, end: {search: 'text'}}\n"
                "2. {search: 'text', extend: 'paragraph'}\n"
                "3. {search: 'text', before_chars: N, after_chars: N}\n"
                "4. {section: 'Heading Name'}"
    )


def validate_operation(operation: Dict[str, Any]) -> Tuple[bool, str]:
    """
    Validate a batch operation dictionary.
    
    Args:
        operation: Operation dictionary to validate
    
    Returns:
        Tuple of (is_valid, error_message)
    """
    op_type = operation.get('type')
    if not op_type:
        return False, "Missing 'type' field"
    
    # Validate required fields for each operation type
    required_fields = {
        'insert_text': ['index', 'text'],
        'delete_text': ['start_index', 'end_index'],
        'replace_text': ['start_index', 'end_index', 'text'],
        'format_text': ['start_index', 'end_index'],
        'insert_table': ['index', 'rows', 'columns'],
        'insert_page_break': ['index'],
        'find_replace': ['find_text', 'replace_text']
    }
    
    if op_type not in required_fields:
        return False, f"Unsupported operation type: {op_type or 'None'}"
    
    for field in required_fields[op_type]:
        if field not in operation:
            return False, f"Missing required field: {field}"
    
    return True, ""

