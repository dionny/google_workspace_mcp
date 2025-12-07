"""
Google Docs MCP Tools

This module provides MCP tools for interacting with Google Docs API and managing Google Docs via Drive.
"""

import logging
import asyncio
import io
from typing import List, Dict, Any, Literal

from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError

# Auth & server utilities
from auth.service_decorator import require_google_service, require_multiple_services
from core.utils import extract_office_xml_text, handle_http_errors
from core.server import server
from core.comments import create_comment_tools

# Import helper functions for document operations
from gdocs.docs_helpers import (
    create_insert_text_request,
    create_delete_range_request,
    create_format_text_request,
    create_clear_formatting_request,
    create_find_replace_request,
    create_insert_table_request,
    create_insert_page_break_request,
    create_insert_horizontal_rule_requests,
    create_insert_section_break_request,
    create_insert_image_request,
    create_insert_footnote_request,
    create_insert_text_in_footnote_request,
    create_bullet_list_request,
    create_paragraph_style_request,
    create_named_range_request,
    create_delete_named_range_request,
    calculate_search_based_indices,
    find_all_occurrences_in_document,
    find_text_in_document,
    SearchPosition,
    OperationType,
    build_operation_result,
    resolve_range,
    extract_text_at_range,
    interpret_escape_sequences,
    get_character_at_index,
)

# Import document structure and table utilities
from gdocs.docs_structure import (
    parse_document_structure,
    find_tables,
    analyze_document_complexity,
    extract_structural_elements,
    build_headings_outline,
    find_section_by_heading,
    get_all_headings,
    find_section_insertion_point,
    find_elements_by_type,
    get_element_ancestors,
    get_heading_siblings,
    get_paragraph_style_at_index,
    is_heading_style,
    extract_text_in_range,
)
from gdocs.docs_tables import extract_table_as_data

# Import operation managers for complex business logic
from gdocs.managers import (
    TableOperationManager,
    HeaderFooterManager,
    ValidationManager,
    BatchOperationManager,
)
from gdocs.managers.history_manager import get_history_manager, UndoCapability
from gdocs.errors import DocsErrorBuilder, format_error

logger = logging.getLogger(__name__)


@server.tool()
@handle_http_errors("search_docs", is_read_only=True, service_type="docs")
@require_google_service("drive", "drive_read")
async def search_docs(
    service: Any,
    user_google_email: str,
    query: str,
    page_size: int = 10,
) -> str:
    """
    Searches for Google Docs by name using Drive API (mimeType filter).

    Returns:
        str: A formatted list of Google Docs matching the search query.
    """
    logger.info(f"[search_docs] Email={user_google_email}, Query='{query}'")

    escaped_query = query.replace("'", "\\'")

    response = await asyncio.to_thread(
        service.files()
        .list(
            q=f"name contains '{escaped_query}' and mimeType='application/vnd.google-apps.document' and trashed=false",
            pageSize=page_size,
            fields="files(id, name, createdTime, modifiedTime, webViewLink)",
        )
        .execute
    )
    files = response.get("files", [])
    if not files:
        return f"No Google Docs found matching '{query}'."

    output = [f"Found {len(files)} Google Docs matching '{query}':"]
    for f in files:
        output.append(
            f"- {f['name']} (ID: {f['id']}) Modified: {f.get('modifiedTime')} Link: {f.get('webViewLink')}"
        )
    return "\n".join(output)


@server.tool()
@handle_http_errors("get_doc_content", is_read_only=True, service_type="docs")
@require_multiple_services(
    [
        {
            "service_type": "drive",
            "scopes": "drive_read",
            "param_name": "drive_service",
        },
        {"service_type": "docs", "scopes": "docs_read", "param_name": "docs_service"},
    ]
)
async def get_doc_content(
    drive_service: Any,
    docs_service: Any,
    user_google_email: str,
    document_id: str,
) -> str:
    """
    Retrieves content of a Google Doc or a Drive file (like .docx) identified by document_id.
    - Native Google Docs: Fetches content via Docs API (tried first for consistency).
    - Office files (.docx, etc.) stored in Drive: Downloads via Drive API and extracts text.

    Returns:
        str: The document content with metadata header.
    """
    logger.info(
        f"[get_doc_content] Invoked. Document/File ID: '{document_id}' for user '{user_google_email}'"
    )

    # Tab header format constant
    TAB_HEADER_FORMAT = "\n--- TAB: {tab_name} ---\n"

    def extract_text_from_elements(elements, tab_name=None, depth=0):
        """Extract text from document elements (paragraphs, tables, etc.)"""
        # Prevent infinite recursion by limiting depth
        if depth > 5:
            return ""
        text_lines = []
        if tab_name:
            text_lines.append(TAB_HEADER_FORMAT.format(tab_name=tab_name))

        for element in elements:
            if "paragraph" in element:
                paragraph = element.get("paragraph", {})
                para_elements = paragraph.get("elements", [])
                current_line_text = ""
                for pe in para_elements:
                    text_run = pe.get("textRun", {})
                    if text_run and "content" in text_run:
                        current_line_text += text_run["content"]
                if current_line_text.strip():
                    text_lines.append(current_line_text)
            elif "table" in element:
                # Handle table content
                table = element.get("table", {})
                table_rows = table.get("tableRows", [])
                for row in table_rows:
                    row_cells = row.get("tableCells", [])
                    for cell in row_cells:
                        cell_content = cell.get("content", [])
                        cell_text = extract_text_from_elements(
                            cell_content, depth=depth + 1
                        )
                        if cell_text.strip():
                            text_lines.append(cell_text)
        return "".join(text_lines)

    def process_tab_hierarchy(tab, level=0):
        """Process a tab and its nested child tabs recursively"""
        tab_text = ""

        if "documentTab" in tab:
            tab_title = tab.get("documentTab", {}).get("title", "Untitled Tab")
            # Add indentation for nested tabs to show hierarchy
            if level > 0:
                tab_title = "    " * level + tab_title
            tab_body = tab.get("documentTab", {}).get("body", {}).get("content", [])
            tab_text += extract_text_from_elements(tab_body, tab_title)

        # Process child tabs (nested tabs)
        child_tabs = tab.get("childTabs", [])
        for child_tab in child_tabs:
            tab_text += process_tab_hierarchy(child_tab, level + 1)

        return tab_text

    # Try Docs API first for consistency with other gdocs tools
    # This ensures users with Docs API access but not Drive API access can still get content
    try:
        logger.info("[get_doc_content] Trying Docs API first for consistency.")
        doc_data = await asyncio.to_thread(
            docs_service.documents()
            .get(documentId=document_id, includeTabsContent=True)
            .execute
        )

        # Successfully got document via Docs API - it's a native Google Doc
        file_name = doc_data.get("title", "Untitled Document")
        mime_type = "application/vnd.google-apps.document"
        web_view_link = f"https://docs.google.com/document/d/{document_id}/edit"

        logger.info(
            f"[get_doc_content] Successfully retrieved '{file_name}' via Docs API."
        )

        processed_text_lines = []

        # Process main document body
        body_elements = doc_data.get("body", {}).get("content", [])
        main_content = extract_text_from_elements(body_elements)
        if main_content.strip():
            processed_text_lines.append(main_content)

        # Process all tabs
        tabs = doc_data.get("tabs", [])
        for tab in tabs:
            tab_content = process_tab_hierarchy(tab)
            if tab_content.strip():
                processed_text_lines.append(tab_content)

        body_text = "".join(processed_text_lines)

        header = (
            f'File: "{file_name}" (ID: {document_id}, Type: {mime_type})\n'
            f"Link: {web_view_link}\n\n--- CONTENT ---\n"
        )
        return header + body_text

    except HttpError as e:
        # Check if this is a "not a Google Doc" type error (400) vs permission/not found errors
        if e.resp.status == 400:
            # The ID might be a non-Google Doc file (like .docx) - fall back to Drive API
            logger.info(
                "[get_doc_content] Docs API returned 400, trying Drive API for non-Google Doc file."
            )
        elif e.resp.status in (403, 404):
            # Permission denied or not found - re-raise to let the decorator handle it
            logger.warning(f"[get_doc_content] Docs API returned {e.resp.status}: {e}")
            raise
        else:
            # Unexpected error - re-raise
            logger.error(f"[get_doc_content] Unexpected Docs API error: {e}")
            raise

    # Fall back to Drive API for non-Google Docs files (.docx, etc.)
    logger.info("[get_doc_content] Falling back to Drive API for non-Google Doc file.")

    file_metadata = await asyncio.to_thread(
        drive_service.files()
        .get(fileId=document_id, fields="id, name, mimeType, webViewLink")
        .execute
    )
    mime_type = file_metadata.get("mimeType", "")
    file_name = file_metadata.get("name", "Unknown File")
    web_view_link = file_metadata.get("webViewLink", "#")

    logger.info(
        f"[get_doc_content] File '{file_name}' (ID: {document_id}) has mimeType: '{mime_type}'"
    )

    export_mime_type_map = {
        # Example: "application/vnd.google-apps.spreadsheet": "text/csv",
        # Native GSuite types that are not Docs would go here if this function
        # was intended to export them. For .docx, direct download is used.
    }
    effective_export_mime = export_mime_type_map.get(mime_type)

    request_obj = (
        drive_service.files().export_media(
            fileId=document_id, mimeType=effective_export_mime
        )
        if effective_export_mime
        else drive_service.files().get_media(fileId=document_id)
    )

    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request_obj)
    loop = asyncio.get_event_loop()
    done = False
    while not done:
        status, done = await loop.run_in_executor(None, downloader.next_chunk)

    file_content_bytes = fh.getvalue()

    office_text = extract_office_xml_text(file_content_bytes, mime_type)
    if office_text:
        body_text = office_text
    else:
        try:
            body_text = file_content_bytes.decode("utf-8")
        except UnicodeDecodeError:
            body_text = (
                f"[Binary or unsupported text encoding for mimeType '{mime_type}' - "
                f"{len(file_content_bytes)} bytes]"
            )

    header = (
        f'File: "{file_name}" (ID: {document_id}, Type: {mime_type})\n'
        f"Link: {web_view_link}\n\n--- CONTENT ---\n"
    )
    return header + body_text


@server.tool()
@handle_http_errors("list_docs_in_folder", is_read_only=True, service_type="docs")
@require_google_service("drive", "drive_read")
async def list_docs_in_folder(
    service: Any, user_google_email: str, folder_id: str = "root", page_size: int = 100
) -> str:
    """
    Lists Google Docs within a specific Drive folder.

    Returns:
        str: A formatted list of Google Docs in the specified folder.
    """
    logger.info(
        f"[list_docs_in_folder] Invoked. Email: '{user_google_email}', Folder ID: '{folder_id}'"
    )

    rsp = await asyncio.to_thread(
        service.files()
        .list(
            q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.document' and trashed=false",
            pageSize=page_size,
            fields="files(id, name, modifiedTime, webViewLink)",
        )
        .execute
    )
    items = rsp.get("files", [])
    if not items:
        return f"No Google Docs found in folder '{folder_id}'."
    out = [f"Found {len(items)} Docs in folder '{folder_id}':"]
    for f in items:
        out.append(
            f"- {f['name']} (ID: {f['id']}) Modified: {f.get('modifiedTime')} Link: {f.get('webViewLink')}"
        )
    return "\n".join(out)


@server.tool()
@handle_http_errors("list_doc_tabs", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def list_doc_tabs(
    service: Any,
    user_google_email: str,
    document_id: str,
) -> str:
    """
    Lists all tabs in a Google Doc, including nested child tabs.

    Multi-tab documents allow organizing content into separate tabs within a single document.
    This tool retrieves information about all tabs including their IDs (needed for targeting
    specific tabs in editing operations), titles, and hierarchy.

    Args:
        document_id: The ID of the Google Doc

    Returns:
        str: A formatted list of tabs with their IDs, titles, and hierarchy.
             The tab_id values can be used with other tools (modify_doc_text, batch_edit_doc, etc.)
             to target specific tabs for editing.
    """
    logger.info(
        f"[list_doc_tabs] Invoked. Email: '{user_google_email}', Document ID: '{document_id}'"
    )

    # Fetch document with tab content
    doc_data = await asyncio.to_thread(
        service.documents()
        .get(documentId=document_id, includeTabsContent=True)
        .execute
    )

    doc_title = doc_data.get("title", "Untitled Document")
    tabs = doc_data.get("tabs", [])

    def process_tab(tab, level=0):
        """Process a tab and its children recursively, returning tab info."""
        results = []
        indent = "  " * level

        tab_props = tab.get("tabProperties", {})
        tab_id = tab_props.get("tabId", "unknown")
        tab_title = tab_props.get("title", "Untitled Tab")

        # Get content length from documentTab
        doc_tab = tab.get("documentTab", {})
        body = doc_tab.get("body", {})
        content = body.get("content", [])

        # Calculate approximate character count
        char_count = 0
        for element in content:
            if "paragraph" in element:
                for pe in element.get("paragraph", {}).get("elements", []):
                    text_run = pe.get("textRun", {})
                    if text_run and "content" in text_run:
                        char_count += len(text_run["content"])

        results.append({
            "level": level,
            "tab_id": tab_id,
            "title": tab_title,
            "char_count": char_count,
            "indent": indent,
        })

        # Process child tabs
        child_tabs = tab.get("childTabs", [])
        for child_tab in child_tabs:
            results.extend(process_tab(child_tab, level + 1))

        return results

    # Process all tabs
    all_tabs = []
    for tab in tabs:
        all_tabs.extend(process_tab(tab))

    if not all_tabs:
        return f"Document '{doc_title}' (ID: {document_id}) has no tabs (single-tab document without explicit tab structure)."

    # Format output
    out = [
        f"Document: '{doc_title}' (ID: {document_id})",
        f"Total tabs: {len(all_tabs)}",
        "",
        "Tabs:",
    ]

    for tab_info in all_tabs:
        indent = tab_info["indent"]
        hierarchy_marker = "└─ " if tab_info["level"] > 0 else ""
        out.append(
            f"{indent}{hierarchy_marker}'{tab_info['title']}' (tab_id: {tab_info['tab_id']}) - ~{tab_info['char_count']} chars"
        )

    out.append("")
    out.append("TIP: Use the tab_id value with modify_doc_text, batch_edit_doc, or find_and_replace_doc to edit specific tabs.")

    return "\n".join(out)


@server.tool()
@handle_http_errors("create_doc", service_type="docs")
@require_google_service("docs", "docs_write")
async def create_doc(
    service: Any,
    user_google_email: str,
    title: str,
    content: str = "",
) -> str:
    """
    Creates a new Google Doc and optionally inserts initial content.

    Returns:
        str: Confirmation message with document ID and link.
    """
    logger.info(f"[create_doc] Invoked. Email: '{user_google_email}', Title='{title}'")

    doc = await asyncio.to_thread(
        service.documents().create(body={"title": title}).execute
    )
    doc_id = doc.get("documentId")
    if content:
        # Interpret escape sequences in content (e.g., \n -> actual newline)
        content = interpret_escape_sequences(content)
        requests = [{"insertText": {"location": {"index": 1}, "text": content}}]
        await asyncio.to_thread(
            service.documents()
            .batchUpdate(documentId=doc_id, body={"requests": requests})
            .execute
        )
    link = f"https://docs.google.com/document/d/{doc_id}/edit"
    msg = f"Created Google Doc '{title}' (ID: {doc_id}) for {user_google_email}. Link: {link}"
    logger.info(
        f"Successfully created Google Doc '{title}' (ID: {doc_id}) for {user_google_email}. Link: {link}"
    )
    return msg


@server.tool()
@handle_http_errors("modify_doc_text", service_type="docs")
@require_google_service("docs", "docs_write")
async def modify_doc_text(
    service: Any,
    user_google_email: str,
    document_id: str,
    start_index: int = None,
    end_index: int = None,
    text: str = None,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    strikethrough: bool = None,
    small_caps: bool = None,
    subscript: bool = None,
    superscript: bool = None,
    font_size: int = None,
    font_family: str = None,
    link: str = None,
    foreground_color: str = None,
    background_color: str = None,
    line_spacing: float = None,
    heading_style: str = None,
    alignment: str = None,
    indent_first_line: float = None,
    indent_start: float = None,
    indent_end: float = None,
    space_above: float = None,
    space_below: float = None,
    search: str = None,
    position: str = None,
    occurrence: int = 1,
    match_case: bool = True,
    heading: str = None,
    section_position: str = None,
    range: Dict[str, Any] = None,
    location: str = None,
    preview: bool = False,
    convert_to_list: str = None,
    code_block: bool = None,
    delete_paragraph: bool = False,
    tab_id: str = None,
) -> str:
    """
    Modifies text in a Google Doc - can insert/replace text and/or apply formatting.

    Supports five positioning modes:
    1. Location-based: Use location='end' or location='start' for common operations
    2. Index-based: Use start_index/end_index to specify exact positions
    3. Search-based: Use search/position to find text and operate relative to it
    4. Heading-based: Use heading/section_position to target a specific section
    5. Range-based: Use range parameter for semantic text selection

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update

        Location-based positioning (simplest):
        location: Semantic location for insertion. Values:
            - "end": Append to document end (most common operation)
            - "start": Insert at document beginning (after initial section break)

        Index-based positioning (traditional):
        start_index: Start position for operation (0-based)
        end_index: End position for text replacement/formatting

        Search-based positioning (replaces ONE occurrence):
        search: Text to search for in the document
        position: Where to insert relative to search result:
            - "before": Insert before the found text
            - "after": Insert after the found text
            - "replace": Replace the found text (ONE occurrence only!)
                         For replacing ALL occurrences, use find_and_replace_doc instead
        occurrence: Which occurrence to target (1=first, 2=second, -1=last). Default: 1
        match_case: Whether to match case exactly. Default: True

        Heading-based positioning (structural):
        heading: Section heading text to target
        section_position: Where to insert within the section:
            - "start": Insert right after the heading
            - "end": Insert at the end of the section (before next same-level heading)

        Range-based positioning (semantic):
        range: Dictionary specifying a semantic range. Supported formats:

            1. Range by search bounds - select from start text to end text:
               {"start": {"search": "Introduction", "occurrence": 1},
                "end": {"search": "Conclusion", "occurrence": 1}}

            2. Search with boundary extension - find text and extend to boundary:
               {"search": "keyword", "extend": "paragraph"}
               extend options: "paragraph", "sentence", "line", "section"

            3. Search with character offsets - include surrounding context:
               {"search": "keyword", "before_chars": 50, "after_chars": 100}

            4. Section reference - select entire section by heading:
               {"section": "The Velocity Trap", "include_heading": False,
                "include_subsections": True}

            All range formats support "match_case": True/False (default True)

        Text and formatting:
        text: New text to insert or replace with
        bold: Whether to make text bold (True/False/None to leave unchanged)
        italic: Whether to make text italic (True/False/None to leave unchanged)
        underline: Whether to underline text (True/False/None to leave unchanged)
        strikethrough: Whether to strikethrough text (True/False/None to leave unchanged)
        small_caps: Whether to make text small caps (True/False/None to leave unchanged)
        subscript: Whether to make text subscript (True/False/None to leave unchanged, mutually exclusive with superscript)
        superscript: Whether to make text superscript (True/False/None to leave unchanged, mutually exclusive with subscript)
        font_size: Font size in points
        font_family: Font family name (e.g., "Arial", "Times New Roman")
        link: URL to create a hyperlink. Use empty string "" to remove an existing link.
            Supports http://, https://, or # (for internal document bookmarks).
        foreground_color: Text color as hex (#FF0000, #F00) or named color (red, blue, green, etc.)
        background_color: Background/highlight color as hex or named color

        Paragraph formatting:
        line_spacing: Line spacing. Accepts multiple intuitive formats:
            - Named strings: 'single', 'double', '1.5x', '1.15x'
            - Decimal multipliers: 1.0 (single), 1.5 (150%), 2.0 (double) - auto-converted
            - Percentage values: 100 (single), 115 (1.15x default), 150 (1.5x), 200 (double)
            Custom percentage values from 50-1000 are supported.
        heading_style: Change paragraph to a named style. Valid values:
            - "HEADING_1" through "HEADING_6": Heading levels
            - "NORMAL_TEXT": Regular paragraph text
            - "TITLE": Document title style
            - "SUBTITLE": Document subtitle style
            Note: This changes the entire paragraph(s) containing the affected range.
        alignment: Paragraph alignment. Valid values:
            - "START": Left-aligned for left-to-right text
            - "CENTER": Center-aligned
            - "END": Right-aligned for left-to-right text
            - "JUSTIFIED": Justified text (even margins on both sides)
            Note: This changes the entire paragraph(s) containing the affected range.
        indent_first_line: First line indentation in points (72 points = 1 inch).
            Use positive values for standard first line indent.
            Use with negative value and indent_start for hanging indent.
            Example: indent_first_line=36 indents first line by 0.5 inch.
        indent_start: Left margin indentation in points for LTR text (72 points = 1 inch).
            This is the paragraph's left edge offset from the page margin.
            Example: indent_start=72 indents the entire paragraph 1 inch from the left.
        indent_end: Right margin indentation in points for LTR text (72 points = 1 inch).
            This is the paragraph's right edge offset from the page margin.
            Example: indent_end=72 indents the paragraph 1 inch from the right margin.
        space_above: Extra space above the paragraph in points (72 points = 1 inch).
            Common values: 0 (none), 6 (small), 12 (medium), 18 (large)
            Example: space_above=12 adds about 1/6 inch space above the paragraph.
        space_below: Extra space below the paragraph in points (72 points = 1 inch).
            Common values: 0 (none), 6 (small), 12 (medium), 18 (large)
            Example: space_below=12 adds about 1/6 inch space below the paragraph.

        List conversion:
        convert_to_list: Convert the affected text range to a bullet list.
            Values: "UNORDERED" (bullet points) or "ORDERED" (numbered list)
            Can be combined with text operations or used alone with range/search positioning.
            Converts existing paragraphs within the range to list items.

        Code block formatting:
        code_block: Apply code block styling to text. When True, applies:
            - Monospace font (Courier New)
            - Light gray background (#f5f5f5)
            Can be combined with text insertion or used alone for formatting existing text.
            Individual font_family or background_color settings override code_block defaults.

        Preview mode:
        preview: If True, returns what would change without actually modifying the document.
            Useful for validating operations before applying them. Default: False

            Range inspection preview: When preview=True is used with positioning parameters
            (range, search, heading, location, or start_index) but WITHOUT text or formatting,
            the function returns information about what range would be selected without
            requiring an operation. This is useful for:
            - Verifying range/search queries before deciding what to do
            - Debugging range-based selections
            - Previewing section content without modification

        Delete paragraph:
        delete_paragraph: When True and deleting text (text=""), also removes the trailing
            newline to delete the entire paragraph/list item. This prevents orphaned bullet
            markers or empty paragraphs after deletion. Only applies when text="" (delete mode).
            Default: False

        Multi-tab document support:
        tab_id: Optional tab ID for multi-tab documents. If not provided, operates on
            the first tab (default behavior). Use list_doc_tabs() to discover tab IDs
            in a multi-tab document.

    Examples:
        # Append to document end (simplest - location-based):
        modify_doc_text(document_id="...", location="end", text="\\n[New content]")

        # Insert at document beginning:
        modify_doc_text(document_id="...", location="start", text="Prepended text\\n")

        # Insert at end of a specific section (heading-based):
        modify_doc_text(document_id="...", heading="The Problem",
                       section_position="end", text="\\n[New content]")

        # Insert right after a heading:
        modify_doc_text(document_id="...", heading="Introduction",
                       section_position="start", text="\\nUpdated intro text.")

        # Insert after specific text (search-based):
        modify_doc_text(document_id="...", search="Chapter 1", position="after",
                       text=" [Updated]")

        # Replace a SPECIFIC occurrence (first by default):
        # Note: To replace ALL occurrences, use find_and_replace_doc instead
        modify_doc_text(document_id="...", search="old heading", position="replace",
                       text="new heading")

        # Replace the SECOND occurrence specifically:
        modify_doc_text(document_id="...", search="TODO", position="replace",
                       text="DONE", occurrence=2)

        # Replace the LAST occurrence:
        modify_doc_text(document_id="...", search="error", position="replace",
                       text="fixed", occurrence=-1)

        # Traditional index-based (still supported):
        modify_doc_text(document_id="...", start_index=10, end_index=20,
                       text="replacement")

        # Insert text WITH formatting (no end_index needed!):
        modify_doc_text(document_id="...", start_index=100, text="IMPORTANT:",
                       bold=True, italic=True)
        # This inserts "IMPORTANT:" at position 100 and automatically formats it.
        # The formatting range is calculated as [start_index, start_index + len(text)].

        # Add a hyperlink to existing text:
        modify_doc_text(document_id="...", search="click here", position="replace",
                       text="click here", link="https://example.com")

        # Insert new linked text:
        modify_doc_text(document_id="...", location="end",
                       text="Visit our website", link="https://example.com")

        # Remove a hyperlink from text (use empty string):
        modify_doc_text(document_id="...", search="linked text", position="replace",
                       text="linked text", link="")

        # Range-based: Replace everything between two search terms:
        modify_doc_text(document_id="...",
                       range={"start": {"search": "Introduction"},
                              "end": {"search": "Conclusion"}},
                       text="[ENTIRE RANGE REPLACED]")

        # Range-based: Select and format entire paragraph containing keyword:
        modify_doc_text(document_id="...",
                       range={"search": "important keyword", "extend": "paragraph"},
                       bold=True)

        # Range-based: Replace entire section under a heading:
        modify_doc_text(document_id="...",
                       range={"section": "The Velocity Trap", "include_heading": False},
                       text="New section content here.")

        # Range-based: Include surrounding context with offsets:
        modify_doc_text(document_id="...",
                       range={"search": "error", "before_chars": 50, "after_chars": 100},
                       italic=True)

        # Preview mode - see what would change without modifying:
        modify_doc_text(document_id="...", search="old text", position="replace",
                       text="new text", preview=True)
        # Returns preview info with current_content, new_content, context, etc.

        # Range inspection preview - preview what range would be selected:
        modify_doc_text(document_id="...",
                       range={"section": "The Problem", "include_heading": False},
                       preview=True)
        # Returns the content of the section without requiring text/formatting

        # Search inspection - preview what a search finds:
        modify_doc_text(document_id="...", search="important keyword", preview=True)
        # Returns the found text and surrounding context

        # Convert existing paragraphs to bullet list:
        modify_doc_text(document_id="...",
                       range={"search": "- Item one", "extend": "paragraph"},
                       convert_to_list="UNORDERED")

        # Convert range of text to numbered list:
        modify_doc_text(document_id="...", start_index=100, end_index=200,
                       convert_to_list="ORDERED")

        # Insert new text AND convert it to a list:
        modify_doc_text(document_id="...", location="end",
                       text="Item 1\\nItem 2\\nItem 3\\n",
                       convert_to_list="UNORDERED")

        # Set line spacing to double using decimal multiplier:
        modify_doc_text(document_id="...", start_index=100, end_index=200,
                       line_spacing=2.0)  # Also accepts: 'double', 200

        # Set line spacing to 1.5x using named string:
        modify_doc_text(document_id="...",
                       range={"section": "Introduction", "include_heading": True},
                       line_spacing='1.5x')  # Also accepts: 1.5, 150

        # Insert new paragraph with single spacing:
        modify_doc_text(document_id="...", location="end",
                       text="\\nNew paragraph with single spacing.\\n",
                       line_spacing='single')  # Also accepts: 1.0, 100

        # Change a paragraph to Heading 2:
        modify_doc_text(document_id="...", search="Section Title",
                       position="replace", text="Section Title",
                       heading_style="HEADING_2")

        # Convert normal text to a title:
        modify_doc_text(document_id="...", start_index=1, end_index=50,
                       heading_style="TITLE")

        # Change heading back to normal text:
        modify_doc_text(document_id="...",
                       range={"search": "Old Heading", "extend": "paragraph"},
                       heading_style="NORMAL_TEXT")

        # Center-align a paragraph:
        modify_doc_text(document_id="...", search="Title Text",
                       position="replace", text="Title Text",
                       alignment="CENTER")

        # Right-align a range of text:
        modify_doc_text(document_id="...", start_index=100, end_index=200,
                       alignment="END")

        # Justify a section:
        modify_doc_text(document_id="...",
                       range={"section": "Introduction", "include_heading": False},
                       alignment="JUSTIFIED")

        # Combine alignment with heading style:
        modify_doc_text(document_id="...", search="Chapter 1",
                       position="replace", text="Chapter 1",
                       heading_style="HEADING_1", alignment="CENTER")

        # Apply first line indent (0.5 inch):
        modify_doc_text(document_id="...", start_index=100, end_index=200,
                       indent_first_line=36)

        # Create a block quote with left/right margins (1 inch each side):
        modify_doc_text(document_id="...",
                       range={"search": "Quote text", "extend": "paragraph"},
                       indent_start=72, indent_end=72)

        # Add paragraph spacing (12pt above and below):
        modify_doc_text(document_id="...", start_index=100, end_index=200,
                       space_above=12, space_below=12)

        # Create hanging indent (for bibliographies/references):
        modify_doc_text(document_id="...",
                       range={"search": "Reference text", "extend": "paragraph"},
                       indent_start=36, indent_first_line=-36)

        # Insert a code block (monospace font with gray background):
        modify_doc_text(document_id="...", location="end",
                       text="\\ndef hello():\\n    print('Hello, World!')\\n",
                       code_block=True)

        # Apply code block formatting to existing text:
        modify_doc_text(document_id="...", search="console.log", position="replace",
                       text="console.log", code_block=True)

        # Code block with custom background color:
        modify_doc_text(document_id="...", location="end",
                       text="const x = 42;", code_block=True, background_color="#e0e0e0")

    Returns:
        str: JSON string with operation details including position shift information.

        Response includes:
        - success (bool): Whether the operation succeeded
        - operation (str): Type of operation ("insert", "replace", "delete", "format")
        - position_shift (int): How much subsequent positions shifted (+ve = right, -ve = left)
        - affected_range (dict): {"start": x, "end": y} of the modified area
        - message (str): Human-readable description
        - link (str): URL to the document
        - inserted_length (int, optional): Length of inserted text (for insert operations)
        - original_length (int, optional): Original text length (for replace operations)
        - new_length (int, optional): New text length (for replace operations)
        - deleted_length (int, optional): Deleted text length (for delete operations)
        - styles_applied (list, optional): List of styles applied (for format operations)
        - resolved_range (dict, optional): For range-based operations, details about how range was resolved

        Example response for insert:
        {
            "success": true,
            "operation": "insert",
            "position_shift": 13,
            "affected_range": {"start": 100, "end": 113},
            "inserted_length": 13,
            "message": "Inserted 13 characters at index 100",
            "link": "https://docs.google.com/document/d/.../edit"
        }

        Use position_shift for efficient follow-up edits:
        ```python
        result = modify_doc_text(doc_id, start_index=100, text="AAA")
        # result["position_shift"] = 3
        # For next edit originally at 200: use 200 + result["position_shift"] = 203
        ```

        When preview=True, returns a preview response instead:
        {
            "preview": true,
            "would_modify": true,
            "operation": "replace",
            "affected_range": {"start": 100, "end": 108},
            "position_shift": -3,
            "current_content": "old text",
            "new_content": "new text",
            "context": {
                "before": "...text before the affected range...",
                "after": "...text after the affected range..."
            },
            "positioning_info": {"search_text": "old text", ...},
            "message": "Would replace 8 characters with 8 characters at index 100",
            "link": "https://docs.google.com/document/d/.../edit"
        }

        When preview=True with positioning but no text/formatting (range inspection):
        {
            "preview": true,
            "would_modify": false,
            "operation": "range_inspection",
            "affected_range": {"start": 100, "end": 250},
            "position_shift": 0,
            "current_content": "The text content at the resolved range...",
            "content_length": 150,
            "context": {
                "before": "...text before the range...",
                "after": "...text after the range..."
            },
            "positioning_info": {"range": {...}, "resolved_start": 100, ...},
            "message": "Range inspection: resolved to indices 100-250",
            "link": "https://docs.google.com/document/d/.../edit"
        }
    """
    logger.info(
        f"[modify_doc_text] Doc={document_id}, location={location}, search={search}, position={position}, "
        f"heading={heading}, section_position={section_position}, range={range is not None}, "
        f"start={start_index}, end={end_index}, text={text is not None}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Check if this is a preview-only range inspection request
    # When preview=True and positioning is specified but no text/formatting,
    # we return information about what range would be selected
    has_positioning = any([
        range is not None,
        search is not None,
        heading is not None,
        location is not None,
        start_index is not None,
    ])
    is_range_inspection_preview = preview and has_positioning

    # Validate that we have something to do (unless this is a range inspection preview)
    if not is_range_inspection_preview and text is None and not any(
        [
            bold is not None,
            italic is not None,
            underline is not None,
            strikethrough is not None,
            small_caps is not None,
            subscript is not None,
            superscript is not None,
            font_size,
            font_family,
            link is not None,
            foreground_color is not None,
            background_color is not None,
            line_spacing is not None,
            heading_style is not None,
            alignment is not None,
            indent_first_line is not None,
            indent_start is not None,
            indent_end is not None,
            space_above is not None,
            space_below is not None,
            convert_to_list is not None,
            code_block is not None,
        ]
    ):
        error = DocsErrorBuilder.missing_required_param(
            param_name="text or formatting or convert_to_list or code_block",
            context_description="for document modification",
            valid_values=[
                "text",
                "bold",
                "italic",
                "underline",
                "strikethrough",
                "small_caps",
                "subscript",
                "superscript",
                "font_size",
                "font_family",
                "link",
                "foreground_color",
                "background_color",
                "line_spacing",
                "heading_style",
                "alignment",
                "indent_first_line",
                "indent_start",
                "indent_end",
                "space_above",
                "space_below",
                "convert_to_list",
                "code_block",
            ],
        )
        return format_error(error)

    # Interpret escape sequences in text (e.g., \n -> actual newline)
    if text is not None:
        text = interpret_escape_sequences(text)

    # Normalize and validate convert_to_list parameter
    # Accept common aliases for better usability
    if convert_to_list is not None:
        list_type_aliases = {
            "bullet": "UNORDERED",
            "bullets": "UNORDERED",
            "unordered": "UNORDERED",
            "numbered": "ORDERED",
            "numbers": "ORDERED",
            "ordered": "ORDERED",
        }
        normalized = list_type_aliases.get(
            convert_to_list.lower(), convert_to_list.upper()
        )
        if normalized not in ["ORDERED", "UNORDERED"]:
            return validator.create_invalid_param_error(
                param_name="convert_to_list",
                received=convert_to_list,
                valid_values=[
                    "ORDERED",
                    "UNORDERED",
                    "bullet",
                    "numbered",
                ],
            )
        convert_to_list = normalized

    # Validate and normalize line_spacing parameter
    if line_spacing is not None:
        # Accept named string values for common spacings
        if isinstance(line_spacing, str):
            line_spacing_map = {
                'single': 100,
                '1': 100,
                '1.0': 100,
                '1.15': 115,
                '1.15x': 115,
                '1.5': 150,
                '1.5x': 150,
                'double': 200,
                '2': 200,
                '2.0': 200,
            }
            normalized = line_spacing_map.get(line_spacing.lower().strip())
            if normalized is None:
                return validator.create_invalid_param_error(
                    param_name="line_spacing",
                    received=str(line_spacing),
                    valid_values=[
                        "Named: 'single', 'double', '1.5x'",
                        "Decimal: 1.0, 1.5, 2.0 (auto-converted to percentage)",
                        "Percentage: 100, 150, 200 (100=single, 150=1.5x, 200=double)",
                    ],
                )
            line_spacing = normalized
        elif isinstance(line_spacing, (int, float)):
            # Auto-convert decimal multipliers to percentage
            # Values < 10 are treated as multipliers (1.0 -> 100, 1.5 -> 150, 2.0 -> 200)
            if line_spacing < 10:
                line_spacing = line_spacing * 100
            # Now validate the percentage range
            if line_spacing < 50 or line_spacing > 1000:
                return validator.create_invalid_param_error(
                    param_name="line_spacing",
                    received=str(line_spacing),
                    valid_values=[
                        "Named: 'single', 'double', '1.5x'",
                        "Decimal: 1.0, 1.5, 2.0 (auto-converted to percentage)",
                        "Percentage: 100, 150, 200 (100=single, 150=1.5x, 200=double)",
                    ],
                )
        else:
            return validator.create_invalid_param_error(
                param_name="line_spacing",
                received=str(line_spacing),
                valid_values=[
                    "Named: 'single', 'double', '1.5x'",
                    "Decimal: 1.0, 1.5, 2.0 (auto-converted to percentage)",
                    "Percentage: 100, 150, 200 (100=single, 150=1.5x, 200=double)",
                ],
            )

    # Validate heading_style parameter
    valid_heading_styles = [
        "HEADING_1",
        "HEADING_2",
        "HEADING_3",
        "HEADING_4",
        "HEADING_5",
        "HEADING_6",
        "NORMAL_TEXT",
        "TITLE",
        "SUBTITLE",
    ]
    if heading_style is not None and heading_style not in valid_heading_styles:
        return validator.create_invalid_param_error(
            param_name="heading_style",
            received=heading_style,
            valid_values=valid_heading_styles,
        )

    # Validate alignment parameter
    valid_alignments = ["START", "CENTER", "END", "JUSTIFIED"]
    if alignment is not None and alignment not in valid_alignments:
        return validator.create_invalid_param_error(
            param_name="alignment", received=alignment, valid_values=valid_alignments
        )

    # Validate indent_first_line parameter (can be negative for hanging indent)
    if indent_first_line is not None:
        if not isinstance(indent_first_line, (int, float)):
            return validator.create_invalid_param_error(
                param_name="indent_first_line",
                received=str(indent_first_line),
                valid_values=["number in points (72 points = 1 inch)"],
            )

    # Validate indent_start parameter (left margin)
    if indent_start is not None:
        if not isinstance(indent_start, (int, float)) or indent_start < 0:
            return validator.create_invalid_param_error(
                param_name="indent_start",
                received=str(indent_start),
                valid_values=["non-negative number in points (72 points = 1 inch)"],
            )

    # Validate indent_end parameter (right margin)
    if indent_end is not None:
        if not isinstance(indent_end, (int, float)) or indent_end < 0:
            return validator.create_invalid_param_error(
                param_name="indent_end",
                received=str(indent_end),
                valid_values=["non-negative number in points (72 points = 1 inch)"],
            )

    # Validate space_above parameter
    if space_above is not None:
        if not isinstance(space_above, (int, float)) or space_above < 0:
            return validator.create_invalid_param_error(
                param_name="space_above",
                received=str(space_above),
                valid_values=["non-negative number in points (72 points = 1 inch)"],
            )

    # Validate space_below parameter
    if space_below is not None:
        if not isinstance(space_below, (int, float)) or space_below < 0:
            return validator.create_invalid_param_error(
                param_name="space_below",
                received=str(space_below),
                valid_values=["non-negative number in points (72 points = 1 inch)"],
            )

    # Apply code_block formatting defaults (monospace font + light gray background)
    # Only set defaults if user hasn't explicitly specified font_family or background_color
    if code_block is True:
        if font_family is None:
            font_family = "Courier New"
        if background_color is None:
            background_color = "#f5f5f5"

    # Determine positioning mode
    use_range_mode = range is not None
    use_search_mode = search is not None
    use_heading_mode = heading is not None
    use_location_mode = location is not None

    # Validate positioning parameters - range mode takes priority, then location, then heading, then search, then index
    if use_range_mode:
        if (
            location is not None
            or heading is not None
            or search is not None
            or start_index is not None
            or end_index is not None
        ):
            logger.warning(
                "Multiple positioning parameters provided; range mode takes precedence"
            )
    elif use_location_mode:
        if location not in ["start", "end"]:
            return validator.create_invalid_param_error(
                param_name="location", received=location, valid_values=["start", "end"]
            )
        if start_index is not None or end_index is not None:
            error = DocsErrorBuilder.conflicting_params(
                params=["location", "start_index", "end_index"],
                message="Cannot use 'location' parameter with explicit 'start_index' or 'end_index'",
            )
            return format_error(error)
        if heading is not None or search is not None:
            logger.warning(
                "Multiple positioning parameters provided; location mode takes precedence"
            )
    elif use_heading_mode:
        if not section_position:
            return validator.create_missing_param_error(
                param_name="section_position",
                context="when using 'heading'",
                valid_values=["start", "end"],
            )
        if section_position not in ["start", "end"]:
            return validator.create_invalid_param_error(
                param_name="section_position",
                received=section_position,
                valid_values=["start", "end"],
            )
        if search is not None or start_index is not None or end_index is not None:
            logger.warning(
                "Multiple positioning parameters provided; heading mode takes precedence"
            )
    elif use_search_mode:
        # Validate search is not empty
        if search == "":
            return validator.create_empty_search_error()
        if not position:
            return validator.create_missing_param_error(
                param_name="position",
                context="when using 'search'",
                valid_values=["before", "after", "replace"],
            )
        if position not in [p.value for p in SearchPosition]:
            return validator.create_invalid_param_error(
                param_name="position",
                received=position,
                valid_values=["before", "after", "replace"],
            )
        if start_index is not None or end_index is not None:
            logger.warning(
                "Both search and index parameters provided; search mode takes precedence"
            )
    else:
        # Traditional index-based mode
        if start_index is None:
            return validator.create_missing_param_error(
                param_name="positioning",
                context="for document modification",
                valid_values=[
                    "location",
                    "range",
                    "heading+section_position",
                    "search+position",
                    "start_index",
                ],
            )
        # Validate index values (non-negative)
        is_valid, index_error = validator.validate_index_range_structured(
            start_index, end_index
        )
        if not is_valid:
            return index_error

    # Track search results for response
    search_info = {}
    range_result_info = None  # For range-based operations
    location_info = None  # For location-based operations

    # If using location mode, resolve to indices by fetching document
    if use_location_mode:
        # Get document to determine total length
        doc_data = await asyncio.to_thread(
            service.documents().get(documentId=document_id).execute
        )

        structure = parse_document_structure(doc_data)
        total_length = structure["total_length"]

        if location == "end":
            # Use total_length - 1 for safe insertion (last valid index)
            resolved_end_index = total_length - 1 if total_length > 1 else 1
            start_index = resolved_end_index
            location_info = {
                "location": "end",
                "resolved_index": resolved_end_index,
                "message": f"Appending at document end (index {resolved_end_index})",
            }
        else:  # location == 'start'
            start_index = 1  # After the initial section break at index 0
            location_info = {
                "location": "start",
                "resolved_index": 1,
                "message": "Inserting at document start (index 1)",
            }
        end_index = None  # Location mode is always insert, not replace
        search_info = location_info

    # If using range mode, resolve the range to indices
    elif use_range_mode:
        # Get document
        doc_data = await asyncio.to_thread(
            service.documents().get(documentId=document_id).execute
        )

        # Resolve the range specification
        range_result = resolve_range(doc_data, range)

        if not range_result.success:
            # Return structured error with range resolution details
            import json

            error_response = {
                "success": False,
                "error": "range_resolution_failed",
                "message": range_result.message,
                "hint": "Check range specification format and search terms",
            }
            return json.dumps(error_response, indent=2)

        start_index = range_result.start_index
        end_index = range_result.end_index
        range_result_info = range_result.to_dict()
        search_info = {
            "range": range,
            "resolved_start": start_index,
            "resolved_end": end_index,
            "message": range_result.message,
        }
        if range_result.matched_start:
            search_info["matched_start"] = range_result.matched_start
        if range_result.matched_end:
            search_info["matched_end"] = range_result.matched_end
        if range_result.extend_type:
            search_info["extend_type"] = range_result.extend_type
        if range_result.section_name:
            search_info["section_name"] = range_result.section_name

    # If using heading mode, find the section and calculate insertion point
    elif use_heading_mode:
        # Get document
        doc_data = await asyncio.to_thread(
            service.documents().get(documentId=document_id).execute
        )

        # Find the insertion point using section navigation
        insertion_index = find_section_insertion_point(
            doc_data, heading, section_position, match_case
        )

        if insertion_index is None:
            # Provide helpful error with available headings
            all_headings = get_all_headings(doc_data)
            heading_list = [h["text"] for h in all_headings] if all_headings else []
            return validator.create_heading_not_found_error(
                heading=heading, available_headings=heading_list, match_case=match_case
            )

        start_index = insertion_index
        end_index = None  # Heading mode is always insert, not replace
        search_info = {
            "heading": heading,
            "section_position": section_position,
            "insertion_index": insertion_index,
            "message": f"Found section '{heading}', inserting at {section_position}",
        }

    # If using search mode, find the text and calculate indices
    elif use_search_mode:
        # Get document to search
        doc_data = await asyncio.to_thread(
            service.documents().get(documentId=document_id).execute
        )

        success, calc_start, calc_end, message = calculate_search_based_indices(
            doc_data, search, position, occurrence, match_case
        )

        if not success:
            # Provide helpful error with occurrence info
            all_occurrences = find_all_occurrences_in_document(
                doc_data, search, match_case
            )
            if all_occurrences:
                # Check if it's an invalid occurrence error
                if "occurrence" in message.lower():
                    return validator.create_invalid_occurrence_error(
                        occurrence=occurrence,
                        total_found=len(all_occurrences),
                        search_text=search,
                    )
                # Multiple occurrences exist but specified one not found
                occurrences_data = [
                    {"index": i + 1, "position": f"{s}-{e}"}
                    for i, (s, e) in enumerate(all_occurrences[:5])
                ]
                return validator.create_ambiguous_search_error(
                    search_text=search,
                    occurrences=occurrences_data,
                    total_count=len(all_occurrences),
                )
            # Text not found at all
            return validator.create_search_not_found_error(
                search_text=search, match_case=match_case
            )

        start_index = calc_start
        end_index = calc_end
        search_info = {
            "search_text": search,
            "position": position,
            "occurrence": occurrence,
            "found_at_index": calc_start,
            "message": message,
        }

        # For "before" and "after" positions, we're just inserting, not replacing
        if position in [SearchPosition.BEFORE.value, SearchPosition.AFTER.value]:
            # Set end_index to None to trigger insert mode rather than replace
            end_index = None

    # Validate text formatting params if provided
    has_formatting = any(
        [
            bold is not None,
            italic is not None,
            underline is not None,
            strikethrough is not None,
            small_caps is not None,
            subscript is not None,
            superscript is not None,
            font_size is not None,
            font_family is not None,
            link is not None,
            foreground_color is not None,
            background_color is not None,
        ]
    )
    formatting_params_list = []
    if bold is not None:
        formatting_params_list.append("bold")
    if italic is not None:
        formatting_params_list.append("italic")
    if underline is not None:
        formatting_params_list.append("underline")
    if strikethrough is not None:
        formatting_params_list.append("strikethrough")
    if small_caps is not None:
        formatting_params_list.append("small_caps")
    if subscript is not None:
        formatting_params_list.append("subscript")
    if superscript is not None:
        formatting_params_list.append("superscript")
    if font_size is not None:
        formatting_params_list.append("font_size")
    if font_family is not None:
        formatting_params_list.append("font_family")
    if link is not None:
        formatting_params_list.append("link")
    if foreground_color is not None:
        formatting_params_list.append("foreground_color")
    if background_color is not None:
        formatting_params_list.append("background_color")
    if code_block is True:
        formatting_params_list.append("code_block")

    if has_formatting:
        is_valid, error_msg = validator.validate_text_formatting_params(
            bold,
            italic,
            underline,
            strikethrough,
            small_caps,
            subscript,
            superscript,
            font_size,
            font_family,
            link,
            foreground_color,
            background_color,
        )
        if not is_valid:
            # Check if this is a color format error and return specific structured error
            if "Invalid color format" in error_msg or "Invalid hex color" in error_msg:
                # Extract color value and param name from error
                if "foreground_color" in error_msg:
                    return validator.create_invalid_color_error(
                        foreground_color, "foreground_color"
                    )
                elif "background_color" in error_msg:
                    return validator.create_invalid_color_error(
                        background_color, "background_color"
                    )
            return validator.create_invalid_param_error(
                param_name="formatting",
                received=str(formatting_params_list),
                valid_values=[
                    "bold (bool)",
                    "italic (bool)",
                    "underline (bool)",
                    "strikethrough (bool)",
                    "small_caps (bool)",
                    "subscript (bool)",
                    "superscript (bool)",
                    "font_size (1-400)",
                    "font_family (string)",
                    "link (URL string)",
                    "foreground_color (hex/#FF0000 or named)",
                    "background_color (hex or named)",
                ],
                context=error_msg,
            )

        # For formatting without text insertion, we need end_index
        # But if text is provided, we can calculate end_index automatically
        if end_index is None and text is None:
            is_valid, structured_error = validator.validate_formatting_range_structured(
                start_index=start_index,
                end_index=end_index,
                text=text,
                formatting_params=formatting_params_list,
            )
            if not is_valid:
                return structured_error

        if end_index is not None:
            is_valid, structured_error = validator.validate_index_range_structured(
                start_index, end_index
            )
            if not is_valid:
                return structured_error

    # Auto-detect heading style inheritance prevention
    # When inserting text (not replacing) without an explicit heading_style,
    # check if the insertion point is in a heading paragraph and auto-apply NORMAL_TEXT
    # to prevent the new text from inheriting the heading style
    auto_normal_text_applied = False
    if text is not None and end_index is None and heading_style is None:
        # We need doc_data to detect the paragraph style at the insertion point
        # It's already fetched for location, range, heading, and search modes
        # For index-based mode, we need to fetch it
        doc_data_for_style_check = locals().get("doc_data")

        if doc_data_for_style_check is None:
            doc_data_for_style_check = await asyncio.to_thread(
                service.documents().get(documentId=document_id).execute
            )
            # Store for later use (e.g., preview mode)
            doc_data = doc_data_for_style_check

        # Check the paragraph style at the insertion point
        if doc_data_for_style_check:
            current_style = get_paragraph_style_at_index(
                doc_data_for_style_check, start_index
            )
            if is_heading_style(current_style):
                # Auto-apply NORMAL_TEXT to prevent heading style inheritance
                heading_style = "NORMAL_TEXT"
                auto_normal_text_applied = True
                logger.info(
                    f"[modify_doc_text] Auto-applying NORMAL_TEXT style to prevent inheritance from {current_style}"
                )

    requests = []
    operations = []

    # Track operation details for response
    operation_type = None
    actual_start_index = start_index
    actual_end_index = end_index
    format_styles = []

    # For delete_paragraph, we need doc_data to check if next char is newline
    # Fetch it if we're doing a delete operation and don't have it yet
    if (
        delete_paragraph
        and text == ""
        and end_index is not None
        and end_index > start_index
    ):
        if doc_data is None:
            doc_data = await asyncio.to_thread(
                service.documents().get(documentId=document_id).execute
            )

    # Handle text insertion/replacement/deletion
    if text is not None:
        if end_index is not None and end_index > start_index:
            if text == "":
                # Empty text with range = delete operation (no insert needed)
                operation_type = OperationType.DELETE

                # Handle delete_paragraph: extend deletion to include trailing newline
                # This removes the entire paragraph/list item, preventing orphaned bullets
                delete_end = end_index
                paragraph_deleted = False
                if delete_paragraph and doc_data is not None:
                    char_at_end = get_character_at_index(doc_data, end_index)
                    if char_at_end == "\n":
                        delete_end = end_index + 1
                        paragraph_deleted = True
                        logger.info(
                            f"[modify_doc_text] delete_paragraph: extending deletion to include trailing newline at index {end_index}"
                        )

                if start_index == 0:
                    # Cannot delete at index 0 (first section break), start from 1
                    actual_start_index = 1
                    requests.append(create_delete_range_request(1, delete_end, tab_id=tab_id))
                    op_msg = f"Deleted text from index 1 to {delete_end}"
                    if paragraph_deleted:
                        op_msg += " (including paragraph break)"
                    operations.append(op_msg)
                else:
                    requests.append(
                        create_delete_range_request(start_index, delete_end, tab_id=tab_id)
                    )
                    op_msg = f"Deleted text from index {start_index} to {delete_end}"
                    if paragraph_deleted:
                        op_msg += " (including paragraph break)"
                    operations.append(op_msg)

                actual_end_index = delete_end
            else:
                # Text replacement
                operation_type = OperationType.REPLACE
                if start_index == 0:
                    # Special case: Cannot delete at index 0 (first section break)
                    # Instead, we insert new text at index 1 and then delete the old text
                    actual_start_index = 1
                    requests.append(create_insert_text_request(1, text, tab_id=tab_id))
                    adjusted_end = end_index + len(text)
                    requests.append(
                        create_delete_range_request(1 + len(text), adjusted_end, tab_id=tab_id)
                    )
                    operations.append(
                        f"Replaced text from index {start_index} to {end_index}"
                    )
                else:
                    # Normal replacement: delete old text, then insert new text
                    requests.extend(
                        [
                            create_delete_range_request(start_index, end_index, tab_id=tab_id),
                            create_insert_text_request(start_index, text, tab_id=tab_id),
                        ]
                    )
                    operations.append(
                        f"Replaced text from index {start_index} to {end_index}"
                    )
        else:
            # Text insertion - validate that text is not empty
            if text == "":
                # Empty text insertion is invalid - return structured error
                error = DocsErrorBuilder.empty_text_insertion()
                return format_error(error)
            operation_type = OperationType.INSERT
            actual_start_index = 1 if start_index == 0 else start_index
            actual_end_index = None  # Insert has no end_index
            requests.append(create_insert_text_request(actual_start_index, text, tab_id=tab_id))
            operations.append(f"Inserted text at index {actual_start_index}")
            search_info["inserted_at_index"] = actual_start_index

    # Clear formatting for plain text insertions (no formatting specified).
    # This ensures inserted text doesn't inherit surrounding formatting.
    if text is not None and operation_type == OperationType.INSERT and not has_formatting:
        format_start = 1 if start_index == 0 else start_index
        format_end = format_start + len(text)
        requests.append(
            create_clear_formatting_request(format_start, format_end, preserve_links=False, tab_id=tab_id)
        )
        operations.append(f"Cleared inherited formatting from range {format_start}-{format_end}")

    # Handle formatting
    if has_formatting:
        # Adjust range for formatting based on text operations
        format_start = start_index
        format_end = end_index

        if text is not None:
            if end_index is not None and end_index > start_index:
                # Text was replaced - format the new text
                format_end = start_index + len(text)
            else:
                # Text was inserted - format the inserted text
                actual_index = 1 if start_index == 0 else start_index
                format_start = actual_index
                format_end = actual_index + len(text)

        # Handle special case for formatting at index 0
        if format_start == 0:
            format_start = 1
        if format_end is not None and format_end <= format_start:
            format_end = format_start + 1

        # Clear existing formatting before applying new styles to inserted text.
        # This prevents the inserted text from inheriting formatting from surrounding text
        # (e.g., if inserting after bold text, new text would otherwise inherit bold).
        # We clear formatting when:
        # 1. Text was inserted (not replaced) AND has formatting to apply
        # 2. code_block=True (regardless of insert/replace)
        should_clear_formatting = (
            (text is not None and operation_type == OperationType.INSERT) or
            code_block is True
        )
        if should_clear_formatting:
            requests.append(
                create_clear_formatting_request(
                    format_start, format_end, preserve_links=(link is not None), tab_id=tab_id
                )
            )

        requests.append(
            create_format_text_request(
                format_start,
                format_end,
                bold,
                italic,
                underline,
                strikethrough,
                small_caps,
                subscript,
                superscript,
                font_size,
                font_family,
                link,
                foreground_color,
                background_color,
                tab_id=tab_id,
            )
        )

        format_details = []
        if bold is not None:
            format_details.append("bold" if bold else "remove bold")
        if italic is not None:
            format_details.append("italic" if italic else "remove italic")
        if underline is not None:
            format_details.append("underline" if underline else "remove underline")
        if strikethrough is not None:
            format_details.append(
                "strikethrough" if strikethrough else "remove strikethrough"
            )
        if small_caps is not None:
            format_details.append("small_caps" if small_caps else "remove small_caps")
        if subscript is not None:
            format_details.append("subscript" if subscript else "remove subscript")
        if superscript is not None:
            format_details.append(
                "superscript" if superscript else "remove superscript"
            )
        if font_size is not None:
            format_details.append(f"font_size={font_size}")
        if font_family is not None:
            format_details.append(f"font_family={font_family}")
        if link is not None:
            format_details.append("link" if link else "remove link")
        if foreground_color is not None:
            format_details.append("foreground_color")
        if background_color is not None:
            format_details.append("background_color")
        if code_block is True:
            format_details.append("code_block")

        format_styles = format_details
        operations.append(
            f"Applied formatting ({', '.join(format_details)}) to range {format_start}-{format_end}"
        )

        # If only formatting (no text operation), set operation type
        if operation_type is None:
            operation_type = OperationType.FORMAT
            actual_start_index = format_start
            actual_end_index = format_end

    # Handle paragraph formatting (line spacing, heading style, alignment, indentation, and spacing)
    has_paragraph_formatting = any([
        line_spacing is not None,
        heading_style is not None,
        alignment is not None,
        indent_first_line is not None,
        indent_start is not None,
        indent_end is not None,
        space_above is not None,
        space_below is not None,
    ])
    if has_paragraph_formatting:
        # Determine the range for paragraph formatting
        para_start = (
            actual_start_index if actual_start_index is not None else start_index
        )
        para_end = actual_end_index if actual_end_index is not None else end_index

        # For text insertion, the paragraph range is the newly inserted text
        if text is not None and operation_type == OperationType.INSERT:
            para_start = actual_start_index
            para_end = actual_start_index + len(text)
        elif text is not None and operation_type == OperationType.REPLACE:
            para_start = actual_start_index
            para_end = actual_start_index + len(text)

        # When applying heading_style to newly inserted multi-paragraph text,
        # behavior depends on whether this is auto-applied NORMAL_TEXT to prevent
        # heading inheritance, or user-specified heading style:
        # - auto_normal_text_applied=True: Apply NORMAL_TEXT to ALL paragraphs
        #   to prevent the heading style from bleeding into subsequent paragraphs
        # - User-specified heading: Only apply to FIRST paragraph
        # line_spacing and alignment should still apply to all paragraphs.
        heading_style_start = para_start
        heading_style_end = para_end

        if heading_style is not None and text is not None:
            # Strip leading newlines - heading style should start after them
            text_stripped = text.lstrip("\n")
            leading_newlines = len(text) - len(text_stripped)
            if leading_newlines > 0:
                heading_style_start = para_start + leading_newlines

            # When auto_normal_text_applied is True, we need to apply NORMAL_TEXT
            # to ALL paragraphs, not just the first one, to prevent heading
            # style inheritance from bleeding through to subsequent paragraphs.
            if not auto_normal_text_applied:
                # User-specified heading: only apply to first paragraph
                first_newline_pos = text_stripped.find("\n")
                if first_newline_pos != -1:
                    # Multi-paragraph text: only apply heading to first paragraph
                    heading_style_end = heading_style_start + first_newline_pos
                else:
                    # Single paragraph text: strip trailing newlines to prevent style bleed
                    trailing_newlines = len(text_stripped) - len(text_stripped.rstrip("\n"))
                    if trailing_newlines > 0:
                        heading_style_end = para_start + len(text) - trailing_newlines
            # else: auto_normal_text_applied - keep heading_style_end = para_end
            # to apply NORMAL_TEXT to all paragraphs

        # Handle special case for paragraph formatting at index 0
        if para_start == 0:
            para_start = 1
        if heading_style_start == 0:
            heading_style_start = 1

        # Validate we have a valid range for paragraph formatting
        if para_start is not None and para_end is not None and para_end > para_start:
            para_format_details = []

            # Apply line_spacing, alignment, indentation, and spacing to full range
            has_general_para_formatting = any([
                line_spacing is not None,
                alignment is not None,
                indent_first_line is not None,
                indent_start is not None,
                indent_end is not None,
                space_above is not None,
                space_below is not None,
            ])
            if has_general_para_formatting:
                general_para_request = create_paragraph_style_request(
                    para_start, para_end, line_spacing, None, alignment,
                    indent_first_line, indent_start, indent_end,
                    space_above, space_below, tab_id=tab_id
                )
                if general_para_request:
                    requests.append(general_para_request)
                    if line_spacing is not None:
                        para_format_details.append(f"line_spacing={line_spacing}%")
                        format_styles.append(f"line_spacing_{line_spacing}")
                    if alignment is not None:
                        para_format_details.append(f"alignment={alignment}")
                        format_styles.append(f"alignment_{alignment}")
                    if indent_first_line is not None:
                        para_format_details.append(f"indent_first_line={indent_first_line}pt")
                        format_styles.append(f"indent_first_line_{indent_first_line}")
                    if indent_start is not None:
                        para_format_details.append(f"indent_start={indent_start}pt")
                        format_styles.append(f"indent_start_{indent_start}")
                    if indent_end is not None:
                        para_format_details.append(f"indent_end={indent_end}pt")
                        format_styles.append(f"indent_end_{indent_end}")
                    if space_above is not None:
                        para_format_details.append(f"space_above={space_above}pt")
                        format_styles.append(f"space_above_{space_above}")
                    if space_below is not None:
                        para_format_details.append(f"space_below={space_below}pt")
                        format_styles.append(f"space_below_{space_below}")
                    operations.append(
                        f"Applied paragraph formatting ({', '.join(para_format_details)}) to range {para_start}-{para_end}"
                    )

            # Apply heading_style - either to first paragraph only (user-specified)
            # or to all paragraphs (auto-applied NORMAL_TEXT to prevent inheritance)
            if heading_style is not None and heading_style_end > heading_style_start:
                heading_request = create_paragraph_style_request(
                    heading_style_start, heading_style_end, None, heading_style, None,
                    tab_id=tab_id
                )
                if heading_request:
                    requests.append(heading_request)
                    format_styles.append(f"heading_style_{heading_style}")
                    if auto_normal_text_applied:
                        operations.append(
                            f"Applied heading_style={heading_style} to all paragraphs to prevent heading inheritance (range {heading_style_start}-{heading_style_end})"
                        )
                    else:
                        operations.append(
                            f"Applied heading_style={heading_style} to first paragraph (range {heading_style_start}-{heading_style_end})"
                        )

            # If only paragraph formatting (no text or text style operation), set operation type
            if operation_type is None:
                operation_type = OperationType.FORMAT
                actual_start_index = para_start
                actual_end_index = para_end
        elif para_start is None or para_end is None or para_end <= para_start:
            # Need a valid range for paragraph formatting when not inserting text
            if text is None:
                # Determine which paragraph formatting parameter was used for error message
                para_type = "paragraph formatting"
                if line_spacing is not None:
                    para_type = "line_spacing"
                elif heading_style is not None:
                    para_type = "heading_style"
                elif alignment is not None:
                    para_type = "alignment"
                elif indent_first_line is not None:
                    para_type = "indent_first_line"
                elif indent_start is not None:
                    para_type = "indent_start"
                elif indent_end is not None:
                    para_type = "indent_end"
                elif space_above is not None:
                    para_type = "space_above"
                elif space_below is not None:
                    para_type = "space_below"
                error = DocsErrorBuilder.missing_required_param(
                    param_name="end_index or range",
                    context_description=f"for {para_type} (need a range of text to format)",
                    valid_values=[
                        "start_index + end_index",
                        "range parameter",
                        "text (for new text insertion)",
                    ],
                )
                return format_error(error)

    # Handle list conversion
    if convert_to_list is not None:
        # Determine the range for list conversion
        list_start = (
            actual_start_index if actual_start_index is not None else start_index
        )
        list_end = actual_end_index if actual_end_index is not None else end_index

        # For text insertion, the list range is the newly inserted text
        if text is not None and operation_type == OperationType.INSERT:
            list_start = actual_start_index
            list_end = actual_start_index + len(text)
        elif text is not None and operation_type == OperationType.REPLACE:
            list_start = actual_start_index
            list_end = actual_start_index + len(text)

        # Validate we have a valid range for list conversion
        if list_start is not None and list_end is not None and list_end > list_start:
            requests.append(
                create_bullet_list_request(list_start, list_end, convert_to_list, tab_id=tab_id)
            )
            list_type_display = (
                "bullet" if convert_to_list == "UNORDERED" else "numbered"
            )
            operations.append(
                f"Converted to {list_type_display} list in range {list_start}-{list_end}"
            )
            format_styles.append(f"convert_to_{list_type_display}_list")

            # If only list conversion (no text or formatting operation), set operation type
            if operation_type is None:
                operation_type = OperationType.FORMAT
                actual_start_index = list_start
                actual_end_index = list_end
        elif list_start is None or list_end is None or list_end <= list_start:
            # Need a valid range for list conversion when not inserting text
            if text is None:
                error = DocsErrorBuilder.missing_required_param(
                    param_name="end_index or range",
                    context_description="for convert_to_list (need a range of text to convert)",
                    valid_values=[
                        "start_index + end_index",
                        "range parameter",
                        "text (for new text insertion)",
                    ],
                )
                return format_error(error)

    # Handle preview mode - return what would change without actually modifying
    import json

    if preview:
        # For preview, we need doc_data to extract current content
        # If we don't have it from positioning resolution, fetch it now
        if not any(
            [use_location_mode, use_range_mode, use_heading_mode, use_search_mode]
        ):
            # Index-based mode - need to fetch document for preview
            doc_data = await asyncio.to_thread(
                service.documents().get(documentId=document_id).execute
            )

        # Calculate what would change
        preview_result = {
            "preview": True,
            "would_modify": True,
            "operation": operation_type.value if operation_type else "unknown",
            "affected_range": {
                "start": actual_start_index,
                "end": actual_end_index if actual_end_index else actual_start_index,
            },
            "positioning_info": search_info if search_info else {},
            "link": f"https://docs.google.com/document/d/{document_id}/edit",
        }

        # Calculate position shift
        text_length = len(text) if text else 0
        if operation_type == OperationType.INSERT:
            preview_result["position_shift"] = text_length
            preview_result["new_content"] = text
            msg = f"Would insert {text_length} characters at index {actual_start_index}"
            if convert_to_list:
                preview_result["convert_to_list"] = convert_to_list
                list_type_display = (
                    "bullet" if convert_to_list == "UNORDERED" else "numbered"
                )
                msg += f" and convert to {list_type_display} list"
            if line_spacing is not None:
                preview_result["line_spacing"] = line_spacing
                msg += f" and set line spacing to {line_spacing}%"
            if heading_style is not None:
                preview_result["heading_style"] = heading_style
                msg += f" and set heading style to {heading_style}"
            preview_result["message"] = msg
        elif operation_type == OperationType.REPLACE:
            original_length = (
                actual_end_index or actual_start_index
            ) - actual_start_index
            preview_result["position_shift"] = text_length - original_length
            preview_result["original_length"] = original_length
            preview_result["new_length"] = text_length
            preview_result["new_content"] = text

            # Extract current content at the range
            if doc_data:
                current = extract_text_at_range(
                    doc_data, actual_start_index, actual_end_index or actual_start_index
                )
                preview_result["current_content"] = current.get("text", "")
                preview_result["context"] = {
                    "before": current.get("context_before", ""),
                    "after": current.get("context_after", ""),
                }
            msg = f"Would replace {original_length} characters with {text_length} characters at index {actual_start_index}"
            if convert_to_list:
                preview_result["convert_to_list"] = convert_to_list
                list_type_display = (
                    "bullet" if convert_to_list == "UNORDERED" else "numbered"
                )
                msg += f" and convert to {list_type_display} list"
            if line_spacing is not None:
                preview_result["line_spacing"] = line_spacing
                msg += f" and set line spacing to {line_spacing}%"
            if heading_style is not None:
                preview_result["heading_style"] = heading_style
                msg += f" and set heading style to {heading_style}"
            preview_result["message"] = msg
        elif operation_type == OperationType.DELETE:
            deleted_length = (
                actual_end_index or actual_start_index
            ) - actual_start_index
            preview_result["position_shift"] = -deleted_length
            preview_result["deleted_length"] = deleted_length

            # Extract current content at the delete range
            if doc_data:
                current = extract_text_at_range(
                    doc_data, actual_start_index, actual_end_index or actual_start_index
                )
                preview_result["current_content"] = current.get("text", "")
                preview_result["context"] = {
                    "before": current.get("context_before", ""),
                    "after": current.get("context_after", ""),
                }
            preview_result["message"] = (
                f"Would delete {deleted_length} characters from index {actual_start_index} to {actual_end_index}"
            )
        elif operation_type == OperationType.FORMAT:
            preview_result["position_shift"] = 0
            preview_result["styles_to_apply"] = format_styles
            if convert_to_list:
                preview_result["convert_to_list"] = convert_to_list
            if line_spacing is not None:
                preview_result["line_spacing"] = line_spacing
            if heading_style is not None:
                preview_result["heading_style"] = heading_style

            # Extract current content at the format range
            if doc_data:
                format_end_idx = actual_end_index or actual_start_index
                current = extract_text_at_range(
                    doc_data, actual_start_index, format_end_idx
                )
                preview_result["current_content"] = current.get("text", "")
                preview_result["context"] = {
                    "before": current.get("context_before", ""),
                    "after": current.get("context_after", ""),
                }
            preview_result["message"] = (
                f"Would apply formatting ({', '.join(format_styles)}) to range {actual_start_index}-{actual_end_index}"
            )
        elif is_range_inspection_preview:
            # Range inspection preview - no operation, just showing what range would be selected
            preview_result["would_modify"] = False
            preview_result["operation"] = "range_inspection"
            preview_result["position_shift"] = 0

            # Extract content at the resolved range
            if doc_data:
                range_end_idx = actual_end_index if actual_end_index else actual_start_index
                if range_end_idx > actual_start_index:
                    current = extract_text_at_range(
                        doc_data, actual_start_index, range_end_idx
                    )
                    preview_result["current_content"] = current.get("text", "")
                    preview_result["content_length"] = len(current.get("text", ""))
                    preview_result["context"] = {
                        "before": current.get("context_before", ""),
                        "after": current.get("context_after", ""),
                    }
                else:
                    # Insertion point, no range selected
                    current = extract_text_at_range(
                        doc_data, max(1, actual_start_index - 25), min(actual_start_index + 25, actual_start_index + 50)
                    )
                    preview_result["current_content"] = ""
                    preview_result["content_length"] = 0
                    preview_result["context"] = {
                        "before": current.get("context_before", ""),
                        "after": current.get("context_after", ""),
                    }
                    preview_result["note"] = "This is an insertion point, not a range selection"

            # Build informative message based on positioning mode
            if use_range_mode:
                preview_result["message"] = (
                    f"Range inspection: resolved to indices {actual_start_index}-{actual_end_index}"
                )
            elif use_search_mode:
                preview_result["message"] = (
                    f"Search '{search}' found at indices {actual_start_index}-{actual_end_index}"
                )
            elif use_heading_mode:
                preview_result["message"] = (
                    f"Heading '{heading}' section {section_position}: insertion point at index {actual_start_index}"
                )
            elif use_location_mode:
                preview_result["message"] = (
                    f"Location '{location}': insertion point at index {actual_start_index}"
                )
            else:
                preview_result["message"] = (
                    f"Index range: {actual_start_index}-{actual_end_index if actual_end_index else actual_start_index}"
                )
        else:
            preview_result["position_shift"] = 0
            preview_result["message"] = "Would perform operation"

        # Add resolved range info for range-based operations
        if range_result_info:
            preview_result["resolved_range"] = range_result_info

        return json.dumps(preview_result, indent=2)

    # Capture text before operation for undo support (delete/replace operations)
    deleted_text_for_undo = None
    original_text_for_undo = None
    if (
        operation_type in [OperationType.DELETE, OperationType.REPLACE]
        and actual_end_index
    ):
        # Ensure we have doc_data for text capture
        try:
            if "doc_data" not in dir() or doc_data is None:
                doc_data = await asyncio.to_thread(
                    service.documents().get(documentId=document_id).execute
                )
            extracted = extract_text_at_range(
                doc_data, actual_start_index, actual_end_index
            )
            captured_text = extracted.get("text", "")
            if operation_type == OperationType.DELETE:
                deleted_text_for_undo = captured_text
            else:
                original_text_for_undo = captured_text
        except Exception as e:
            logger.warning(f"Failed to capture text for undo: {e}")

    await asyncio.to_thread(
        service.documents()
        .batchUpdate(documentId=document_id, body={"requests": requests})
        .execute
    )

    # Record operation for undo history (automatic tracking)
    try:
        history_manager = get_history_manager()
        # Map OperationType to string for history manager
        op_type_map = {
            OperationType.INSERT: "insert_text",
            OperationType.DELETE: "delete_text",
            OperationType.REPLACE: "replace_text",
            OperationType.FORMAT: "format_text",
        }
        history_op_type = op_type_map.get(operation_type, "unknown")

        # Calculate position shift
        position_shift = 0
        if operation_type == OperationType.INSERT:
            position_shift = len(text) if text else 0
        elif operation_type == OperationType.DELETE:
            position_shift = (
                -(actual_end_index - actual_start_index) if actual_end_index else 0
            )
        elif operation_type == OperationType.REPLACE:
            old_len = (actual_end_index - actual_start_index) if actual_end_index else 0
            new_len = len(text) if text else 0
            position_shift = new_len - old_len

        # Determine undo capability
        undo_capability = UndoCapability.FULL
        undo_notes = None
        if operation_type == OperationType.FORMAT:
            undo_capability = UndoCapability.NONE
            undo_notes = (
                "Format undo requires capturing original formatting (not yet supported)"
            )

        history_manager.record_operation(
            document_id=document_id,
            operation_type=history_op_type,
            operation_params={
                "start_index": actual_start_index,
                "end_index": actual_end_index,
                "text": text,
            },
            start_index=actual_start_index,
            end_index=actual_end_index,
            position_shift=position_shift,
            deleted_text=deleted_text_for_undo,
            original_text=original_text_for_undo,
            undo_capability=undo_capability,
            undo_notes=undo_notes,
        )
        logger.debug(
            f"Recorded operation for undo: {history_op_type} at {actual_start_index}"
        )
    except Exception as e:
        logger.warning(f"Failed to record operation for undo history: {e}")

    # Build structured operation result

    # Determine final operation type for combined operations
    if text is not None and has_formatting:
        # Combined insert/replace + format - report the text operation as primary
        pass  # operation_type already set from text operation

    # Build the operation result
    op_result = build_operation_result(
        operation_type=operation_type,
        start_index=actual_start_index,
        end_index=actual_end_index,
        text=text,
        document_id=document_id,
        extra_info=search_info if search_info else None,
        styles_applied=format_styles if format_styles else None,
    )

    # Convert to dict and return as JSON
    result_dict = op_result.to_dict()

    # Add resolved range info for range-based operations
    if range_result_info:
        result_dict["resolved_range"] = range_result_info

    # Add auto-normal text info if it was applied
    if auto_normal_text_applied:
        result_dict["auto_normal_text_applied"] = True
        result_dict["auto_normal_text_reason"] = (
            "Prevented heading style inheritance - insertion point was in a heading paragraph"
        )

    return json.dumps(result_dict, indent=2)


@server.tool()
@handle_http_errors("find_and_replace_doc", service_type="docs")
@require_google_service("docs", "docs_write")
async def find_and_replace_doc(
    service: Any,
    user_google_email: str,
    document_id: str,
    find_text: str = None,
    replace_text: str = None,
    match_case: bool = False,
    # Parameter aliases for consistency with other tools
    search: str = None,  # Alias for find_text (matches format_all_occurrences)
    text: str = None,  # Alias for replace_text (matches modify_doc_text)
    preview: bool = False,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    strikethrough: bool = None,
    font_size: int = None,
    font_family: str = None,
    link: str = None,
    foreground_color: str = None,
    background_color: str = None,
    tab_id: str = None,
) -> str:
    """
    Replaces ALL occurrences of text throughout a Google Doc.

    Use this tool when you want to replace EVERY instance of text in the document.
    For replacing a SPECIFIC occurrence (first, second, or last), use modify_doc_text
    with search + position='replace' instead.

    Comparison:
    | Tool                  | Replaces       | Use Case                              |
    |-----------------------|----------------|---------------------------------------|
    | find_and_replace_doc  | ALL occurrences| Global search-replace (like Ctrl+H)   |
    | modify_doc_text       | ONE occurrence | Replace specific instance by position |

    Examples:
        # Replace ALL "TODO" with "DONE":
        find_and_replace_doc(document_id="...", find_text="TODO", replace_text="DONE")

        # Preview what would be replaced before committing:
        find_and_replace_doc(document_id="...", find_text="TODO", replace_text="DONE", preview=True)

        # Replace and format - make all "TODO" become bold red "DONE":
        find_and_replace_doc(document_id="...", find_text="TODO", replace_text="DONE",
                            bold=True, foreground_color="red")

        # Replace and make text a clickable link:
        find_and_replace_doc(document_id="...", find_text="our website",
                            replace_text="Example.com", link="https://example.com")

        # Replace only the FIRST "TODO": use modify_doc_text instead
        modify_doc_text(document_id="...", search="TODO", position="replace", text="DONE")

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        find_text: Text to search for (will replace ALL matches). Alias: 'search'
        replace_text: Text to replace with (default: empty string for deletion). Alias: 'text'
        match_case: Whether to match case exactly (default: False)
        preview: If True, returns what would be replaced without making changes (default: False)
        search: Alias for find_text (for consistency with format_all_occurrences)
        text: Alias for replace_text (for consistency with modify_doc_text)
        bold: Whether to make replaced text bold (True/False/None to leave unchanged)
        italic: Whether to make replaced text italic (True/False/None to leave unchanged)
        underline: Whether to underline replaced text (True/False/None to leave unchanged)
        strikethrough: Whether to strikethrough replaced text (True/False/None to leave unchanged)
        font_size: Font size in points for replaced text
        font_family: Font family name for replaced text (e.g., "Arial", "Times New Roman")
        link: URL to create a hyperlink on replaced text
        foreground_color: Text color as hex (#FF0000) or named color (red, blue, green, etc.)
        background_color: Background/highlight color as hex or named color
        tab_id: Optional tab ID for multi-tab documents. If not provided, replaces in all tabs
            (default behavior for replaceAllText). Use list_doc_tabs() to discover tab IDs.
            Note: The underlying Google Docs API replaceAllText operation affects all tabs by
            default, or can be restricted to specific tabs using tabsCriteria.

    Returns:
        str: JSON string with operation details.

        When preview=False (default), returns structured operation result:
        {
            "success": true,
            "operation": "find_replace",
            "occurrences_replaced": 3,
            "find_text": "TODO",
            "replace_text": "DONE",
            "match_case": false,
            "position_shift_per_replacement": 1,
            "total_position_shift": 3,
            "affected_ranges": [
                {"index": 1, "original_range": {"start": 15, "end": 19}},
                {"index": 2, "original_range": {"start": 50, "end": 54}},
                ...
            ],
            "formatting_applied": ["bold", "foreground_color"],  // if formatting was requested
            "message": "Replaced 3 occurrence(s) of 'TODO' with 'DONE'",
            "link": "https://docs.google.com/document/d/.../edit"
        }

        When preview=True, returns a preview response:
        {
            "preview": true,
            "would_modify": true,
            "occurrences_found": 3,
            "find_text": "TODO",
            "replace_text": "DONE",
            "matches": [
                {
                    "index": 1,
                    "range": {"start": 15, "end": 19},
                    "context": {"before": "...", "after": "..."}
                },
                ...
            ],
            "position_shift_per_replacement": -1,
            "total_position_shift": -3,
            "message": "Would replace 3 occurrences..."
        }
    """
    # Resolve parameter aliases for API consistency
    # 'search' is an alias for 'find_text' (matches format_all_occurrences)
    # 'text' is an alias for 'replace_text' (matches modify_doc_text)
    if search is not None and find_text is None:
        find_text = search
    if text is not None and replace_text is None:
        replace_text = text

    # Ensure replace_text has a default if not provided
    if replace_text is None:
        replace_text = ""
    else:
        # Interpret escape sequences in replace_text (e.g., \n -> actual newline)
        replace_text = interpret_escape_sequences(replace_text)

    logger.info(
        f"[find_and_replace_doc] Doc={document_id}, find='{find_text}', replace='{replace_text}', preview={preview}"
    )

    import json
    from gdocs.managers.validation_manager import ValidationManager

    # Check if any formatting was requested
    has_formatting = any(
        [
            bold is not None,
            italic is not None,
            underline is not None,
            strikethrough is not None,
            font_size is not None,
            font_family is not None,
            link is not None,
            foreground_color is not None,
            background_color is not None,
        ]
    )

    # Build list of formatting parameters for response
    formatting_applied = []
    if bold is not None:
        formatting_applied.append("bold" if bold else "remove bold")
    if italic is not None:
        formatting_applied.append("italic" if italic else "remove italic")
    if underline is not None:
        formatting_applied.append("underline" if underline else "remove underline")
    if strikethrough is not None:
        formatting_applied.append(
            "strikethrough" if strikethrough else "remove strikethrough"
        )
    if font_size is not None:
        formatting_applied.append(f"font_size={font_size}")
    if font_family is not None:
        formatting_applied.append(f"font_family={font_family}")
    if link is not None:
        formatting_applied.append("link" if link else "remove link")
    if foreground_color is not None:
        formatting_applied.append("foreground_color")
    if background_color is not None:
        formatting_applied.append("background_color")

    # Validate font_size if provided
    validator = ValidationManager()
    if font_size is not None and font_size <= 0:
        return validator.create_invalid_param_error(
            param_name="font_size",
            received=font_size,
            valid_values=["positive integer (e.g., 10, 12, 14)"],
        )

    # Validate find_text is not empty
    if not find_text:
        error = DocsErrorBuilder.invalid_param_value(
            param_name="find_text",
            received_value="(empty string)",
            valid_values=["non-empty string"],
            context_description="find_text cannot be empty",
        )
        return format_error(error)

    # Handle preview mode - find all occurrences without modifying
    if preview:
        doc_data = await asyncio.to_thread(
            service.documents().get(documentId=document_id).execute
        )

        all_occurrences = find_all_occurrences_in_document(
            doc_data, find_text, match_case
        )
        link = f"https://docs.google.com/document/d/{document_id}/edit"

        # Calculate position shift
        len_diff = len(replace_text) - len(find_text)

        # Build matches list with context
        matches = []
        for idx, (start, end) in enumerate(all_occurrences, 1):
            match_info = {
                "index": idx,
                "range": {"start": start, "end": end},
            }
            # Extract context for each match
            context = extract_text_at_range(doc_data, start, end)
            if context.get("found"):
                match_info["text"] = context.get("text", "")
                match_info["context"] = {
                    "before": context.get("context_before", ""),
                    "after": context.get("context_after", ""),
                }
            matches.append(match_info)

        preview_result = {
            "preview": True,
            "would_modify": len(all_occurrences) > 0,
            "occurrences_found": len(all_occurrences),
            "find_text": find_text,
            "replace_text": replace_text,
            "match_case": match_case,
            "matches": matches,
            "position_shift_per_replacement": len_diff,
            "total_position_shift": len_diff * len(all_occurrences),
            "link": link,
        }

        if len(all_occurrences) == 0:
            preview_result["message"] = (
                f"No occurrences of '{find_text}' found in document"
            )
        else:
            preview_result["message"] = (
                f"Would replace {len(all_occurrences)} occurrence(s) of '{find_text}' with '{replace_text}'"
            )
            if has_formatting:
                preview_result["message"] += (
                    f" and apply formatting ({', '.join(formatting_applied)})"
                )

        # Include formatting info in preview if formatting was requested
        if has_formatting:
            preview_result["formatting_requested"] = formatting_applied

        return json.dumps(preview_result, indent=2)

    # For non-preview mode, first get document to find all occurrence positions
    # This allows us to report affected ranges in the structured response
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )
    all_occurrences = find_all_occurrences_in_document(doc_data, find_text, match_case)

    # Build matches list with original positions (before replacement)
    len_diff = len(replace_text) - len(find_text)
    matches = []
    for idx, (start, end) in enumerate(all_occurrences, 1):
        matches.append(
            {
                "index": idx,
                "original_range": {"start": start, "end": end},
            }
        )

    # Execute the actual replacement
    # Note: tab_id is passed as a single-element list (tab_ids) since replaceAllText
    # uses tabsCriteria which accepts multiple tab IDs
    tab_ids = [tab_id] if tab_id else None
    requests = [create_find_replace_request(find_text, replace_text, match_case, tab_ids=tab_ids)]

    result = await asyncio.to_thread(
        service.documents()
        .batchUpdate(documentId=document_id, body={"requests": requests})
        .execute
    )

    # Extract number of replacements from response
    replacements = 0
    if "replies" in result and result["replies"]:
        reply = result["replies"][0]
        if "replaceAllText" in reply:
            replacements = reply["replaceAllText"].get("occurrencesChanged", 0)

    # Apply formatting if requested and replacements were made
    occurrences_formatted = 0
    if has_formatting and replacements > 0 and replace_text:
        # Fetch the updated document to find positions of replaced text
        updated_doc_data = await asyncio.to_thread(
            service.documents().get(documentId=document_id).execute
        )

        # Find all occurrences of the replacement text
        replaced_occurrences = find_all_occurrences_in_document(
            updated_doc_data, replace_text, match_case
        )

        # Build formatting requests for each occurrence
        # First clear existing formatting, then apply the requested formatting
        # This ensures the replaced text only has the explicitly specified formatting
        format_requests = []
        for start, end in replaced_occurrences:
            # Clear existing formatting first (preserve links if user is setting a link)
            clear_request = create_clear_formatting_request(
                start, end, preserve_links=(link is not None), tab_id=tab_id
            )
            format_requests.append(clear_request)

            # Then apply the requested formatting
            format_request = create_format_text_request(
                start,
                end,
                bold,
                italic,
                underline,
                strikethrough,
                None,
                None,
                None,  # small_caps, subscript, superscript not supported
                font_size,
                font_family,
                link,
                foreground_color,
                background_color,
                tab_id=tab_id,
            )
            if format_request:
                format_requests.append(format_request)

        # Apply formatting in a batch update
        if format_requests:
            await asyncio.to_thread(
                service.documents()
                .batchUpdate(documentId=document_id, body={"requests": format_requests})
                .execute
            )
            occurrences_formatted = len(replaced_occurrences)

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    # Use len(matches) for consistency - occurrences_replaced should match len(affected_ranges)
    # The API's replacements count may differ slightly due to different matching algorithms
    reported_count = len(matches)
    message = (
        f"Replaced {reported_count} occurrence(s) of '{find_text}' with '{replace_text}'"
    )

    if occurrences_formatted > 0:
        message += f" and applied formatting to {occurrences_formatted} occurrence(s)"

    # Add hint when only 1 replacement was made - user might have wanted targeted replacement
    if reported_count == 1 and not has_formatting:
        message += ". Tip: For single replacements, you can also use modify_doc_text with search + position='replace' to target a specific occurrence."

    # Build structured response consistent with other tools
    operation_result = {
        "success": True,
        "operation": "find_replace",
        "occurrences_replaced": reported_count,
        "find_text": find_text,
        "replace_text": replace_text,
        "match_case": match_case,
        "position_shift_per_replacement": len_diff,
        "total_position_shift": len_diff * reported_count,
        "affected_ranges": matches,
        "message": message,
        "link": link,
    }

    # Include formatting info if formatting was applied
    if has_formatting:
        operation_result["formatting_applied"] = formatting_applied
        operation_result["occurrences_formatted"] = occurrences_formatted

    return json.dumps(operation_result, indent=2)


@server.tool()
@handle_http_errors("format_all_occurrences", service_type="docs")
@require_google_service("docs", "docs_write")
async def format_all_occurrences(
    service: Any,
    user_google_email: str,
    document_id: str,
    search: str,
    match_case: bool = True,
    preview: bool = False,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    strikethrough: bool = None,
    small_caps: bool = None,
    subscript: bool = None,
    superscript: bool = None,
    font_size: int = None,
    font_family: str = None,
    link: str = None,
    foreground_color: str = None,
    background_color: str = None,
) -> str:
    """
    Formats ALL occurrences of text in a Google Doc without changing the text itself.

    Use this tool when you want to apply formatting to every instance of a word or phrase
    without replacing it. This is more efficient than calling modify_doc_text multiple times.

    Comparison:
    | Tool                    | Modifies Text | Formats | Targets          |
    |-------------------------|---------------|---------|------------------|
    | format_all_occurrences  | No            | Yes     | ALL occurrences  |
    | find_and_replace_doc    | Yes           | Yes     | ALL occurrences  |
    | modify_doc_text         | Optional      | Yes     | ONE occurrence   |

    Examples:
        # Make all "TODO" bold and red:
        format_all_occurrences(document_id="...", search="TODO",
                              bold=True, foreground_color="red")

        # Preview what would be formatted:
        format_all_occurrences(document_id="...", search="TODO",
                              bold=True, preview=True)

        # Highlight all mentions of a term with background color:
        format_all_occurrences(document_id="...", search="important",
                              background_color="yellow")

        # Make all instances of a name a clickable link:
        format_all_occurrences(document_id="...", search="Google",
                              link="https://google.com")

        # Apply multiple formatting styles at once:
        format_all_occurrences(document_id="...", search="WARNING",
                              bold=True, italic=True, foreground_color="orange",
                              font_size=14)

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        search: Text to search for (will format ALL matches)
        match_case: Whether to match case exactly (default: True)
        preview: If True, returns what would be formatted without making changes (default: False)
        bold: Whether to make text bold (True/False/None to leave unchanged)
        italic: Whether to make text italic (True/False/None to leave unchanged)
        underline: Whether to underline text (True/False/None to leave unchanged)
        strikethrough: Whether to strikethrough text (True/False/None to leave unchanged)
        small_caps: Whether to make text small caps (True/False/None to leave unchanged)
        subscript: Whether to make text subscript (True/False/None to leave unchanged)
        superscript: Whether to make text superscript (True/False/None to leave unchanged)
        font_size: Font size in points
        font_family: Font family name (e.g., "Arial", "Times New Roman")
        link: URL to create a hyperlink. Use empty string "" to remove existing link.
        foreground_color: Text color as hex (#FF0000) or named color (red, blue, green, etc.)
        background_color: Background/highlight color as hex or named color

    Returns:
        str: JSON string with operation details.

        When preview=False (default), returns structured operation result:
        {
            "success": true,
            "operation": "format_all",
            "occurrences_formatted": 3,
            "search": "TODO",
            "match_case": true,
            "affected_ranges": [
                {"index": 1, "range": {"start": 15, "end": 19}},
                {"index": 2, "range": {"start": 50, "end": 54}},
                ...
            ],
            "formatting_applied": ["bold", "foreground_color"],
            "message": "Applied formatting to 3 occurrence(s) of 'TODO'",
            "link": "https://docs.google.com/document/d/.../edit"
        }

        When preview=True, returns a preview response:
        {
            "preview": true,
            "would_modify": true,
            "occurrences_found": 3,
            "search": "TODO",
            "matches": [
                {
                    "index": 1,
                    "range": {"start": 15, "end": 19},
                    "text": "TODO",
                    "context": {"before": "...", "after": "..."}
                },
                ...
            ],
            "formatting_to_apply": ["bold", "foreground_color"],
            "message": "Would format 3 occurrences of 'TODO'"
        }
    """
    import json

    logger.info(
        f"[format_all_occurrences] Doc={document_id}, search='{search}', preview={preview}"
    )

    validator = ValidationManager()

    # Validate document_id
    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Validate search is not empty
    if not search:
        error = DocsErrorBuilder.invalid_param_value(
            param_name="search",
            received_value="(empty string)",
            valid_values=["non-empty string"],
            context_description="search text cannot be empty",
        )
        return format_error(error)

    # Check if any formatting was requested
    has_formatting = any(
        [
            bold is not None,
            italic is not None,
            underline is not None,
            strikethrough is not None,
            small_caps is not None,
            subscript is not None,
            superscript is not None,
            font_size is not None,
            font_family is not None,
            link is not None,
            foreground_color is not None,
            background_color is not None,
        ]
    )

    if not has_formatting:
        error = DocsErrorBuilder.missing_required_param(
            param_name="formatting option",
            context_description="for formatting operation",
            valid_values=[
                "bold",
                "italic",
                "underline",
                "strikethrough",
                "small_caps",
                "subscript",
                "superscript",
                "font_size",
                "font_family",
                "link",
                "foreground_color",
                "background_color",
            ],
        )
        return format_error(error)

    # Validate font_size if provided
    if font_size is not None and font_size <= 0:
        return validator.create_invalid_param_error(
            param_name="font_size",
            received=font_size,
            valid_values=["positive integer (e.g., 10, 12, 14)"],
        )

    # Validate subscript/superscript mutual exclusivity
    if subscript and superscript:
        return validator.create_invalid_param_error(
            param_name="subscript/superscript",
            received="both True",
            valid_values=["subscript=True OR superscript=True, not both"],
        )

    # Build list of formatting parameters for response
    formatting_applied = []
    if bold is not None:
        formatting_applied.append("bold" if bold else "remove bold")
    if italic is not None:
        formatting_applied.append("italic" if italic else "remove italic")
    if underline is not None:
        formatting_applied.append("underline" if underline else "remove underline")
    if strikethrough is not None:
        formatting_applied.append(
            "strikethrough" if strikethrough else "remove strikethrough"
        )
    if small_caps is not None:
        formatting_applied.append("small_caps" if small_caps else "remove small_caps")
    if subscript is not None:
        formatting_applied.append("subscript" if subscript else "remove subscript")
    if superscript is not None:
        formatting_applied.append(
            "superscript" if superscript else "remove superscript"
        )
    if font_size is not None:
        formatting_applied.append(f"font_size={font_size}")
    if font_family is not None:
        formatting_applied.append(f"font_family={font_family}")
    if link is not None:
        formatting_applied.append("link" if link else "remove link")
    if foreground_color is not None:
        formatting_applied.append("foreground_color")
    if background_color is not None:
        formatting_applied.append("background_color")

    # Get document and find all occurrences
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )
    all_occurrences = find_all_occurrences_in_document(doc_data, search, match_case)

    doc_link = f"https://docs.google.com/document/d/{document_id}/edit"

    # Handle preview mode
    if preview:
        matches = []
        for idx, (start, end) in enumerate(all_occurrences, 1):
            match_info = {
                "index": idx,
                "range": {"start": start, "end": end},
            }
            context = extract_text_at_range(doc_data, start, end)
            if context.get("found"):
                match_info["text"] = context.get("text", "")
                match_info["context"] = {
                    "before": context.get("context_before", ""),
                    "after": context.get("context_after", ""),
                }
            matches.append(match_info)

        preview_result = {
            "preview": True,
            "would_modify": len(all_occurrences) > 0,
            "occurrences_found": len(all_occurrences),
            "search": search,
            "match_case": match_case,
            "matches": matches,
            "formatting_to_apply": formatting_applied,
            "link": doc_link,
        }

        if len(all_occurrences) == 0:
            preview_result["message"] = (
                f"No occurrences of '{search}' found in document"
            )
        else:
            preview_result["message"] = (
                f"Would format {len(all_occurrences)} occurrence(s) of '{search}' with {', '.join(formatting_applied)}"
            )

        return json.dumps(preview_result, indent=2)

    # No occurrences found
    if not all_occurrences:
        return json.dumps(
            {
                "success": True,
                "operation": "format_all",
                "occurrences_formatted": 0,
                "search": search,
                "match_case": match_case,
                "affected_ranges": [],
                "formatting_applied": formatting_applied,
                "message": f"No occurrences of '{search}' found in document",
                "link": doc_link,
            },
            indent=2,
        )

    # Build formatting requests for each occurrence
    format_requests = []
    affected_ranges = []
    for idx, (start, end) in enumerate(all_occurrences, 1):
        format_request = create_format_text_request(
            start,
            end,
            bold,
            italic,
            underline,
            strikethrough,
            small_caps,
            subscript,
            superscript,
            font_size,
            font_family,
            link,
            foreground_color,
            background_color,
        )
        if format_request:
            format_requests.append(format_request)
            affected_ranges.append(
                {"index": idx, "range": {"start": start, "end": end}}
            )

    # Apply formatting in a single batch update
    if format_requests:
        await asyncio.to_thread(
            service.documents()
            .batchUpdate(documentId=document_id, body={"requests": format_requests})
            .execute
        )

    operation_result = {
        "success": True,
        "operation": "format_all",
        "occurrences_formatted": len(format_requests),
        "search": search,
        "match_case": match_case,
        "affected_ranges": affected_ranges,
        "formatting_applied": formatting_applied,
        "message": f"Applied formatting ({', '.join(formatting_applied)}) to {len(format_requests)} occurrence(s) of '{search}'",
        "link": doc_link,
    }

    return json.dumps(operation_result, indent=2)


@server.tool()
@handle_http_errors("auto_linkify_doc", service_type="docs")
@require_google_service("docs", "docs_write")
async def auto_linkify_doc(
    service: Any,
    user_google_email: str,
    document_id: str,
    auto_detect: bool = True,
    url_pattern: str = None,
    exclude_already_linked: bool = True,
    preview: bool = False,
) -> str:
    """
    Automatically converts plain text URLs into clickable hyperlinks throughout a Google Doc.

    This tool scans the document for URLs that are not already hyperlinked and converts
    them into clickable links. It's useful for cleaning up documents that contain
    pasted URLs that weren't automatically linked.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update

        Detection options:
        auto_detect: If True (default), automatically detect URLs using standard patterns.
            Detects http://, https://, and www. URLs.
        url_pattern: Optional custom regex pattern to match URLs. If provided, this is used
            instead of the default URL detection. Use for specialized URL formats.

        Filtering options:
        exclude_already_linked: If True (default), skip text that already has a hyperlink.
            Set to False to re-apply links (useful for fixing broken links).

        Preview mode:
        preview: If True, returns what URLs would be linked without making changes.
            Useful for reviewing before applying. Default: False

    Examples:
        # Auto-detect and linkify all plain URLs in document:
        auto_linkify_doc(document_id="...")

        # Preview what would be linked:
        auto_linkify_doc(document_id="...", preview=True)

        # Re-apply links even to already linked text (useful for fixing):
        auto_linkify_doc(document_id="...", exclude_already_linked=False)

        # Use custom pattern for specific URL format:
        auto_linkify_doc(document_id="...",
                        url_pattern=r"https://github\\.com/[\\w-]+/[\\w-]+")

    Returns:
        str: JSON string with operation details.

        When preview=False (default), returns structured operation result:
        {
            "success": true,
            "operation": "auto_linkify",
            "urls_linked": 5,
            "urls_found": 5,
            "urls_skipped": 0,
            "affected_ranges": [
                {"url": "https://example.com", "range": {"start": 15, "end": 35}},
                ...
            ],
            "message": "Linked 5 URL(s) in document",
            "link": "https://docs.google.com/document/d/.../edit"
        }

        When preview=True, returns a preview response:
        {
            "preview": true,
            "would_modify": true,
            "urls_found": 5,
            "urls_to_link": [
                {"url": "https://example.com", "range": {"start": 15, "end": 35}},
                ...
            ],
            "urls_already_linked": 2,
            "message": "Would link 5 URL(s)"
        }
    """
    import json
    import re
    from gdocs.docs_helpers import (
        extract_document_text_with_indices,
        create_format_text_request,
    )

    logger.info(
        f"[auto_linkify_doc] Doc={document_id}, auto_detect={auto_detect}, preview={preview}"
    )

    # Define URL regex patterns
    # This pattern matches:
    # - http:// and https:// URLs
    # - www. URLs (will be prefixed with https://)
    # The pattern avoids capturing trailing punctuation that's not part of the URL
    DEFAULT_URL_PATTERN = (
        r"(?:https?://|www\.)"  # Protocol or www.
        r"[a-zA-Z0-9]"  # Must start with alphanumeric
        r"(?:[a-zA-Z0-9\-._~:/?#\[\]@!$&\'()*+,;=%]*[a-zA-Z0-9/])?"  # URL chars, must end with alphanumeric or /
    )

    # Compile the pattern
    if url_pattern:
        try:
            pattern = re.compile(url_pattern)
        except re.error as e:
            error = DocsErrorBuilder.invalid_param_value(
                param_name="url_pattern",
                received_value=url_pattern,
                valid_values=["valid regular expression"],
                context_description=f"Regex compilation failed: {e}",
            )
            return format_error(error)
    else:
        pattern = re.compile(DEFAULT_URL_PATTERN, re.IGNORECASE)

    # Get document data
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Extract text with indices
    text_segments = extract_document_text_with_indices(doc_data)

    # Build full text and index mapping for regex matching
    full_text = ""
    index_map = []  # Maps position in full_text to document index

    for segment_text, start_idx, _ in text_segments:
        for i, char in enumerate(segment_text):
            index_map.append(start_idx + i)
            full_text += char

    # Find all URLs in the document
    found_urls = []
    for match in pattern.finditer(full_text):
        text_start = match.start()
        text_end = match.end()
        url_text = match.group()

        # Map back to document indices
        if text_start < len(index_map) and text_end <= len(index_map):
            doc_start = index_map[text_start]
            doc_end = index_map[text_end - 1] + 1

            # Normalize URL (add https:// to www. URLs)
            normalized_url = url_text
            if url_text.lower().startswith("www."):
                normalized_url = "https://" + url_text

            found_urls.append(
                {
                    "url": normalized_url,
                    "original_text": url_text,
                    "range": {"start": doc_start, "end": doc_end},
                }
            )

    # If exclude_already_linked, check which URLs are already hyperlinked
    urls_to_link = []
    urls_already_linked = []

    if exclude_already_linked and found_urls:
        # Extract existing links from document
        def extract_links_from_elements(elements, links_set):
            """Recursively extract all link ranges from document elements."""
            for element in elements:
                if "paragraph" in element:
                    paragraph = element["paragraph"]
                    for para_element in paragraph.get("elements", []):
                        if "textRun" in para_element:
                            text_run = para_element["textRun"]
                            text_style = text_run.get("textStyle", {})
                            if "link" in text_style:
                                start_idx = para_element.get("startIndex", 0)
                                end_idx = para_element.get("endIndex", 0)
                                links_set.add((start_idx, end_idx))
                elif "table" in element:
                    table = element["table"]
                    for row in table.get("tableRows", []):
                        for cell in row.get("tableCells", []):
                            cell_content = cell.get("content", [])
                            extract_links_from_elements(cell_content, links_set)

        existing_links = set()
        body = doc_data.get("body", {})
        content = body.get("content", [])
        extract_links_from_elements(content, existing_links)

        # Check each found URL against existing links
        for url_info in found_urls:
            url_start = url_info["range"]["start"]
            url_end = url_info["range"]["end"]

            # Check if this URL overlaps with any existing link
            is_already_linked = False
            for link_start, link_end in existing_links:
                # Check for overlap
                if url_start < link_end and url_end > link_start:
                    is_already_linked = True
                    break

            if is_already_linked:
                urls_already_linked.append(url_info)
            else:
                urls_to_link.append(url_info)
    else:
        urls_to_link = found_urls

    doc_link = f"https://docs.google.com/document/d/{document_id}/edit"

    # Handle preview mode
    if preview:
        preview_result = {
            "preview": True,
            "would_modify": len(urls_to_link) > 0,
            "urls_found": len(found_urls),
            "urls_to_link": urls_to_link,
            "urls_already_linked": len(urls_already_linked),
            "link": doc_link,
        }

        if len(urls_to_link) == 0:
            if len(found_urls) == 0:
                preview_result["message"] = "No URLs found in document"
            else:
                preview_result["message"] = (
                    f"All {len(found_urls)} URL(s) are already linked"
                )
        else:
            preview_result["message"] = f"Would link {len(urls_to_link)} URL(s)"
            if len(urls_already_linked) > 0:
                preview_result["message"] += (
                    f" ({len(urls_already_linked)} already linked, will be skipped)"
                )

        return json.dumps(preview_result, indent=2)

    # No URLs to link
    if not urls_to_link:
        return json.dumps(
            {
                "success": True,
                "operation": "auto_linkify",
                "urls_linked": 0,
                "urls_found": len(found_urls),
                "urls_skipped": len(urls_already_linked),
                "affected_ranges": [],
                "message": "No URLs found to link"
                if len(found_urls) == 0
                else f"All {len(found_urls)} URL(s) are already linked",
                "link": doc_link,
            },
            indent=2,
        )

    # Build formatting requests to apply links
    format_requests = []
    affected_ranges = []

    for url_info in urls_to_link:
        start = url_info["range"]["start"]
        end = url_info["range"]["end"]
        url = url_info["url"]

        # Create format request with link
        format_request = create_format_text_request(
            start,
            end,
            link=url,  # Apply the URL as a hyperlink
        )
        if format_request:
            format_requests.append(format_request)
            affected_ranges.append(
                {
                    "url": url,
                    "original_text": url_info["original_text"],
                    "range": {"start": start, "end": end},
                }
            )

    # Apply links in a single batch update
    if format_requests:
        await asyncio.to_thread(
            service.documents()
            .batchUpdate(documentId=document_id, body={"requests": format_requests})
            .execute
        )

    operation_result = {
        "success": True,
        "operation": "auto_linkify",
        "urls_linked": len(format_requests),
        "urls_found": len(found_urls),
        "urls_skipped": len(urls_already_linked),
        "affected_ranges": affected_ranges,
        "message": f"Linked {len(format_requests)} URL(s) in document",
        "link": doc_link,
    }

    if len(urls_already_linked) > 0:
        operation_result["message"] += (
            f" ({len(urls_already_linked)} already linked, skipped)"
        )

    return json.dumps(operation_result, indent=2)


@server.tool()
@handle_http_errors("insert_doc_elements", service_type="docs")
@require_google_service("docs", "docs_write")
async def insert_doc_elements(
    service: Any,
    user_google_email: str,
    document_id: str,
    element_type: str,
    index: int = None,
    location: str = None,
    rows: int = None,
    columns: int = None,
    list_type: str = None,
    text: str = None,
    items: list = None,
    nesting_levels: list = None,
    section_type: str = None,
) -> str:
    """
    Inserts structural elements like tables, lists, page breaks, or horizontal rules into a Google Doc.

    SIMPLIFIED USAGE:
    - Use location='end' to append element at end of document (recommended)
    - Use location='start' to insert element at beginning of document
    - Provide explicit 'index' only for precise positioning

    For lists, you can insert multiple items at once using the 'items' parameter:
        insert_doc_elements(element_type='list', location='end', list_type='ORDERED',
                           items=['First item', 'Second item', 'Third item'])

    For nested lists (hierarchical lists with sub-items), use the 'nesting_levels' parameter:
        insert_doc_elements(element_type='list', location='end', list_type='UNORDERED',
                           items=['Main item', 'Sub item 1', 'Sub item 2', 'Another main'],
                           nesting_levels=[0, 1, 1, 0])

    Nesting levels: 0 = top level, 1 = first indent, 2 = second indent, etc. (max 8 levels)

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        element_type: Type of element to insert ("table", "list", "page_break", "horizontal_rule", "section_break")
        index: Position to insert element (optional, mutually exclusive with location)
        location: Semantic location - "start" or "end" (mutually exclusive with index)
        rows: Number of rows for table (required for table)
        columns: Number of columns for table (required for table)
        list_type: Type of list ("UNORDERED", "ORDERED") (required for list)
        text: Text for single list item (for backwards compatibility)
        items: List of strings for multiple list items (preferred over 'text')
        nesting_levels: List of integers specifying nesting level for each item (0-8).
                       Must match length of 'items'. Default is 0 (top level) for all items.
        section_type: Type of section break ("NEXT_PAGE" or "CONTINUOUS"). Default is "NEXT_PAGE".
                     NEXT_PAGE starts the new section on the next page.
                     CONTINUOUS starts the new section immediately after the previous section.

    Returns:
        str: Confirmation message with insertion details
    """
    logger.info(
        f"[insert_doc_elements] Doc={document_id}, type={element_type}, index={index}, location={location}"
    )

    # Input validation
    validator = ValidationManager()

    # Validate location parameter if provided
    if location is not None and location not in ["start", "end"]:
        return validator.create_invalid_param_error(
            param_name="location", received=location, valid_values=["start", "end"]
        )

    # Check for mutually exclusive positioning parameters
    if index is not None and location is not None:
        return (
            "ERROR: Cannot specify both 'index' and 'location'. Use one or the other."
        )

    # Require at least one positioning parameter
    if index is None and location is None:
        return (
            "ERROR: Must specify positioning. Use either 'index' for exact position "
            "or 'location' ('start' or 'end') for semantic positioning."
        )

    # Resolve insertion index
    resolved_index = index
    location_description = None

    if index is not None:
        # Explicit index provided
        # Handle the special case where we can't insert at the first section break
        if index == 0:
            logger.debug("Adjusting index from 0 to 1 to avoid first section break")
            resolved_index = 1
        location_description = f"at index {resolved_index}"
    else:
        # Location-based positioning - fetch document first
        try:
            doc_data = await asyncio.to_thread(
                service.documents().get(documentId=document_id).execute
            )
        except Exception as e:
            return f"ERROR: Failed to fetch document for index calculation: {str(e)}"

        if location == "start":
            # Insert at document start (after initial section break)
            resolved_index = 1
            location_description = "at start of document"
        else:
            # location == 'end': append to end of document
            structure = parse_document_structure(doc_data)
            total_length = structure["total_length"]
            resolved_index = total_length - 1 if total_length > 1 else 1
            location_description = "at end of document"

    requests = []

    if element_type == "table":
        if not rows or not columns:
            return validator.create_missing_param_error(
                param_name="rows and columns",
                context="for table insertion",
                valid_values=["rows: positive integer", "columns: positive integer"],
            )

        requests.append(create_insert_table_request(resolved_index, rows, columns))
        description = f"table ({rows}x{columns})"

    elif element_type == "list":
        if not list_type:
            return validator.create_missing_param_error(
                param_name="list_type",
                context="for list insertion",
                valid_values=["UNORDERED", "ORDERED"],
            )

        valid_list_types = ["ORDERED", "UNORDERED"]
        if list_type.upper() not in valid_list_types:
            return validator.create_invalid_param_error(
                param_name="list_type",
                received=list_type,
                valid_values=valid_list_types,
            )

        # Determine list items to insert
        list_items = []
        item_nesting = []
        if items is not None:
            if not isinstance(items, list):
                return validator.create_invalid_param_error(
                    param_name="items",
                    received=type(items).__name__,
                    valid_values=["list of strings"],
                )
            if len(items) == 0:
                return validator.create_invalid_param_error(
                    param_name="items",
                    received="empty list",
                    valid_values=["non-empty list of strings"],
                )
            # Validate all items are strings
            for i, item in enumerate(items):
                if not isinstance(item, str):
                    return validator.create_invalid_param_error(
                        param_name=f"items[{i}]",
                        received=type(item).__name__,
                        valid_values=["string"],
                    )
            list_items = [interpret_escape_sequences(item) for item in items]

            # Handle nesting_levels parameter
            if nesting_levels is not None:
                if not isinstance(nesting_levels, list):
                    return validator.create_invalid_param_error(
                        param_name="nesting_levels",
                        received=type(nesting_levels).__name__,
                        valid_values=["list of integers"],
                    )
                if len(nesting_levels) != len(items):
                    return validator.create_invalid_param_error(
                        param_name="nesting_levels",
                        received=f"list of length {len(nesting_levels)}",
                        valid_values=[
                            f"list of length {len(items)} (must match items length)"
                        ],
                    )
                # Validate all nesting levels are valid integers
                for i, level in enumerate(nesting_levels):
                    if not isinstance(level, int):
                        return validator.create_invalid_param_error(
                            param_name=f"nesting_levels[{i}]",
                            received=type(level).__name__,
                            valid_values=["integer (0-8)"],
                        )
                    if level < 0 or level > 8:
                        return validator.create_invalid_param_error(
                            param_name=f"nesting_levels[{i}]",
                            received=str(level),
                            valid_values=["integer from 0 to 8"],
                        )
                item_nesting = nesting_levels
            else:
                # Default to level 0 for all items
                item_nesting = [0] * len(items)
        elif text:
            list_items = [interpret_escape_sequences(text)]
            item_nesting = [0]
        else:
            list_items = ["List item"]
            item_nesting = [0]

        # Build combined text with tab prefixes for nesting levels
        # The Google Docs API determines nesting level by counting leading tabs
        nested_items = []
        for item_text, level in zip(list_items, item_nesting):
            prefix = "\t" * level
            nested_items.append(prefix + item_text)

        combined_text = "\n".join(nested_items) + "\n"
        total_length = len(combined_text) - 1  # Exclude final newline from range

        # Insert all text first, then apply bullet formatting to the entire range
        requests.extend(
            [
                create_insert_text_request(resolved_index, combined_text),
                create_bullet_list_request(
                    resolved_index, resolved_index + total_length, list_type
                ),
            ]
        )

        item_count = len(list_items)
        has_nested = any(level > 0 for level in item_nesting)
        if has_nested:
            max_level = max(item_nesting)
            description = f"{list_type.lower()} list with {item_count} item{'s' if item_count > 1 else ''} (nested, max depth {max_level})"
        else:
            description = f"{list_type.lower()} list with {item_count} item{'s' if item_count > 1 else ''}"

    elif element_type == "page_break":
        requests.append(create_insert_page_break_request(resolved_index))
        description = "page break"

    elif element_type == "horizontal_rule":
        # Uses table-based workaround (Google Docs API has no native horizontal rule support)
        requests.extend(create_insert_horizontal_rule_requests(resolved_index))
        description = "horizontal rule"

    elif element_type == "section_break":
        # section_type defaults to "NEXT_PAGE" if not specified
        # Valid section types: "NEXT_PAGE" (starts on next page), "CONTINUOUS" (starts immediately)
        valid_section_types = ["NEXT_PAGE", "CONTINUOUS"]
        section_type_value = section_type if section_type else "NEXT_PAGE"

        if section_type_value not in valid_section_types:
            return validator.create_invalid_param_error(
                param_name="section_type",
                received=section_type_value,
                valid_values=valid_section_types,
            )

        requests.append(create_insert_section_break_request(resolved_index, section_type_value))
        description = f"section break ({section_type_value.lower().replace('_', ' ')})"

    else:
        return validator.create_invalid_param_error(
            param_name="element_type",
            received=element_type,
            valid_values=["table", "list", "page_break", "horizontal_rule", "section_break"],
        )

    await asyncio.to_thread(
        service.documents()
        .batchUpdate(documentId=document_id, body={"requests": requests})
        .execute
    )

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Inserted {description} {location_description} in document {document_id}. Link: {link}"


@server.tool()
@handle_http_errors("insert_doc_image", service_type="docs")
@require_multiple_services(
    [
        {"service_type": "docs", "scopes": "docs_write", "param_name": "docs_service"},
        {
            "service_type": "drive",
            "scopes": "drive_read",
            "param_name": "drive_service",
        },
    ]
)
async def insert_doc_image(
    docs_service: Any,
    drive_service: Any,
    user_google_email: str,
    document_id: str,
    image_source: str = None,
    index: int = None,
    after_heading: str = None,
    location: str = None,
    width: int = None,
    height: int = None,
    # Parameter alias for intuitive naming
    image_url: str = None,  # Alias for image_source (more intuitive for URL-based images)
) -> str:
    """
    Inserts an image into a Google Doc from Drive or a URL.

    SIMPLIFIED USAGE - No pre-flight call needed:
    - Use location='end' to append image at end of document (recommended)
    - Use location='start' to insert image at beginning of document
    - Use 'after_heading' to insert image after a specific heading
    - Provide explicit 'index' only for precise positioning

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        image_source: Drive file ID or public image URL
        index: Document position (optional, mutually exclusive with location/after_heading)
        after_heading: Insert after this heading (optional, case-insensitive)
        location: Semantic location - "start" or "end" (mutually exclusive with index/after_heading)
        width: Image width in points (optional)
        height: Image height in points (optional)
        image_url: Alias for image_source (more intuitive for URL-based images)

    Returns:
        str: Confirmation message with insertion details
    """
    # Resolve parameter alias: 'image_url' is an alias for 'image_source'
    if image_url is not None and image_source is None:
        image_source = image_url

    logger.info(
        f"[insert_doc_image] Doc={document_id}, source={image_source}, index={index}, after_heading={after_heading}, location={location}"
    )

    # Input validation
    validator = ValidationManager()

    # Validate image_source is provided (either directly or via alias)
    if not image_source:
        return validator.create_invalid_param_error(
            param_name="image_source",
            received="None",
            valid_values=["Drive file ID", "public image URL"],
            context="image_source (or image_url) is required",
        )

    # Validate location parameter if provided
    if location is not None and location not in ["start", "end"]:
        return validator.create_invalid_param_error(
            param_name="location", received=location, valid_values=["start", "end"]
        )

    # Check for mutually exclusive positioning parameters
    positioning_params = [
        ("index", index),
        ("after_heading", after_heading),
        ("location", location),
    ]
    provided_params = [name for name, value in positioning_params if value is not None]
    if len(provided_params) > 1:
        return (
            f"ERROR: Cannot specify multiple positioning parameters. "
            f"Got: {', '.join(provided_params)}. Use only one of: index, after_heading, or location."
        )

    # Resolve insertion index
    resolved_index = index
    location_description = None

    if index is not None:
        # Explicit index provided - validate it
        is_valid, error_msg = validator.validate_index(index, "Index")
        if not is_valid:
            return f"ERROR: {error_msg}"
        # Handle the special case where we can't insert at the first section break
        if index == 0:
            logger.debug("Adjusting index from 0 to 1 to avoid first section break")
            resolved_index = 1
        location_description = f"at explicit index {resolved_index}"
    else:
        # Auto-detect insertion point - fetch document first
        try:
            doc_data = await asyncio.to_thread(
                docs_service.documents().get(documentId=document_id).execute
            )
        except Exception as e:
            return f"ERROR: Failed to fetch document for index calculation: {str(e)}"

        if after_heading is not None:
            # Insert after specified heading
            insertion_point = find_section_insertion_point(
                doc_data, after_heading, position="end"
            )
            if insertion_point is None:
                # List available headings to help user
                headings = get_all_headings(doc_data)
                if headings:
                    heading_list = ", ".join([f'"{h["text"]}"' for h in headings[:5]])
                    more = (
                        f" (and {len(headings) - 5} more)" if len(headings) > 5 else ""
                    )
                    return (
                        f"ERROR: Heading '{after_heading}' not found. "
                        f"Available headings: {heading_list}{more}"
                    )
                else:
                    return (
                        f"ERROR: Heading '{after_heading}' not found. "
                        "Document has no headings."
                    )
            resolved_index = insertion_point
            location_description = f"after heading '{after_heading}'"
        elif location == "start":
            # Insert at document start (after initial section break)
            resolved_index = 1
            location_description = "at start of document"
        else:
            # Default or location='end': append to end of document
            structure = parse_document_structure(doc_data)
            total_length = structure["total_length"]
            # Use total_length - 1 for safe insertion point
            resolved_index = total_length - 1 if total_length > 1 else 1
            location_description = "at end of document"

    logger.debug(
        f"[insert_doc_image] Resolved index: {resolved_index} ({location_description})"
    )

    # Determine if source is a Drive file ID or URL
    is_drive_file = not (
        image_source.startswith("http://") or image_source.startswith("https://")
    )

    if is_drive_file:
        # Verify Drive file exists and get metadata
        try:
            file_metadata = await asyncio.to_thread(
                drive_service.files()
                .get(fileId=image_source, fields="id, name, mimeType")
                .execute
            )
            mime_type = file_metadata.get("mimeType", "")
            if not mime_type.startswith("image/"):
                return validator.create_image_error(
                    image_source=image_source, actual_mime_type=mime_type
                )

            image_uri = f"https://drive.google.com/uc?id={image_source}"
            source_description = f"Drive file {file_metadata.get('name', image_source)}"
        except Exception as e:
            return validator.create_image_error(
                image_source=image_source, error_detail=str(e)
            )
    else:
        image_uri = image_source
        source_description = "URL image"

    # Use helper to create image request
    requests = [create_insert_image_request(resolved_index, image_uri, width, height)]

    await asyncio.to_thread(
        docs_service.documents()
        .batchUpdate(documentId=document_id, body={"requests": requests})
        .execute
    )

    size_info = ""
    if width or height:
        size_info = f" (size: {width or 'auto'}x{height or 'auto'} points)"

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return (
        f"Inserted {source_description}{size_info} {location_description} "
        f"(index {resolved_index}) in document {document_id}. Link: {link}"
    )


@server.tool()
@handle_http_errors("insert_doc_footnote", service_type="docs")
@require_google_service("docs", "docs_write")
async def insert_doc_footnote(
    service: Any,
    user_google_email: str,
    document_id: str,
    footnote_text: str,
    index: int = None,
    location: str = None,
    search: str = None,
    position: str = None,
    occurrence: int = 1,
    match_case: bool = False,
) -> str:
    """
    Inserts a footnote into a Google Doc at the specified position.

    A footnote reference (superscript number) is inserted at the specified location,
    and the footnote content is added to the footnote section at the bottom of the page.

    POSITIONING OPTIONS (use exactly one):
    - location='end' to insert footnote at end of document (recommended for appending)
    - location='start' to insert footnote at beginning of document
    - index=N for explicit position (character index)
    - search + position to insert relative to found text

    Note: Footnotes can only be inserted in the document body, not in headers,
    footers, or other footnotes.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        footnote_text: Text content for the footnote
        index: Position to insert footnote reference (optional, mutually exclusive with location/search)
        location: Semantic location - "start" or "end" (mutually exclusive with index/search)
        search: Text to search for to determine insertion point (use with position)
        position: Where to insert relative to search text - "before" or "after" (required when using search)
        occurrence: Which occurrence of search text (1=first, 2=second, -1=last). Default: 1
        match_case: Whether search should be case-sensitive. Default: false

    Returns:
        str: Confirmation message with footnote details and document link
    """
    logger.info(
        f"[insert_doc_footnote] Doc={document_id}, index={index}, location={location}, search={search}"
    )

    # Input validation
    validator = ValidationManager()

    # Validate footnote_text
    if not footnote_text:
        return validator.create_missing_param_error(
            param_name="footnote_text",
            context="for footnote insertion",
            valid_values=["non-empty string with footnote content"],
        )

    # Validate location parameter if provided
    if location is not None and location not in ["start", "end"]:
        return validator.create_invalid_param_error(
            param_name="location", received=location, valid_values=["start", "end"]
        )

    # Count positioning parameters provided
    positioning_params = [
        ("index", index is not None),
        ("location", location is not None),
        ("search", search is not None),
    ]
    provided_count = sum(1 for _, provided in positioning_params if provided)

    if provided_count == 0:
        return (
            "ERROR: Must specify positioning. Use one of:\n"
            "  - index=N for explicit position\n"
            "  - location='start' or 'end' for semantic positioning\n"
            "  - search='text' with position='before'/'after' for search-based positioning"
        )
    if provided_count > 1:
        provided_names = [name for name, provided in positioning_params if provided]
        return (
            f"ERROR: Cannot specify multiple positioning parameters. "
            f"You provided: {', '.join(provided_names)}. Use only one."
        )

    # Validate search-based positioning parameters
    if search is not None:
        if position is None:
            return validator.create_missing_param_error(
                param_name="position",
                context="when using search-based positioning",
                valid_values=["before", "after"],
            )
        if position not in ["before", "after"]:
            return validator.create_invalid_param_error(
                param_name="position",
                received=position,
                valid_values=["before", "after"],
            )

    # Fetch document to resolve positioning
    try:
        doc_data = await asyncio.to_thread(
            service.documents().get(documentId=document_id).execute
        )
    except Exception as e:
        return f"ERROR: Failed to fetch document: {str(e)}"

    # Resolve insertion index based on positioning mode
    resolved_index = None
    location_description = None

    if index is not None:
        # Explicit index provided
        resolved_index = max(1, index)  # Cannot insert before first section break
        location_description = f"at index {resolved_index}"
    elif location == "start":
        resolved_index = 1  # After initial section break
        location_description = "at start of document"
    elif location == "end":
        structure = parse_document_structure(doc_data)
        total_length = structure["total_length"]
        resolved_index = max(1, total_length - 1)
        location_description = "at end of document"
    elif search is not None:
        # Search-based positioning
        success, found_start, found_end, msg = calculate_search_based_indices(
            doc_data, search, position, occurrence, match_case
        )
        if not success:
            return f"ERROR: {msg}"
        resolved_index = found_start if position == "before" else found_end
        location_description = f"{position} '{search}'"
        if occurrence != 1:
            location_description += f" (occurrence {occurrence})"

    # Interpret escape sequences in footnote text
    processed_text = interpret_escape_sequences(footnote_text)

    # Build and execute requests
    # Step 1: Create the footnote (this inserts a footnote reference and creates the footnote segment)
    create_footnote_request = create_insert_footnote_request(resolved_index)

    try:
        result = await asyncio.to_thread(
            service.documents()
            .batchUpdate(
                documentId=document_id,
                body={"requests": [create_footnote_request]}
            )
            .execute
        )
    except Exception as e:
        error_msg = str(e)
        if "Invalid requests" in error_msg or "footerHeaderOrFootnote" in error_msg.lower():
            return (
                "ERROR: Cannot insert footnote at this location. Footnotes can only be "
                "inserted in the document body, not in headers, footers, or other footnotes."
            )
        return f"ERROR: Failed to create footnote: {error_msg}"

    # Get the footnote ID from the response
    footnote_id = None
    if "replies" in result:
        for reply in result["replies"]:
            if "createFootnote" in reply:
                footnote_id = reply["createFootnote"].get("footnoteId")
                break

    if not footnote_id:
        return "ERROR: Footnote was created but footnote ID was not returned. The footnote reference was inserted but content could not be added."

    # Step 2: Insert text into the footnote
    # Footnotes start with a space followed by newline, so we insert at index 1
    insert_text_request = create_insert_text_in_footnote_request(
        footnote_id=footnote_id,
        index=1,  # Insert after the initial space
        text=processed_text
    )

    try:
        await asyncio.to_thread(
            service.documents()
            .batchUpdate(
                documentId=document_id,
                body={"requests": [insert_text_request]}
            )
            .execute
        )
    except Exception as e:
        return (
            f"WARNING: Footnote reference was created {location_description}, "
            f"but failed to add footnote text: {str(e)}. "
            "You may need to manually add the footnote content."
        )

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    text_preview = processed_text[:50] + "..." if len(processed_text) > 50 else processed_text
    return (
        f"Inserted footnote {location_description} in document {document_id}. "
        f"Footnote content: \"{text_preview}\". Link: {link}"
    )


@server.tool()
@handle_http_errors("update_doc_headers_footers", service_type="docs")
@require_google_service("docs", "docs_write")
async def update_doc_headers_footers(
    service: Any,
    user_google_email: str,
    document_id: str,
    section_type: str = None,
    content: str = None,
    header_footer_type: str = "DEFAULT",
    create_if_missing: bool = True,
    # Parameter alias for consistency with other tools
    text: str = None,  # Alias for content (matches modify_doc_text, insert_doc_elements)
    # New parameters for combined header+footer updates
    header_content: str = None,  # Set header content directly (use instead of section_type+content)
    footer_content: str = None,  # Set footer content directly (use instead of section_type+content)
) -> str:
    """
    Updates headers or footers in a Google Doc. Creates the header/footer if it doesn't exist.

    Supports two usage modes:
    1. Single section: Use section_type + content to update one section
    2. Combined: Use header_content and/or footer_content to update both in one call

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        section_type: Type of section to update ("header" or "footer") - used with content param
        content: Text content for the header/footer - used with section_type param
        header_footer_type: Type of header/footer ("DEFAULT", "FIRST_PAGE", "EVEN_PAGE")
        create_if_missing: If true, creates the header/footer if it doesn't exist (default: true)
        text: Alias for content (for consistency with modify_doc_text, insert_doc_elements)
        header_content: Header text - set directly without section_type (can combine with footer_content)
        footer_content: Footer text - set directly without section_type (can combine with header_content)

    Returns:
        str: Confirmation message with update details

    Examples:
        # Single section (original API):
        update_doc_headers_footers(doc_id, section_type="header", content="My Header")

        # Combined header and footer in one call:
        update_doc_headers_footers(doc_id, header_content="My Header", footer_content="My Footer")

        # Just header using new parameter:
        update_doc_headers_footers(doc_id, header_content="My Header")
    """
    # Resolve parameter alias: 'text' is an alias for 'content'
    if text is not None and content is None:
        content = text

    logger.info(f"[update_doc_headers_footers] Doc={document_id}, type={section_type}, header_content={header_content is not None}, footer_content={footer_content is not None}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Determine which mode we're in: combined (header_content/footer_content) or single (section_type+content)
    use_combined_mode = header_content is not None or footer_content is not None

    if use_combined_mode:
        # Combined mode: use header_content and/or footer_content
        # Cannot mix with section_type+content
        if section_type is not None and content is not None:
            return validator.create_invalid_param_error(
                param_name="parameters",
                received="both section_type+content and header_content/footer_content",
                valid_values=["Use either section_type+content OR header_content/footer_content, not both"],
                context="Choose one usage mode: single section (section_type+content) or combined (header_content/footer_content)",
            )

        # Validate header_footer_type
        if header_footer_type not in ["DEFAULT", "FIRST_PAGE", "EVEN_PAGE"]:
            return validator.create_invalid_param_error(
                param_name="header_footer_type",
                received=header_footer_type,
                valid_values=["DEFAULT", "FIRST_PAGE", "EVEN_PAGE"],
            )

        # Validate that at least one content is provided (should always be true given how we got here)
        if header_content is None and footer_content is None:
            return validator.create_invalid_param_error(
                param_name="header_content/footer_content",
                received="both None",
                valid_values=["at least one of header_content or footer_content"],
            )

        # Use HeaderFooterManager to handle the complex logic
        header_footer_manager = HeaderFooterManager(service)

        results = []
        errors = []

        # Update header if provided
        if header_content is not None:
            is_valid, error_msg = validator.validate_text_content(header_content)
            if not is_valid:
                return validator.create_invalid_param_error(
                    param_name="header_content",
                    received=repr(header_content)[:50],
                    valid_values=["non-empty string"],
                    context=error_msg,
                )

            success, message = await header_footer_manager.update_header_footer_content(
                document_id, "header", header_content, header_footer_type, create_if_missing
            )
            if success:
                results.append("header")
            else:
                errors.append(f"header: {message}")

        # Update footer if provided
        if footer_content is not None:
            is_valid, error_msg = validator.validate_text_content(footer_content)
            if not is_valid:
                return validator.create_invalid_param_error(
                    param_name="footer_content",
                    received=repr(footer_content)[:50],
                    valid_values=["non-empty string"],
                    context=error_msg,
                )

            success, message = await header_footer_manager.update_header_footer_content(
                document_id, "footer", footer_content, header_footer_type, create_if_missing
            )
            if success:
                results.append("footer")
            else:
                errors.append(f"footer: {message}")

        link = f"https://docs.google.com/document/d/{document_id}/edit"

        if errors and not results:
            # All operations failed
            return validator.create_api_error(
                operation="update_header_footer",
                error_message="; ".join(errors),
                document_id=document_id,
            )
        elif errors:
            # Partial success
            return f"Partially updated: {', '.join(results)}. Errors: {'; '.join(errors)}. Link: {link}"
        else:
            # All succeeded
            return f"Updated {' and '.join(results)} in document {document_id}. Link: {link}"

    else:
        # Single section mode: original behavior with section_type + content
        if section_type is None:
            return validator.create_invalid_param_error(
                param_name="section_type",
                received="None",
                valid_values=["header", "footer"],
                context="Either use section_type+content, or use header_content/footer_content",
            )

        is_valid, error_msg = validator.validate_header_footer_params(
            section_type, header_footer_type
        )
        if not is_valid:
            if "section_type" in error_msg.lower():
                return validator.create_invalid_param_error(
                    param_name="section_type",
                    received=section_type,
                    valid_values=["header", "footer"],
                )
            else:
                return validator.create_invalid_param_error(
                    param_name="header_footer_type",
                    received=header_footer_type,
                    valid_values=["DEFAULT", "FIRST_PAGE", "EVEN_PAGE"],
                )

        is_valid, error_msg = validator.validate_text_content(content)
        if not is_valid:
            return validator.create_invalid_param_error(
                param_name="content",
                received=repr(content)[:50],
                valid_values=["non-empty string"],
                context=error_msg,
            )

        # Use HeaderFooterManager to handle the complex logic
        header_footer_manager = HeaderFooterManager(service)

        success, message = await header_footer_manager.update_header_footer_content(
            document_id, section_type, content, header_footer_type, create_if_missing
        )

        if success:
            link = f"https://docs.google.com/document/d/{document_id}/edit"
            return f"{message}. Link: {link}"
        else:
            return validator.create_api_error(
                operation="update_header_footer",
                error_message=message,
                document_id=document_id,
            )


@server.tool()
@handle_http_errors("get_doc_headers_footers", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def get_doc_headers_footers(
    service: Any,
    user_google_email: str,
    document_id: str,
    section_type: str = None,
) -> str:
    """
    Get the content of headers and/or footers in a Google Doc.

    Use this to read the current text content of document headers and footers
    before updating them, or to check if headers/footers exist.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to read
        section_type: Optional filter - "header" to get only headers, "footer" to get
            only footers, or None/omit to get both (default: both)

    Returns:
        str: JSON-formatted string containing header/footer content and metadata

    Example response:
        {
            "has_headers": true,
            "has_footers": true,
            "headers": {
                "kix.abc123": {
                    "type": "DEFAULT",
                    "content": "My Document Header",
                    "is_empty": false
                }
            },
            "footers": {
                "kix.xyz789": {
                    "type": "DEFAULT",
                    "content": "Page 1",
                    "is_empty": false
                }
            }
        }
    """
    import json

    logger.info(f"[get_doc_headers_footers] Doc={document_id}, section_type={section_type}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    if section_type is not None and section_type not in ["header", "footer"]:
        return validator.create_invalid_param_error(
            param_name="section_type",
            received=section_type,
            valid_values=["header", "footer", "None (for both)"],
        )

    # Use HeaderFooterManager to get the information
    header_footer_manager = HeaderFooterManager(service)
    info = await header_footer_manager.get_header_footer_info(document_id)

    if "error" in info:
        return validator.create_api_error(
            operation="get_headers_footers",
            error_message=info["error"],
            document_id=document_id,
        )

    # Extract full text content for each header/footer
    doc = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    result = {
        "has_headers": info.get("has_headers", False),
        "has_footers": info.get("has_footers", False),
        "headers": {},
        "footers": {},
    }

    # Process headers if requested (or if no filter specified)
    if section_type is None or section_type == "header":
        for header_id, header_data in doc.get("headers", {}).items():
            content = _extract_section_text(header_data)
            result["headers"][header_id] = {
                "type": _infer_header_footer_type(header_id),
                "content": content,
                "is_empty": not content.strip(),
            }

    # Process footers if requested (or if no filter specified)
    if section_type is None or section_type == "footer":
        for footer_id, footer_data in doc.get("footers", {}).items():
            content = _extract_section_text(footer_data)
            result["footers"][footer_id] = {
                "type": _infer_header_footer_type(footer_id),
                "content": content,
                "is_empty": not content.strip(),
            }

    # Filter result based on section_type
    if section_type == "header":
        del result["footers"]
        del result["has_footers"]
    elif section_type == "footer":
        del result["headers"]
        del result["has_headers"]

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return json.dumps(result, indent=2) + f"\n\nDocument link: {link}"


def _extract_section_text(section_data: dict) -> str:
    """Extract plain text content from a header/footer section."""
    text_content = ""
    for element in section_data.get("content", []):
        if "paragraph" in element:
            para = element["paragraph"]
            for para_element in para.get("elements", []):
                if "textRun" in para_element:
                    text_content += para_element["textRun"].get("content", "")
    # Remove trailing newline that Google Docs adds
    return text_content.rstrip("\n")


def _infer_header_footer_type(section_id: str) -> str:
    """Infer the header/footer type from its section ID."""
    section_id_lower = section_id.lower()
    if "first" in section_id_lower:
        return "FIRST_PAGE"
    elif "even" in section_id_lower:
        return "EVEN_PAGE"
    else:
        return "DEFAULT"


@server.tool()
@handle_http_errors("batch_edit_doc", service_type="docs")
@require_google_service("docs", "docs_write")
async def batch_edit_doc(
    service: Any,
    user_google_email: str,
    document_id: str,
    operations: List[Dict[str, Any]],
    auto_adjust_positions: bool = True,
    preview: bool = False,
    tab_id: str = None,
) -> str:
    """
    Execute multiple document operations atomically with search-based positioning.

    This is the RECOMMENDED batch editing tool for all batch operations.

    Features:
    - Search-based positioning (insert before/after search text)
    - Automatic position adjustment for sequential operations
    - Per-operation results with position shift tracking
    - Atomic execution (all succeed or all fail)
    - Accepts BOTH naming conventions (e.g., "insert" or "insert_text")
    - Preview mode to see what would change without modifying the document

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        operations: List of operations. Each operation can use either:
            - Location-based: {"type": "insert", "location": "end", "text": "Appended text"}
            - Index-based: {"type": "insert_text", "index": 100, "text": "Hello"}
            - Search-based: {"type": "insert", "search": "Conclusion", "position": "before", "text": "New text"}

        auto_adjust_positions: If True (default), automatically adjusts positions
            for subsequent operations based on cumulative shifts from earlier operations.
        preview: If True, returns what would be modified without actually executing
            the operations. Useful for validating operations before committing changes.
        tab_id: Optional tab ID for multi-tab documents. If not provided, operates on
            the first tab (default behavior). Use list_doc_tabs() to discover tab IDs.

    Supported operation types (both naming styles work):
        - insert / insert_text: Insert text at position
        - delete / delete_text: Delete text range
        - replace / replace_text: Replace text range with new text
        - format / format_text: Apply formatting (bold, italic, underline, font_size, font_family)
        - insert_table: Insert table at position (requires: index/location, rows, columns)
          Example: {"type": "insert_table", "location": "end", "rows": 3, "columns": 4}
        - insert_page_break: Insert page break
        - find_replace: Find and replace all occurrences

    Location-based positioning options:
        - location: "start" (insert at document beginning) or "end" (append at document end)

    Search-based positioning options:
        - search: Text to find in the document
        - position: "before" (insert before match), "after" (insert after), "replace" (replace match)
        - occurrence: Which occurrence to target (1=first, 2=second, -1=last)
        - all_occurrences: If True, apply operation to ALL occurrences (not just first)
        - match_case: Whether to match case exactly (default: True)
        - extend: Extend position to boundary: "paragraph", "sentence", or "line"
          When extend is used:
          - position="after" + extend="paragraph": Insert after the END of the paragraph
          - position="before" + extend="paragraph": Insert before the START of the paragraph
          - position="replace" + extend="paragraph": Replace the ENTIRE paragraph

    Example operations:
        [
            {"type": "insert", "search": "Conclusion", "position": "before",
             "text": "\\n\\nNew section content.\\n"},
            {"type": "format", "search": "Important Note", "position": "replace",
             "bold": True, "font_size": 14},
            {"type": "insert_text", "index": 1, "text": "Header text\\n"},
            {"type": "find_replace", "find_text": "old term", "replace_text": "new term"}
        ]

    Example with location-based positioning (append at end or insert at start):
        [
            {"type": "insert", "location": "end", "text": "\\n[Appended at end]"},
            {"type": "insert", "location": "start", "text": "[Prepended at start]\\n"}
        ]

    Example with all_occurrences (format ALL matching text):
        [
            {"type": "format", "search": "TODO", "all_occurrences": True,
             "bold": True, "foreground_color": "red"},
            {"type": "replace", "search": "DRAFT", "all_occurrences": True,
             "text": "FINAL"}
        ]

    Example with extend (insert after entire paragraph, not just after matched text):
        [
            {"type": "insert", "search": "Introduction", "position": "after",
             "extend": "paragraph", "text": "\\n\\nNew paragraph after Introduction."},
            {"type": "delete", "search": "deprecated", "position": "replace",
             "extend": "sentence"}
        ]

    Returns:
        JSON string with detailed results including:
        - success: Overall success status
        - operations_completed: Number of operations executed
        - results: Per-operation details with position shifts
        - total_position_shift: Cumulative position change
        - document_link: Link to edited document

        Example response:
        {
            "success": true,
            "operations_completed": 2,
            "total_operations": 2,
            "results": [
                {
                    "index": 0,
                    "type": "insert",
                    "success": true,
                    "description": "insert 'New section...' at 150",
                    "position_shift": 20,
                    "affected_range": {"start": 150, "end": 170},
                    "resolved_index": 150
                },
                {
                    "index": 1,
                    "type": "format",
                    "success": true,
                    "description": "format 180-195 (bold=True)",
                    "position_shift": 0,
                    "affected_range": {"start": 180, "end": 195}
                }
            ],
            "total_position_shift": 20,
            "message": "Successfully executed 2 operation(s)",
            "document_link": "https://docs.google.com/document/d/.../edit"
        }

        Using position_shift for chained operations:
        ```python
        # First operation at index 100 inserts 15 chars
        result1 = batch_edit_doc(doc_id, [{"type": "insert", "index": 100, "text": "15 char string."}])
        # result1["results"][0]["position_shift"] = 15

        # For a subsequent edit originally targeting index 200:
        # new_index = 200 + result1["total_position_shift"] = 215
        ```

        Preview mode - see what would change without modifying:
        ```python
        result = batch_edit_doc(doc_id, operations, preview=True)
        # result["preview"] = true
        # result["would_modify"] = true
        # result["operations"] = [detailed preview of each operation]
        ```
    """
    import json

    logger.debug(
        f"[batch_edit_doc] Doc={document_id}, operations={len(operations)}, preview={preview}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, error_msg = validator.validate_document_id(document_id)
    if not is_valid:
        return json.dumps({"success": False, "error": error_msg}, indent=2)

    if not operations or not isinstance(operations, list):
        return json.dumps(
            {"success": False, "error": "Operations must be a non-empty list"}, indent=2
        )

    # Use BatchOperationManager with enhanced search support
    batch_manager = BatchOperationManager(service, tab_id=tab_id)

    result = await batch_manager.execute_batch_with_search(
        document_id,
        operations,
        auto_adjust_positions=auto_adjust_positions,
        preview_only=preview,
    )

    return json.dumps(result.to_dict(), indent=2)


# Detail level type for get_doc_info
DetailLevel = Literal["summary", "structure", "tables", "headings", "all"]


@server.tool()
@handle_http_errors("get_doc_info", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def get_doc_info(
    service: Any,
    user_google_email: str,
    document_id: str,
    detail: DetailLevel = "all",
) -> str:
    """
    Get comprehensive document information with configurable detail level.

    This is the recommended tool for understanding document structure before
    making modifications. It combines information from document analysis
    and structural elements into a single response.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to analyze
        detail: Level of detail to return:
            - "summary": Quick stats, element counts, safe insertion indices (fast)
            - "structure": Full element hierarchy with positions
            - "tables": Table-focused view with dimensions and positions
            - "headings": Just the headings outline for navigation
            - "all": Everything combined (default)

    Returns:
        str: JSON containing document information based on requested detail level

    Example Usage:
        # Quick check before table insertion:
        get_doc_info(doc_id, detail="summary")

        # Navigate document structure:
        get_doc_info(doc_id, detail="headings")

        # Full analysis:
        get_doc_info(doc_id)  # Returns all info
    """
    import json

    logger.debug(f"[get_doc_info] Doc={document_id}, detail={detail}")

    # Input validation
    validator = ValidationManager()
    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Get the document once
    doc = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    result = {
        "title": doc.get("title", "Untitled"),
        "document_id": document_id,
        "detail_level": detail,
    }

    # Summary info (always included for context)
    complexity = analyze_document_complexity(doc)
    result["total_length"] = complexity.get("total_length", 1)
    result["safe_insertion_index"] = max(1, complexity.get("total_length", 1) - 1)

    if detail in ("summary", "all"):
        # Quick stats and safe insertion indices
        result["statistics"] = {
            "total_elements": complexity.get("total_elements", 0),
            "paragraphs": complexity.get("paragraphs", 0),
            "tables": complexity.get("tables", 0),
            "lists": complexity.get("lists", 0),
            "complexity_score": complexity.get("complexity_score", "simple"),
        }

    if detail in ("structure", "all"):
        # Full element hierarchy
        structure = parse_document_structure(doc)
        elements = []
        for element in structure["body"]:
            elem_info = {
                "type": element["type"],
                "start_index": element["start_index"],
                "end_index": element["end_index"],
            }
            if element["type"] == "paragraph":
                elem_info["text_preview"] = element.get("text", "")[:100]
            elif element["type"] == "table":
                elem_info["rows"] = element["rows"]
                elem_info["columns"] = element["columns"]
            elements.append(elem_info)
        result["elements"] = elements

    if detail in ("tables", "all"):
        # Table-focused view
        tables = find_tables(doc)
        table_list = []
        for i, table in enumerate(tables):
            table_data = extract_table_as_data(table)
            table_list.append(
                {
                    "index": i,
                    "position": {
                        "start": table["start_index"],
                        "end": table["end_index"],
                    },
                    "dimensions": {"rows": table["rows"], "columns": table["columns"]},
                    "preview": table_data[:3] if table_data else [],  # First 3 rows
                }
            )
        result["tables"] = table_list if table_list else []

    if detail in ("headings", "all"):
        # Headings outline for navigation
        all_elements = extract_structural_elements(doc)
        headings_outline = build_headings_outline(all_elements)
        result["headings_outline"] = _clean_outline(headings_outline)

        # Also include flat heading list for quick reference
        headings = [e for e in all_elements if e["type"].startswith("heading")]
        result["headings"] = [
            {
                "level": h.get("level", 0),
                "text": h.get("text", ""),
                "start_index": h["start_index"],
                "end_index": h["end_index"],
            }
            for h in headings
        ]

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Document info for {document_id} (detail={detail}):\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


def _clean_outline(outline: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Clean outline for JSON output, removing internal fields."""
    clean = []
    for item in outline:
        clean_item = {
            "level": item["level"],
            "text": item["text"],
            "start_index": item["start_index"],
            "end_index": item["end_index"],
            "children": _clean_outline(item.get("children", [])),
        }
        clean.append(clean_item)
    return clean


@server.tool()
@handle_http_errors("get_doc_section", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def get_doc_section(
    service: Any,
    user_google_email: str,
    document_id: str,
    heading: str,
    include_subsections: bool = True,
    match_case: bool = False,
) -> str:
    """
    Get content of a specific section by heading.

    A section includes all content from the heading until the next heading
    of the same or higher level (smaller number = higher level).

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to read
        heading: The heading text to search for
        include_subsections: Whether to include nested subsection headings in the response
        match_case: Whether to match case exactly when finding the heading

    Returns:
        str: JSON containing:
            - heading: The matched heading text
            - level: The heading level (1-6)
            - start_index: Start position of the section
            - end_index: End position of the section
            - content: Full text content of the section
            - subsections: List of subsection headings (if include_subsections=True)

    Example:
        get_doc_section(document_id="...", heading="The Problem")

        Response:
        {
            "heading": "The Problem",
            "level": 2,
            "start_index": 151,
            "end_index": 450,
            "content": "Full text of the section...",
            "subsections": [
                {"heading": "Sub-problem 1", "level": 3, "start_index": 200, "end_index": 300}
            ]
        }
    """
    logger.debug(f"[get_doc_section] Doc={document_id}, heading={heading}")

    # Input validation
    validator = ValidationManager()
    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    if not heading or not heading.strip():
        return validator.create_missing_param_error(
            param_name="heading",
            context="for section retrieval",
            valid_values=["non-empty heading text"],
        )

    # Get the document
    doc = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Find the section
    section = find_section_by_heading(doc, heading, match_case)

    if section is None:
        # Provide helpful error with available headings
        all_headings = get_all_headings(doc)
        heading_list = [h["text"] for h in all_headings] if all_headings else []
        return validator.create_heading_not_found_error(
            heading=heading, available_headings=heading_list, match_case=match_case
        )

    # Build result
    result = {
        "heading": section["heading"],
        "level": section["level"],
        "start_index": section["start_index"],
        "end_index": section["end_index"],
        "content": section["content"],
    }

    if include_subsections:
        result["subsections"] = section["subsections"]

    import json

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Section '{heading}' in document {document_id}:\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


@server.tool()
@handle_http_errors("delete_doc_section", service_type="docs")
@require_google_service("docs", "docs_write")
async def delete_doc_section(
    service: Any,
    user_google_email: str,
    document_id: str,
    heading: str,
    include_heading: bool = True,
    match_case: bool = False,
    preview: bool = False,
) -> str:
    """
    Delete a section from a Google Doc by heading.

    A section includes all content from the heading until the next heading
    of the same or higher level (smaller number = higher level).

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to modify
        heading: The heading text to search for
        include_heading: Whether to delete the heading itself (default True).
            Set to False to keep the heading but delete its content.
        match_case: Whether to match case exactly when finding the heading
        preview: If True, shows what would be deleted without modifying the document

    Returns:
        str: JSON containing:
            - deleted: True if section was deleted (or would be in preview mode)
            - heading: The matched heading text
            - level: The heading level (1-6)
            - start_index: Start position of the deleted range
            - end_index: End position of the deleted range
            - characters_deleted: Number of characters removed
            - subsections_deleted: Number of subsection headings that were removed
            - preview: True if this was a preview operation

    Example:
        # Delete entire section including heading
        delete_doc_section(document_id="...", heading="Old Section")

        # Delete section content but keep the heading
        delete_doc_section(
            document_id="...", heading="Section Name", include_heading=False
        )

        # Preview what would be deleted
        delete_doc_section(document_id="...", heading="Section Name", preview=True)
    """
    logger.debug(
        f"[delete_doc_section] Doc={document_id}, heading={heading}, "
        f"include_heading={include_heading}"
    )

    # Input validation
    validator = ValidationManager()
    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    if not heading or not heading.strip():
        return validator.create_missing_param_error(
            param_name="heading",
            context="for section deletion",
            valid_values=["non-empty heading text"],
        )

    # Get the document
    doc = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Find the section
    section = find_section_by_heading(doc, heading, match_case)

    if section is None:
        # Provide helpful error with available headings
        all_headings = get_all_headings(doc)
        heading_list = [h["text"] for h in all_headings] if all_headings else []
        return validator.create_heading_not_found_error(
            heading=heading, available_headings=heading_list, match_case=match_case
        )

    # Determine deletion range
    if include_heading:
        delete_start = section["start_index"]
    else:
        # Find where the heading ends to keep it
        elements = extract_structural_elements(doc)
        heading_end = section["start_index"]
        for elem in elements:
            if elem["type"].startswith("heading") or elem["type"] == "title":
                elem_text = elem["text"] if match_case else elem["text"].lower()
                search_text = heading if match_case else heading.lower()
                if elem_text.strip() == search_text.strip():
                    heading_end = elem["end_index"]
                    break
        delete_start = heading_end

    delete_end = section["end_index"]

    # Check if the section extends to the end of the document body.
    # Google Docs API doesn't allow deleting the final newline character that
    # terminates the document body segment. If we're deleting to the end,
    # we need to exclude that final newline (subtract 1 from end_index).
    body = doc.get("body", {})
    body_content = body.get("content", [])
    if body_content:
        doc_end_index = body_content[-1].get("endIndex", 0)
        if delete_end == doc_end_index:
            delete_end = delete_end - 1
            logger.debug(
                f"[delete_doc_section] Section at document end, "
                f"adjusted delete_end from {doc_end_index} to {delete_end}"
            )

    # Calculate characters to be deleted
    characters_to_delete = delete_end - delete_start
    subsections_count = len(section.get("subsections", []))

    import json

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    # Build result
    result = {
        "heading": section["heading"],
        "level": section["level"],
        "start_index": delete_start,
        "end_index": delete_end,
        "characters_deleted": characters_to_delete,
        "subsections_deleted": subsections_count,
        "include_heading": include_heading,
        "link": link,
    }

    # Handle preview mode
    if preview:
        result["preview"] = True
        result["deleted"] = False
        result["would_delete"] = True
        content = section["content"]
        suffix = "..." if len(content) > 500 else ""
        result["content_preview"] = content[:500] + suffix
        return (
            f"Preview - Would delete section '{heading}':\n\n"
            f"{json.dumps(result, indent=2)}"
        )

    # Perform the deletion
    if characters_to_delete > 0:
        requests = [create_delete_range_request(delete_start, delete_end)]
        await asyncio.to_thread(
            service.documents()
            .batchUpdate(documentId=document_id, body={"requests": requests})
            .execute
        )

    result["deleted"] = True
    result["preview"] = False

    return (
        f"Deleted section '{heading}' from document {document_id}:\n\n"
        f"{json.dumps(result, indent=2)}\n\nLink: {link}"
    )


@server.tool()
@handle_http_errors("create_table_with_data", service_type="docs")
@require_google_service("docs", "docs_write")
async def create_table_with_data(
    service: Any,
    user_google_email: str,
    document_id: str,
    table_data: List[List[str]],
    index: int = None,
    after_heading: str = None,
    location: str = None,
    bold_headers: bool = True,
) -> str:
    """
    Creates a table and populates it with data in one reliable operation.

    SIMPLIFIED USAGE - No pre-flight call needed:
    - Use location='end' to append table at end of document (recommended)
    - Use location='start' to insert table at beginning of document
    - Use 'after_heading' to insert table after a specific heading
    - Provide explicit 'index' only for precise positioning

    EXAMPLE DATA FORMAT:
    table_data = [
        ["Header1", "Header2", "Header3"],    # Row 0 - headers
        ["Data1", "Data2", "Data3"],          # Row 1 - first data row
        ["Data4", "Data5", "Data6"]           # Row 2 - second data row
    ]

    USAGE PATTERNS:
    1. Append to end: create_table_with_data(doc_id, data, location="end")
    2. Insert at start: create_table_with_data(doc_id, data, location="start")
    3. After heading: create_table_with_data(doc_id, data, after_heading="Data Section")
    4. Explicit index: create_table_with_data(doc_id, data, index=42)

    DATA FORMAT REQUIREMENTS:
    - Must be 2D list of strings only
    - Each inner list = one table row
    - All rows MUST have same number of columns
    - Use empty strings "" for empty cells, never None
    - Use debug_table_structure after creation to verify results

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        table_data: 2D list of strings [["col1", "col2"], ["row1", "row2"]]
        index: Document position (optional, mutually exclusive with location/after_heading)
        after_heading: Insert after this heading (optional, case-insensitive)
        location: Semantic location - "start" or "end" (mutually exclusive with index/after_heading)
        bold_headers: Whether to make first row bold (default: true)

    Returns:
        str: Confirmation with table details and link
    """
    logger.debug(
        f"[create_table_with_data] Doc={document_id}, "
        f"index={index}, after_heading={after_heading}, location={location}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, error_msg = validator.validate_document_id(document_id)
    if not is_valid:
        return f"ERROR: {error_msg}"

    is_valid, error_msg = validator.validate_table_data(table_data)
    if not is_valid:
        return f"ERROR: {error_msg}"

    # Validate location parameter if provided
    if location is not None and location not in ["start", "end"]:
        return validator.create_invalid_param_error(
            param_name="location", received=location, valid_values=["start", "end"]
        )

    # Check for mutually exclusive positioning parameters
    positioning_params = [
        ("index", index),
        ("after_heading", after_heading),
        ("location", location),
    ]
    provided_params = [name for name, value in positioning_params if value is not None]
    if len(provided_params) > 1:
        return (
            f"ERROR: Cannot specify multiple positioning parameters. "
            f"Got: {', '.join(provided_params)}. Use only one of: index, after_heading, or location."
        )

    # Resolve insertion index
    resolved_index = index
    location_description = None

    if index is not None:
        # Explicit index provided - validate it
        is_valid, error_msg = validator.validate_index(index, "Index")
        if not is_valid:
            return f"ERROR: {error_msg}"
        location_description = f"at explicit index {index}"
    else:
        # Auto-detect insertion point - fetch document first
        try:
            doc_data = await asyncio.to_thread(
                service.documents().get(documentId=document_id).execute
            )
        except Exception as e:
            return f"ERROR: Failed to fetch document for index calculation: {str(e)}"

        if after_heading is not None:
            # Insert after specified heading
            insertion_point = find_section_insertion_point(
                doc_data, after_heading, position="end"
            )
            if insertion_point is None:
                # List available headings to help user
                headings = get_all_headings(doc_data)
                if headings:
                    heading_list = ", ".join([f'"{h["text"]}"' for h in headings[:5]])
                    more = (
                        f" (and {len(headings) - 5} more)" if len(headings) > 5 else ""
                    )
                    return (
                        f"ERROR: Heading '{after_heading}' not found. "
                        f"Available headings: {heading_list}{more}"
                    )
                else:
                    return (
                        f"ERROR: Heading '{after_heading}' not found. "
                        "Document has no headings."
                    )
            resolved_index = insertion_point
            location_description = f"after heading '{after_heading}'"
        elif location == "start":
            # Insert at document start (after initial section break)
            resolved_index = 1
            location_description = "at start of document"
        else:
            # Default or location='end': append to end of document
            structure = parse_document_structure(doc_data)
            total_length = structure["total_length"]
            # Use total_length - 1 for safe insertion point
            resolved_index = total_length - 1 if total_length > 1 else 1
            location_description = "at end of document"

    logger.debug(
        f"[create_table_with_data] Resolved index: {resolved_index} "
        f"({location_description})"
    )

    # Use TableOperationManager to handle the complex logic
    table_manager = TableOperationManager(service)

    # Try to create table; retry with index-1 if at document boundary
    success, message, metadata = await table_manager.create_and_populate_table(
        document_id, table_data, resolved_index, bold_headers
    )

    # If it failed due to index at/beyond document end, retry with adjusted idx
    if not success and "must be less than the end index" in message:
        logger.debug(
            f"Index {resolved_index} is at document boundary, "
            f"retrying with index {resolved_index - 1}"
        )
        success, message, metadata = await table_manager.create_and_populate_table(
            document_id, table_data, resolved_index - 1, bold_headers
        )

    if success:
        link = f"https://docs.google.com/document/d/{document_id}/edit"
        rows = metadata.get("rows", 0)
        columns = metadata.get("columns", 0)

        return (
            f"SUCCESS: {message}. Table: {rows}x{columns}, "
            f"inserted {location_description}. Link: {link}"
        )
    else:
        return f"ERROR: {message}"


@server.tool()
@handle_http_errors("debug_table_structure", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def debug_table_structure(
    service: Any,
    user_google_email: str,
    document_id: str,
    table_index: int = 0,
) -> str:
    """
    ESSENTIAL DEBUGGING TOOL - Use this whenever tables don't work as expected.

    USE THIS IMMEDIATELY WHEN:
    - Table population put data in wrong cells
    - You get "table not found" errors
    - Data appears concatenated in first cell
    - Need to understand existing table structure
    - Planning to use populate_existing_table

    WHAT THIS SHOWS YOU:
    - Exact table dimensions (rows × columns)
    - Each cell's position coordinates (row,col)
    - Current content in each cell
    - Insertion indices for each cell
    - Table boundaries and ranges

    HOW TO READ THE OUTPUT:
    - "dimensions": "2x3" = 2 rows, 3 columns
    - "position": "(0,0)" = first row, first column
    - "current_content": What's actually in each cell right now
    - "insertion_index": Where new text would be inserted in that cell

    WORKFLOW INTEGRATION:
    1. After creating table → Use this to verify structure
    2. Before populating → Use this to plan your data format
    3. After population fails → Use this to see what went wrong
    4. When debugging → Compare your data array to actual table structure

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to inspect
        table_index: Which table to debug. Supports:
            - Positive indices: 0 = first table, 1 = second table, etc.
            - Negative indices: -1 = last table, -2 = second-to-last, etc.

    Returns:
        str: Detailed JSON structure showing table layout, cell positions, and current content
    """
    logger.debug(
        f"[debug_table_structure] Doc={document_id}, table_index={table_index}"
    )

    # Get the document
    doc = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Find tables
    tables = find_tables(doc)

    # Handle negative indices (Python-style: -1 = last, -2 = second-to-last)
    original_index = table_index
    if table_index < 0:
        table_index = len(tables) + table_index

    # Validate index is in bounds
    if table_index < 0 or table_index >= len(tables):
        validator = ValidationManager()
        return validator.create_table_not_found_error(
            table_index=original_index, total_tables=len(tables)
        )

    table_info = tables[table_index]

    import json

    # Extract detailed cell information
    debug_info = {
        "table_index": table_index,
        "total_tables": len(tables),
        "dimensions": f"{table_info['rows']}x{table_info['columns']}",
        "table_range": f"[{table_info['start_index']}-{table_info['end_index']}]",
        "cells": [],
    }
    # Include original index if it was negative (for clarity)
    if original_index < 0:
        debug_info["requested_index"] = original_index

    for row_idx, row in enumerate(table_info["cells"]):
        row_info = []
        for col_idx, cell in enumerate(row):
            row_span = cell.get("row_span", 1)
            column_span = cell.get("column_span", 1)

            cell_debug = {
                "position": f"({row_idx},{col_idx})",
                "range": f"[{cell['start_index']}-{cell['end_index']}]",
                "insertion_index": cell.get("insertion_index", "N/A"),
                "current_content": repr(cell.get("content", "")),
                "content_elements_count": len(cell.get("content_elements", [])),
            }

            # Add span info if cell is merged (spans > 1 cell)
            if row_span > 1 or column_span > 1:
                cell_debug["merged"] = True
                cell_debug["spans"] = f"{row_span}x{column_span} (rows x cols)"
            elif row_span == 0 or column_span == 0:
                # This cell is covered by another merged cell
                cell_debug["covered_by_merge"] = True

            row_info.append(cell_debug)
        debug_info["cells"].append(row_info)

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Table structure debug for table {table_index}:\n\n{json.dumps(debug_info, indent=2)}\n\nLink: {link}"


@server.tool()
@handle_http_errors("modify_table", service_type="docs")
@require_google_service("docs", "docs_write")
async def modify_table(
    service: Any,
    user_google_email: str,
    document_id: str,
    table_index: int = 0,
    operations: List[Dict[str, Any]] = None,
) -> str:
    """
    Modify an existing table's structure by adding/removing rows/columns or updating cell content.

    This tool supports multiple operations to modify table structure:
    - insert_row: Add a new row to the table
    - delete_row: Remove a row from the table
    - insert_column: Add a new column to the table
    - delete_column: Remove a column from the table
    - update_cell: Update content of a specific cell
    - merge_cells: Merge multiple cells into one
    - unmerge_cells: Split merged cells back into individual cells

    OPERATION FORMATS:

    1. insert_row:
       {"action": "insert_row", "row": 1, "insert_below": true}
       - row: Reference row index (0-based)
       - insert_below: If true, insert below the row; if false, insert above (default: true)

    2. delete_row:
       {"action": "delete_row", "row": 2}
       - row: Row index to delete (0-based)

    3. insert_column:
       {"action": "insert_column", "column": 1, "insert_right": true}
       - column: Reference column index (0-based)
       - insert_right: If true, insert to the right; if false, insert to left (default: true)

    4. delete_column:
       {"action": "delete_column", "column": 2}
       - column: Column index to delete (0-based)

    5. update_cell:
       {"action": "update_cell", "row": 0, "column": 1, "text": "New Value"}
       - row, column: Cell coordinates (0-based)
       - text: New text content for the cell (replaces existing content)

    6. delete_table:
       {"action": "delete_table"}
       - Deletes the entire table from the document
       - No additional parameters required
       - WARNING: This operation cannot be undone via this API

    7. merge_cells:
       {"action": "merge_cells", "row": 0, "column": 0, "row_span": 1, "column_span": 2}
       - row, column: Upper-left cell coordinates of the merge area (0-based)
       - row_span: Number of rows to merge (minimum 1)
       - column_span: Number of columns to merge (minimum 1)
       - Text content from all merged cells is concatenated in the resulting cell
       - NOTE: row_span=1 and column_span=1 means single cell (no merge)

    8. unmerge_cells:
       {"action": "unmerge_cells", "row": 0, "column": 0, "row_span": 1, "column_span": 2}
       - row, column: Upper-left cell coordinates of the area to unmerge (0-based)
       - row_span: Number of rows in the range
       - column_span: Number of columns in the range
       - If cells are not merged, this operation has no effect

    9. format_cell:
       {"action": "format_cell", "row": 0, "column": 0, "background_color": "#FF0000"}
       - row, column: Cell coordinates (0-based)
       - row_span: Number of rows to format (default: 1)
       - column_span: Number of columns to format (default: 1)
       - background_color: Cell background color (hex like "#FF0000" or named like "red")
       - border_color: Border color for all sides (hex or named)
       - border_width: Border width in points
       - border_dash_style: SOLID, DOT, DASH, DASH_DOT, LONG_DASH, LONG_DASH_DOT
       - border_top, border_bottom, border_left, border_right: Override specific borders
         as dict with {color, width, dash_style}
       - padding_top, padding_bottom, padding_left, padding_right: Padding in points
       - content_alignment: Vertical content alignment (TOP, MIDDLE, BOTTOM)

    10. resize_column:
        {"action": "resize_column", "column": 0, "width": 100}
        - column: Column index (0-based) or list of column indices to resize
        - width: Width in points (minimum 5 points per Google API requirement)
        - width_type: "FIXED_WIDTH" (default) or "EVENLY_DISTRIBUTED"
        - Single column: {"action": "resize_column", "column": 0, "width": 150}
        - Multiple columns: {"action": "resize_column", "column": [0, 2], "width": 100}

    USAGE EXAMPLES:

    # Add a row at the end of a 3-row table
    operations=[{"action": "insert_row", "row": 2, "insert_below": true}]

    # Delete the second row
    operations=[{"action": "delete_row", "row": 1}]

    # Add a column to the right of the first column
    operations=[{"action": "insert_column", "column": 0, "insert_right": true}]

    # Update a cell's content
    operations=[{"action": "update_cell", "row": 1, "column": 2, "text": "Updated"}]

    # Merge first row cells to create a header spanning 3 columns
    operations=[{"action": "merge_cells", "row": 0, "column": 0, "row_span": 1, "column_span": 3}]

    # Unmerge previously merged cells
    operations=[{"action": "unmerge_cells", "row": 0, "column": 0, "row_span": 1, "column_span": 3}]

    # Format a cell with background color
    operations=[{"action": "format_cell", "row": 0, "column": 0, "background_color": "#FFFF00"}]

    # Format a cell with red borders
    operations=[{"action": "format_cell", "row": 1, "column": 1, "border_color": "red", "border_width": 2}]

    # Format multiple cells (2x3 range) with centered content
    operations=[{"action": "format_cell", "row": 0, "column": 0, "row_span": 2, "column_span": 3,
                 "background_color": "#E0E0E0", "content_alignment": "MIDDLE"}]

    # Format cell with custom top border only
    operations=[{"action": "format_cell", "row": 0, "column": 0,
                 "border_top": {"color": "blue", "width": 3, "dash_style": "SOLID"}}]

    # Resize a single column to 150 points
    operations=[{"action": "resize_column", "column": 0, "width": 150}]

    # Resize multiple columns at once
    operations=[{"action": "resize_column", "column": [0, 2], "width": 100}]

    # Multiple operations (executed sequentially)
    operations=[
        {"action": "insert_row", "row": 0, "insert_below": true},
        {"action": "update_cell", "row": 1, "column": 0, "text": "New Row Data"}
    ]

    # Delete the entire table
    operations=[{"action": "delete_table"}]

    IMPORTANT NOTES:
    - Operations are executed sequentially in the order provided
    - After structural changes (insert/delete row/column), the table structure is
      refreshed before the next operation to ensure correct indices
    - Row and column indices are 0-based (first row = 0, first column = 0)
    - Use debug_table_structure before and after modifications to verify results
    - delete_table removes the entire table; any subsequent operations in the same
      call will fail as the table no longer exists
    - merge_cells requires a rectangular range; non-rectangular ranges will fail

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document containing the table
        table_index: Which table to modify. Supports:
            - Positive indices: 0 = first table (default), 1 = second table, etc.
            - Negative indices: -1 = last table, -2 = second-to-last, etc.
        operations: List of operation dictionaries describing modifications

    Returns:
        str: Summary of operations performed with success/failure status
    """
    logger.debug(
        f"[modify_table] Doc={document_id}, table_index={table_index}, operations={operations}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, error_msg = validator.validate_document_id(document_id)
    if not is_valid:
        return f"ERROR: {error_msg}"

    if not operations or not isinstance(operations, list):
        return "ERROR: 'operations' must be a non-empty list of operation dictionaries"

    # Valid operation actions
    valid_actions = {
        "insert_row",
        "delete_row",
        "insert_column",
        "delete_column",
        "update_cell",
        "delete_table",
        "merge_cells",
        "unmerge_cells",
        "format_cell",
        "resize_column",
    }

    # Validate all operations first
    for i, op in enumerate(operations):
        if not isinstance(op, dict):
            return f"ERROR: Operation {i} must be a dictionary, got {type(op).__name__}"

        action = op.get("action")
        if not action:
            return f"ERROR: Operation {i} missing required 'action' field"

        if action not in valid_actions:
            return (
                f"ERROR: Operation {i} has invalid action '{action}'. "
                f"Valid actions: {', '.join(sorted(valid_actions))}"
            )

        # Validate required fields per action
        if action == "insert_row":
            if "row" not in op:
                return f"ERROR: Operation {i} (insert_row) missing required 'row' field"
            if not isinstance(op["row"], int) or op["row"] < 0:
                return f"ERROR: Operation {i} (insert_row) 'row' must be a non-negative integer"

        elif action == "delete_row":
            if "row" not in op:
                return f"ERROR: Operation {i} (delete_row) missing required 'row' field"
            if not isinstance(op["row"], int) or op["row"] < 0:
                return f"ERROR: Operation {i} (delete_row) 'row' must be a non-negative integer"

        elif action == "insert_column":
            if "column" not in op:
                return f"ERROR: Operation {i} (insert_column) missing required 'column' field"
            if not isinstance(op["column"], int) or op["column"] < 0:
                return f"ERROR: Operation {i} (insert_column) 'column' must be a non-negative integer"

        elif action == "delete_column":
            if "column" not in op:
                return f"ERROR: Operation {i} (delete_column) missing required 'column' field"
            if not isinstance(op["column"], int) or op["column"] < 0:
                return f"ERROR: Operation {i} (delete_column) 'column' must be a non-negative integer"

        elif action == "update_cell":
            if "row" not in op or "column" not in op:
                return f"ERROR: Operation {i} (update_cell) missing required 'row' and/or 'column' fields"
            if not isinstance(op["row"], int) or op["row"] < 0:
                return f"ERROR: Operation {i} (update_cell) 'row' must be a non-negative integer"
            if not isinstance(op["column"], int) or op["column"] < 0:
                return f"ERROR: Operation {i} (update_cell) 'column' must be a non-negative integer"
            if "text" not in op:
                return (
                    f"ERROR: Operation {i} (update_cell) missing required 'text' field"
                )

        elif action == "merge_cells":
            if "row" not in op or "column" not in op:
                return f"ERROR: Operation {i} (merge_cells) missing required 'row' and/or 'column' fields"
            if not isinstance(op["row"], int) or op["row"] < 0:
                return f"ERROR: Operation {i} (merge_cells) 'row' must be a non-negative integer"
            if not isinstance(op["column"], int) or op["column"] < 0:
                return f"ERROR: Operation {i} (merge_cells) 'column' must be a non-negative integer"
            if "row_span" not in op or "column_span" not in op:
                return f"ERROR: Operation {i} (merge_cells) missing required 'row_span' and/or 'column_span' fields"
            if not isinstance(op["row_span"], int) or op["row_span"] < 1:
                return f"ERROR: Operation {i} (merge_cells) 'row_span' must be a positive integer (minimum 1)"
            if not isinstance(op["column_span"], int) or op["column_span"] < 1:
                return f"ERROR: Operation {i} (merge_cells) 'column_span' must be a positive integer (minimum 1)"

        elif action == "unmerge_cells":
            if "row" not in op or "column" not in op:
                return f"ERROR: Operation {i} (unmerge_cells) missing required 'row' and/or 'column' fields"
            if not isinstance(op["row"], int) or op["row"] < 0:
                return f"ERROR: Operation {i} (unmerge_cells) 'row' must be a non-negative integer"
            if not isinstance(op["column"], int) or op["column"] < 0:
                return f"ERROR: Operation {i} (unmerge_cells) 'column' must be a non-negative integer"
            if "row_span" not in op or "column_span" not in op:
                return f"ERROR: Operation {i} (unmerge_cells) missing required 'row_span' and/or 'column_span' fields"
            if not isinstance(op["row_span"], int) or op["row_span"] < 1:
                return f"ERROR: Operation {i} (unmerge_cells) 'row_span' must be a positive integer (minimum 1)"
            if not isinstance(op["column_span"], int) or op["column_span"] < 1:
                return f"ERROR: Operation {i} (unmerge_cells) 'column_span' must be a positive integer (minimum 1)"

        elif action == "format_cell":
            if "row" not in op or "column" not in op:
                return f"ERROR: Operation {i} (format_cell) missing required 'row' and/or 'column' fields"
            if not isinstance(op["row"], int) or op["row"] < 0:
                return f"ERROR: Operation {i} (format_cell) 'row' must be a non-negative integer"
            if not isinstance(op["column"], int) or op["column"] < 0:
                return f"ERROR: Operation {i} (format_cell) 'column' must be a non-negative integer"
            # Validate optional row_span and column_span
            if "row_span" in op and (not isinstance(op["row_span"], int) or op["row_span"] < 1):
                return f"ERROR: Operation {i} (format_cell) 'row_span' must be a positive integer (minimum 1)"
            if "column_span" in op and (not isinstance(op["column_span"], int) or op["column_span"] < 1):
                return f"ERROR: Operation {i} (format_cell) 'column_span' must be a positive integer (minimum 1)"
            # Validate content_alignment if provided
            if "content_alignment" in op:
                valid_alignments = {"TOP", "MIDDLE", "BOTTOM"}
                if op["content_alignment"] not in valid_alignments:
                    return f"ERROR: Operation {i} (format_cell) 'content_alignment' must be one of: {', '.join(valid_alignments)}"
            # Validate border_dash_style if provided
            if "border_dash_style" in op:
                valid_dash_styles = {"SOLID", "DOT", "DASH", "DASH_DOT", "LONG_DASH", "LONG_DASH_DOT"}
                if op["border_dash_style"] not in valid_dash_styles:
                    return f"ERROR: Operation {i} (format_cell) 'border_dash_style' must be one of: {', '.join(valid_dash_styles)}"
            # Check that at least one formatting option is provided
            format_options = [
                "background_color", "border_color", "border_width", "border_dash_style",
                "border_top", "border_bottom", "border_left", "border_right",
                "padding_top", "padding_bottom", "padding_left", "padding_right",
                "content_alignment"
            ]
            if not any(opt in op for opt in format_options):
                return f"ERROR: Operation {i} (format_cell) must specify at least one formatting option"

        elif action == "resize_column":
            if "column" not in op:
                return f"ERROR: Operation {i} (resize_column) missing required 'column' field"
            # Column can be int or list of ints
            col_val = op["column"]
            if isinstance(col_val, int):
                if col_val < 0:
                    return f"ERROR: Operation {i} (resize_column) 'column' must be a non-negative integer"
            elif isinstance(col_val, list):
                if not col_val:
                    return f"ERROR: Operation {i} (resize_column) 'column' list must not be empty"
                for idx, c in enumerate(col_val):
                    if not isinstance(c, int) or c < 0:
                        return f"ERROR: Operation {i} (resize_column) 'column[{idx}]' must be a non-negative integer"
            else:
                return f"ERROR: Operation {i} (resize_column) 'column' must be an integer or list of integers"
            if "width" not in op:
                return f"ERROR: Operation {i} (resize_column) missing required 'width' field"
            if not isinstance(op["width"], (int, float)) or op["width"] < 5:
                return f"ERROR: Operation {i} (resize_column) 'width' must be at least 5 points"
            # Validate width_type if provided
            if "width_type" in op:
                valid_width_types = {"FIXED_WIDTH", "EVENLY_DISTRIBUTED"}
                if op["width_type"] not in valid_width_types:
                    return f"ERROR: Operation {i} (resize_column) 'width_type' must be one of: {', '.join(valid_width_types)}"

    # Import helper functions for table operations
    from gdocs.docs_helpers import (
        create_insert_table_row_request,
        create_delete_table_row_request,
        create_insert_table_column_request,
        create_delete_table_column_request,
        create_delete_range_request,
        create_insert_text_request,
        create_merge_table_cells_request,
        create_unmerge_table_cells_request,
        create_update_table_cell_style_request,
        create_update_table_column_properties_request,
    )

    # Track results
    results = []

    # Store original index for error messages
    original_table_index = table_index

    # Execute operations sequentially
    for i, op in enumerate(operations):
        action = op["action"]

        try:
            # Refresh table structure before each operation
            doc = await asyncio.to_thread(
                service.documents().get(documentId=document_id).execute
            )
            tables = find_tables(doc)

            # Handle negative indices (Python-style: -1 = last, -2 = second-to-last)
            resolved_index = table_index
            if table_index < 0:
                resolved_index = len(tables) + table_index

            # Validate index is in bounds
            if resolved_index < 0 or resolved_index >= len(tables):
                return validator.create_table_not_found_error(
                    table_index=original_table_index, total_tables=len(tables)
                )

            table_info = tables[resolved_index]
            table_start = table_info["start_index"]
            num_rows = table_info["rows"]
            num_cols = table_info["columns"]

            if action == "insert_row":
                row_idx = op["row"]
                insert_below = op.get("insert_below", True)

                if row_idx >= num_rows:
                    results.append(
                        f"Op {i} (insert_row): FAILED - row {row_idx} out of bounds (table has {num_rows} rows)"
                    )
                    continue

                request = create_insert_table_row_request(
                    table_start_index=table_start,
                    row_index=row_idx,
                    insert_below=insert_below,
                )

                await asyncio.to_thread(
                    service.documents()
                    .batchUpdate(documentId=document_id, body={"requests": [request]})
                    .execute
                )

                position = "below" if insert_below else "above"
                results.append(
                    f"Op {i} (insert_row): SUCCESS - inserted row {position} row {row_idx}"
                )

            elif action == "delete_row":
                row_idx = op["row"]

                if row_idx >= num_rows:
                    results.append(
                        f"Op {i} (delete_row): FAILED - row {row_idx} out of bounds (table has {num_rows} rows)"
                    )
                    continue

                if num_rows <= 1:
                    results.append(
                        f"Op {i} (delete_row): FAILED - cannot delete the only row in the table"
                    )
                    continue

                request = create_delete_table_row_request(
                    table_start_index=table_start, row_index=row_idx
                )

                await asyncio.to_thread(
                    service.documents()
                    .batchUpdate(documentId=document_id, body={"requests": [request]})
                    .execute
                )

                results.append(f"Op {i} (delete_row): SUCCESS - deleted row {row_idx}")

            elif action == "insert_column":
                col_idx = op["column"]
                insert_right = op.get("insert_right", True)

                if col_idx >= num_cols:
                    results.append(
                        f"Op {i} (insert_column): FAILED - column {col_idx} out of bounds (table has {num_cols} columns)"
                    )
                    continue

                request = create_insert_table_column_request(
                    table_start_index=table_start,
                    row_index=0,
                    column_index=col_idx,
                    insert_right=insert_right,
                )

                await asyncio.to_thread(
                    service.documents()
                    .batchUpdate(documentId=document_id, body={"requests": [request]})
                    .execute
                )

                position = "right of" if insert_right else "left of"
                results.append(
                    f"Op {i} (insert_column): SUCCESS - inserted column {position} column {col_idx}"
                )

            elif action == "delete_column":
                col_idx = op["column"]

                if col_idx >= num_cols:
                    results.append(
                        f"Op {i} (delete_column): FAILED - column {col_idx} out of bounds (table has {num_cols} columns)"
                    )
                    continue

                if num_cols <= 1:
                    results.append(
                        f"Op {i} (delete_column): FAILED - cannot delete the only column in the table"
                    )
                    continue

                request = create_delete_table_column_request(
                    table_start_index=table_start, row_index=0, column_index=col_idx
                )

                await asyncio.to_thread(
                    service.documents()
                    .batchUpdate(documentId=document_id, body={"requests": [request]})
                    .execute
                )

                results.append(
                    f"Op {i} (delete_column): SUCCESS - deleted column {col_idx}"
                )

            elif action == "update_cell":
                row_idx = op["row"]
                col_idx = op["column"]
                new_text = op["text"]

                if row_idx >= num_rows:
                    results.append(
                        f"Op {i} (update_cell): FAILED - row {row_idx} out of bounds (table has {num_rows} rows)"
                    )
                    continue

                if col_idx >= num_cols:
                    results.append(
                        f"Op {i} (update_cell): FAILED - column {col_idx} out of bounds (table has {num_cols} columns)"
                    )
                    continue

                cells = table_info.get("cells", [])
                if row_idx >= len(cells) or col_idx >= len(cells[row_idx]):
                    results.append(
                        f"Op {i} (update_cell): FAILED - cell ({row_idx}, {col_idx}) not accessible"
                    )
                    continue

                cell = cells[row_idx][col_idx]
                cell_start = cell["start_index"]
                current_content = cell.get("content", "")

                # Build requests to replace cell content
                requests = []

                # Check if there's actual content to delete
                if current_content and len(current_content.strip()) > 0:
                    # Find actual text content boundaries (skip structural elements)
                    insertion_index = cell.get("insertion_index")
                    if insertion_index and current_content:
                        # Delete the text content
                        text_end = insertion_index + len(current_content.rstrip("\n"))
                        if text_end > insertion_index:
                            requests.append(
                                create_delete_range_request(insertion_index, text_end)
                            )

                # Insert new text at the cell insertion point
                insertion_index = cell.get("insertion_index", cell_start + 1)

                if new_text:
                    requests.append(
                        create_insert_text_request(insertion_index, new_text)
                    )

                if requests:
                    # Execute delete first (if any), then insert
                    for req in requests:
                        await asyncio.to_thread(
                            service.documents()
                            .batchUpdate(
                                documentId=document_id, body={"requests": [req]}
                            )
                            .execute
                        )

                results.append(
                    f"Op {i} (update_cell): SUCCESS - updated cell ({row_idx}, {col_idx})"
                )

            elif action == "delete_table":
                # Get the table's full range
                table_end = table_info["end_index"]

                # Delete the entire table using deleteContentRange
                request = create_delete_range_request(table_start, table_end)

                await asyncio.to_thread(
                    service.documents()
                    .batchUpdate(documentId=document_id, body={"requests": [request]})
                    .execute
                )

                results.append(
                    f"Op {i} (delete_table): SUCCESS - deleted table {table_index}"
                )
                # Note: Any subsequent operations on this table will fail since it no longer exists

            elif action == "merge_cells":
                row_idx = op["row"]
                col_idx = op["column"]
                row_span = op["row_span"]
                column_span = op["column_span"]

                # Validate bounds
                if row_idx >= num_rows:
                    results.append(
                        f"Op {i} (merge_cells): FAILED - row {row_idx} out of bounds (table has {num_rows} rows)"
                    )
                    continue

                if col_idx >= num_cols:
                    results.append(
                        f"Op {i} (merge_cells): FAILED - column {col_idx} out of bounds (table has {num_cols} columns)"
                    )
                    continue

                # Validate that merge range doesn't exceed table bounds
                if row_idx + row_span > num_rows:
                    results.append(
                        f"Op {i} (merge_cells): FAILED - row_span {row_span} starting at row {row_idx} exceeds table bounds (table has {num_rows} rows)"
                    )
                    continue

                if col_idx + column_span > num_cols:
                    results.append(
                        f"Op {i} (merge_cells): FAILED - column_span {column_span} starting at column {col_idx} exceeds table bounds (table has {num_cols} columns)"
                    )
                    continue

                request = create_merge_table_cells_request(
                    table_start_index=table_start,
                    row_index=row_idx,
                    column_index=col_idx,
                    row_span=row_span,
                    column_span=column_span,
                )

                await asyncio.to_thread(
                    service.documents()
                    .batchUpdate(documentId=document_id, body={"requests": [request]})
                    .execute
                )

                results.append(
                    f"Op {i} (merge_cells): SUCCESS - merged cells from ({row_idx}, {col_idx}) spanning {row_span} rows and {column_span} columns"
                )

            elif action == "unmerge_cells":
                row_idx = op["row"]
                col_idx = op["column"]
                row_span = op["row_span"]
                column_span = op["column_span"]

                # Validate bounds
                if row_idx >= num_rows:
                    results.append(
                        f"Op {i} (unmerge_cells): FAILED - row {row_idx} out of bounds (table has {num_rows} rows)"
                    )
                    continue

                if col_idx >= num_cols:
                    results.append(
                        f"Op {i} (unmerge_cells): FAILED - column {col_idx} out of bounds (table has {num_cols} columns)"
                    )
                    continue

                # Validate that unmerge range doesn't exceed table bounds
                if row_idx + row_span > num_rows:
                    results.append(
                        f"Op {i} (unmerge_cells): FAILED - row_span {row_span} starting at row {row_idx} exceeds table bounds (table has {num_rows} rows)"
                    )
                    continue

                if col_idx + column_span > num_cols:
                    results.append(
                        f"Op {i} (unmerge_cells): FAILED - column_span {column_span} starting at column {col_idx} exceeds table bounds (table has {num_cols} columns)"
                    )
                    continue

                request = create_unmerge_table_cells_request(
                    table_start_index=table_start,
                    row_index=row_idx,
                    column_index=col_idx,
                    row_span=row_span,
                    column_span=column_span,
                )

                await asyncio.to_thread(
                    service.documents()
                    .batchUpdate(documentId=document_id, body={"requests": [request]})
                    .execute
                )

                results.append(
                    f"Op {i} (unmerge_cells): SUCCESS - unmerged cells from ({row_idx}, {col_idx}) spanning {row_span} rows and {column_span} columns"
                )

            elif action == "format_cell":
                row_idx = op["row"]
                col_idx = op["column"]
                row_span = op.get("row_span", 1)
                column_span = op.get("column_span", 1)

                # Validate bounds
                if row_idx >= num_rows:
                    results.append(
                        f"Op {i} (format_cell): FAILED - row {row_idx} out of bounds (table has {num_rows} rows)"
                    )
                    continue

                if col_idx >= num_cols:
                    results.append(
                        f"Op {i} (format_cell): FAILED - column {col_idx} out of bounds (table has {num_cols} columns)"
                    )
                    continue

                # Validate that format range doesn't exceed table bounds
                if row_idx + row_span > num_rows:
                    results.append(
                        f"Op {i} (format_cell): FAILED - row_span {row_span} starting at row {row_idx} exceeds table bounds (table has {num_rows} rows)"
                    )
                    continue

                if col_idx + column_span > num_cols:
                    results.append(
                        f"Op {i} (format_cell): FAILED - column_span {column_span} starting at column {col_idx} exceeds table bounds (table has {num_cols} columns)"
                    )
                    continue

                # Build the request with all provided formatting options
                request = create_update_table_cell_style_request(
                    table_start_index=table_start,
                    row_index=row_idx,
                    column_index=col_idx,
                    row_span=row_span,
                    column_span=column_span,
                    background_color=op.get("background_color"),
                    border_color=op.get("border_color"),
                    border_width=op.get("border_width"),
                    border_dash_style=op.get("border_dash_style"),
                    border_top=op.get("border_top"),
                    border_bottom=op.get("border_bottom"),
                    border_left=op.get("border_left"),
                    border_right=op.get("border_right"),
                    padding_top=op.get("padding_top"),
                    padding_bottom=op.get("padding_bottom"),
                    padding_left=op.get("padding_left"),
                    padding_right=op.get("padding_right"),
                    content_alignment=op.get("content_alignment"),
                )

                await asyncio.to_thread(
                    service.documents()
                    .batchUpdate(documentId=document_id, body={"requests": [request]})
                    .execute
                )

                # Build a description of what was formatted
                format_desc_parts = []
                if op.get("background_color"):
                    format_desc_parts.append(f"background={op['background_color']}")
                if op.get("border_color") or op.get("border_width"):
                    format_desc_parts.append("borders")
                if any(op.get(f"border_{side}") for side in ["top", "bottom", "left", "right"]):
                    format_desc_parts.append("custom borders")
                if any(op.get(f"padding_{side}") for side in ["top", "bottom", "left", "right"]):
                    format_desc_parts.append("padding")
                if op.get("content_alignment"):
                    format_desc_parts.append(f"alignment={op['content_alignment']}")

                format_desc = ", ".join(format_desc_parts) if format_desc_parts else "style"
                cell_range = f"({row_idx}, {col_idx})"
                if row_span > 1 or column_span > 1:
                    cell_range = f"({row_idx}, {col_idx}) to ({row_idx + row_span - 1}, {col_idx + column_span - 1})"

                results.append(
                    f"Op {i} (format_cell): SUCCESS - formatted cell(s) {cell_range} with {format_desc}"
                )

            elif action == "resize_column":
                col_val = op["column"]
                width = op["width"]
                width_type = op.get("width_type", "FIXED_WIDTH")

                # Convert single column to list
                if isinstance(col_val, int):
                    column_indices = [col_val]
                else:
                    column_indices = col_val

                # Validate all columns are within bounds
                out_of_bounds = [c for c in column_indices if c >= num_cols]
                if out_of_bounds:
                    results.append(
                        f"Op {i} (resize_column): FAILED - column(s) {out_of_bounds} out of bounds (table has {num_cols} columns)"
                    )
                    continue

                request = create_update_table_column_properties_request(
                    table_start_index=table_start,
                    column_indices=column_indices,
                    width=width,
                    width_type=width_type,
                )

                await asyncio.to_thread(
                    service.documents()
                    .batchUpdate(documentId=document_id, body={"requests": [request]})
                    .execute
                )

                col_desc = f"column {column_indices[0]}" if len(column_indices) == 1 else f"columns {column_indices}"
                results.append(
                    f"Op {i} (resize_column): SUCCESS - resized {col_desc} to {width}pt ({width_type})"
                )

        except Exception as e:
            error_msg = str(e)
            logger.error(f"[modify_table] Operation {i} ({action}) failed: {error_msg}")
            results.append(f"Op {i} ({action}): FAILED - {error_msg}")

    # Build summary
    link = f"https://docs.google.com/document/d/{document_id}/edit"
    success_count = sum(1 for r in results if "SUCCESS" in r)

    summary = f"Table modification complete. {success_count}/{len(operations)} operations succeeded.\n\n"
    summary += "Results:\n" + "\n".join(f"  - {r}" for r in results)
    summary += f"\n\nLink: {link}"

    return summary


@server.tool()
@handle_http_errors("export_doc_to_pdf", service_type="drive")
@require_google_service("drive", "drive_file")
async def export_doc_to_pdf(
    service: Any,
    user_google_email: str,
    document_id: str,
    pdf_filename: str = None,
    folder_id: str = None,
) -> str:
    """
    Exports a Google Doc to PDF format and saves it to Google Drive.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the Google Doc to export
        pdf_filename: Name for the PDF file (optional - if not provided, uses original name + "_PDF")
        folder_id: Drive folder ID to save PDF in (optional - if not provided, saves in root)

    Returns:
        str: Confirmation message with PDF file details and links
    """
    logger.info(
        f"[export_doc_to_pdf] Email={user_google_email}, Doc={document_id}, pdf_filename={pdf_filename}, folder_id={folder_id}"
    )

    # Input validation
    validator = ValidationManager()

    # Get file metadata first to validate it's a Google Doc
    try:
        file_metadata = await asyncio.to_thread(
            service.files()
            .get(fileId=document_id, fields="id, name, mimeType, webViewLink")
            .execute
        )
    except Exception as e:
        return validator.create_pdf_export_error(
            document_id=document_id, stage="access", error_detail=str(e)
        )

    mime_type = file_metadata.get("mimeType", "")
    original_name = file_metadata.get("name", "Unknown Document")
    web_view_link = file_metadata.get("webViewLink", "#")

    # Verify it's a Google Doc
    if mime_type != "application/vnd.google-apps.document":
        return validator.create_invalid_document_type_error(
            document_id=document_id, file_name=original_name, actual_mime_type=mime_type
        )

    logger.info(f"[export_doc_to_pdf] Exporting '{original_name}' to PDF")

    # Export the document as PDF
    try:
        request_obj = service.files().export_media(
            fileId=document_id, mimeType="application/pdf"
        )

        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request_obj)

        done = False
        while not done:
            _, done = await asyncio.to_thread(downloader.next_chunk)

        pdf_content = fh.getvalue()
        pdf_size = len(pdf_content)

    except Exception as e:
        return validator.create_pdf_export_error(
            document_id=document_id, stage="export", error_detail=str(e)
        )

    # Determine PDF filename
    if not pdf_filename:
        pdf_filename = f"{original_name}_PDF.pdf"
    elif not pdf_filename.endswith(".pdf"):
        pdf_filename += ".pdf"

    # Upload PDF to Drive
    try:
        # Reuse the existing BytesIO object by resetting to the beginning
        fh.seek(0)
        # Create media upload object
        media = MediaIoBaseUpload(fh, mimetype="application/pdf", resumable=True)

        # Prepare file metadata for upload
        file_metadata = {"name": pdf_filename, "mimeType": "application/pdf"}

        # Add parent folder if specified
        if folder_id:
            file_metadata["parents"] = [folder_id]

        # Upload the file
        uploaded_file = await asyncio.to_thread(
            service.files()
            .create(
                body=file_metadata,
                media_body=media,
                fields="id, name, webViewLink, parents",
                supportsAllDrives=True,
            )
            .execute
        )

        pdf_file_id = uploaded_file.get("id")
        pdf_web_link = uploaded_file.get("webViewLink", "#")
        pdf_parents = uploaded_file.get("parents", [])

        logger.info(
            f"[export_doc_to_pdf] Successfully uploaded PDF to Drive: {pdf_file_id}"
        )

        folder_info = ""
        if folder_id:
            folder_info = f" in folder {folder_id}"
        elif pdf_parents:
            folder_info = f" in folder {pdf_parents[0]}"

        return f"Successfully exported '{original_name}' to PDF and saved to Drive as '{pdf_filename}' (ID: {pdf_file_id}, {pdf_size:,} bytes){folder_info}. PDF: {pdf_web_link} | Original: {web_view_link}"

    except Exception as e:
        return validator.create_pdf_export_error(
            document_id=document_id,
            stage="upload",
            error_detail=f"{str(e)}. PDF was generated successfully ({pdf_size:,} bytes) but could not be saved to Drive.",
        )


@server.tool()
@handle_http_errors("preview_search_results", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def preview_search_results(
    service: Any,
    user_google_email: str,
    document_id: str,
    search_text: str = None,
    match_case: bool = True,
    context_chars: int = 50,
    # Parameter alias for consistency with other tools
    search: str = None,  # Alias for search_text (matches modify_doc_text, format_all_occurrences)
) -> str:
    """
    Preview what text will be matched by a search operation before modifying.

    USE THIS TOOL BEFORE:
    - Using modify_doc_text with search parameter
    - Using batch_edit_doc with search-based operations
    - Any operation where you need to verify search targets

    This tool shows ALL occurrences of search text in the document with
    surrounding context, helping you:
    - Verify the correct text will be matched
    - Identify which occurrence to target
    - See if search text spans fragmented text runs
    - Understand character ranges before making changes

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to search
        search_text: Text to search for in the document
        match_case: Whether to match case exactly (default: True)
        context_chars: Characters of context to show before/after each match (default: 50)
        search: Alias for search_text (for consistency with modify_doc_text, format_all_occurrences)

    Returns:
        str: JSON containing:
            - total_matches: Number of occurrences found
            - matches: List of match details including:
                - occurrence: Which occurrence (1, 2, 3, etc.)
                - start_index: Character position where match starts
                - end_index: Character position where match ends
                - context_before: Text immediately before the match
                - matched_text: The actual matched text
                - context_after: Text immediately after the match
            - search_text: The text that was searched for
            - match_case: Whether case-sensitive search was used

    Example Response:
        {
            "total_matches": 3,
            "search_text": "TODO",
            "match_case": true,
            "matches": [
                {
                    "occurrence": 1,
                    "start_index": 150,
                    "end_index": 154,
                    "context_before": "...needs review. ",
                    "matched_text": "TODO",
                    "context_after": ": Fix this bug..."
                },
                {
                    "occurrence": 2,
                    "start_index": 450,
                    "end_index": 454,
                    "context_before": "...later. ",
                    "matched_text": "TODO",
                    "context_after": ": Add tests..."
                }
            ]
        }

    Usage Tips:
        - Use occurrence number from results when calling modify_doc_text
        - Context helps verify you're targeting the right match
        - If no matches found, try match_case=False for case-insensitive search
    """
    import json

    # Resolve parameter alias for API consistency
    # 'search' is an alias for 'search_text' (matches modify_doc_text, format_all_occurrences)
    if search is not None and search_text is None:
        search_text = search

    logger.debug(
        f"[preview_search_results] Doc={document_id}, search='{search_text}', match_case={match_case}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    if not search_text or not search_text.strip():
        error = DocsErrorBuilder.missing_required_param(
            param_name="search_text",
            context_description="for search preview",
            valid_values=["non-empty search string"],
        )
        return format_error(error)

    # Get the document
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Find all occurrences
    all_matches = find_all_occurrences_in_document(doc_data, search_text, match_case)

    if not all_matches:
        # No matches - provide helpful response
        result = {
            "total_matches": 0,
            "search_text": search_text,
            "match_case": match_case,
            "matches": [],
            "message": f"No matches found for '{search_text}'"
            + (" (case-sensitive)" if match_case else " (case-insensitive)"),
            "hint": "Try match_case=False for case-insensitive search"
            if match_case
            else "Verify the search text exists in the document",
        }
        return json.dumps(result, indent=2)

    # Extract full document text for context retrieval
    from gdocs.docs_helpers import extract_document_text_with_indices

    text_segments = extract_document_text_with_indices(doc_data)

    # Build full text and index mapping
    full_text = ""
    index_map = []  # Maps position in full_text to document index
    reverse_map = {}  # Maps document index to position in full_text

    for segment_text, start_idx, _ in text_segments:
        for i, char in enumerate(segment_text):
            doc_idx = start_idx + i
            reverse_map[doc_idx] = len(full_text)
            index_map.append(doc_idx)
            full_text += char

    # Build match details with context
    matches_with_context = []

    for i, (start_idx, end_idx) in enumerate(all_matches):
        # Find positions in full_text
        text_start = reverse_map.get(start_idx)
        text_end = reverse_map.get(end_idx - 1, reverse_map.get(end_idx))

        if text_start is not None and text_end is not None:
            text_end += 1  # Adjust to exclusive end

            # Extract context
            context_start = max(0, text_start - context_chars)
            context_end = min(len(full_text), text_end + context_chars)

            context_before = full_text[context_start:text_start]
            matched_text = full_text[text_start:text_end]
            context_after = full_text[text_end:context_end]

            # Add ellipsis indicators if truncated
            if context_start > 0:
                context_before = "..." + context_before
            if context_end < len(full_text):
                context_after = context_after + "..."

            match_info = {
                "occurrence": i + 1,
                "start_index": start_idx,
                "end_index": end_idx,
                "context_before": context_before,
                "matched_text": matched_text,
                "context_after": context_after,
            }
        else:
            # Fallback if mapping failed
            match_info = {
                "occurrence": i + 1,
                "start_index": start_idx,
                "end_index": end_idx,
                "context_before": "[context unavailable]",
                "matched_text": search_text,
                "context_after": "[context unavailable]",
            }

        matches_with_context.append(match_info)

    result = {
        "total_matches": len(all_matches),
        "search_text": search_text,
        "match_case": match_case,
        "matches": matches_with_context,
    }

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Search preview for document {document_id}:\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


@server.tool()
@handle_http_errors("find_doc_elements", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def find_doc_elements(
    service: Any,
    user_google_email: str,
    document_id: str,
    element_type: str,
) -> str:
    """
    Find all elements of a specific type in a Google Doc.

    Use this tool to locate all tables, headings, lists, or other structural
    elements in a document. Returns positions and metadata for each element.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to search
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
        str: JSON containing:
            - count: Number of elements found
            - element_type: The type that was searched
            - elements: List of element details including positions and metadata

    Example Response for element_type='table':
        {
            "count": 2,
            "element_type": "table",
            "elements": [
                {"type": "table", "start_index": 150, "end_index": 350, "rows": 3, "columns": 4},
                {"type": "table", "start_index": 500, "end_index": 650, "rows": 2, "columns": 3}
            ]
        }

    Example Response for element_type='heading2':
        {
            "count": 3,
            "element_type": "heading2",
            "elements": [
                {"type": "heading2", "text": "Introduction", "start_index": 10, "end_index": 22, "level": 2},
                {"type": "heading2", "text": "Methods", "start_index": 150, "end_index": 158, "level": 2},
                {"type": "heading2", "text": "Results", "start_index": 450, "end_index": 458, "level": 2}
            ]
        }

    Use Cases:
        - Find all tables before inserting data
        - Navigate to specific heading levels
        - Count structural elements in a document
        - Locate all lists for formatting changes
    """
    import json

    logger.debug(f"[find_doc_elements] Doc={document_id}, element_type={element_type}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    if not element_type or not element_type.strip():
        error = DocsErrorBuilder.missing_required_param(
            param_name="element_type",
            context_description="for element search",
            valid_values=[
                "table",
                "heading",
                "heading1-6",
                "paragraph",
                "bullet_list",
                "numbered_list",
                "list",
            ],
        )
        return format_error(error)

    # Get the document
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Find elements
    elements = find_elements_by_type(doc_data, element_type)

    # Clean up elements for output
    clean_elements = []
    for elem in elements:
        clean_elem = {
            "type": elem["type"],
            "start_index": elem["start_index"],
            "end_index": elem["end_index"],
        }
        if "text" in elem:
            clean_elem["text"] = elem["text"]
        if "level" in elem:
            clean_elem["level"] = elem["level"]
        if "rows" in elem:
            clean_elem["rows"] = elem["rows"]
            clean_elem["columns"] = elem.get("columns", 0)
        if "items" in elem:
            clean_elem["item_count"] = len(elem["items"])
            clean_elem["items"] = [
                {
                    "text": item["text"],
                    "start_index": item["start_index"],
                    "end_index": item["end_index"],
                }
                for item in elem["items"]
            ]
        clean_elements.append(clean_elem)

    result = {
        "count": len(clean_elements),
        "element_type": element_type,
        "elements": clean_elements,
    }

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Found {len(clean_elements)} element(s) of type '{element_type}' in document {document_id}:\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


@server.tool()
@handle_http_errors("get_element_context", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def get_element_context(
    service: Any,
    user_google_email: str,
    document_id: str,
    index: int,
) -> str:
    """
    Find the parent section/heading hierarchy for an element at a given position.

    Use this tool to determine which section(s) contain a specific position in
    the document. Returns the full hierarchy from document root to the most
    specific containing heading.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to analyze
        index: Character position in the document to analyze

    Returns:
        str: JSON containing:
            - index: The position that was queried
            - ancestors: List of containing headings from root to leaf, each with:
                - type: Heading type (e.g., 'heading1', 'heading2')
                - text: Heading text
                - level: Heading level (0=title, 1-6 for headings)
                - start_index: Start position of the heading
                - end_index: End position of the heading
                - section_end: Where this section ends
            - depth: Number of heading levels containing this position
            - innermost_section: The most specific heading containing the position

    Example Response:
        {
            "index": 500,
            "ancestors": [
                {"type": "heading1", "text": "Introduction", "level": 1,
                 "start_index": 10, "end_index": 22, "section_end": 800},
                {"type": "heading2", "text": "Background", "level": 2,
                 "start_index": 150, "end_index": 162, "section_end": 600}
            ],
            "depth": 2,
            "innermost_section": {"type": "heading2", "text": "Background", ...}
        }

    Use Cases:
        - Determine which section contains a search result
        - Navigate document structure from a specific position
        - Find context for error messages about specific positions
        - Build breadcrumb navigation for document positions
    """
    import json

    logger.debug(f"[get_element_context] Doc={document_id}, index={index}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    if index < 0:
        error = DocsErrorBuilder.invalid_param(
            param_name="index",
            received_value=str(index),
            valid_values=["non-negative integer"],
            context_description="document position",
        )
        return format_error(error)

    # Get the document
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Get ancestors
    ancestors = get_element_ancestors(doc_data, index)

    result = {
        "index": index,
        "ancestors": ancestors,
        "depth": len(ancestors),
        "innermost_section": ancestors[-1] if ancestors else None,
    }

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    if ancestors:
        innermost = ancestors[-1]
        summary = f"Position {index} is within section '{innermost['text']}' (level {innermost['level']}) with {len(ancestors)} containing section(s)"
    else:
        summary = f"Position {index} is not within any heading section"

    return f"{summary}:\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


@server.tool()
@handle_http_errors("get_text_formatting", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def get_text_formatting(
    service: Any,
    user_google_email: str,
    document_id: str,
    start_index: int = None,
    end_index: int = None,
    search: str = None,
    occurrence: int = 1,
    match_case: bool = True,
) -> str:
    """
    Get the text formatting/style attributes at a specific location or range.

    Use this tool to inspect the current formatting of text, which is useful for:
    - Copying formatting from one location to another ("format painter" workflow)
    - Checking if text has specific styles applied
    - Debugging formatting issues
    - Building reports of document formatting patterns

    Supports two positioning modes:
    1. Index-based: Use start_index/end_index for exact positions
    2. Search-based: Use search to find text and get its formatting

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to analyze

        Index-based positioning:
        start_index: Starting character position (inclusive)
        end_index: Ending character position (exclusive). If not provided,
                   returns formatting for just the character at start_index.

        Search-based positioning:
        search: Text to search for in the document. Returns formatting for
                the found text.
        occurrence: Which occurrence to target (1=first, 2=second, -1=last). Default: 1
        match_case: Whether to match case exactly. Default: True

    Returns:
        str: JSON containing:
            - start_index: Starting position queried
            - end_index: Ending position queried
            - text: The text content at the specified range
            - formatting: Array of formatting spans, each containing:
                - start_index: Start of this formatting span
                - end_index: End of this formatting span
                - text: Text content of this span
                - bold: Boolean if text is bold
                - italic: Boolean if text is italic
                - underline: Boolean if text is underlined
                - strikethrough: Boolean if text has strikethrough
                - small_caps: Boolean if text is in small caps
                - baseline_offset: "SUBSCRIPT", "SUPERSCRIPT", or "NONE"
                - font_size: Font size in points (if set)
                - font_family: Font family name (if set)
                - foreground_color: Text color as hex string (if set)
                - background_color: Background color as hex string (if set)
                - link_url: URL if text is a hyperlink (if set)
            - has_mixed_formatting: Boolean indicating if the range has different
                                    formatting across different spans
            - search_info: (only for search-based) Info about the search match

    Example Response (index-based):
        {
            "start_index": 100,
            "end_index": 120,
            "text": "formatted text here",
            "formatting": [...],
            "has_mixed_formatting": true
        }

    Example Response (search-based):
        {
            "start_index": 100,
            "end_index": 120,
            "text": "formatted text here",
            "formatting": [...],
            "has_mixed_formatting": true,
            "search_info": {
                "search_text": "formatted text here",
                "occurrence": 1,
                "match_case": true
            }
        }

    Use Cases:
        - Copy formatting: Get formatting from source, apply to destination
        - Check if text is bold/italic/etc before toggling
        - Find all text with specific formatting
        - Debug why text looks different than expected
    """
    import json

    logger.debug(
        f"[get_text_formatting] Doc={document_id}, start={start_index}, end={end_index}, search={search}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Determine positioning mode
    use_search_mode = search is not None
    use_index_mode = start_index is not None

    # Validate that exactly one positioning mode is used
    if use_search_mode and use_index_mode:
        error = DocsErrorBuilder.invalid_param(
            param_name="positioning",
            received_value="both search and start_index",
            valid_values=["search OR start_index (not both)"],
            context_description="Use either search-based or index-based positioning, not both",
        )
        return format_error(error)

    if not use_search_mode and not use_index_mode:
        error = DocsErrorBuilder.invalid_param(
            param_name="positioning",
            received_value="neither search nor start_index",
            valid_values=["search", "start_index"],
            context_description="Must provide either search text or start_index",
        )
        return format_error(error)

    # Get the document first (needed for both modes)
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    search_info = None

    if use_search_mode:
        # Search-based positioning
        if search == "":
            return validator.create_empty_search_error()

        result = find_text_in_document(doc_data, search, occurrence, match_case)
        if result is None:
            all_occurrences = find_all_occurrences_in_document(doc_data, search, match_case)
            if all_occurrences:
                if occurrence != 1:
                    return validator.create_invalid_occurrence_error(
                        occurrence=occurrence,
                        total_found=len(all_occurrences),
                        search_text=search,
                    )
                occurrences_data = [
                    {"index": i + 1, "position": f"{s}-{e}"}
                    for i, (s, e) in enumerate(all_occurrences[:5])
                ]
                return validator.create_ambiguous_search_error(
                    search_text=search,
                    occurrences=occurrences_data,
                    total_count=len(all_occurrences),
                )
            return validator.create_search_not_found_error(
                search_text=search, match_case=match_case
            )

        start_index, end_index = result
        search_info = {
            "search_text": search,
            "occurrence": occurrence,
            "match_case": match_case,
        }
    else:
        # Index-based positioning
        if start_index < 0:
            error = DocsErrorBuilder.invalid_param(
                param_name="start_index",
                received_value=str(start_index),
                valid_values=["non-negative integer"],
                context_description="document position",
            )
            return format_error(error)

        # Default end_index to start_index + 1 if not provided
        if end_index is None:
            end_index = start_index + 1

        if end_index <= start_index:
            error = DocsErrorBuilder.invalid_param(
                param_name="end_index",
                received_value=str(end_index),
                valid_values=[f"integer greater than start_index ({start_index})"],
                context_description="document position range",
            )
            return format_error(error)

    # Extract formatting from the specified range
    formatting_spans = _extract_text_formatting_from_range(
        doc_data, start_index, end_index
    )

    # Collect all text content
    all_text = "".join(span.get("text", "") for span in formatting_spans)

    # Determine if formatting is mixed (more than one unique style)
    has_mixed = _has_mixed_formatting(formatting_spans)

    result = {
        "start_index": start_index,
        "end_index": end_index,
        "text": all_text,
        "formatting": formatting_spans,
        "has_mixed_formatting": has_mixed,
    }

    # Add search_info if using search-based positioning
    if search_info:
        result["search_info"] = search_info

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    if formatting_spans:
        summary = f"Found {len(formatting_spans)} formatting span(s) in range {start_index}-{end_index}"
        if search_info:
            summary += f" for search '{search_info['search_text']}'"
    else:
        summary = f"No text content found in range {start_index}-{end_index}"

    return f"{summary}:\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


def _extract_text_formatting_from_range(
    doc_data: dict, start_index: int, end_index: int
) -> list:
    """
    Extract text formatting information from a document range.

    Args:
        doc_data: Raw document data from Google Docs API
        start_index: Start position (inclusive)
        end_index: End position (exclusive)

    Returns:
        List of formatting span dictionaries
    """
    formatting_spans = []
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
                    text_run = para_elem["textRun"]
                    content_text = text_run.get("content", "")
                    text_style = text_run.get("textStyle", {})

                    # Calculate the actual range within our query range
                    actual_start = max(pe_start, start_index)
                    actual_end = min(pe_end, end_index)

                    # Extract the substring of text that falls within our range
                    text_offset_start = actual_start - pe_start
                    text_offset_end = actual_end - pe_start
                    span_text = content_text[text_offset_start:text_offset_end]

                    # Build the formatting info for this span
                    span_info = _build_formatting_info(
                        text_style, actual_start, actual_end, span_text
                    )
                    formatting_spans.append(span_info)

        elif "table" in element:
            # Handle table content
            table = element["table"]
            for row in table.get("tableRows", []):
                for cell in row.get("tableCells", []):
                    cell_start = cell.get("startIndex", 0)
                    cell_end = cell.get("endIndex", 0)

                    # Skip cells completely outside our range
                    if cell_end <= start_index or cell_start >= end_index:
                        continue

                    # Process cell content
                    for cell_element in cell.get("content", []):
                        if "paragraph" in cell_element:
                            paragraph = cell_element["paragraph"]
                            for para_elem in paragraph.get("elements", []):
                                pe_start = para_elem.get("startIndex", 0)
                                pe_end = para_elem.get("endIndex", 0)

                                if pe_end <= start_index or pe_start >= end_index:
                                    continue

                                if "textRun" in para_elem:
                                    text_run = para_elem["textRun"]
                                    content_text = text_run.get("content", "")
                                    text_style = text_run.get("textStyle", {})

                                    actual_start = max(pe_start, start_index)
                                    actual_end = min(pe_end, end_index)

                                    text_offset_start = actual_start - pe_start
                                    text_offset_end = actual_end - pe_start
                                    span_text = content_text[
                                        text_offset_start:text_offset_end
                                    ]

                                    span_info = _build_formatting_info(
                                        text_style, actual_start, actual_end, span_text
                                    )
                                    formatting_spans.append(span_info)

    return formatting_spans


def _build_formatting_info(
    text_style: dict, start_index: int, end_index: int, text: str
) -> dict:
    """
    Build a formatting info dictionary from a text style.

    Args:
        text_style: The textStyle object from Google Docs API
        start_index: Start position of this span
        end_index: End position of this span
        text: Text content of this span

    Returns:
        Dictionary with normalized formatting information
    """
    # Normalize baseline_offset: BASELINE_OFFSET_UNSPECIFIED means inherited/normal
    raw_baseline_offset = text_style.get("baselineOffset", "NONE")
    if raw_baseline_offset == "BASELINE_OFFSET_UNSPECIFIED":
        raw_baseline_offset = "NONE"

    info = {
        "start_index": start_index,
        "end_index": end_index,
        "text": text,
        "bold": text_style.get("bold", False),
        "italic": text_style.get("italic", False),
        "underline": text_style.get("underline", False),
        "strikethrough": text_style.get("strikethrough", False),
        "small_caps": text_style.get("smallCaps", False),
        "baseline_offset": raw_baseline_offset,
    }

    # Font size
    font_size = text_style.get("fontSize")
    if font_size:
        info["font_size"] = font_size.get("magnitude")

    # Font family
    weighted_font = text_style.get("weightedFontFamily")
    if weighted_font:
        info["font_family"] = weighted_font.get("fontFamily")

    # Foreground color
    fg_color = text_style.get("foregroundColor")
    if fg_color:
        color_hex = _color_to_hex(fg_color)
        if color_hex:
            info["foreground_color"] = color_hex

    # Background color
    bg_color = text_style.get("backgroundColor")
    if bg_color:
        color_hex = _color_to_hex(bg_color)
        if color_hex:
            info["background_color"] = color_hex

    # Link
    link = text_style.get("link")
    if link:
        info["link_url"] = link.get("url", "")

    return info


def _color_to_hex(color_obj: dict) -> str:
    """
    Convert a Google Docs color object to hex string.

    Args:
        color_obj: Color object from Google Docs API with 'color' key containing
                   'rgbColor' with red, green, blue values (0-1 floats)

    Returns:
        Hex color string (e.g., "#FF0000") or empty string if not valid
    """
    if not color_obj:
        return ""

    color = color_obj.get("color", {})
    rgb_color = color.get("rgbColor", {})

    if not rgb_color:
        return ""

    red = int(rgb_color.get("red", 0) * 255)
    green = int(rgb_color.get("green", 0) * 255)
    blue = int(rgb_color.get("blue", 0) * 255)

    return f"#{red:02X}{green:02X}{blue:02X}"


def _has_mixed_formatting(formatting_spans: list) -> bool:
    """
    Determine if a list of formatting spans has mixed formatting.

    Args:
        formatting_spans: List of formatting span dictionaries

    Returns:
        True if there are different formatting styles across spans
    """
    if len(formatting_spans) <= 1:
        return False

    # Compare all spans to the first one
    first = formatting_spans[0]
    style_keys = [
        "bold",
        "italic",
        "underline",
        "strikethrough",
        "small_caps",
        "baseline_offset",
        "font_size",
        "font_family",
        "foreground_color",
        "background_color",
        "link_url",
    ]

    for span in formatting_spans[1:]:
        for key in style_keys:
            if span.get(key) != first.get(key):
                return True

    return False


@server.tool()
@handle_http_errors("navigate_heading_siblings", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def navigate_heading_siblings(
    service: Any,
    user_google_email: str,
    document_id: str,
    heading: str = None,
    match_case: bool = False,
    # Parameter alias for clarity
    current_heading: str = None,  # Alias for heading (clarifies this is the heading to find siblings FOR)
) -> str:
    """
    Find the previous and next headings at the same level as a specified heading.

    Use this tool to navigate between sibling sections in a document,
    helping move to previous or next sections at the same hierarchy level.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to analyze
        heading: Text of the heading to find siblings for
        match_case: Whether to match case exactly when finding the heading
        current_heading: Alias for heading (clarifies this is the heading to find siblings FOR)

    Returns:
        str: JSON containing:
            - found: Boolean indicating if the heading was found
            - heading: The matched heading info
            - level: The heading level
            - previous: Previous sibling heading info, or null if first at this level
            - next: Next sibling heading info, or null if last at this level
            - siblings_count: Total count of headings at this level
            - position_in_siblings: 1-based position among siblings (e.g., 2 of 5)

    Example Response:
        {
            "found": true,
            "heading": {"type": "heading2", "text": "Methods", "level": 2,
                       "start_index": 150, "end_index": 158},
            "level": 2,
            "previous": {"type": "heading2", "text": "Introduction", "level": 2,
                        "start_index": 10, "end_index": 22},
            "next": {"type": "heading2", "text": "Results", "level": 2,
                    "start_index": 450, "end_index": 458},
            "siblings_count": 4,
            "position_in_siblings": 2
        }

    Use Cases:
        - Navigate sequentially through document sections
        - Find adjacent sections for comparison or movement
        - Understand document structure at a specific level
        - Build next/previous navigation for section editing
    """
    import json

    # Resolve parameter alias: 'current_heading' is an alias for 'heading'
    if current_heading is not None and heading is None:
        heading = current_heading

    logger.debug(
        f"[navigate_heading_siblings] Doc={document_id}, heading={heading}, match_case={match_case}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    if not heading or not heading.strip():
        error = DocsErrorBuilder.missing_required_param(
            param_name="heading",
            context_description="for sibling navigation",
            valid_values=["non-empty heading text"],
        )
        return format_error(error)

    # Get the document
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Get siblings
    result = get_heading_siblings(doc_data, heading, match_case)

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    if not result["found"]:
        # Provide helpful error with available headings
        all_headings = get_all_headings(doc_data)
        heading_list = [h["text"] for h in all_headings] if all_headings else []
        return validator.create_heading_not_found_error(
            heading=heading, available_headings=heading_list, match_case=match_case
        )

    summary_parts = [f"Heading '{result['heading']['text']}' (level {result['level']})"]
    summary_parts.append(
        f"Position {result['position_in_siblings']} of {result['siblings_count']} headings at this level"
    )
    if result["previous"]:
        summary_parts.append(f"Previous: '{result['previous']['text']}'")
    else:
        summary_parts.append("No previous sibling (first at this level)")
    if result["next"]:
        summary_parts.append(f"Next: '{result['next']['text']}'")
    else:
        summary_parts.append("No next sibling (last at this level)")

    return f"{' | '.join(summary_parts)}:\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


# ============================================================================
# SMART CONTENT EXTRACTION TOOLS
# ============================================================================


@server.tool()
@handle_http_errors("extract_links", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def extract_links(
    service: Any,
    user_google_email: str,
    document_id: str,
    include_section_context: bool = True,
) -> str:
    """
    Extract all hyperlinks from a Google Doc with their text and URLs.

    This tool identifies all linked text in the document and returns structured
    data about each link including:
    - The linked text (anchor text)
    - The destination URL
    - Position in the document
    - Section context (which heading it appears under)

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to extract links from
        include_section_context: Whether to include which section each link is in (default: True)

    Returns:
        str: JSON containing:
            - total_links: Number of links found
            - links: List of link details including:
                - text: The anchor text of the link
                - url: The destination URL
                - start_index: Character position where link starts
                - end_index: Character position where link ends
                - section: (if include_section_context) The heading this link is under
            - document_link: Link to open the document

    Example Response:
        {
            "total_links": 5,
            "links": [
                {
                    "text": "Google",
                    "url": "https://google.com",
                    "start_index": 150,
                    "end_index": 156,
                    "section": "References"
                }
            ],
            "document_link": "https://docs.google.com/document/d/..."
        }
    """
    import json

    logger.debug(f"[extract_links] Doc={document_id}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Get the document
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Get headings for section context
    headings = []
    if include_section_context:
        headings = get_all_headings(doc_data)

    def find_section_for_index(idx: int) -> str:
        """Find which section heading a given index falls under."""
        if not headings:
            return ""
        current_section = ""
        for heading in headings:
            if heading["start_index"] <= idx:
                current_section = heading["text"]
            else:
                break
        return current_section

    # Traverse document to find all links
    links = []
    body = doc_data.get("body", {})
    content = body.get("content", [])

    def process_elements(elements: list, depth: int = 0) -> None:
        """Recursively process elements to extract links."""
        if depth > 10:  # Prevent infinite recursion
            return

        for element in elements:
            if "paragraph" in element:
                paragraph = element["paragraph"]
                for para_elem in paragraph.get("elements", []):
                    if "textRun" in para_elem:
                        text_run = para_elem["textRun"]
                        text_style = text_run.get("textStyle", {})

                        # Check for link in text style
                        if "link" in text_style:
                            link_info = text_style["link"]
                            url = link_info.get("url", "")

                            # Skip internal document bookmarks
                            if url and not url.startswith("#"):
                                link_text = text_run.get("content", "").strip()
                                start_idx = para_elem.get("startIndex", 0)
                                end_idx = para_elem.get("endIndex", 0)

                                link_entry = {
                                    "text": link_text,
                                    "url": url,
                                    "start_index": start_idx,
                                    "end_index": end_idx,
                                }

                                if include_section_context:
                                    link_entry["section"] = find_section_for_index(
                                        start_idx
                                    )

                                links.append(link_entry)

            elif "table" in element:
                # Process table cells
                table = element["table"]
                for row in table.get("tableRows", []):
                    for cell in row.get("tableCells", []):
                        process_elements(cell.get("content", []), depth + 1)

    process_elements(content)

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    result = {"total_links": len(links), "links": links, "document_link": link}

    return json.dumps(result, indent=2)


@server.tool()
@handle_http_errors("extract_images", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def extract_images(
    service: Any,
    user_google_email: str,
    document_id: str,
    include_section_context: bool = True,
) -> str:
    """
    Extract all images from a Google Doc with their URLs and positions.

    This tool identifies all inline images in the document and returns
    structured data about each image including:
    - Image URL (for viewing/downloading)
    - Dimensions (width, height)
    - Position in the document
    - Section context (which heading it appears under)

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to extract images from
        include_section_context: Whether to include which section each image is in (default: True)

    Returns:
        str: JSON containing:
            - total_images: Number of images found
            - images: List of image details including:
                - object_id: Internal ID of the image object
                - content_uri: URL to access the image content
                - source_uri: Original source URL (if available)
                - width_pt: Width in points
                - height_pt: Height in points
                - start_index: Character position of the image
                - section: (if include_section_context) The heading this image is under
            - document_link: Link to open the document

    Example Response:
        {
            "total_images": 3,
            "images": [
                {
                    "object_id": "kix.abc123",
                    "content_uri": "https://lh3.googleusercontent.com/...",
                    "width_pt": 400,
                    "height_pt": 300,
                    "start_index": 250,
                    "section": "Diagrams"
                }
            ],
            "document_link": "https://docs.google.com/document/d/..."
        }
    """
    import json

    logger.debug(f"[extract_images] Doc={document_id}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Get the document
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Get inline objects registry
    inline_objects = doc_data.get("inlineObjects", {})

    # Get headings for section context
    headings = []
    if include_section_context:
        headings = get_all_headings(doc_data)

    def find_section_for_index(idx: int) -> str:
        """Find which section heading a given index falls under."""
        if not headings:
            return ""
        current_section = ""
        for heading in headings:
            if heading["start_index"] <= idx:
                current_section = heading["text"]
            else:
                break
        return current_section

    # Map to track inline object references and their positions
    image_refs = {}  # object_id -> start_index

    body = doc_data.get("body", {})
    content = body.get("content", [])

    def process_elements(elements: list, depth: int = 0) -> None:
        """Recursively process elements to find inline object references."""
        if depth > 10:
            return

        for element in elements:
            if "paragraph" in element:
                paragraph = element["paragraph"]
                for para_elem in paragraph.get("elements", []):
                    if "inlineObjectElement" in para_elem:
                        obj_elem = para_elem["inlineObjectElement"]
                        obj_id = obj_elem.get("inlineObjectId")
                        if obj_id:
                            start_idx = para_elem.get("startIndex", 0)
                            image_refs[obj_id] = start_idx

            elif "table" in element:
                table = element["table"]
                for row in table.get("tableRows", []):
                    for cell in row.get("tableCells", []):
                        process_elements(cell.get("content", []), depth + 1)

    process_elements(content)

    # Build image list from inline objects registry
    images = []
    for obj_id, obj_data in inline_objects.items():
        props = obj_data.get("inlineObjectProperties", {})
        embedded_obj = props.get("embeddedObject", {})

        # Get image properties
        image_props = embedded_obj.get("imageProperties", {})
        content_uri = image_props.get("contentUri", "")
        source_uri = image_props.get("sourceUri", "")

        # Get size
        size = embedded_obj.get("size", {})
        width = size.get("width", {})
        height = size.get("height", {})

        width_pt = (
            width.get("magnitude", 0)
            if width.get("unit") == "PT"
            else width.get("magnitude", 0)
        )
        height_pt = (
            height.get("magnitude", 0)
            if height.get("unit") == "PT"
            else height.get("magnitude", 0)
        )

        start_idx = image_refs.get(obj_id, 0)

        image_entry = {
            "object_id": obj_id,
            "content_uri": content_uri,
            "width_pt": width_pt,
            "height_pt": height_pt,
            "start_index": start_idx,
        }

        if source_uri:
            image_entry["source_uri"] = source_uri

        if include_section_context:
            image_entry["section"] = find_section_for_index(start_idx)

        images.append(image_entry)

    # Sort by position in document
    images.sort(key=lambda x: x["start_index"])

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    result = {"total_images": len(images), "images": images, "document_link": link}

    return json.dumps(result, indent=2)


@server.tool()
@handle_http_errors("extract_code_blocks", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def extract_code_blocks(
    service: Any,
    user_google_email: str,
    document_id: str,
    include_section_context: bool = True,
) -> str:
    """
    Extract code-formatted text blocks from a Google Doc.

    This tool identifies text that appears to be code based on formatting:
    - Monospace fonts (Courier New, Consolas, Monaco, etc.)
    - Background-colored text blocks
    - Consecutive monospace-formatted lines

    Note: Google Docs doesn't have native code blocks like Markdown, so
    detection is heuristic-based on common code formatting conventions.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to extract code from
        include_section_context: Whether to include which section each code block is in (default: True)

    Returns:
        str: JSON containing:
            - total_code_blocks: Number of code blocks found
            - code_blocks: List of code block details including:
                - content: The code text
                - font_family: The font used (e.g., "Courier New")
                - start_index: Character position where code starts
                - end_index: Character position where code ends
                - has_background: Whether background color is set
                - section: (if include_section_context) The heading this code is under
            - document_link: Link to open the document

    Example Response:
        {
            "total_code_blocks": 2,
            "code_blocks": [
                {
                    "content": "def hello():\\n    print('Hello')",
                    "font_family": "Courier New",
                    "start_index": 150,
                    "end_index": 185,
                    "has_background": true,
                    "section": "Implementation"
                }
            ],
            "document_link": "https://docs.google.com/document/d/..."
        }
    """
    import json

    logger.debug(f"[extract_code_blocks] Doc={document_id}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Get the document
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Common monospace fonts used for code
    MONOSPACE_FONTS = {
        "courier new",
        "consolas",
        "monaco",
        "menlo",
        "source code pro",
        "fira code",
        "jetbrains mono",
        "roboto mono",
        "ubuntu mono",
        "droid sans mono",
        "liberation mono",
        "dejavu sans mono",
        "lucida console",
        "andale mono",
        "courier",
    }

    # Get headings for section context
    headings = []
    if include_section_context:
        headings = get_all_headings(doc_data)

    def find_section_for_index(idx: int) -> str:
        """Find which section heading a given index falls under."""
        if not headings:
            return ""
        current_section = ""
        for heading in headings:
            if heading["start_index"] <= idx:
                current_section = heading["text"]
            else:
                break
        return current_section

    def is_code_formatted(text_style: dict) -> tuple[bool, str, bool]:
        """Check if text style indicates code formatting.

        Returns:
            Tuple of (is_code, font_family, has_background)
        """
        # Check font family
        font_info = text_style.get("weightedFontFamily", {})
        font_family = font_info.get("fontFamily", "").lower()

        is_monospace = any(mono in font_family for mono in MONOSPACE_FONTS)

        # Check background color
        bg_color = text_style.get("backgroundColor", {})
        has_background = bool(bg_color.get("color", {}))

        return is_monospace, font_info.get("fontFamily", ""), has_background

    # Collect code-formatted text runs
    code_runs = []

    body = doc_data.get("body", {})
    content = body.get("content", [])

    def process_elements(elements: list, depth: int = 0) -> None:
        """Recursively process elements to find code-formatted text."""
        if depth > 10:
            return

        for element in elements:
            if "paragraph" in element:
                paragraph = element["paragraph"]
                for para_elem in paragraph.get("elements", []):
                    if "textRun" in para_elem:
                        text_run = para_elem["textRun"]
                        text_style = text_run.get("textStyle", {})
                        text_content = text_run.get("content", "")

                        is_code, font_family, has_background = is_code_formatted(
                            text_style
                        )

                        if is_code and text_content.strip():
                            code_runs.append(
                                {
                                    "content": text_content,
                                    "font_family": font_family,
                                    "start_index": para_elem.get("startIndex", 0),
                                    "end_index": para_elem.get("endIndex", 0),
                                    "has_background": has_background,
                                }
                            )

            elif "table" in element:
                table = element["table"]
                for row in table.get("tableRows", []):
                    for cell in row.get("tableCells", []):
                        process_elements(cell.get("content", []), depth + 1)

    process_elements(content)

    # Merge consecutive code runs into blocks
    code_blocks = []
    current_block = None

    for run in sorted(code_runs, key=lambda x: x["start_index"]):
        if current_block is None:
            current_block = run.copy()
        elif run["start_index"] <= current_block["end_index"] + 1:
            # Consecutive or overlapping - merge
            current_block["content"] += run["content"]
            current_block["end_index"] = max(
                current_block["end_index"], run["end_index"]
            )
            current_block["has_background"] = (
                current_block["has_background"] or run["has_background"]
            )
        else:
            # Gap - save current block and start new one
            if include_section_context:
                current_block["section"] = find_section_for_index(
                    current_block["start_index"]
                )
            code_blocks.append(current_block)
            current_block = run.copy()

    # Don't forget the last block
    if current_block:
        if include_section_context:
            current_block["section"] = find_section_for_index(
                current_block["start_index"]
            )
        code_blocks.append(current_block)

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    result = {
        "total_code_blocks": len(code_blocks),
        "code_blocks": code_blocks,
        "document_link": link,
    }

    return json.dumps(result, indent=2)


@server.tool()
@handle_http_errors("extract_document_summary", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def extract_document_summary(
    service: Any,
    user_google_email: str,
    document_id: str,
    max_preview_words: int = 30,
) -> str:
    """
    Generate a structural summary/outline of a Google Doc.

    This tool provides a high-level overview of the document including:
    - Hierarchical heading structure (outline)
    - Section statistics (word counts, element counts)
    - Document metadata
    - Content distribution

    Useful for:
    - Understanding document structure before making changes
    - Finding specific sections by browsing the outline
    - Assessing document size and complexity

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to summarize
        max_preview_words: Maximum words to include in section previews (default: 30)

    Returns:
        str: JSON containing:
            - title: Document title
            - total_characters: Total character count
            - total_headings: Number of headings
            - total_paragraphs: Number of paragraphs
            - total_tables: Number of tables
            - total_lists: Number of lists
            - outline: Hierarchical heading structure with:
                - text: Heading text
                - level: Heading level (1-6)
                - start_index: Position in document
                - preview: First N words of section content
                - children: Nested subheadings
            - document_link: Link to open the document

    Example Response:
        {
            "title": "Project Proposal",
            "total_characters": 15420,
            "total_headings": 8,
            "total_paragraphs": 42,
            "total_tables": 2,
            "total_lists": 5,
            "outline": [
                {
                    "text": "Introduction",
                    "level": 1,
                    "start_index": 1,
                    "preview": "This document outlines the proposed approach for...",
                    "children": [
                        {
                            "text": "Background",
                            "level": 2,
                            ...
                        }
                    ]
                }
            ],
            "document_link": "https://docs.google.com/document/d/..."
        }
    """
    import json

    logger.debug(f"[extract_document_summary] Doc={document_id}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Get the document
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Get structural elements
    elements = extract_structural_elements(doc_data)

    # Count elements by type
    counts = {"headings": 0, "paragraphs": 0, "tables": 0, "lists": 0}

    for elem in elements:
        elem_type = elem.get("type", "")
        if elem_type.startswith("heading") or elem_type == "title":
            counts["headings"] += 1
        elif elem_type == "paragraph":
            counts["paragraphs"] += 1
        elif elem_type == "table":
            counts["tables"] += 1
        elif elem_type in ("bullet_list", "numbered_list"):
            counts["lists"] += 1

    # Calculate total characters
    body = doc_data.get("body", {})
    content = body.get("content", [])
    total_chars = 0
    if content:
        total_chars = content[-1].get("endIndex", 0) - 1  # Subtract 1 for start index

    # Build outline with previews
    outline = build_headings_outline(elements)

    def get_section_preview(heading_start: int, heading_end: int) -> str:
        """Get preview text from section content."""
        preview_parts = []
        word_count = 0

        for elem in elements:
            if elem["start_index"] > heading_end:
                # We're past this heading
                if elem.get("type", "").startswith("heading"):
                    # Hit next heading, stop
                    break

            if elem["start_index"] > heading_start:
                text = elem.get("text", "")
                if text:
                    words = text.split()
                    for word in words:
                        if word_count >= max_preview_words:
                            break
                        preview_parts.append(word)
                        word_count += 1
                    if word_count >= max_preview_words:
                        break

        preview = " ".join(preview_parts)
        if word_count >= max_preview_words:
            preview += "..."
        return preview

    def add_previews_to_outline(outline_items: list, elements_list: list) -> None:
        """Recursively add preview text to outline items."""
        for i, item in enumerate(outline_items):
            # Find section end (next sibling heading start or document end)
            section_end = total_chars
            if i + 1 < len(outline_items):
                section_end = outline_items[i + 1]["start_index"]

            item["preview"] = get_section_preview(item["start_index"], section_end)

            if item.get("children"):
                add_previews_to_outline(item["children"], elements_list)

    add_previews_to_outline(outline, elements)

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    result = {
        "title": doc_data.get("title", ""),
        "total_characters": total_chars,
        "total_headings": counts["headings"],
        "total_paragraphs": counts["paragraphs"],
        "total_tables": counts["tables"],
        "total_lists": counts["lists"],
        "outline": outline,
        "document_link": link,
    }

    return json.dumps(result, indent=2)


@server.tool()
@handle_http_errors("get_doc_statistics", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def get_doc_statistics(
    service: Any,
    user_google_email: str,
    document_id: str,
    include_breakdown: bool = False,
) -> str:
    """
    Get document statistics including word count, character count, and structural metrics.

    This tool provides comprehensive statistics about a Google Doc including:
    - Word count (total and by section if breakdown requested)
    - Character count (with and without spaces)
    - Paragraph count
    - Sentence count (approximate)
    - Structural element counts (headings, tables, lists, images)
    - Reading time estimate

    Useful for:
    - Checking if document meets length requirements
    - Tracking writing progress
    - Understanding document complexity
    - Comparing document sizes

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to analyze
        include_breakdown: If True, include word counts per section (default: False)

    Returns:
        str: JSON containing:
            - title: Document title
            - word_count: Total number of words
            - character_count: Total characters (including spaces)
            - character_count_no_spaces: Characters excluding spaces
            - paragraph_count: Number of paragraphs
            - sentence_count: Approximate number of sentences
            - page_count_estimate: Estimated page count (based on ~500 words/page)
            - reading_time_minutes: Estimated reading time (based on 200 words/min)
            - structure: Object containing counts of structural elements
                - headings: Number of headings
                - tables: Number of tables
                - lists: Number of lists
                - images: Number of inline images
            - section_breakdown: (if include_breakdown=True) Word counts per heading section
            - document_link: Link to open the document

    Example Response:
        {
            "title": "Project Proposal",
            "word_count": 2500,
            "character_count": 15420,
            "character_count_no_spaces": 12920,
            "paragraph_count": 42,
            "sentence_count": 125,
            "page_count_estimate": 5,
            "reading_time_minutes": 13,
            "structure": {
                "headings": 8,
                "tables": 2,
                "lists": 5,
                "images": 3
            },
            "document_link": "https://docs.google.com/document/d/..."
        }
    """
    import json
    import re

    logger.debug(f"[get_doc_statistics] Doc={document_id}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Get the document
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Extract all text from the document body
    body = doc_data.get("body", {})
    content = body.get("content", [])

    def extract_all_text(elements: list) -> str:
        """Extract all text from document elements."""
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
                        text_parts.append(extract_all_text(cell_content))
        return "".join(text_parts)

    full_text = extract_all_text(content)

    # Calculate basic statistics
    # Word count: split on whitespace and filter empty strings
    words = [w for w in full_text.split() if w.strip()]
    word_count = len(words)

    # Character counts
    char_count = len(full_text)
    char_count_no_spaces = len(full_text.replace(" ", "").replace("\t", "").replace("\n", ""))

    # Sentence count (approximate - count sentence-ending punctuation)
    sentence_endings = re.findall(r"[.!?]+", full_text)
    sentence_count = len(sentence_endings)

    # Paragraph count - count structural paragraphs
    paragraph_count = 0
    for element in content:
        if "paragraph" in element:
            # Only count non-empty paragraphs
            para = element["paragraph"]
            para_text = ""
            for elem in para.get("elements", []):
                if "textRun" in elem:
                    para_text += elem["textRun"].get("content", "")
            if para_text.strip():
                paragraph_count += 1

    # Count structural elements
    elements = extract_structural_elements(doc_data)
    structure_counts = {"headings": 0, "tables": 0, "lists": 0, "images": 0}

    for elem in elements:
        elem_type = elem.get("type", "")
        if elem_type.startswith("heading") or elem_type == "title":
            structure_counts["headings"] += 1
        elif elem_type == "table":
            structure_counts["tables"] += 1
        elif elem_type in ("bullet_list", "numbered_list"):
            structure_counts["lists"] += 1

    # Count inline images
    for element in content:
        if "paragraph" in element:
            para = element["paragraph"]
            for elem in para.get("elements", []):
                if "inlineObjectElement" in elem:
                    structure_counts["images"] += 1

    # Calculate estimates
    # Average page is ~500 words (double-spaced, 12pt font)
    page_count_estimate = max(1, round(word_count / 500))

    # Average reading speed is ~200-250 words per minute
    reading_time_minutes = max(1, round(word_count / 200))

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    result = {
        "title": doc_data.get("title", ""),
        "word_count": word_count,
        "character_count": char_count,
        "character_count_no_spaces": char_count_no_spaces,
        "paragraph_count": paragraph_count,
        "sentence_count": sentence_count,
        "page_count_estimate": page_count_estimate,
        "reading_time_minutes": reading_time_minutes,
        "structure": structure_counts,
        "document_link": link,
    }

    # Add section breakdown if requested
    if include_breakdown:
        headings = [e for e in elements if e.get("type", "").startswith("heading") or e.get("type") == "title"]

        section_breakdown = []
        for i, heading in enumerate(headings):
            section_start = heading["start_index"]
            # Section ends at next heading or end of document
            if i + 1 < len(headings):
                section_end = headings[i + 1]["start_index"]
            else:
                section_end = content[-1].get("endIndex", section_start) if content else section_start

            # Extract text for this section
            section_text = ""
            for element in content:
                elem_start = element.get("startIndex", 0)
                elem_end = element.get("endIndex", 0)

                # Skip elements outside this section
                if elem_end <= section_start or elem_start >= section_end:
                    continue

                if "paragraph" in element:
                    para = element["paragraph"]
                    for elem in para.get("elements", []):
                        if "textRun" in elem:
                            text = elem["textRun"].get("content", "")
                            text_start = elem.get("startIndex", 0)
                            text_end = elem.get("endIndex", 0)

                            # Calculate overlap with section
                            overlap_start = max(section_start, text_start)
                            overlap_end = min(section_end, text_end)

                            if overlap_start < overlap_end:
                                # Calculate character offsets within the text
                                char_start = overlap_start - text_start
                                char_end = overlap_end - text_start
                                section_text += text[char_start:char_end]

            section_words = [w for w in section_text.split() if w.strip()]
            section_breakdown.append({
                "heading": heading.get("text", "").strip(),
                "level": heading.get("level", 1),
                "word_count": len(section_words),
            })

        result["section_breakdown"] = section_breakdown

    return json.dumps(result, indent=2)


# Create comment management tools for documents
_comment_tools = create_comment_tools("document", "document_id")

# Extract and register the functions
read_doc_comments = _comment_tools["read_comments"]
create_doc_comment = _comment_tools["create_comment"]
reply_to_comment = _comment_tools["reply_to_comment"]
resolve_comment = _comment_tools["resolve_comment"]


# =============================================================================
# Operation History and Undo Tools
# =============================================================================


@server.tool()
async def get_doc_operation_history(
    document_id: str,
    user_google_email: str = None,
    limit: int = 10,
    include_undone: bool = False,
) -> str:
    """
    Get the operation history for a Google Doc.

    Returns recent operations performed on this document through MCP tools,
    showing what can be undone. Use this before attempting undo to understand
    what's available and what limitations apply.

    ⚠️  CRITICAL LIMITATIONS:
    • SESSION-ONLY: History is stored in memory and is LOST when the MCP
      server restarts. This is NOT persistent storage.
    • MCP-ONLY: Only tracks operations made through this MCP server.
      Edits made in browser, mobile app, or other integrations are NOT tracked.
    • PER-DOCUMENT: Each document has independent history (max 50 operations).

    When to use this tool:
    • Before attempting undo, to see what operations are available
    • To audit what changes the MCP has made to a document in this session
    • To check undo_capability before deciding whether to attempt undo

    When NOT to use this tool:
    • After server restart (history will be empty)
    • To find changes made outside MCP (use Google Docs version history instead)

    Args:
        document_id: ID of the document
        user_google_email: Google email of the user (accepted for API consistency, not used)
        limit: Maximum number of operations to return (default: 10, max: 50)
        include_undone: Whether to include already-undone operations (default: False)

    Returns:
        str: JSON containing:
            - document_id: The document ID
            - operations: List of operations, most recent first, each containing:
                - id: Unique operation ID
                - timestamp: When the operation was performed (ISO 8601)
                - operation_type: Type of operation (insert_text, delete_text, etc.)
                - start_index: Start position of the operation
                - end_index: End position (for range operations)
                - position_shift: How much positions shifted
                - undo_capability: "full", "partial", or "none"
                - undone: Whether this operation has been undone
                - undo_notes: Any notes about undo limitations
            - total_operations: Total number of tracked operations
            - undoable_count: Number of operations that can be undone
    """
    import json

    logger.debug(f"[get_doc_operation_history] Doc={document_id}, limit={limit}")

    # Validate inputs
    if not document_id or not document_id.strip():
        return json.dumps(
            {"success": False, "error": "document_id is required"}, indent=2
        )

    limit = min(max(1, limit), 50)  # Clamp between 1 and 50

    manager = get_history_manager()
    operations = manager.get_history(
        document_id, limit=limit, include_undone=include_undone
    )

    # Count undoable operations
    undoable_count = sum(
        1
        for op in operations
        if not op.get("undone") and op.get("undo_capability") != "none"
    )

    result = {
        "document_id": document_id,
        "operations": operations,
        "total_operations": len(operations),
        "undoable_count": undoable_count,
    }

    return json.dumps(result, indent=2)


@server.tool()
@handle_http_errors("undo_doc_operation", service_type="docs")
@require_google_service("docs", "docs_write")
async def undo_doc_operation(
    service: Any,
    user_google_email: str,
    document_id: str,
    operation_id: str = None,
) -> str:
    """
    Undo the last operation performed on a Google Doc.

    Reverses the most recent undoable operation by executing a compensating
    operation (e.g., re-inserting deleted text, deleting inserted text).

    ⚠️  CRITICAL LIMITATIONS - Undo is FRAGILE:
    • SESSION-ONLY: History is LOST on server restart. No undo available after restart.
    • EXTERNAL EDITS BREAK UNDO: If the document was modified outside MCP
      (browser, mobile, another user) since the operation, undo will likely
      fail or corrupt the document. Indices shift when content changes.
    • LIMITED OPERATION SUPPORT:
      - FULL undo: insert_text, delete_text, replace_text, page_break
      - PARTIAL undo: format_text (only if original formatting was captured)
      - NO undo: find_replace, insert_table, insert_image, complex operations

    When to use:
    • Immediately after an MCP operation made a mistake (same session)
    • When you're confident no external edits occurred since the operation
    • For simple text operations (insert, delete, replace)

    When NOT to use:
    • After server restart (no history exists)
    • If the document may have been edited externally since the operation
    • For complex operations like find_replace or table insertions
    • When significant time has passed (prefer Google Docs version history)

    SAFER ALTERNATIVE: For important documents or after external edits,
    guide the user to use Google Docs' built-in version history
    (File > Version history) which tracks ALL changes reliably.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document
        operation_id: Specific operation ID to undo (optional, defaults to last undoable)

    Returns:
        str: JSON containing:
            - success: Whether the undo was successful
            - message: Description of what was undone
            - operation_id: ID of the undone operation
            - reverse_operation: Details of the reverse operation executed
            - error: Error message if undo failed
    """
    import json

    logger.info(f"[undo_doc_operation] Doc={document_id}, OpID={operation_id}")

    # Validate document_id
    validator = ValidationManager()
    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    manager = get_history_manager()

    # Generate the undo operation
    undo_result = manager.generate_undo_operation(document_id)

    if not undo_result.success:
        return json.dumps(
            {
                "success": False,
                "message": undo_result.message,
                "error": undo_result.error,
            },
            indent=2,
        )

    # Execute the reverse operation
    reverse_op = undo_result.reverse_operation
    op_type = reverse_op.get("type")

    try:
        requests = []

        if op_type == "insert_text":
            requests.append(
                create_insert_text_request(reverse_op["index"], reverse_op["text"])
            )
        elif op_type == "delete_text":
            requests.append(
                create_delete_range_request(
                    reverse_op["start_index"], reverse_op["end_index"]
                )
            )
        elif op_type == "replace_text":
            # Replace = delete + insert
            requests.append(
                create_delete_range_request(
                    reverse_op["start_index"], reverse_op["end_index"]
                )
            )
            requests.append(
                create_insert_text_request(
                    reverse_op["start_index"], reverse_op["text"]
                )
            )
        elif op_type == "format_text":
            requests.append(
                create_format_text_request(
                    reverse_op["start_index"],
                    reverse_op["end_index"],
                    reverse_op.get("bold"),
                    reverse_op.get("italic"),
                    reverse_op.get("underline"),
                    reverse_op.get("font_size"),
                    reverse_op.get("font_family"),
                )
            )
        else:
            return json.dumps(
                {
                    "success": False,
                    "message": f"Unknown reverse operation type: {op_type}",
                    "error": "Cannot execute undo",
                },
                indent=2,
            )

        # Execute the batch update
        await asyncio.to_thread(
            service.documents()
            .batchUpdate(documentId=document_id, body={"requests": requests})
            .execute
        )

        # Mark the operation as undone
        manager.mark_undone(document_id, undo_result.operation_id)

        return json.dumps(
            {
                "success": True,
                "message": "Successfully undone operation",
                "operation_id": undo_result.operation_id,
                "reverse_operation": {
                    "type": op_type,
                    "description": reverse_op.get("description", ""),
                },
                "document_link": f"https://docs.google.com/document/d/{document_id}/edit",
            },
            indent=2,
        )

    except Exception as e:
        logger.error(f"Failed to execute undo: {str(e)}")
        return json.dumps(
            {
                "success": False,
                "message": "Failed to execute undo operation",
                "operation_id": undo_result.operation_id,
                "error": str(e),
            },
            indent=2,
        )


@server.tool()
async def clear_doc_history(
    document_id: str,
) -> str:
    """
    Clear the operation history for a Google Doc.

    Removes all tracked operations for the specified document, making undo
    unavailable for previous operations. This is useful when you want to
    "commit" your changes and prevent accidental undo, or to free memory.

    Note: History is already lost on server restart, so this is mainly useful
    within a long-running session where you want to explicitly discard undo
    capability for a document.

    When to use:
    • After completing a set of changes you want to keep permanently
    • To free memory if tracking many operations on a document
    • When starting a new logical "unit of work" on a document

    Args:
        document_id: ID of the document

    Returns:
        str: JSON containing:
            - success: Whether the history was cleared
            - message: Description of the result
            - document_id: The document ID
    """
    import json

    logger.info(f"[clear_doc_history] Doc={document_id}")

    if not document_id or not document_id.strip():
        return json.dumps(
            {"success": False, "error": "document_id is required"}, indent=2
        )

    manager = get_history_manager()
    cleared = manager.clear_history(document_id)

    return json.dumps(
        {
            "success": True,
            "message": "Cleared history for document"
            if cleared
            else "No history existed for document",
            "document_id": document_id,
        },
        indent=2,
    )


@server.tool()
async def get_history_stats() -> str:
    """
    Get statistics about tracked operation history across all documents.

    Provides an overview of the in-memory history tracking system, showing
    how many documents and operations are currently tracked in this session.

    Note: All statistics reset to zero when the MCP server restarts since
    history is stored in memory only. This tool is useful for debugging
    or understanding current session state.

    When to use:
    • Debugging: to verify operations are being tracked
    • Memory monitoring: to see how many operations are stored
    • Before bulk operations: to understand current state

    Returns:
        str: JSON containing:
            - documents_tracked: Number of documents with tracked history
            - total_operations: Total operations across all documents
            - undone_operations: Number of operations that have been undone
            - operations_per_document: Dictionary mapping document IDs to operation counts
    """
    import json

    logger.debug("[get_history_stats]")

    manager = get_history_manager()
    stats = manager.get_stats()

    return json.dumps(stats, indent=2)


@server.tool()
@handle_http_errors("record_doc_operation", service_type="docs")
@require_google_service("docs", "docs_read")
async def record_doc_operation(
    service: Any,
    user_google_email: str,
    document_id: str,
    operation_type: str,
    start_index: int,
    end_index: int = None,
    text: str = None,
    position_shift: int = 0,
    capture_deleted_text: bool = True,
) -> str:
    """
    Manually record an operation for undo tracking.

    NOTE: Operations through modify_doc_text and batch_edit_doc are now
    automatically recorded. This tool is only needed for advanced scenarios
    or custom integrations that bypass the standard tools.

    ⚠️  LIMITATIONS (same as all undo functionality):
    • SESSION-ONLY: Recorded history is LOST on server restart.
    • FRAGILE: Undo fails if document is edited externally after recording.

    Call this BEFORE performing the operation to capture the original text
    that will be deleted or replaced.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document
        operation_type: Type of operation. Supported values:
            - "insert_text": Inserting new text (full undo support)
            - "delete_text": Deleting existing text (full undo support)
            - "replace_text": Replacing text with new text (full undo support)
            - "format_text": Applying formatting (partial undo - needs original format)
        start_index: Start position of the operation
        end_index: End position for range operations (delete/replace/format)
        text: Text being inserted or replacing original (for insert/replace)
        position_shift: How much positions will shift (calculated if not provided)
        capture_deleted_text: If True, captures text at range before operation (for undo)

    Returns:
        str: JSON containing:
            - success: Whether the operation was recorded
            - operation_id: Unique ID for the recorded operation
            - undo_capability: "full", "partial", or "none"
            - deleted_text: Text that was captured (if applicable)
            - message: Description of what was recorded
    """
    import json

    logger.info(
        f"[record_doc_operation] Doc={document_id}, Type={operation_type}, Start={start_index}"
    )

    # Validate inputs
    validator = ValidationManager()
    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    valid_types = ["insert_text", "delete_text", "replace_text", "format_text"]
    if operation_type not in valid_types:
        return json.dumps(
            {
                "success": False,
                "error": f"Invalid operation_type. Must be one of: {', '.join(valid_types)}",
            },
            indent=2,
        )

    if start_index < 0:
        return json.dumps(
            {"success": False, "error": "start_index must be non-negative"}, indent=2
        )

    # Determine undo capability
    undo_capability = UndoCapability.FULL
    undo_notes = None
    deleted_text = None
    original_text = None

    # Capture text if needed for undo
    if capture_deleted_text and operation_type in ["delete_text", "replace_text"]:
        if end_index is None or end_index <= start_index:
            return json.dumps(
                {
                    "success": False,
                    "error": "end_index is required and must be greater than start_index for delete/replace operations",
                },
                indent=2,
            )

        # Fetch document to capture the text
        doc_data = await asyncio.to_thread(
            service.documents().get(documentId=document_id).execute
        )

        # Extract the text at the range
        extracted = extract_text_at_range(doc_data, start_index, end_index)
        captured_text = extracted.get("text", "")

        if operation_type == "delete_text":
            deleted_text = captured_text
        else:  # replace_text
            original_text = captured_text

    # Calculate position shift if not provided
    if position_shift == 0:
        if operation_type == "insert_text":
            position_shift = len(text) if text else 0
        elif operation_type == "delete_text":
            position_shift = -(end_index - start_index) if end_index else 0
        elif operation_type == "replace_text":
            old_len = (end_index - start_index) if end_index else 0
            new_len = len(text) if text else 0
            position_shift = new_len - old_len

    # Handle operation-specific undo limitations
    if operation_type == "format_text":
        undo_capability = UndoCapability.NONE
        undo_notes = (
            "Format undo requires capturing original formatting (not yet supported)"
        )

    # Record the operation
    manager = get_history_manager()
    snapshot = manager.record_operation(
        document_id=document_id,
        operation_type=operation_type,
        operation_params={
            "start_index": start_index,
            "end_index": end_index,
            "text": text,
        },
        start_index=start_index,
        end_index=end_index,
        position_shift=position_shift,
        deleted_text=deleted_text,
        original_text=original_text,
        undo_capability=undo_capability,
        undo_notes=undo_notes,
    )

    result = {
        "success": True,
        "operation_id": snapshot.id,
        "operation_type": operation_type,
        "undo_capability": undo_capability.value,
        "message": f"Recorded {operation_type} operation",
    }

    if deleted_text is not None:
        result["deleted_text_preview"] = (
            deleted_text[:100] + "..." if len(deleted_text) > 100 else deleted_text
        )
        result["deleted_text_length"] = len(deleted_text)

    if original_text is not None:
        result["original_text_preview"] = (
            original_text[:100] + "..." if len(original_text) > 100 else original_text
        )
        result["original_text_length"] = len(original_text)

    return json.dumps(result, indent=2)


@server.tool()
@handle_http_errors("clear_doc_formatting", is_read_only=False, service_type="docs")
@require_google_service("docs", "docs_write")
async def clear_doc_formatting(
    service: Any,
    user_google_email: str,
    document_id: str,
    start_index: int = None,
    end_index: int = None,
    search: str = None,
    occurrence: int = 1,
    match_case: bool = True,
    range: Dict[str, Any] = None,
    location: str = None,
    preserve_links: bool = False,
    preview: bool = False,
) -> str:
    """
    Clears/removes all text formatting from a range in a Google Doc, returning it to default style.

    Use this tool when you want to strip formatting from text - commonly needed after pasting
    content from other sources that brings unwanted formatting.

    This removes:
    - Bold, italic, underline, strikethrough
    - Small caps, subscript, superscript
    - Text color (foreground and background/highlight)
    - Hyperlinks (unless preserve_links=True)

    Note: Font size and font family are NOT reset, as these inherit from paragraph/document
    styles and there's no universal "default" value.

    Supports multiple positioning modes (same as modify_doc_text):
    1. Index-based: Use start_index and end_index for exact positions
    2. Search-based: Use search parameter to find and clear formatting from specific text
    3. Range-based: Use range parameter for semantic text selection
    4. Location-based: Use location='all' to clear formatting from entire document body

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update

        Index-based positioning:
        start_index: Start position for operation (0-based)
        end_index: End position (required for index-based mode)

        Search-based positioning:
        search: Text to search for in the document (clears formatting from found text)
        occurrence: Which occurrence to target (1=first, 2=second, -1=last). Default: 1
        match_case: Whether to match case exactly. Default: True

        Range-based positioning:
        range: Dictionary specifying a semantic range. Supported formats:
            1. Range by search bounds:
               {"start": {"search": "Introduction"}, "end": {"search": "Conclusion"}}
            2. Search with boundary extension:
               {"search": "keyword", "extend": "paragraph"}
            3. Section reference:
               {"section": "Section Title", "include_heading": True}

        Location-based positioning:
        location: Use "all" to clear formatting from entire document body

        Options:
        preserve_links: If True, hyperlinks will not be removed (default: False)
        preview: If True, returns what would change without modifying (default: False)

    Examples:
        # Clear formatting from a specific range (index-based):
        clear_doc_formatting(document_id="...", start_index=10, end_index=50)

        # Clear formatting from specific found text (search-based):
        clear_doc_formatting(document_id="...", search="some formatted text")

        # Clear formatting from entire paragraph containing a keyword:
        clear_doc_formatting(document_id="...",
                           range={"search": "keyword", "extend": "paragraph"})

        # Clear formatting from entire document:
        clear_doc_formatting(document_id="...", location="all")

        # Clear formatting but keep hyperlinks:
        clear_doc_formatting(document_id="...", start_index=10, end_index=50,
                           preserve_links=True)

        # Preview what would be cleared:
        clear_doc_formatting(document_id="...", search="formatted text", preview=True)

    Returns:
        str: JSON string with operation details.

        When preview=False (default), returns structured operation result:
        {
            "success": true,
            "operation": "clear_formatting",
            "affected_range": {"start": 10, "end": 50},
            "formatting_cleared": ["bold", "italic", "underline", ...],
            "links_preserved": false,
            "message": "Cleared formatting from 40 characters (indices 10-50)",
            "link": "https://docs.google.com/document/d/.../edit"
        }

        When preview=True, returns a preview response:
        {
            "preview": true,
            "would_modify": true,
            "affected_range": {"start": 10, "end": 50},
            "current_content": "the formatted text",
            "context": {"before": "...", "after": "..."},
            "formatting_to_clear": ["bold", "italic", ...],
            "message": "Would clear formatting from 40 characters"
        }
    """
    import json

    logger.info(
        f"[clear_doc_formatting] Doc={document_id}, location={location}, search={search}, "
        f"range={range is not None}, start={start_index}, end={end_index}"
    )

    validator = ValidationManager()

    # Validate document_id
    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Determine positioning mode
    use_range_mode = range is not None
    use_search_mode = search is not None
    use_location_mode = location is not None

    # Validate positioning parameters
    if use_location_mode:
        if location != "all":
            return validator.create_invalid_param_error(
                param_name="location", received=location, valid_values=["all"]
            )
        if (
            start_index is not None
            or end_index is not None
            or search is not None
            or range is not None
        ):
            logger.warning(
                "Multiple positioning parameters provided; location mode takes precedence"
            )
    elif use_range_mode:
        if start_index is not None or end_index is not None or search is not None:
            logger.warning(
                "Multiple positioning parameters provided; range mode takes precedence"
            )
    elif use_search_mode:
        if search == "":
            return validator.create_empty_search_error()
        if start_index is not None or end_index is not None:
            logger.warning(
                "Both search and index parameters provided; search mode takes precedence"
            )
    else:
        # Index-based mode
        if start_index is None:
            return validator.create_missing_param_error(
                param_name="positioning",
                context="for clearing formatting",
                valid_values=[
                    "location='all'",
                    "range parameter",
                    "search",
                    "start_index + end_index",
                ],
            )
        if end_index is None:
            return validator.create_missing_param_error(
                param_name="end_index",
                context="for index-based formatting clear",
                valid_values=["integer greater than start_index"],
            )
        is_valid, index_error = validator.validate_index_range_structured(
            start_index, end_index
        )
        if not is_valid:
            return index_error

    # Fetch document to resolve positions
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )
    doc_link = f"https://docs.google.com/document/d/{document_id}/edit"

    # Resolve indices based on positioning mode
    search_info = {}

    if use_location_mode:
        # Clear entire document body
        structure = parse_document_structure(doc_data)
        start_index = 1  # Start after initial section break
        end_index = structure["total_length"] - 1  # Before final newline
        search_info = {
            "location": "all",
            "resolved_start": start_index,
            "resolved_end": end_index,
            "message": f"Clearing formatting from entire document body (indices {start_index}-{end_index})",
        }

    elif use_range_mode:
        # Resolve range specification
        range_result = resolve_range(doc_data, range)

        if not range_result.success:
            error_response = {
                "success": False,
                "error": "range_resolution_failed",
                "message": range_result.message,
                "hint": "Check range specification format and search terms",
            }
            return json.dumps(error_response, indent=2)

        start_index = range_result.start_index
        end_index = range_result.end_index
        search_info = {
            "range": range,
            "resolved_start": start_index,
            "resolved_end": end_index,
            "message": range_result.message,
        }

    elif use_search_mode:
        # Find the search text
        success, calc_start, calc_end, message = calculate_search_based_indices(
            doc_data, search, SearchPosition.REPLACE.value, occurrence, match_case
        )

        if not success:
            all_occurrences = find_all_occurrences_in_document(
                doc_data, search, match_case
            )
            if all_occurrences:
                if "occurrence" in message.lower():
                    return validator.create_invalid_occurrence_error(
                        occurrence=occurrence,
                        total_found=len(all_occurrences),
                        search_text=search,
                    )
                occurrences_data = [
                    {"index": i + 1, "position": f"{s}-{e}"}
                    for i, (s, e) in enumerate(all_occurrences[:5])
                ]
                return validator.create_ambiguous_search_error(
                    search_text=search,
                    occurrences=occurrences_data,
                    total_count=len(all_occurrences),
                )
            return validator.create_search_not_found_error(
                search_text=search, match_case=match_case
            )

        start_index = calc_start
        end_index = calc_end
        search_info = {
            "search_text": search,
            "occurrence": occurrence,
            "found_at_range": f"{calc_start}-{calc_end}",
            "message": message,
        }

    # Validate we have a valid range
    if start_index is None or end_index is None or end_index <= start_index:
        return json.dumps(
            {
                "success": False,
                "error": "invalid_range",
                "message": "Could not determine a valid range for clearing formatting",
                "hint": "Ensure start_index < end_index or provide valid search/range parameters",
            },
            indent=2,
        )

    # Build list of formatting that will be cleared
    formatting_cleared = [
        "bold",
        "italic",
        "underline",
        "strikethrough",
        "small_caps",
        "subscript",
        "superscript",
        "foreground_color",
        "background_color",
    ]
    if not preserve_links:
        formatting_cleared.append("link")

    # Handle preview mode
    if preview:
        # Extract current content at the range
        current = extract_text_at_range(doc_data, start_index, end_index)

        preview_result = {
            "preview": True,
            "would_modify": True,
            "affected_range": {"start": start_index, "end": end_index},
            "formatting_to_clear": formatting_cleared,
            "links_preserved": preserve_links,
            "link": doc_link,
        }

        if current.get("found"):
            preview_result["current_content"] = current.get("text", "")
            preview_result["context"] = {
                "before": current.get("context_before", ""),
                "after": current.get("context_after", ""),
            }

        if search_info:
            preview_result["positioning_info"] = search_info

        char_count = end_index - start_index
        preview_result["message"] = (
            f"Would clear formatting from {char_count} characters (indices {start_index}-{end_index})"
        )

        return json.dumps(preview_result, indent=2)

    # Create and execute the clear formatting request
    clear_request = create_clear_formatting_request(
        start_index, end_index, preserve_links
    )

    await asyncio.to_thread(
        service.documents()
        .batchUpdate(documentId=document_id, body={"requests": [clear_request]})
        .execute
    )

    # Build success response
    char_count = end_index - start_index
    result = {
        "success": True,
        "operation": "clear_formatting",
        "affected_range": {"start": start_index, "end": end_index},
        "formatting_cleared": formatting_cleared,
        "links_preserved": preserve_links,
        "message": f"Cleared formatting from {char_count} characters (indices {start_index}-{end_index})",
        "link": doc_link,
    }

    if search_info:
        result["positioning_info"] = search_info

    return json.dumps(result, indent=2)


# =============================================================================
# Named Range Tools
# =============================================================================


@server.tool()
@handle_http_errors("create_doc_named_range", service_type="docs")
@require_google_service("docs", "docs_write")
async def create_doc_named_range(
    service: Any,
    user_google_email: str,
    document_id: str,
    name: str,
    start_index: int = None,
    end_index: int = None,
    search: str = None,
    occurrence: int = 1,
    match_case: bool = True,
    location: Literal["start", "end"] = None,
) -> str:
    """
    Create a named range in a Google Doc to mark a section for later reference.

    Named ranges allow you to mark sections of a document that can be referenced
    programmatically. The range indices automatically update as document content
    changes, eliminating the need for manual position tracking.

    POSITIONING OPTIONS (mutually exclusive):
    1. Index-based: Provide start_index and end_index directly
    2. Search-based: Use search parameter to find text and create range around it
    3. Location-based: Use location="start" or location="end" for document boundaries

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document
        name: Name for the range (1-256 characters). Names don't need to be unique;
              multiple ranges can share the same name.
        start_index: Start index of the range (inclusive). Required for index-based positioning.
        end_index: End index of the range (exclusive). Required for index-based positioning.
        search: Text to search for. The named range will be created around the found text.
        occurrence: Which occurrence of search text (1=first, 2=second, -1=last). Default: 1
        match_case: Whether to match case when searching. Default: True
        location: Create range at "start" or "end" of document (creates a zero-width marker)

    Returns:
        JSON string with created named range details including the named range ID.

    Examples:
        # Create named range by indices
        create_doc_named_range(document_id="abc123", name="introduction",
                               start_index=1, end_index=150)

        # Create named range around search text
        create_doc_named_range(document_id="abc123", name="section_header",
                               search="Chapter 1: Introduction")

        # Create named range around second occurrence of text
        create_doc_named_range(document_id="abc123", name="second_todo",
                               search="TODO", occurrence=2)

        # Create marker at end of document
        create_doc_named_range(document_id="abc123", name="document_end",
                               location="end")

    Notes:
        - Named ranges are visible to anyone with API access to the document
        - Named ranges do not duplicate when content is copied
        - Use list_doc_named_ranges to see all named ranges
        - Use delete_doc_named_range to remove a named range
    """
    import json

    logger.debug(
        f"[create_doc_named_range] Doc={document_id}, name={name}, "
        f"start={start_index}, end={end_index}, search={search}, location={location}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Validate name
    if not name:
        return validator.create_missing_param_error(
            param_name="name",
            context="for creating named range",
            valid_values=["non-empty string (1-256 characters)"],
        )
    if len(name) > 256:
        return validator.create_invalid_param_error(
            param_name="name",
            received=f"string with {len(name)} characters",
            valid_values=["string with 1-256 characters"],
            context="Named range names must be 256 characters or fewer",
        )

    # Determine positioning mode
    use_index_mode = start_index is not None or end_index is not None
    use_search_mode = search is not None
    use_location_mode = location is not None

    # Check for conflicting modes
    modes_used = sum([use_index_mode, use_search_mode, use_location_mode])
    if modes_used == 0:
        return validator.create_missing_param_error(
            param_name="positioning",
            context="for creating named range",
            valid_values=[
                "start_index + end_index",
                "search",
                "location='start' or location='end'",
            ],
        )
    if modes_used > 1:
        return validator.create_invalid_param_error(
            param_name="positioning",
            received="multiple positioning methods",
            valid_values=[
                "Use only one: index-based (start_index/end_index), "
                "search-based (search), or location-based (location)"
            ],
            context="Positioning methods are mutually exclusive",
        )

    # Fetch document to resolve positions
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )
    doc_link = f"https://docs.google.com/document/d/{document_id}/edit"

    # Resolve indices based on positioning mode
    resolved_start = None
    resolved_end = None
    positioning_info = {}

    if use_location_mode:
        structure = parse_document_structure(doc_data)
        if location == "start":
            resolved_start = 1
            resolved_end = 1
            positioning_info = {"location": "start", "message": "Range at document start"}
        elif location == "end":
            resolved_start = structure["total_length"] - 1
            resolved_end = structure["total_length"] - 1
            positioning_info = {"location": "end", "message": "Range at document end"}

    elif use_search_mode:
        if search == "":
            return validator.create_empty_search_error()

        result = find_text_in_document(doc_data, search, occurrence, match_case)
        if result is None:
            all_occurrences = find_all_occurrences_in_document(doc_data, search, match_case)
            if all_occurrences:
                if occurrence != 1:
                    return validator.create_invalid_occurrence_error(
                        occurrence=occurrence,
                        total_found=len(all_occurrences),
                        search_text=search,
                    )
                occurrences_data = [
                    {"index": i + 1, "position": f"{s}-{e}"}
                    for i, (s, e) in enumerate(all_occurrences[:5])
                ]
                return validator.create_ambiguous_search_error(
                    search_text=search,
                    occurrences=occurrences_data,
                    total_count=len(all_occurrences),
                )
            return validator.create_search_not_found_error(
                search_text=search, match_case=match_case
            )

        resolved_start, resolved_end = result
        positioning_info = {
            "search_text": search,
            "occurrence": occurrence,
            "found_at_range": f"{resolved_start}-{resolved_end}",
        }

    elif use_index_mode:
        if start_index is None:
            return validator.create_missing_param_error(
                param_name="start_index",
                context="for index-based positioning",
                valid_values=["integer >= 1"],
            )
        if end_index is None:
            return validator.create_missing_param_error(
                param_name="end_index",
                context="for index-based positioning",
                valid_values=["integer > start_index"],
            )

        is_valid, index_error = validator.validate_index_range_structured(
            start_index, end_index
        )
        if not is_valid:
            return index_error

        resolved_start = start_index
        resolved_end = end_index
        positioning_info = {"start_index": start_index, "end_index": end_index}

    # Create the named range request
    request = create_named_range_request(name, resolved_start, resolved_end)

    # Execute the request
    result = await asyncio.to_thread(
        service.documents()
        .batchUpdate(documentId=document_id, body={"requests": [request]})
        .execute
    )

    # Extract the named range ID from the response
    replies = result.get("replies", [])
    named_range_id = None
    if replies and "createNamedRange" in replies[0]:
        named_range_id = replies[0]["createNamedRange"].get("namedRangeId")

    response = {
        "success": True,
        "named_range_id": named_range_id,
        "name": name,
        "range": {"start_index": resolved_start, "end_index": resolved_end},
        "message": f"Created named range '{name}' at indices {resolved_start}-{resolved_end}",
        "link": doc_link,
    }

    if positioning_info:
        response["positioning_info"] = positioning_info

    return json.dumps(response, indent=2)


@server.tool()
@handle_http_errors("list_doc_named_ranges", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def list_doc_named_ranges(
    service: Any,
    user_google_email: str,
    document_id: str,
    name_filter: str = None,
) -> str:
    """
    List all named ranges in a Google Doc.

    Named ranges mark sections of a document that can be referenced programmatically.
    This tool retrieves all named ranges with their IDs, names, and position information.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document
        name_filter: Optional filter to show only ranges with this name

    Returns:
        JSON string with list of named ranges including:
        - named_range_id: Unique ID of the range
        - name: Name of the range
        - ranges: List of range positions (a name can have multiple ranges)

    Examples:
        # List all named ranges
        list_doc_named_ranges(document_id="abc123")

        # List only ranges with specific name
        list_doc_named_ranges(document_id="abc123", name_filter="section_header")

    Notes:
        - Multiple ranges can share the same name
        - Named ranges are visible to anyone with API access to the document
        - Range positions automatically update as document content changes
    """
    import json

    logger.debug(f"[list_doc_named_ranges] Doc={document_id}, filter={name_filter}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Fetch document with tabs content to get named ranges from all tabs
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id, includeTabsContent=True).execute
    )
    doc_link = f"https://docs.google.com/document/d/{document_id}/edit"

    def extract_named_ranges_from_dict(named_ranges_dict, tab_id=None):
        """Extract named ranges from a namedRanges dictionary."""
        result = []
        for range_name, range_data in named_ranges_dict.items():
            # Apply filter if provided
            if name_filter and range_name != name_filter:
                continue

            for nr in range_data.get("namedRanges", []):
                nr_id = nr.get("namedRangeId")
                nr_ranges = []
                for r in nr.get("ranges", []):
                    range_info = {
                        "start_index": r.get("startIndex"),
                        "end_index": r.get("endIndex"),
                        "segment_id": r.get("segmentId"),
                    }
                    nr_ranges.append(range_info)

                entry = {
                    "named_range_id": nr_id,
                    "name": range_name,
                    "ranges": nr_ranges,
                }
                if tab_id:
                    entry["tab_id"] = tab_id
                result.append(entry)
        return result

    def process_tabs_recursively(tabs, results):
        """Process tabs and child tabs to extract all named ranges."""
        for tab in tabs:
            tab_props = tab.get("tabProperties", {})
            tab_id = tab_props.get("tabId")
            doc_tab = tab.get("documentTab", {})
            tab_named_ranges = doc_tab.get("namedRanges", {})
            results.extend(extract_named_ranges_from_dict(tab_named_ranges, tab_id))
            # Process child tabs
            child_tabs = tab.get("childTabs", [])
            if child_tabs:
                process_tabs_recursively(child_tabs, results)

    named_ranges_list = []

    # With includeTabsContent=True, named ranges are in tabs, not at root level
    tabs = doc_data.get("tabs", [])
    if tabs:
        process_tabs_recursively(tabs, named_ranges_list)
    else:
        # Fallback for documents without tabs structure (shouldn't happen with includeTabsContent=True)
        root_named_ranges = doc_data.get("namedRanges", {})
        named_ranges_list = extract_named_ranges_from_dict(root_named_ranges)

    # Sort by name for consistent output
    named_ranges_list.sort(key=lambda x: (x["name"], x["named_range_id"] or ""))

    response = {
        "success": True,
        "document_id": document_id,
        "named_ranges_count": len(named_ranges_list),
        "named_ranges": named_ranges_list,
        "link": doc_link,
    }

    if name_filter:
        response["filter"] = name_filter

    if not named_ranges_list:
        if name_filter:
            response["message"] = f"No named ranges found with name '{name_filter}'"
        else:
            response["message"] = "No named ranges found in document"
    else:
        response["message"] = f"Found {len(named_ranges_list)} named range(s)"

    return json.dumps(response, indent=2)


@server.tool()
@handle_http_errors("delete_doc_named_range", service_type="docs")
@require_google_service("docs", "docs_write")
async def delete_doc_named_range(
    service: Any,
    user_google_email: str,
    document_id: str,
    named_range_id: str = None,
    name: str = None,
) -> str:
    """
    Delete a named range from a Google Doc.

    Can delete by either the specific named_range_id (deletes one range) or
    by name (deletes all ranges with that name).

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document
        named_range_id: ID of the specific named range to delete (use list_doc_named_ranges to find)
        name: Name of ranges to delete (deletes ALL ranges with this name)

    Returns:
        JSON string confirming deletion.

    Examples:
        # Delete by ID (specific range)
        delete_doc_named_range(document_id="abc123",
                               named_range_id="kix.abc123def456")

        # Delete by name (all ranges with this name)
        delete_doc_named_range(document_id="abc123", name="old_marker")

    Notes:
        - Deleting by name removes ALL ranges with that name
        - Use list_doc_named_ranges to find range IDs
        - Deleting a named range does not affect document content
    """
    import json

    logger.debug(
        f"[delete_doc_named_range] Doc={document_id}, id={named_range_id}, name={name}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    if not named_range_id and not name:
        return validator.create_missing_param_error(
            param_name="named_range_id or name",
            context="for deleting named range",
            valid_values=["named_range_id (specific range) or name (all ranges with name)"],
        )

    if named_range_id and name:
        return validator.create_invalid_param_error(
            param_name="identifier",
            received="both named_range_id and name",
            valid_values=["Provide only one: named_range_id OR name"],
            context="Cannot specify both named_range_id and name",
        )

    doc_link = f"https://docs.google.com/document/d/{document_id}/edit"

    # Fetch document to verify named range exists (using tabs content for multi-tab support)
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id, includeTabsContent=True).execute
    )

    # Find all named ranges in the document (from all tabs)
    def find_all_named_ranges(doc_data):
        """Extract all named ranges from document, including from all tabs."""
        all_ranges = {}  # name -> list of (ID, tab_id) tuples
        all_ids = {}  # id -> tab_id mapping

        def extract_from_dict(named_ranges_dict, tab_id=None):
            for range_name, range_data in named_ranges_dict.items():
                for nr in range_data.get("namedRanges", []):
                    nr_id = nr.get("namedRangeId")
                    if nr_id:
                        all_ids[nr_id] = tab_id
                        if range_name not in all_ranges:
                            all_ranges[range_name] = []
                        all_ranges[range_name].append((nr_id, tab_id))

        def process_tabs(tabs):
            for tab in tabs:
                tab_props = tab.get("tabProperties", {})
                tab_id = tab_props.get("tabId")
                doc_tab = tab.get("documentTab", {})
                tab_named_ranges = doc_tab.get("namedRanges", {})
                extract_from_dict(tab_named_ranges, tab_id)
                child_tabs = tab.get("childTabs", [])
                if child_tabs:
                    process_tabs(child_tabs)

        tabs = doc_data.get("tabs", [])
        if tabs:
            process_tabs(tabs)
        else:
            # Fallback for non-tabs response
            extract_from_dict(doc_data.get("namedRanges", {}))

        return all_ranges, all_ids

    all_ranges, all_ids = find_all_named_ranges(doc_data)

    # Verify the named range exists
    if named_range_id:
        if named_range_id not in all_ids:
            available_ids = list(all_ids)[:5]  # Show up to 5 example IDs
            return json.dumps({
                "error": True,
                "code": "NAMED_RANGE_NOT_FOUND",
                "message": f"No named range with ID '{named_range_id}' found in document",
                "suggestion": "Use list_doc_named_ranges to see available named ranges and their IDs",
                "available_named_range_ids": available_ids if available_ids else [],
                "link": doc_link,
            }, indent=2)
    else:  # deleting by name
        if name not in all_ranges:
            available_names = list(all_ranges.keys())[:5]  # Show up to 5 example names
            return json.dumps({
                "error": True,
                "code": "NAMED_RANGE_NOT_FOUND",
                "message": f"No named range with name '{name}' found in document",
                "suggestion": "Use list_doc_named_ranges to see available named ranges",
                "available_named_range_names": available_names if available_names else [],
                "link": doc_link,
            }, indent=2)

    # Determine tabs_criteria for multi-tab support
    tabs_criteria = None
    if named_range_id:
        # Get the tab_id for this specific named range
        tab_id = all_ids.get(named_range_id)
        if tab_id:
            tabs_criteria = {"tabIds": [tab_id]}
    else:
        # Deleting by name - collect all unique tab_ids for ranges with this name
        tab_ids = set()
        for _, tab_id in all_ranges.get(name, []):
            if tab_id:
                tab_ids.add(tab_id)
        if tab_ids:
            tabs_criteria = {"tabIds": list(tab_ids)}

    # Create the delete request
    try:
        request = create_delete_named_range_request(
            named_range_id=named_range_id, name=name, tabs_criteria=tabs_criteria
        )
    except ValueError as e:
        return json.dumps({"success": False, "error": str(e)}, indent=2)

    # Execute the request
    await asyncio.to_thread(
        service.documents()
        .batchUpdate(documentId=document_id, body={"requests": [request]})
        .execute
    )

    response = {
        "success": True,
        "link": doc_link,
    }

    if named_range_id:
        response["deleted_range_id"] = named_range_id
        response["message"] = f"Deleted named range with ID '{named_range_id}'"
    else:
        response["deleted_by_name"] = name
        response["message"] = f"Deleted all named ranges with name '{name}'"

    return json.dumps(response, indent=2)


@server.tool()
@handle_http_errors("convert_list_type", service_type="docs")
@require_google_service("docs", "docs_write")
async def convert_list_type(
    service: Any,
    user_google_email: str,
    document_id: str,
    list_type: str,
    list_index: int = 0,
    search: str = None,
    start_index: int = None,
    end_index: int = None,
    preview: bool = False,
) -> str:
    """
    Convert an existing list between bullet (unordered) and numbered (ordered) types.

    Use this tool to change a bullet list to a numbered list or vice versa without
    having to delete and recreate the list.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document containing the list
        list_type: Target list type to convert to. Values:
            - "ORDERED" or "numbered": Convert to numbered list (1, 2, 3...)
            - "UNORDERED" or "bullet": Convert to bullet list
        list_index: Which list to convert (0-based, default 0 for first list).
            Used when not specifying search or index range.
        search: Text to search for within a list item. If provided, converts the
            list containing that text.
        start_index: Explicit start position of the range to convert (optional).
        end_index: Explicit end position of the range to convert (optional).
        preview: If True, show what would be converted without making changes.

    Returns:
        JSON string with conversion result or preview information.

    Examples:
        # Convert the first list to numbered
        convert_list_type(document_id="abc123", list_type="numbered")

        # Convert the second list to bullets
        convert_list_type(document_id="abc123", list_type="bullet", list_index=1)

        # Convert list containing specific text to numbered
        convert_list_type(document_id="abc123", list_type="ORDERED",
                         search="step one")

        # Preview what would be converted
        convert_list_type(document_id="abc123", list_type="bullet", preview=True)

    Notes:
        - This tool finds existing lists and re-applies bullet formatting with
          the new type
        - All items in the affected list will be converted
        - Use find_doc_elements with element_type='list' to see all lists first
    """
    import json

    logger.debug(
        f"[convert_list_type] Doc={document_id}, list_type={list_type}, "
        f"list_index={list_index}, search={search}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Normalize list_type parameter
    list_type_aliases = {
        "bullet": "UNORDERED",
        "bullets": "UNORDERED",
        "unordered": "UNORDERED",
        "numbered": "ORDERED",
        "numbers": "ORDERED",
        "ordered": "ORDERED",
    }
    normalized_type = list_type_aliases.get(list_type.lower(), list_type.upper())

    if normalized_type not in ["ORDERED", "UNORDERED"]:
        error = DocsErrorBuilder.invalid_param_value(
            param_name="list_type",
            received=list_type,
            valid_values=["ORDERED", "UNORDERED", "numbered", "bullet"],
        )
        return format_error(error)

    doc_link = f"https://docs.google.com/document/d/{document_id}/edit"

    # Get document content
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Find all lists in the document
    all_lists = find_elements_by_type(doc_data, "list")

    if not all_lists:
        return json.dumps({
            "error": True,
            "code": "NO_LISTS_FOUND",
            "message": "No lists found in the document",
            "suggestion": "Use insert_doc_elements or modify_doc_text with convert_to_list to create a list first",
            "link": doc_link,
        }, indent=2)

    # Find the target list
    target_list = None

    if search:
        # Find list containing the search text
        search_lower = search.lower()
        for lst in all_lists:
            for item in lst.get("items", []):
                if search_lower in item.get("text", "").lower():
                    target_list = lst
                    break
            if target_list:
                break

        if not target_list:
            return json.dumps({
                "error": True,
                "code": "LIST_NOT_FOUND",
                "message": f"No list containing text '{search}' found",
                "suggestion": "Check the search text or use find_doc_elements "
                              "with element_type='list' to see all lists",
                "available_lists": len(all_lists),
                "link": doc_link,
            }, indent=2)

    elif start_index is not None and end_index is not None:
        # Find list overlapping with the given range
        for lst in all_lists:
            if lst["start_index"] < end_index and lst["end_index"] > start_index:
                target_list = lst
                break

        if not target_list:
            return json.dumps({
                "error": True,
                "code": "LIST_NOT_FOUND",
                "message": f"No list found in range {start_index}-{end_index}",
                "suggestion": "Use find_doc_elements with element_type='list' to see list positions",
                "available_lists": len(all_lists),
                "link": doc_link,
            }, indent=2)

    else:
        # Use list_index
        if list_index < 0 or list_index >= len(all_lists):
            return json.dumps({
                "error": True,
                "code": "INVALID_LIST_INDEX",
                "message": f"List index {list_index} is out of range. Document has {len(all_lists)} list(s).",
                "suggestion": f"Use list_index between 0 and {len(all_lists) - 1}",
                "available_lists": len(all_lists),
                "link": doc_link,
            }, indent=2)
        target_list = all_lists[list_index]

    # Determine current list type for display
    current_type = target_list.get("list_type", target_list.get("type", "unknown"))
    if "bullet" in current_type or "unordered" in current_type.lower():
        current_type_display = "bullet"
        current_type_normalized = "UNORDERED"
    else:
        current_type_display = "numbered"
        current_type_normalized = "ORDERED"

    target_type_display = "numbered" if normalized_type == "ORDERED" else "bullet"

    # Check if already the target type
    if current_type_normalized == normalized_type:
        return json.dumps({
            "success": True,
            "message": f"List is already a {target_type_display} list",
            "no_change_needed": True,
            "list_range": {
                "start_index": target_list["start_index"],
                "end_index": target_list["end_index"],
            },
            "items_count": len(target_list.get("items", [])),
            "link": doc_link,
        }, indent=2)

    # Build preview info
    preview_info = {
        "current_type": current_type_display,
        "target_type": target_type_display,
        "list_range": {
            "start_index": target_list["start_index"],
            "end_index": target_list["end_index"],
        },
        "items_count": len(target_list.get("items", [])),
        "items": [
            {
                "text": item.get("text", "")[:50] + ("..." if len(item.get("text", "")) > 50 else ""),
                "start_index": item["start_index"],
                "end_index": item["end_index"],
            }
            for item in target_list.get("items", [])[:5]  # Show up to 5 items in preview
        ],
    }

    if preview:
        return json.dumps({
            "preview": True,
            "message": f"Would convert {current_type_display} list to {target_type_display} list",
            **preview_info,
            "link": doc_link,
        }, indent=2)

    # Create the conversion request using createParagraphBullets
    request = create_bullet_list_request(
        target_list["start_index"],
        target_list["end_index"],
        normalized_type
    )

    # Execute the request
    await asyncio.to_thread(
        service.documents()
        .batchUpdate(documentId=document_id, body={"requests": [request]})
        .execute
    )

    return json.dumps({
        "success": True,
        "message": f"Converted {current_type_display} list to {target_type_display} list",
        "converted_from": current_type_display,
        "converted_to": target_type_display,
        "list_range": {
            "start_index": target_list["start_index"],
            "end_index": target_list["end_index"],
        },
        "items_count": len(target_list.get("items", [])),
        "link": doc_link,
    }, indent=2)


@server.tool()
@handle_http_errors("append_to_list", service_type="docs")
@require_google_service("docs", "docs_write")
async def append_to_list(
    service: Any,
    user_google_email: str,
    document_id: str,
    items: list,
    list_index: int = 0,
    search: str = None,
    nesting_levels: list = None,
    preview: bool = False,
) -> str:
    """
    Append items to an existing list in a Google Doc.

    This tool adds new items to the end of an existing list without creating a new list.
    The appended items will automatically continue the list's formatting (bullet or numbered).

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document containing the list
        items: List of strings to append as new list items
        list_index: Which list to append to (0-based, default 0 for first list).
            Used when not specifying search.
        search: Text to search for within a list item. If provided, appends to the
            list containing that text.
        nesting_levels: Optional list of integers specifying nesting level for each item (0-8).
            Must match length of 'items'. Default is 0 (top level) for all items.
            Use higher numbers for sub-items (1 = first indent, 2 = second indent, etc.)
        preview: If True, show what would be appended without making changes.

    Returns:
        JSON string with operation result or preview information.

    Examples:
        # Append single item to first list
        append_to_list(document_id="abc123", items=["New item"])

        # Append multiple items to the second list
        append_to_list(document_id="abc123", items=["Item A", "Item B"], list_index=1)

        # Append to list containing specific text
        append_to_list(document_id="abc123", items=["Follow-up task"],
                      search="existing task")

        # Append nested items
        append_to_list(document_id="abc123",
                      items=["Main point", "Sub point 1", "Sub point 2", "Another main"],
                      nesting_levels=[0, 1, 1, 0])

        # Preview what would be appended
        append_to_list(document_id="abc123", items=["Test item"], preview=True)

    Notes:
        - The appended items will inherit the list type (bullet or numbered) of the target list
        - Use find_doc_elements with element_type='list' to see all lists first
        - For creating new lists, use insert_doc_elements with element_type='list'
    """
    import json

    logger.debug(
        f"[append_to_list] Doc={document_id}, items={items}, "
        f"list_index={list_index}, search={search}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Validate items parameter
    if not isinstance(items, list):
        return validator.create_invalid_param_error(
            param_name="items",
            received=type(items).__name__,
            valid_values=["list of strings"],
        )
    if len(items) == 0:
        return validator.create_invalid_param_error(
            param_name="items",
            received="empty list",
            valid_values=["non-empty list of strings"],
        )
    # Validate all items are strings
    for i, item in enumerate(items):
        if not isinstance(item, str):
            return validator.create_invalid_param_error(
                param_name=f"items[{i}]",
                received=type(item).__name__,
                valid_values=["string"],
            )

    # Process items with escape sequences
    processed_items = [interpret_escape_sequences(item) for item in items]

    # Handle nesting_levels parameter
    item_nesting = []
    if nesting_levels is not None:
        if not isinstance(nesting_levels, list):
            return validator.create_invalid_param_error(
                param_name="nesting_levels",
                received=type(nesting_levels).__name__,
                valid_values=["list of integers"],
            )
        if len(nesting_levels) != len(items):
            return validator.create_invalid_param_error(
                param_name="nesting_levels",
                received=f"list of length {len(nesting_levels)}",
                valid_values=[f"list of length {len(items)} (must match items length)"],
            )
        # Validate all nesting levels are valid integers
        for i, level in enumerate(nesting_levels):
            if not isinstance(level, int):
                return validator.create_invalid_param_error(
                    param_name=f"nesting_levels[{i}]",
                    received=type(level).__name__,
                    valid_values=["integer (0-8)"],
                )
            if level < 0 or level > 8:
                return validator.create_invalid_param_error(
                    param_name=f"nesting_levels[{i}]",
                    received=str(level),
                    valid_values=["integer from 0 to 8"],
                )
        item_nesting = nesting_levels
    else:
        # Default to level 0 for all items
        item_nesting = [0] * len(items)

    doc_link = f"https://docs.google.com/document/d/{document_id}/edit"

    # Get document content
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Find all lists in the document
    all_lists = find_elements_by_type(doc_data, "list")

    if not all_lists:
        return json.dumps({
            "error": True,
            "code": "NO_LISTS_FOUND",
            "message": "No lists found in the document",
            "suggestion": "Use insert_doc_elements with element_type='list' to create a list first",
            "link": doc_link,
        }, indent=2)

    # Find the target list
    target_list = None

    if search:
        # Find list containing the search text
        search_lower = search.lower()
        for lst in all_lists:
            for item in lst.get("items", []):
                if search_lower in item.get("text", "").lower():
                    target_list = lst
                    break
            if target_list:
                break

        if not target_list:
            return json.dumps({
                "error": True,
                "code": "LIST_NOT_FOUND",
                "message": f"No list containing text '{search}' found",
                "suggestion": "Check the search text or use find_doc_elements "
                              "with element_type='list' to see all lists",
                "available_lists": len(all_lists),
                "link": doc_link,
            }, indent=2)

    else:
        # Use list_index
        if list_index < 0 or list_index >= len(all_lists):
            return json.dumps({
                "error": True,
                "code": "INVALID_LIST_INDEX",
                "message": f"List index {list_index} is out of range. Document has {len(all_lists)} list(s).",
                "suggestion": f"Use list_index between 0 and {len(all_lists) - 1}",
                "available_lists": len(all_lists),
                "link": doc_link,
            }, indent=2)
        target_list = all_lists[list_index]

    # Determine list type for applying bullet formatting
    list_type_raw = target_list.get("list_type", target_list.get("type", "bullet"))
    if "bullet" in list_type_raw or "unordered" in list_type_raw.lower():
        list_type = "UNORDERED"
        list_type_display = "bullet"
    else:
        list_type = "ORDERED"
        list_type_display = "numbered"

    # Build the text to insert with tab prefixes for nesting levels
    nested_items = []
    for item_text, level in zip(processed_items, item_nesting):
        prefix = "\t" * level
        nested_items.append(prefix + item_text)

    combined_text = "\n".join(nested_items) + "\n"
    text_length = len(combined_text) - 1  # Exclude final newline from bullet range

    # Calculate insertion point: at the end of the last list item
    # The end_index points to after the last character of the list
    # We need to insert at the end of the list content
    insertion_index = target_list["end_index"]

    # Build preview info
    has_nested = any(level > 0 for level in item_nesting)
    max_level = max(item_nesting) if has_nested else 0
    preview_info = {
        "target_list": {
            "type": list_type_display,
            "start_index": target_list["start_index"],
            "end_index": target_list["end_index"],
            "current_items_count": len(target_list.get("items", [])),
        },
        "items_to_append": len(processed_items),
        "insertion_index": insertion_index,
        "nesting": {
            "has_nested_items": has_nested,
            "max_depth": max_level,
        } if has_nested else None,
        "items_preview": [
            {
                "text": item[:50] + ("..." if len(item) > 50 else ""),
                "nesting_level": level,
            }
            for item, level in zip(processed_items, item_nesting)
        ][:5],  # Show up to 5 items in preview
    }

    if preview:
        return json.dumps({
            "preview": True,
            "message": f"Would append {len(processed_items)} item(s) to {list_type_display} list",
            **preview_info,
            "link": doc_link,
        }, indent=2)

    # Create the requests:
    # 1. Insert the text at the end of the list
    # 2. Apply bullet formatting to the new text
    requests = [
        create_insert_text_request(insertion_index, combined_text),
        create_bullet_list_request(
            insertion_index,
            insertion_index + text_length,
            list_type
        ),
    ]

    # Execute the requests
    await asyncio.to_thread(
        service.documents()
        .batchUpdate(documentId=document_id, body={"requests": requests})
        .execute
    )

    return json.dumps({
        "success": True,
        "message": f"Appended {len(processed_items)} item(s) to {list_type_display} list",
        "list_type": list_type_display,
        "items_appended": len(processed_items),
        "new_list_range": {
            "start_index": target_list["start_index"],
            "end_index": insertion_index + len(combined_text),
        },
        "link": doc_link,
    }, indent=2)


@server.tool()
@handle_http_errors("copy_doc_section", service_type="docs")
@require_google_service("docs", "docs_write")
async def copy_doc_section(
    service: Any,
    user_google_email: str,
    document_id: str,
    # Source specification (one of these methods)
    heading: str = None,
    start_index: int = None,
    end_index: int = None,
    search: str = None,
    search_end: str = None,
    # Destination specification
    destination_index: int = None,
    destination_location: str = None,
    destination_after_heading: str = None,
    # Options
    include_heading: bool = True,
    match_case: bool = False,
    preserve_formatting: bool = True,
    preview: bool = False,
) -> str:
    """
    Copy a section or text range from one location to another in a Google Doc.

    This tool allows you to duplicate content within a document, preserving
    text and optionally its formatting. Useful for template-based editing,
    duplicating sections, or reorganizing document content.

    SOURCE SPECIFICATION (use ONE of these methods):
    1. By heading: Use 'heading' to copy an entire section (heading + content until next same-level heading)
    2. By indices: Use 'start_index' and 'end_index' for precise range selection
    3. By search: Use 'search' and optionally 'search_end' to find text boundaries

    DESTINATION SPECIFICATION (use ONE of these methods):
    1. By index: Use 'destination_index' to specify exact insertion point
    2. By location: Use 'destination_location' with values 'start' or 'end'
    3. After heading: Use 'destination_after_heading' to insert after a section heading

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to modify
        heading: Copy a section by its heading text (includes all content until next same/higher level heading)
        start_index: Start position for explicit range copy
        end_index: End position for explicit range copy
        search: Search text to find start of range to copy
        search_end: Search text to find end of range (if omitted with search, copies just the found text)
        destination_index: Exact index where content should be inserted
        destination_location: 'start' or 'end' of document
        destination_after_heading: Insert after this heading's content
        include_heading: When copying by heading, whether to include the heading itself (default True)
        match_case: Whether to match case exactly when searching for headings/text
        preserve_formatting: Whether to preserve text formatting in the copy (default True)
        preview: If True, shows what would be copied without modifying the document

    Returns:
        str: JSON containing:
            - success: True if copy was successful
            - source: Information about what was copied (indices, text preview)
            - destination: Where content was inserted
            - characters_copied: Number of characters copied
            - formatting_spans: Number of formatting spans preserved (if preserve_formatting=True)
            - preview: True if this was a preview operation
            - link: Link to the document

    Example - Copy a section to the end:
        copy_doc_section(
            document_id="...",
            heading="Template Section",
            destination_location="end"
        )

    Example - Copy a text range by search:
        copy_doc_section(
            document_id="...",
            search="[START]",
            search_end="[END]",
            destination_after_heading="New Section"
        )

    Example - Copy by indices:
        copy_doc_section(
            document_id="...",
            start_index=100,
            end_index=500,
            destination_index=1000
        )
    """
    import json

    logger.debug(
        f"[copy_doc_section] Doc={document_id}, heading={heading}, "
        f"start_index={start_index}, end_index={end_index}, search={search}"
    )

    # Input validation
    validator = ValidationManager()
    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Validate source specification - exactly one method should be used
    source_methods = [
        heading is not None,
        (start_index is not None and end_index is not None),
        search is not None,
    ]
    if sum(source_methods) == 0:
        return validator.create_missing_param_error(
            param_name="source specification",
            context="for copy operation",
            valid_values=["heading", "start_index + end_index", "search"],
        )
    if sum(source_methods) > 1:
        return json.dumps({
            "error": "INVALID_PARAMETERS",
            "message": "Multiple source methods specified. Use only one of: heading, start_index+end_index, or search",
            "hint": "Choose one method to specify the source content to copy"
        }, indent=2)

    # Validate destination specification - exactly one method should be used
    dest_methods = [
        destination_index is not None,
        destination_location is not None,
        destination_after_heading is not None,
    ]
    if sum(dest_methods) == 0:
        return validator.create_missing_param_error(
            param_name="destination specification",
            context="for copy operation",
            valid_values=["destination_index", "destination_location", "destination_after_heading"],
        )
    if sum(dest_methods) > 1:
        return json.dumps({
            "error": "INVALID_PARAMETERS",
            "message": "Multiple destination methods specified. Use only one of: destination_index, destination_location, or destination_after_heading",
            "hint": "Choose one method to specify where the content should be copied to"
        }, indent=2)

    # Validate destination_location if provided
    if destination_location and destination_location not in ("start", "end"):
        return json.dumps({
            "error": "INVALID_PARAMETERS",
            "message": f"Invalid destination_location: '{destination_location}'",
            "valid_values": ["start", "end"],
            "hint": "Use 'start' to insert at beginning or 'end' to append at document end"
        }, indent=2)

    # Get the document
    doc = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Determine source range
    copy_start = None
    copy_end = None
    source_info = {}

    if heading:
        # Copy by heading/section
        section = find_section_by_heading(doc, heading, match_case)
        if section is None:
            all_headings = get_all_headings(doc)
            heading_list = [h["text"] for h in all_headings] if all_headings else []
            return validator.create_heading_not_found_error(
                heading=heading, available_headings=heading_list, match_case=match_case
            )

        if include_heading:
            copy_start = section["start_index"]
        else:
            # Find where the heading ends
            elements = extract_structural_elements(doc)
            for elem in elements:
                if elem["type"].startswith("heading") or elem["type"] == "title":
                    elem_text = elem["text"] if match_case else elem["text"].lower()
                    search_text = heading if match_case else heading.lower()
                    if elem_text.strip() == search_text.strip():
                        copy_start = elem["end_index"]
                        break
            if copy_start is None:
                copy_start = section["start_index"]

        copy_end = section["end_index"]
        source_info = {
            "method": "heading",
            "heading": section["heading"],
            "level": section["level"],
            "include_heading": include_heading,
            "subsections_count": len(section.get("subsections", [])),
        }

    elif start_index is not None and end_index is not None:
        # Copy by explicit indices
        if start_index < 1:
            return json.dumps({
                "error": "INVALID_INDEX",
                "message": f"start_index must be >= 1, got {start_index}",
                "hint": "Document indices start at 1"
            }, indent=2)
        if end_index <= start_index:
            return json.dumps({
                "error": "INVALID_RANGE",
                "message": f"end_index ({end_index}) must be greater than start_index ({start_index})",
                "hint": "Specify a valid range with end_index > start_index"
            }, indent=2)

        copy_start = start_index
        copy_end = end_index
        source_info = {
            "method": "indices",
            "start_index": start_index,
            "end_index": end_index,
        }

    elif search:
        # Copy by search text
        # find_all_occurrences_in_document returns List[Tuple[int, int]] - (start, end) pairs
        positions = find_all_occurrences_in_document(doc, search, match_case)
        if not positions:
            return json.dumps({
                "error": "TEXT_NOT_FOUND",
                "message": f"Search text not found: '{search}'",
                "match_case": match_case,
                "hint": "Check your search text or try with match_case=False"
            }, indent=2)

        copy_start = positions[0][0]  # First tuple's start index

        if search_end:
            # Find end marker
            end_positions = find_all_occurrences_in_document(doc, search_end, match_case)
            if not end_positions:
                return json.dumps({
                    "error": "TEXT_NOT_FOUND",
                    "message": f"End search text not found: '{search_end}'",
                    "match_case": match_case,
                    "hint": "Check your search_end text or try with match_case=False"
                }, indent=2)
            # Find the first occurrence of search_end that comes after our start
            copy_end = None
            for start, end in end_positions:
                if end > copy_start:
                    copy_end = end
                    break
            if copy_end is None:
                return json.dumps({
                    "error": "INVALID_RANGE",
                    "message": f"End marker '{search_end}' not found after start marker '{search}'",
                    "hint": "Ensure search_end appears after search in the document"
                }, indent=2)
        else:
            copy_end = positions[0][1]  # First tuple's end index

        source_info = {
            "method": "search",
            "search": search,
            "search_end": search_end,
        }

    # Check if section extends to document end and adjust
    body = doc.get("body", {})
    body_content = body.get("content", [])
    doc_end_index = body_content[-1].get("endIndex", 0) if body_content else 1

    # Adjust copy_end to not include the final document terminator if at the end
    if copy_end >= doc_end_index:
        copy_end = doc_end_index - 1

    # Extract the text to copy
    text_to_copy = extract_text_in_range(doc, copy_start, copy_end)

    if not text_to_copy:
        return json.dumps({
            "error": "EMPTY_CONTENT",
            "message": "No content to copy in the specified range",
            "source_range": {"start": copy_start, "end": copy_end},
            "hint": "Check your source specification - the range may be empty"
        }, indent=2)

    # Extract formatting spans if needed
    formatting_spans = []
    if preserve_formatting:
        formatting_spans = _extract_text_formatting_from_range(doc, copy_start, copy_end)

    # Determine destination index
    dest_index = None

    if destination_index is not None:
        dest_index = destination_index
    elif destination_location == "start":
        dest_index = 1
    elif destination_location == "end":
        dest_index = doc_end_index - 1
    elif destination_after_heading:
        dest_section = find_section_by_heading(doc, destination_after_heading, match_case)
        if dest_section is None:
            all_headings = get_all_headings(doc)
            heading_list = [h["text"] for h in all_headings] if all_headings else []
            return validator.create_heading_not_found_error(
                heading=destination_after_heading,
                available_headings=heading_list,
                match_case=match_case
            )
        # Insert at the end of the destination section
        dest_index = dest_section["end_index"]
        # Adjust if at document end
        if dest_index >= doc_end_index:
            dest_index = doc_end_index - 1

    # Build result info
    doc_link = f"https://docs.google.com/document/d/{document_id}/edit"
    text_preview = text_to_copy[:200] + ("..." if len(text_to_copy) > 200 else "")

    result = {
        "source": {
            **source_info,
            "start_index": copy_start,
            "end_index": copy_end,
            "characters": len(text_to_copy),
            "text_preview": text_preview,
        },
        "destination": {
            "index": dest_index,
            "method": (
                "index" if destination_index is not None
                else "location" if destination_location is not None
                else "after_heading"
            ),
        },
        "characters_copied": len(text_to_copy),
        "formatting_spans": len(formatting_spans) if preserve_formatting else 0,
        "preserve_formatting": preserve_formatting,
        "link": doc_link,
    }

    if destination_location:
        result["destination"]["location"] = destination_location
    if destination_after_heading:
        result["destination"]["after_heading"] = destination_after_heading

    # Handle preview mode
    if preview:
        result["preview"] = True
        result["would_copy"] = True
        result["success"] = False
        return (
            f"Preview - Would copy {len(text_to_copy)} characters:\n\n"
            f"{json.dumps(result, indent=2)}"
        )

    # Build the requests for the copy operation
    requests = []

    # 1. Insert the text at the destination
    requests.append(create_insert_text_request(dest_index, text_to_copy))

    # 2. Clear formatting first if we're going to apply specific formatting
    # This prevents style inheritance issues
    if preserve_formatting and formatting_spans:
        text_length = len(text_to_copy)
        requests.append(create_clear_formatting_request(dest_index, dest_index + text_length))

        # 3. Re-apply the original formatting at the new location
        for span in formatting_spans:
            # Calculate the new position for this span
            # Original span offset from copy_start
            span_offset = span["start_index"] - copy_start
            span_length = span["end_index"] - span["start_index"]

            new_start = dest_index + span_offset
            new_end = new_start + span_length

            # Build format request from span info
            format_req = create_format_text_request(
                start_index=new_start,
                end_index=new_end,
                bold=span.get("bold") if span.get("bold") else None,
                italic=span.get("italic") if span.get("italic") else None,
                underline=span.get("underline") if span.get("underline") else None,
                strikethrough=span.get("strikethrough") if span.get("strikethrough") else None,
                small_caps=span.get("small_caps") if span.get("small_caps") else None,
                font_size=span.get("font_size"),
                font_family=span.get("font_family"),
                foreground_color=span.get("foreground_color"),
                background_color=span.get("background_color"),
                link=span.get("link_url"),
                subscript=span.get("baseline_offset") == "SUBSCRIPT",
                superscript=span.get("baseline_offset") == "SUPERSCRIPT",
            )
            if format_req:
                requests.append(format_req)

    # Execute the batch update
    await asyncio.to_thread(
        service.documents()
        .batchUpdate(documentId=document_id, body={"requests": requests})
        .execute
    )

    result["success"] = True
    result["preview"] = False

    return (
        f"Copied {len(text_to_copy)} characters to destination:\n\n"
        f"{json.dumps(result, indent=2)}\n\nLink: {doc_link}"
    )
