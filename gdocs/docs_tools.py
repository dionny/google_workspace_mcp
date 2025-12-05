"""
Google Docs MCP Tools

This module provides MCP tools for interacting with Google Docs API and managing Google Docs via Drive.
"""
import logging
import asyncio
import io
from typing import List, Dict, Any

from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

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
    create_find_replace_request,
    create_insert_table_request,
    create_insert_page_break_request,
    create_insert_image_request,
    create_bullet_list_request,
    calculate_search_based_indices,
    find_all_occurrences_in_document,
    SearchPosition,
    OperationType,
    build_operation_result,
    resolve_range,
    RangeResult,
    extract_text_at_range,
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
    get_heading_siblings
)
from gdocs.docs_tables import (
    extract_table_as_data
)

# Import operation managers for complex business logic
from gdocs.managers import (
    TableOperationManager,
    HeaderFooterManager,
    ValidationManager,
    BatchOperationManager
)
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
        service.files().list(
            q=f"name contains '{escaped_query}' and mimeType='application/vnd.google-apps.document' and trashed=false",
            pageSize=page_size,
            fields="files(id, name, createdTime, modifiedTime, webViewLink)"
        ).execute
    )
    files = response.get('files', [])
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
@require_multiple_services([
    {"service_type": "drive", "scopes": "drive_read", "param_name": "drive_service"},
    {"service_type": "docs", "scopes": "docs_read", "param_name": "docs_service"}
])
async def get_doc_content(
    drive_service: Any,
    docs_service: Any,
    user_google_email: str,
    document_id: str,
) -> str:
    """
    Retrieves content of a Google Doc or a Drive file (like .docx) identified by document_id.
    - Native Google Docs: Fetches content via Docs API.
    - Office files (.docx, etc.) stored in Drive: Downloads via Drive API and extracts text.

    Returns:
        str: The document content with metadata header.
    """
    logger.info(f"[get_doc_content] Invoked. Document/File ID: '{document_id}' for user '{user_google_email}'")

    # Step 2: Get file metadata from Drive
    file_metadata = await asyncio.to_thread(
        drive_service.files().get(
            fileId=document_id, fields="id, name, mimeType, webViewLink"
        ).execute
    )
    mime_type = file_metadata.get("mimeType", "")
    file_name = file_metadata.get("name", "Unknown File")
    web_view_link = file_metadata.get("webViewLink", "#")

    logger.info(f"[get_doc_content] File '{file_name}' (ID: {document_id}) has mimeType: '{mime_type}'")

    body_text = "" # Initialize body_text

    # Step 3: Process based on mimeType
    if mime_type == "application/vnd.google-apps.document":
        logger.info("[get_doc_content] Processing as native Google Doc.")
        doc_data = await asyncio.to_thread(
            docs_service.documents().get(
                documentId=document_id,
                includeTabsContent=True
            ).execute
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
                if 'paragraph' in element:
                    paragraph = element.get('paragraph', {})
                    para_elements = paragraph.get('elements', [])
                    current_line_text = ""
                    for pe in para_elements:
                        text_run = pe.get('textRun', {})
                        if text_run and 'content' in text_run:
                            current_line_text += text_run['content']
                    if current_line_text.strip():
                        text_lines.append(current_line_text)
                elif 'table' in element:
                    # Handle table content
                    table = element.get('table', {})
                    table_rows = table.get('tableRows', [])
                    for row in table_rows:
                        row_cells = row.get('tableCells', [])
                        for cell in row_cells:
                            cell_content = cell.get('content', [])
                            cell_text = extract_text_from_elements(cell_content, depth=depth + 1)
                            if cell_text.strip():
                                text_lines.append(cell_text)
            return "".join(text_lines)

        def process_tab_hierarchy(tab, level=0):
            """Process a tab and its nested child tabs recursively"""
            tab_text = ""

            if 'documentTab' in tab:
                tab_title = tab.get('documentTab', {}).get('title', 'Untitled Tab')
                # Add indentation for nested tabs to show hierarchy
                if level > 0:
                    tab_title = "    " * level + tab_title
                tab_body = tab.get('documentTab', {}).get('body', {}).get('content', [])
                tab_text += extract_text_from_elements(tab_body, tab_title)

            # Process child tabs (nested tabs)
            child_tabs = tab.get('childTabs', [])
            for child_tab in child_tabs:
                tab_text += process_tab_hierarchy(child_tab, level + 1)

            return tab_text

        processed_text_lines = []

        # Process main document body
        body_elements = doc_data.get('body', {}).get('content', [])
        main_content = extract_text_from_elements(body_elements)
        if main_content.strip():
            processed_text_lines.append(main_content)

        # Process all tabs
        tabs = doc_data.get('tabs', [])
        for tab in tabs:
            tab_content = process_tab_hierarchy(tab)
            if tab_content.strip():
                processed_text_lines.append(tab_content)

        body_text = "".join(processed_text_lines)
    else:
        logger.info(f"[get_doc_content] Processing as Drive file (e.g., .docx, other). MimeType: {mime_type}")

        export_mime_type_map = {
                # Example: "application/vnd.google-apps.spreadsheet"z: "text/csv",
                # Native GSuite types that are not Docs would go here if this function
                # was intended to export them. For .docx, direct download is used.
        }
        effective_export_mime = export_mime_type_map.get(mime_type)

        request_obj = (
            drive_service.files().export_media(fileId=document_id, mimeType=effective_export_mime)
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
        f'Link: {web_view_link}\n\n--- CONTENT ---\n'
    )
    return header + body_text

@server.tool()
@handle_http_errors("list_docs_in_folder", is_read_only=True, service_type="docs")
@require_google_service("drive", "drive_read")
async def list_docs_in_folder(
    service: Any,
    user_google_email: str,
    folder_id: str = 'root',
    page_size: int = 100
) -> str:
    """
    Lists Google Docs within a specific Drive folder.

    Returns:
        str: A formatted list of Google Docs in the specified folder.
    """
    logger.info(f"[list_docs_in_folder] Invoked. Email: '{user_google_email}', Folder ID: '{folder_id}'")

    rsp = await asyncio.to_thread(
        service.files().list(
            q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.document' and trashed=false",
            pageSize=page_size,
            fields="files(id, name, modifiedTime, webViewLink)"
        ).execute
    )
    items = rsp.get('files', [])
    if not items:
        return f"No Google Docs found in folder '{folder_id}'."
    out = [f"Found {len(items)} Docs in folder '{folder_id}':"]
    for f in items:
        out.append(f"- {f['name']} (ID: {f['id']}) Modified: {f.get('modifiedTime')} Link: {f.get('webViewLink')}")
    return "\n".join(out)

@server.tool()
@handle_http_errors("create_doc", service_type="docs")
@require_google_service("docs", "docs_write")
async def create_doc(
    service: Any,
    user_google_email: str,
    title: str,
    content: str = '',
) -> str:
    """
    Creates a new Google Doc and optionally inserts initial content.

    Returns:
        str: Confirmation message with document ID and link.
    """
    logger.info(f"[create_doc] Invoked. Email: '{user_google_email}', Title='{title}'")

    doc = await asyncio.to_thread(service.documents().create(body={'title': title}).execute)
    doc_id = doc.get('documentId')
    if content:
        requests = [{'insertText': {'location': {'index': 1}, 'text': content}}]
        await asyncio.to_thread(service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute)
    link = f"https://docs.google.com/document/d/{doc_id}/edit"
    msg = f"Created Google Doc '{title}' (ID: {doc_id}) for {user_google_email}. Link: {link}"
    logger.info(f"Successfully created Google Doc '{title}' (ID: {doc_id}) for {user_google_email}. Link: {link}")
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
    font_size: int = None,
    font_family: str = None,
    link: str = None,
    foreground_color: str = None,
    background_color: str = None,
    search: str = None,
    position: str = None,
    occurrence: int = 1,
    match_case: bool = True,
    heading: str = None,
    section_position: str = None,
    range: Dict[str, Any] = None,
    location: str = None,
    preview: bool = False,
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
        font_size: Font size in points
        font_family: Font family name (e.g., "Arial", "Times New Roman")
        link: URL to create a hyperlink. Use empty string "" to remove an existing link.
            Supports http://, https://, or # (for internal document bookmarks).
        foreground_color: Text color as hex (#FF0000, #F00) or named color (red, blue, green, etc.)
        background_color: Background/highlight color as hex or named color

        Preview mode:
        preview: If True, returns what would change without actually modifying the document.
            Useful for validating operations before applying them. Default: False

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
        - legacy_message (str): Backward-compatible text summary

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
    """
    logger.info(f"[modify_doc_text] Doc={document_id}, location={location}, search={search}, position={position}, "
                f"heading={heading}, section_position={section_position}, range={range is not None}, "
                f"start={start_index}, end={end_index}, text={text is not None}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Validate that we have something to do
    if text is None and not any([bold is not None, italic is not None, underline is not None, font_size, font_family, link is not None]):
        error = DocsErrorBuilder.missing_required_param(
            param_name="text or formatting",
            context_description="for document modification",
            valid_values=["text", "bold", "italic", "underline", "font_size", "font_family", "link"]
        )
        return format_error(error)

    # Determine positioning mode
    use_range_mode = range is not None
    use_search_mode = search is not None
    use_heading_mode = heading is not None
    use_location_mode = location is not None

    # Validate positioning parameters - range mode takes priority, then location, then heading, then search, then index
    if use_range_mode:
        if location is not None or heading is not None or search is not None or start_index is not None or end_index is not None:
            logger.warning("Multiple positioning parameters provided; range mode takes precedence")
    elif use_location_mode:
        if location not in ['start', 'end']:
            return validator.create_invalid_param_error(
                param_name="location",
                received=location,
                valid_values=["start", "end"]
            )
        if start_index is not None or end_index is not None:
            error = DocsErrorBuilder.conflicting_params(
                params=["location", "start_index", "end_index"],
                message="Cannot use 'location' parameter with explicit 'start_index' or 'end_index'"
            )
            return format_error(error)
        if heading is not None or search is not None:
            logger.warning("Multiple positioning parameters provided; location mode takes precedence")
    elif use_heading_mode:
        if not section_position:
            return validator.create_missing_param_error(
                param_name="section_position",
                context="when using 'heading'",
                valid_values=["start", "end"]
            )
        if section_position not in ['start', 'end']:
            return validator.create_invalid_param_error(
                param_name="section_position",
                received=section_position,
                valid_values=["start", "end"]
            )
        if search is not None or start_index is not None or end_index is not None:
            logger.warning("Multiple positioning parameters provided; heading mode takes precedence")
    elif use_search_mode:
        if not position:
            return validator.create_missing_param_error(
                param_name="position",
                context="when using 'search'",
                valid_values=["before", "after", "replace"]
            )
        if position not in [p.value for p in SearchPosition]:
            return validator.create_invalid_param_error(
                param_name="position",
                received=position,
                valid_values=["before", "after", "replace"]
            )
        if start_index is not None or end_index is not None:
            logger.warning("Both search and index parameters provided; search mode takes precedence")
    else:
        # Traditional index-based mode
        if start_index is None:
            return validator.create_missing_param_error(
                param_name="positioning",
                context="for document modification",
                valid_values=["location", "range", "heading+section_position", "search+position", "start_index"]
            )

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
        total_length = structure['total_length']

        if location == 'end':
            # Use total_length - 1 for safe insertion (last valid index)
            resolved_end_index = total_length - 1 if total_length > 1 else 1
            start_index = resolved_end_index
            location_info = {
                'location': 'end',
                'resolved_index': resolved_end_index,
                'message': f"Appending at document end (index {resolved_end_index})"
            }
        else:  # location == 'start'
            start_index = 1  # After the initial section break at index 0
            location_info = {
                'location': 'start',
                'resolved_index': 1,
                'message': "Inserting at document start (index 1)"
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
                "hint": "Check range specification format and search terms"
            }
            return json.dumps(error_response, indent=2)

        start_index = range_result.start_index
        end_index = range_result.end_index
        range_result_info = range_result.to_dict()
        search_info = {
            'range': range,
            'resolved_start': start_index,
            'resolved_end': end_index,
            'message': range_result.message
        }
        if range_result.matched_start:
            search_info['matched_start'] = range_result.matched_start
        if range_result.matched_end:
            search_info['matched_end'] = range_result.matched_end
        if range_result.extend_type:
            search_info['extend_type'] = range_result.extend_type
        if range_result.section_name:
            search_info['section_name'] = range_result.section_name

    # If using heading mode, find the section and calculate insertion point
    elif use_heading_mode:
        # Get document
        doc_data = await asyncio.to_thread(
            service.documents().get(documentId=document_id).execute
        )

        # Find the insertion point using section navigation
        insertion_index = find_section_insertion_point(doc_data, heading, section_position, match_case)

        if insertion_index is None:
            # Provide helpful error with available headings
            all_headings = get_all_headings(doc_data)
            heading_list = [h['text'] for h in all_headings] if all_headings else []
            return validator.create_heading_not_found_error(
                heading=heading,
                available_headings=heading_list,
                match_case=match_case
            )

        start_index = insertion_index
        end_index = None  # Heading mode is always insert, not replace
        search_info = {
            'heading': heading,
            'section_position': section_position,
            'insertion_index': insertion_index,
            'message': f"Found section '{heading}', inserting at {section_position}"
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
            all_occurrences = find_all_occurrences_in_document(doc_data, search, match_case)
            if all_occurrences:
                # Check if it's an invalid occurrence error
                if "occurrence" in message.lower():
                    return validator.create_invalid_occurrence_error(
                        occurrence=occurrence,
                        total_found=len(all_occurrences),
                        search_text=search
                    )
                # Multiple occurrences exist but specified one not found
                occurrences_data = [
                    {"index": i + 1, "position": f"{s}-{e}"}
                    for i, (s, e) in enumerate(all_occurrences[:5])
                ]
                return validator.create_ambiguous_search_error(
                    search_text=search,
                    occurrences=occurrences_data,
                    total_count=len(all_occurrences)
                )
            # Text not found at all
            return validator.create_search_not_found_error(
                search_text=search,
                match_case=match_case
            )

        start_index = calc_start
        end_index = calc_end
        search_info = {
            'search_text': search,
            'position': position,
            'occurrence': occurrence,
            'found_at_index': calc_start,
            'message': message
        }

        # For "before" and "after" positions, we're just inserting, not replacing
        if position in [SearchPosition.BEFORE.value, SearchPosition.AFTER.value]:
            # Set end_index to None to trigger insert mode rather than replace
            end_index = None

    # Validate text formatting params if provided
    has_formatting = any([
        bold is not None, italic is not None, underline is not None, strikethrough is not None,
        font_size, font_family, link is not None, foreground_color is not None, background_color is not None
    ])
    formatting_params_list = []
    if bold is not None:
        formatting_params_list.append("bold")
    if italic is not None:
        formatting_params_list.append("italic")
    if underline is not None:
        formatting_params_list.append("underline")
    if strikethrough is not None:
        formatting_params_list.append("strikethrough")
    if font_size:
        formatting_params_list.append("font_size")
    if font_family:
        formatting_params_list.append("font_family")
    if link is not None:
        formatting_params_list.append("link")
    if foreground_color is not None:
        formatting_params_list.append("foreground_color")
    if background_color is not None:
        formatting_params_list.append("background_color")

    if has_formatting:
        is_valid, error_msg = validator.validate_text_formatting_params(bold, italic, underline, font_size, font_family, link)
        if not is_valid:
            return validator.create_invalid_param_error(
                param_name="formatting",
                received=str(formatting_params_list),
                valid_values=["bold (bool)", "italic (bool)", "underline (bool)", "strikethrough (bool)", "font_size (1-400)", "font_family (string)", "link (URL string)", "foreground_color (hex/#FF0000 or named)", "background_color (hex or named)"],
                context=error_msg
            )

        # For formatting without text insertion, we need end_index
        # But if text is provided, we can calculate end_index automatically
        if end_index is None and text is None:
            is_valid, structured_error = validator.validate_formatting_range_structured(
                start_index=start_index,
                end_index=end_index,
                text=text,
                formatting_params=formatting_params_list
            )
            if not is_valid:
                return structured_error

        if end_index is not None:
            is_valid, structured_error = validator.validate_index_range_structured(start_index, end_index)
            if not is_valid:
                return structured_error

    requests = []
    operations = []

    # Track operation details for response
    operation_type = None
    actual_start_index = start_index
    actual_end_index = end_index
    format_styles = []

    # Handle text insertion/replacement/deletion
    if text is not None:
        if end_index is not None and end_index > start_index:
            if text == "":
                # Empty text with range = delete operation (no insert needed)
                operation_type = OperationType.DELETE
                if start_index == 0:
                    # Cannot delete at index 0 (first section break), start from 1
                    actual_start_index = 1
                    requests.append(create_delete_range_request(1, end_index))
                    operations.append(f"Deleted text from index 1 to {end_index}")
                else:
                    requests.append(create_delete_range_request(start_index, end_index))
                    operations.append(f"Deleted text from index {start_index} to {end_index}")
            else:
                # Text replacement
                operation_type = OperationType.REPLACE
                if start_index == 0:
                    # Special case: Cannot delete at index 0 (first section break)
                    # Instead, we insert new text at index 1 and then delete the old text
                    actual_start_index = 1
                    requests.append(create_insert_text_request(1, text))
                    adjusted_end = end_index + len(text)
                    requests.append(create_delete_range_request(1 + len(text), adjusted_end))
                    operations.append(f"Replaced text from index {start_index} to {end_index}")
                else:
                    # Normal replacement: delete old text, then insert new text
                    requests.extend([
                        create_delete_range_request(start_index, end_index),
                        create_insert_text_request(start_index, text)
                    ])
                    operations.append(f"Replaced text from index {start_index} to {end_index}")
        else:
            # Text insertion
            operation_type = OperationType.INSERT
            actual_start_index = 1 if start_index == 0 else start_index
            actual_end_index = None  # Insert has no end_index
            requests.append(create_insert_text_request(actual_start_index, text))
            operations.append(f"Inserted text at index {actual_start_index}")
            search_info['inserted_at_index'] = actual_start_index

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

        requests.append(create_format_text_request(
            format_start, format_end, bold, italic, underline, strikethrough,
            font_size, font_family, link, foreground_color, background_color
        ))

        format_details = []
        if bold is not None:
            format_details.append("bold")
        if italic is not None:
            format_details.append("italic")
        if underline is not None:
            format_details.append("underline")
        if strikethrough is not None:
            format_details.append("strikethrough")
        if font_size:
            format_details.append("font_size")
        if font_family:
            format_details.append("font_family")
        if link is not None:
            format_details.append("link")
        if foreground_color is not None:
            format_details.append("foreground_color")
        if background_color is not None:
            format_details.append("background_color")

        format_styles = format_details
        operations.append(f"Applied formatting ({', '.join(format_details)}) to range {format_start}-{format_end}")

        # If only formatting (no text operation), set operation type
        if operation_type is None:
            operation_type = OperationType.FORMAT
            actual_start_index = format_start
            actual_end_index = format_end

    # Handle preview mode - return what would change without actually modifying
    import json
    if preview:
        # For preview, we need doc_data to extract current content
        # If we don't have it from positioning resolution, fetch it now
        if not any([use_location_mode, use_range_mode, use_heading_mode, use_search_mode]):
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
                "end": actual_end_index if actual_end_index else actual_start_index
            },
            "positioning_info": search_info if search_info else {},
            "link": f"https://docs.google.com/document/d/{document_id}/edit"
        }

        # Calculate position shift
        text_length = len(text) if text else 0
        if operation_type == OperationType.INSERT:
            preview_result["position_shift"] = text_length
            preview_result["new_content"] = text
            preview_result["message"] = f"Would insert {text_length} characters at index {actual_start_index}"
        elif operation_type == OperationType.REPLACE:
            original_length = (actual_end_index or actual_start_index) - actual_start_index
            preview_result["position_shift"] = text_length - original_length
            preview_result["original_length"] = original_length
            preview_result["new_length"] = text_length
            preview_result["new_content"] = text

            # Extract current content at the range
            if doc_data:
                current = extract_text_at_range(doc_data, actual_start_index, actual_end_index or actual_start_index)
                preview_result["current_content"] = current.get("text", "")
                preview_result["context"] = {
                    "before": current.get("context_before", ""),
                    "after": current.get("context_after", "")
                }
            preview_result["message"] = f"Would replace {original_length} characters with {text_length} characters at index {actual_start_index}"
        elif operation_type == OperationType.DELETE:
            deleted_length = (actual_end_index or actual_start_index) - actual_start_index
            preview_result["position_shift"] = -deleted_length
            preview_result["deleted_length"] = deleted_length

            # Extract current content at the delete range
            if doc_data:
                current = extract_text_at_range(doc_data, actual_start_index, actual_end_index or actual_start_index)
                preview_result["current_content"] = current.get("text", "")
                preview_result["context"] = {
                    "before": current.get("context_before", ""),
                    "after": current.get("context_after", "")
                }
            preview_result["message"] = f"Would delete {deleted_length} characters from index {actual_start_index} to {actual_end_index}"
        elif operation_type == OperationType.FORMAT:
            preview_result["position_shift"] = 0
            preview_result["styles_to_apply"] = format_styles

            # Extract current content at the format range
            if doc_data:
                format_end_idx = actual_end_index or actual_start_index
                current = extract_text_at_range(doc_data, actual_start_index, format_end_idx)
                preview_result["current_content"] = current.get("text", "")
                preview_result["context"] = {
                    "before": current.get("context_before", ""),
                    "after": current.get("context_after", "")
                }
            preview_result["message"] = f"Would apply formatting ({', '.join(format_styles)}) to range {actual_start_index}-{actual_end_index}"
        else:
            preview_result["position_shift"] = 0
            preview_result["message"] = "Would perform operation"

        # Add resolved range info for range-based operations
        if range_result_info:
            preview_result['resolved_range'] = range_result_info

        return json.dumps(preview_result, indent=2)

    await asyncio.to_thread(
        service.documents().batchUpdate(
            documentId=document_id,
            body={'requests': requests}
        ).execute
    )

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
        styles_applied=format_styles if format_styles else None
    )

    # Convert to dict and return as JSON
    result_dict = op_result.to_dict()

    # Add resolved range info for range-based operations
    if range_result_info:
        result_dict['resolved_range'] = range_result_info

    # Add legacy message for backward compatibility
    operation_summary = "; ".join(operations)
    result_dict['legacy_message'] = operation_summary

    return json.dumps(result_dict, indent=2)


@server.tool()
@handle_http_errors("find_and_replace_doc", service_type="docs")
@require_google_service("docs", "docs_write")
async def find_and_replace_doc(
    service: Any,
    user_google_email: str,
    document_id: str,
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    preview: bool = False,
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

        # Replace only the FIRST "TODO": use modify_doc_text instead
        modify_doc_text(document_id="...", search="TODO", position="replace", text="DONE")

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        find_text: Text to search for (will replace ALL matches)
        replace_text: Text to replace with
        match_case: Whether to match case exactly (default: False)
        preview: If True, returns what would be replaced without making changes (default: False)

    Returns:
        str: JSON string with operation details.

        When preview=False (default), returns replacement count.

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
    logger.info(f"[find_and_replace_doc] Doc={document_id}, find='{find_text}', replace='{replace_text}', preview={preview}")

    import json

    # Handle preview mode - find all occurrences without modifying
    if preview:
        doc_data = await asyncio.to_thread(
            service.documents().get(documentId=document_id).execute
        )

        all_occurrences = find_all_occurrences_in_document(doc_data, find_text, match_case)
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
                    "after": context.get("context_after", "")
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
            preview_result["message"] = f"No occurrences of '{find_text}' found in document"
        else:
            preview_result["message"] = f"Would replace {len(all_occurrences)} occurrence(s) of '{find_text}' with '{replace_text}'"

        return json.dumps(preview_result, indent=2)

    # Execute the actual replacement
    requests = [create_find_replace_request(find_text, replace_text, match_case)]

    result = await asyncio.to_thread(
        service.documents().batchUpdate(
            documentId=document_id,
            body={'requests': requests}
        ).execute
    )

    # Extract number of replacements from response
    replacements = 0
    if 'replies' in result and result['replies']:
        reply = result['replies'][0]
        if 'replaceAllText' in reply:
            replacements = reply['replaceAllText'].get('occurrencesChanged', 0)

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    message = f"Replaced {replacements} occurrence(s) of '{find_text}' with '{replace_text}' in document {document_id}. Link: {link}"

    # Add hint when only 1 replacement was made - user might have wanted targeted replacement
    if replacements == 1:
        message += "\n\nTip: For single replacements, you can also use modify_doc_text with search + position='replace' to target a specific occurrence."

    return message


@server.tool()
@handle_http_errors("insert_doc_elements", service_type="docs")
@require_google_service("docs", "docs_write")
async def insert_doc_elements(
    service: Any,
    user_google_email: str,
    document_id: str,
    element_type: str,
    index: int,
    rows: int = None,
    columns: int = None,
    list_type: str = None,
    text: str = None,
) -> str:
    """
    Inserts structural elements like tables, lists, or page breaks into a Google Doc.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        element_type: Type of element to insert ("table", "list", "page_break")
        index: Position to insert element (0-based)
        rows: Number of rows for table (required for table)
        columns: Number of columns for table (required for table)
        list_type: Type of list ("UNORDERED", "ORDERED") (required for list)
        text: Initial text content for list items

    Returns:
        str: Confirmation message with insertion details
    """
    logger.info(f"[insert_doc_elements] Doc={document_id}, type={element_type}, index={index}")

    # Input validation
    validator = ValidationManager()

    # Handle the special case where we can't insert at the first section break
    # If index is 0, bump it to 1 to avoid the section break
    if index == 0:
        logger.debug("Adjusting index from 0 to 1 to avoid first section break")
        index = 1

    requests = []

    if element_type == "table":
        if not rows or not columns:
            return validator.create_missing_param_error(
                param_name="rows and columns",
                context="for table insertion",
                valid_values=["rows: positive integer", "columns: positive integer"]
            )

        requests.append(create_insert_table_request(index, rows, columns))
        description = f"table ({rows}x{columns})"

    elif element_type == "list":
        if not list_type:
            return validator.create_missing_param_error(
                param_name="list_type",
                context="for list insertion",
                valid_values=["UNORDERED", "ORDERED"]
            )

        if not text:
            text = "List item"

        # Insert text first, then create list
        requests.extend([
            create_insert_text_request(index, text + '\n'),
            create_bullet_list_request(index, index + len(text), list_type)
        ])
        description = f"{list_type.lower()} list"

    elif element_type == "page_break":
        requests.append(create_insert_page_break_request(index))
        description = "page break"

    else:
        return validator.create_invalid_param_error(
            param_name="element_type",
            received=element_type,
            valid_values=["table", "list", "page_break"]
        )

    await asyncio.to_thread(
        service.documents().batchUpdate(
            documentId=document_id,
            body={'requests': requests}
        ).execute
    )

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Inserted {description} at index {index} in document {document_id}. Link: {link}"

@server.tool()
@handle_http_errors("insert_doc_image", service_type="docs")
@require_multiple_services([
    {"service_type": "docs", "scopes": "docs_write", "param_name": "docs_service"},
    {"service_type": "drive", "scopes": "drive_read", "param_name": "drive_service"}
])
async def insert_doc_image(
    docs_service: Any,
    drive_service: Any,
    user_google_email: str,
    document_id: str,
    image_source: str,
    index: int,
    width: int = 0,
    height: int = 0,
) -> str:
    """
    Inserts an image into a Google Doc from Drive or a URL.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        image_source: Drive file ID or public image URL
        index: Position to insert image (0-based)
        width: Image width in points (optional)
        height: Image height in points (optional)

    Returns:
        str: Confirmation message with insertion details
    """
    logger.info(f"[insert_doc_image] Doc={document_id}, source={image_source}, index={index}")

    # Input validation
    validator = ValidationManager()

    # Handle the special case where we can't insert at the first section break
    # If index is 0, bump it to 1 to avoid the section break
    if index == 0:
        logger.debug("Adjusting index from 0 to 1 to avoid first section break")
        index = 1

    # Determine if source is a Drive file ID or URL
    is_drive_file = not (image_source.startswith('http://') or image_source.startswith('https://'))

    if is_drive_file:
        # Verify Drive file exists and get metadata
        try:
            file_metadata = await asyncio.to_thread(
                drive_service.files().get(
                    fileId=image_source,
                    fields="id, name, mimeType"
                ).execute
            )
            mime_type = file_metadata.get('mimeType', '')
            if not mime_type.startswith('image/'):
                return validator.create_image_error(
                    image_source=image_source,
                    actual_mime_type=mime_type
                )

            image_uri = f"https://drive.google.com/uc?id={image_source}"
            source_description = f"Drive file {file_metadata.get('name', image_source)}"
        except Exception as e:
            return validator.create_image_error(
                image_source=image_source,
                error_detail=str(e)
            )
    else:
        image_uri = image_source
        source_description = "URL image"

    # Use helper to create image request
    requests = [create_insert_image_request(index, image_uri, width, height)]

    await asyncio.to_thread(
        docs_service.documents().batchUpdate(
            documentId=document_id,
            body={'requests': requests}
        ).execute
    )

    size_info = ""
    if width or height:
        size_info = f" (size: {width or 'auto'}x{height or 'auto'} points)"

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Inserted {source_description}{size_info} at index {index} in document {document_id}. Link: {link}"

@server.tool()
@handle_http_errors("update_doc_headers_footers", service_type="docs")
@require_google_service("docs", "docs_write")
async def update_doc_headers_footers(
    service: Any,
    user_google_email: str,
    document_id: str,
    section_type: str,
    content: str,
    header_footer_type: str = "DEFAULT",
) -> str:
    """
    Updates headers or footers in a Google Doc.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        section_type: Type of section to update ("header" or "footer")
        content: Text content for the header/footer
        header_footer_type: Type of header/footer ("DEFAULT", "FIRST_PAGE_ONLY", "EVEN_PAGE")

    Returns:
        str: Confirmation message with update details
    """
    logger.info(f"[update_doc_headers_footers] Doc={document_id}, type={section_type}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    is_valid, error_msg = validator.validate_header_footer_params(section_type, header_footer_type)
    if not is_valid:
        if "section_type" in error_msg.lower():
            return validator.create_invalid_param_error(
                param_name="section_type",
                received=section_type,
                valid_values=["header", "footer"]
            )
        else:
            return validator.create_invalid_param_error(
                param_name="header_footer_type",
                received=header_footer_type,
                valid_values=["DEFAULT", "FIRST_PAGE_ONLY", "EVEN_PAGE"]
            )

    is_valid, error_msg = validator.validate_text_content(content)
    if not is_valid:
        return validator.create_invalid_param_error(
            param_name="content",
            received=repr(content)[:50],
            valid_values=["non-empty string"],
            context=error_msg
        )

    # Use HeaderFooterManager to handle the complex logic
    header_footer_manager = HeaderFooterManager(service)

    success, message = await header_footer_manager.update_header_footer_content(
        document_id, section_type, content, header_footer_type
    )

    if success:
        link = f"https://docs.google.com/document/d/{document_id}/edit"
        return f"{message}. Link: {link}"
    else:
        return validator.create_api_error(
            operation="update_header_footer",
            error_message=message,
            document_id=document_id
        )

@server.tool()
@handle_http_errors("batch_update_doc", service_type="docs")
@require_google_service("docs", "docs_write")
async def batch_update_doc(
    service: Any,
    user_google_email: str,
    document_id: str,
    operations: List[Dict[str, Any]],
) -> str:
    """
    DEPRECATED: Use batch_edit_doc instead.

    This tool is maintained for backward compatibility. The new batch_edit_doc tool
    combines all features from batch_update_doc and batch_modify_doc with:
    - Search-based positioning (insert before/after search text)
    - Automatic position adjustment
    - Both naming conventions (insert/insert_text, etc.)

    Executes multiple document operations in a single atomic batch update.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        operations: List of operation dictionaries. Each operation should contain:
                   - type: Operation type ('insert_text', 'delete_text', 'replace_text', 'format_text', 'insert_table', 'insert_page_break')
                   - Additional parameters specific to each operation type

    Example operations:
        [
            {"type": "insert_text", "index": 1, "text": "Hello World"},
            {"type": "format_text", "start_index": 1, "end_index": 12, "bold": true},
            {"type": "insert_table", "index": 20, "rows": 2, "columns": 3}
        ]

    Returns:
        str: JSON string with operation results including:
            - success (bool): Whether the batch succeeded
            - operations_count (int): Number of operations executed
            - total_position_shift (int): Cumulative position change from all operations
            - per_operation_shifts (list): Position shift for each operation
            - message (str): Human-readable summary
            - document_link (str): Link to the document

        Example response:
        {
            "success": true,
            "operations_count": 3,
            "total_position_shift": 15,
            "per_operation_shifts": [11, 0, 4],
            "message": "Successfully executed 3 operations",
            "document_link": "https://docs.google.com/document/d/.../edit"
        }

        Use total_position_shift for efficient follow-up edits:
        If you had a position at index 100 before these operations,
        the new position is 100 + total_position_shift.
    """
    import json

    logger.debug(f"[batch_update_doc] Doc={document_id}, operations={len(operations)}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    is_valid, error_msg = validator.validate_batch_operations(operations)
    if not is_valid:
        return validator.create_invalid_param_error(
            param_name="operations",
            received=f"list with {len(operations) if isinstance(operations, list) else 'invalid'} items",
            valid_values=["list of operation dicts with 'type' field"],
            context=error_msg
        )

    # Use BatchOperationManager to handle the complex logic
    batch_manager = BatchOperationManager(service)

    success, message, metadata = await batch_manager.execute_batch_operations(
        document_id, operations
    )

    if success:
        result = {
            "success": True,
            "operations_count": metadata.get('operations_count', len(operations)),
            "total_position_shift": metadata.get('total_position_shift', 0),
            "per_operation_shifts": metadata.get('per_operation_shifts', []),
            "message": message,
            "document_link": metadata.get('document_link', f"https://docs.google.com/document/d/{document_id}/edit")
        }
        return json.dumps(result, indent=2)
    else:
        error_result = {
            "success": False,
            "error": message,
            "document_id": document_id
        }
        return json.dumps(error_result, indent=2)


@server.tool()
@handle_http_errors("batch_modify_doc", service_type="docs")
@require_google_service("docs", "docs_write")
async def batch_modify_doc(
    service: Any,
    user_google_email: str,
    document_id: str,
    operations: List[Dict[str, Any]],
    auto_adjust_positions: bool = True,
) -> str:
    """
    DEPRECATED: Use batch_edit_doc instead.

    This tool is maintained for backward compatibility. The new batch_edit_doc tool
    provides identical functionality with a clearer name.

    Execute multiple document operations atomically with search-based positioning.

    This tool enables efficient batch editing with:
    - Search-based positioning (insert before/after search text)
    - Automatic position adjustment for sequential operations
    - Per-operation results with position shift tracking
    - Atomic execution (all succeed or all fail)

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        operations: List of operations. Each operation can use either:
            - Index-based: {"type": "insert_text", "index": 100, "text": "Hello"}
            - Search-based: {"type": "insert", "search": "Conclusion", "position": "before", "text": "New text"}

        auto_adjust_positions: If True (default), automatically adjusts positions
            for subsequent operations based on cumulative shifts from earlier operations.

    Supported operation types:
        - insert/insert_text: Insert text at position
        - delete/delete_text: Delete text range
        - replace/replace_text: Replace text range with new text
        - format/format_text: Apply formatting (bold, italic, underline, font_size, font_family)
        - insert_table: Insert table at position
        - insert_page_break: Insert page break
        - find_replace: Find and replace all occurrences

    Search-based positioning options:
        - search: Text to find in the document
        - position: "before" (insert before match), "after" (insert after), "replace" (replace match)
        - occurrence: Which occurrence to target (1=first, 2=second, -1=last)
        - match_case: Whether to match case exactly (default: True)

    Example operations:
        [
            {"type": "insert", "search": "Conclusion", "position": "before",
             "text": "\\n\\nNew section content.\\n"},
            {"type": "format", "search": "Important Note", "position": "replace",
             "bold": True, "font_size": 14},
            {"type": "insert_text", "index": 1, "text": "Header text\\n"},
            {"type": "find_replace", "find_text": "old term", "replace_text": "new term"}
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
        result1 = batch_modify_doc(doc_id, [{"type": "insert_text", "index": 100, "text": "15 char string."}])
        # result1["results"][0]["position_shift"] = 15

        # For a subsequent edit originally targeting index 200:
        # new_index = 200 + result1["total_position_shift"] = 215
        ```
    """
    import json

    logger.debug(f"[batch_modify_doc] Doc={document_id}, operations={len(operations)}")

    # Input validation
    validator = ValidationManager()

    is_valid, error_msg = validator.validate_document_id(document_id)
    if not is_valid:
        return json.dumps({"success": False, "error": error_msg}, indent=2)

    if not operations or not isinstance(operations, list):
        return json.dumps({
            "success": False,
            "error": "Operations must be a non-empty list"
        }, indent=2)

    # Use BatchOperationManager with enhanced search support
    batch_manager = BatchOperationManager(service)

    result = await batch_manager.execute_batch_with_search(
        document_id,
        operations,
        auto_adjust_positions=auto_adjust_positions
    )

    return json.dumps(result.to_dict(), indent=2)


@server.tool()
@handle_http_errors("batch_edit_doc", service_type="docs")
@require_google_service("docs", "docs_write")
async def batch_edit_doc(
    service: Any,
    user_google_email: str,
    document_id: str,
    operations: List[Dict[str, Any]],
    auto_adjust_positions: bool = True,
) -> str:
    """
    Execute multiple document operations atomically with search-based positioning.

    This is the RECOMMENDED batch editing tool, combining features from both
    batch_update_doc and batch_modify_doc. Use this tool for all batch operations.

    Features:
    - Search-based positioning (insert before/after search text)
    - Automatic position adjustment for sequential operations
    - Per-operation results with position shift tracking
    - Atomic execution (all succeed or all fail)
    - Accepts BOTH naming conventions (e.g., "insert" or "insert_text")

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to update
        operations: List of operations. Each operation can use either:
            - Index-based: {"type": "insert_text", "index": 100, "text": "Hello"}
            - Search-based: {"type": "insert", "search": "Conclusion", "position": "before", "text": "New text"}

        auto_adjust_positions: If True (default), automatically adjusts positions
            for subsequent operations based on cumulative shifts from earlier operations.

    Supported operation types (both naming styles work):
        - insert / insert_text: Insert text at position
        - delete / delete_text: Delete text range
        - replace / replace_text: Replace text range with new text
        - format / format_text: Apply formatting (bold, italic, underline, font_size, font_family)
        - insert_table: Insert table at position
        - insert_page_break: Insert page break
        - find_replace: Find and replace all occurrences

    Search-based positioning options:
        - search: Text to find in the document
        - position: "before" (insert before match), "after" (insert after), "replace" (replace match)
        - occurrence: Which occurrence to target (1=first, 2=second, -1=last)
        - match_case: Whether to match case exactly (default: True)

    Example operations:
        [
            {"type": "insert", "search": "Conclusion", "position": "before",
             "text": "\\n\\nNew section content.\\n"},
            {"type": "format", "search": "Important Note", "position": "replace",
             "bold": True, "font_size": 14},
            {"type": "insert_text", "index": 1, "text": "Header text\\n"},
            {"type": "find_replace", "find_text": "old term", "replace_text": "new term"}
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
    """
    import json

    logger.debug(f"[batch_edit_doc] Doc={document_id}, operations={len(operations)}")

    # Input validation
    validator = ValidationManager()

    is_valid, error_msg = validator.validate_document_id(document_id)
    if not is_valid:
        return json.dumps({"success": False, "error": error_msg}, indent=2)

    if not operations or not isinstance(operations, list):
        return json.dumps({
            "success": False,
            "error": "Operations must be a non-empty list"
        }, indent=2)

    # Use BatchOperationManager with enhanced search support
    batch_manager = BatchOperationManager(service)

    result = await batch_manager.execute_batch_with_search(
        document_id,
        operations,
        auto_adjust_positions=auto_adjust_positions
    )

    return json.dumps(result.to_dict(), indent=2)


# Detail level type for get_doc_info
from typing import Literal
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
        'title': doc.get('title', 'Untitled'),
        'document_id': document_id,
        'detail_level': detail
    }

    # Summary info (always included for context)
    complexity = analyze_document_complexity(doc)
    result['total_length'] = complexity.get('total_length', 1)
    result['safe_insertion_index'] = max(1, complexity.get('total_length', 1) - 1)

    if detail in ("summary", "all"):
        # Quick stats and safe insertion indices
        result['statistics'] = {
            'total_elements': complexity.get('total_elements', 0),
            'paragraphs': complexity.get('paragraphs', 0),
            'tables': complexity.get('tables', 0),
            'lists': complexity.get('lists', 0),
            'complexity_score': complexity.get('complexity_score', 'simple')
        }

    if detail in ("structure", "all"):
        # Full element hierarchy
        structure = parse_document_structure(doc)
        elements = []
        for element in structure['body']:
            elem_info = {
                'type': element['type'],
                'start_index': element['start_index'],
                'end_index': element['end_index']
            }
            if element['type'] == 'paragraph':
                elem_info['text_preview'] = element.get('text', '')[:100]
            elif element['type'] == 'table':
                elem_info['rows'] = element['rows']
                elem_info['columns'] = element['columns']
            elements.append(elem_info)
        result['elements'] = elements

    if detail in ("tables", "all"):
        # Table-focused view
        tables = find_tables(doc)
        table_list = []
        for i, table in enumerate(tables):
            table_data = extract_table_as_data(table)
            table_list.append({
                'index': i,
                'position': {'start': table['start_index'], 'end': table['end_index']},
                'dimensions': {'rows': table['rows'], 'columns': table['columns']},
                'preview': table_data[:3] if table_data else []  # First 3 rows
            })
        result['tables'] = table_list if table_list else []

    if detail in ("headings", "all"):
        # Headings outline for navigation
        all_elements = extract_structural_elements(doc)
        headings_outline = build_headings_outline(all_elements)
        result['headings_outline'] = _clean_outline(headings_outline)

        # Also include flat heading list for quick reference
        headings = [e for e in all_elements if e['type'].startswith('heading')]
        result['headings'] = [
            {'level': h.get('level', 0), 'text': h.get('text', ''),
             'start_index': h['start_index'], 'end_index': h['end_index']}
            for h in headings
        ]

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Document info for {document_id} (detail={detail}):\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


@server.tool()
@handle_http_errors("inspect_doc_structure", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def inspect_doc_structure(
    service: Any,
    user_google_email: str,
    document_id: str,
    detailed: bool = False,
) -> str:
    """
    [LEGACY] Find safe insertion points and document structure.

    DEPRECATED: Use get_doc_info() instead for a unified interface.
    - get_doc_info(doc_id, detail="summary") - equivalent to inspect_doc_structure(detailed=False)
    - get_doc_info(doc_id, detail="all") - equivalent to inspect_doc_structure(detailed=True)

    This function is maintained for backward compatibility.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to inspect
        detailed: Whether to return detailed structure information

    Returns:
        str: JSON string containing document structure and safe insertion indices
    """
    logger.debug(f"[inspect_doc_structure] Doc={document_id}, detailed={detailed}")

    # Get the document
    doc = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    if detailed:
        # Return full parsed structure
        structure = parse_document_structure(doc)

        # Simplify for JSON serialization
        result = {
            'title': structure['title'],
            'total_length': structure['total_length'],
            'statistics': {
                'elements': len(structure['body']),
                'tables': len(structure['tables']),
                'paragraphs': sum(1 for e in structure['body'] if e.get('type') == 'paragraph'),
                'has_headers': bool(structure['headers']),
                'has_footers': bool(structure['footers'])
            },
            'elements': []
        }

        # Add element summaries
        for element in structure['body']:
            elem_summary = {
                'type': element['type'],
                'start_index': element['start_index'],
                'end_index': element['end_index']
            }

            if element['type'] == 'table':
                elem_summary['rows'] = element['rows']
                elem_summary['columns'] = element['columns']
                elem_summary['cell_count'] = len(element.get('cells', []))
            elif element['type'] == 'paragraph':
                elem_summary['text_preview'] = element.get('text', '')[:100]

            result['elements'].append(elem_summary)

        # Add table details
        if structure['tables']:
            result['tables'] = []
            for i, table in enumerate(structure['tables']):
                table_data = extract_table_as_data(table)
                result['tables'].append({
                    'index': i,
                    'position': {'start': table['start_index'], 'end': table['end_index']},
                    'dimensions': {'rows': table['rows'], 'columns': table['columns']},
                    'preview': table_data[:3] if table_data else []  # First 3 rows
                })

    else:
        # Return basic analysis
        result = analyze_document_complexity(doc)

        # Add table information
        tables = find_tables(doc)
        if tables:
            result['table_details'] = []
            for i, table in enumerate(tables):
                result['table_details'].append({
                    'index': i,
                    'rows': table['rows'],
                    'columns': table['columns'],
                    'start_index': table['start_index'],
                    'end_index': table['end_index']
                })

    import json
    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Document structure analysis for {document_id}:\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


@server.tool()
@handle_http_errors("get_doc_structure", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def get_doc_structure(
    service: Any,
    user_google_email: str,
    document_id: str,
    element_types: List[str] = None,
) -> str:
    """
    [LEGACY] Get a hierarchical view of the document's structural elements.

    DEPRECATED: Use get_doc_info() instead for a unified interface.
    - get_doc_info(doc_id, detail="structure") - for full element hierarchy
    - get_doc_info(doc_id, detail="headings") - for headings outline only

    Note: get_doc_info does not support element_types filtering. If you need
    specific element type filtering, continue using this function.

    This function is maintained for backward compatibility.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document to analyze
        element_types: Optional list of element types to filter (e.g., ["heading1", "heading2", "table"])
                      Supported types: heading1-6, title, paragraph, table, bullet_list, numbered_list

    Returns:
        str: JSON containing:
            - elements: List of structural elements with type, text, and positions
            - headings_outline: Hierarchical outline of headings
            - statistics: Count of each element type
    """
    logger.debug(f"[get_doc_structure] Doc={document_id}, filters={element_types}")

    # Input validation
    validator = ValidationManager()
    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    # Get the document
    doc = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Extract structural elements
    elements = extract_structural_elements(doc)

    # Filter by element types if specified
    if element_types:
        elements = [e for e in elements if e['type'] in element_types]

    # Build headings outline
    all_elements = extract_structural_elements(doc)  # Need full list for outline
    headings_outline = build_headings_outline(all_elements)

    # Calculate statistics
    statistics = {}
    for elem in all_elements:
        elem_type = elem['type']
        statistics[elem_type] = statistics.get(elem_type, 0) + 1

    # Clean up elements for output (remove internal fields)
    clean_elements = []
    for elem in elements:
        clean_elem = {
            'type': elem['type'],
            'start_index': elem['start_index'],
            'end_index': elem['end_index']
        }
        if 'text' in elem:
            clean_elem['text'] = elem['text']
        if 'level' in elem:
            clean_elem['level'] = elem['level']
        if 'rows' in elem:
            clean_elem['rows'] = elem['rows']
            clean_elem['columns'] = elem.get('columns', 0)
        if 'items' in elem:
            clean_elem['items'] = [
                {'text': item['text'], 'start_index': item['start_index'], 'end_index': item['end_index']}
                for item in elem['items']
            ]
        clean_elements.append(clean_elem)

    # Build result
    result = {
        'elements': clean_elements,
        'headings_outline': _clean_outline(headings_outline),
        'statistics': statistics
    }

    import json
    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Document structure for {document_id}:\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


def _clean_outline(outline: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Clean outline for JSON output, removing internal fields."""
    clean = []
    for item in outline:
        clean_item = {
            'level': item['level'],
            'text': item['text'],
            'start_index': item['start_index'],
            'end_index': item['end_index'],
            'children': _clean_outline(item.get('children', []))
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
            valid_values=["non-empty heading text"]
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
        heading_list = [h['text'] for h in all_headings] if all_headings else []
        return validator.create_heading_not_found_error(
            heading=heading,
            available_headings=heading_list,
            match_case=match_case
        )

    # Build result
    result = {
        'heading': section['heading'],
        'level': section['level'],
        'start_index': section['start_index'],
        'end_index': section['end_index'],
        'content': section['content']
    }

    if include_subsections:
        result['subsections'] = section['subsections']

    import json
    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Section '{heading}' in document {document_id}:\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


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
    bold_headers: bool = True,
) -> str:
    """
    Creates a table and populates it with data in one reliable operation.

    SIMPLIFIED USAGE - No pre-flight call needed:
    - Omit 'index' to append table at end of document (recommended)
    - Use 'after_heading' to insert table after a specific heading
    - Provide explicit 'index' only for precise positioning

    EXAMPLE DATA FORMAT:
    table_data = [
        ["Header1", "Header2", "Header3"],    # Row 0 - headers
        ["Data1", "Data2", "Data3"],          # Row 1 - first data row
        ["Data4", "Data5", "Data6"]           # Row 2 - second data row
    ]

    USAGE PATTERNS:
    1. Append to end (simplest): create_table_with_data(doc_id, data)
    2. After heading: create_table_with_data(doc_id, data, after_heading="Data Section")
    3. Explicit index: create_table_with_data(doc_id, data, index=42)

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
        index: Document position (optional, defaults to end of document)
        after_heading: Insert after this heading (optional, case-insensitive)
        bold_headers: Whether to make first row bold (default: true)

    Returns:
        str: Confirmation with table details and link
    """
    logger.debug(
        f"[create_table_with_data] Doc={document_id}, "
        f"index={index}, after_heading={after_heading}"
    )

    # Input validation
    validator = ValidationManager()

    is_valid, error_msg = validator.validate_document_id(document_id)
    if not is_valid:
        return f"ERROR: {error_msg}"

    is_valid, error_msg = validator.validate_table_data(table_data)
    if not is_valid:
        return f"ERROR: {error_msg}"

    # Resolve insertion index
    resolved_index = index
    location_description = None

    if index is not None and after_heading is not None:
        return (
            "ERROR: Cannot specify both 'index' and 'after_heading'. "
            "Use one or the other."
        )

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
            return (
                f"ERROR: Failed to fetch document for index calculation: "
                f"{str(e)}"
            )

        if after_heading is not None:
            # Insert after specified heading
            insertion_point = find_section_insertion_point(
                doc_data, after_heading, position='end'
            )
            if insertion_point is None:
                # List available headings to help user
                headings = get_all_headings(doc_data)
                if headings:
                    heading_list = ", ".join(
                        [f'"{h["text"]}"' for h in headings[:5]]
                    )
                    more = (
                        f" (and {len(headings) - 5} more)"
                        if len(headings) > 5 else ""
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
        else:
            # Default: append to end of document
            structure = parse_document_structure(doc_data)
            total_length = structure['total_length']
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
        rows = metadata.get('rows', 0)
        columns = metadata.get('columns', 0)

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
        table_index: Which table to debug (0 = first table, 1 = second table, etc.)

    Returns:
        str: Detailed JSON structure showing table layout, cell positions, and current content
    """
    logger.debug(f"[debug_table_structure] Doc={document_id}, table_index={table_index}")

    # Get the document
    doc = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Find tables
    tables = find_tables(doc)
    if table_index >= len(tables):
        validator = ValidationManager()
        return validator.create_table_not_found_error(
            table_index=table_index,
            total_tables=len(tables)
        )

    table_info = tables[table_index]

    import json

    # Extract detailed cell information
    debug_info = {
        'table_index': table_index,
        'dimensions': f"{table_info['rows']}x{table_info['columns']}",
        'table_range': f"[{table_info['start_index']}-{table_info['end_index']}]",
        'cells': []
    }

    for row_idx, row in enumerate(table_info['cells']):
        row_info = []
        for col_idx, cell in enumerate(row):
            cell_debug = {
                'position': f"({row_idx},{col_idx})",
                'range': f"[{cell['start_index']}-{cell['end_index']}]",
                'insertion_index': cell.get('insertion_index', 'N/A'),
                'current_content': repr(cell.get('content', '')),
                'content_elements_count': len(cell.get('content_elements', []))
            }
            row_info.append(cell_debug)
        debug_info['cells'].append(row_info)

    link = f"https://docs.google.com/document/d/{document_id}/edit"
    return f"Table structure debug for table {table_index}:\n\n{json.dumps(debug_info, indent=2)}\n\nLink: {link}"

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
    logger.info(f"[export_doc_to_pdf] Email={user_google_email}, Doc={document_id}, pdf_filename={pdf_filename}, folder_id={folder_id}")

    # Input validation
    validator = ValidationManager()

    # Get file metadata first to validate it's a Google Doc
    try:
        file_metadata = await asyncio.to_thread(
            service.files().get(
                fileId=document_id,
                fields="id, name, mimeType, webViewLink"
            ).execute
        )
    except Exception as e:
        return validator.create_pdf_export_error(
            document_id=document_id,
            stage="access",
            error_detail=str(e)
        )

    mime_type = file_metadata.get("mimeType", "")
    original_name = file_metadata.get("name", "Unknown Document")
    web_view_link = file_metadata.get("webViewLink", "#")

    # Verify it's a Google Doc
    if mime_type != "application/vnd.google-apps.document":
        return validator.create_invalid_document_type_error(
            document_id=document_id,
            file_name=original_name,
            actual_mime_type=mime_type
        )

    logger.info(f"[export_doc_to_pdf] Exporting '{original_name}' to PDF")

    # Export the document as PDF
    try:
        request_obj = service.files().export_media(
            fileId=document_id,
            mimeType='application/pdf'
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
            document_id=document_id,
            stage="export",
            error_detail=str(e)
        )

    # Determine PDF filename
    if not pdf_filename:
        pdf_filename = f"{original_name}_PDF.pdf"
    elif not pdf_filename.endswith('.pdf'):
        pdf_filename += '.pdf'

    # Upload PDF to Drive
    try:
        # Reuse the existing BytesIO object by resetting to the beginning
        fh.seek(0)
        # Create media upload object
        media = MediaIoBaseUpload(
            fh,
            mimetype='application/pdf',
            resumable=True
        )
        
        # Prepare file metadata for upload
        file_metadata = {
            'name': pdf_filename,
            'mimeType': 'application/pdf'
        }
        
        # Add parent folder if specified
        if folder_id:
            file_metadata['parents'] = [folder_id]
        
        # Upload the file
        uploaded_file = await asyncio.to_thread(
            service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id, name, webViewLink, parents',
                supportsAllDrives=True
            ).execute
        )
        
        pdf_file_id = uploaded_file.get('id')
        pdf_web_link = uploaded_file.get('webViewLink', '#')
        pdf_parents = uploaded_file.get('parents', [])
        
        logger.info(f"[export_doc_to_pdf] Successfully uploaded PDF to Drive: {pdf_file_id}")
        
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
            error_detail=f"{str(e)}. PDF was generated successfully ({pdf_size:,} bytes) but could not be saved to Drive."
        )


@server.tool()
@handle_http_errors("preview_search_results", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def preview_search_results(
    service: Any,
    user_google_email: str,
    document_id: str,
    search_text: str,
    match_case: bool = True,
    context_chars: int = 50,
) -> str:
    """
    Preview what text will be matched by a search operation before modifying.

    USE THIS TOOL BEFORE:
    - Using modify_doc_text with search parameter
    - Using batch_modify_doc with search-based operations
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

    logger.debug(f"[preview_search_results] Doc={document_id}, search='{search_text}', match_case={match_case}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    if not search_text or not search_text.strip():
        error = DocsErrorBuilder.missing_required_param(
            param_name="search_text",
            context_description="for search preview",
            valid_values=["non-empty search string"]
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
            "message": f"No matches found for '{search_text}'" +
                      (" (case-sensitive)" if match_case else " (case-insensitive)"),
            "hint": "Try match_case=False for case-insensitive search" if match_case else
                   "Verify the search text exists in the document"
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
                "context_after": context_after
            }
        else:
            # Fallback if mapping failed
            match_info = {
                "occurrence": i + 1,
                "start_index": start_idx,
                "end_index": end_idx,
                "context_before": "[context unavailable]",
                "matched_text": search_text,
                "context_after": "[context unavailable]"
            }

        matches_with_context.append(match_info)

    result = {
        "total_matches": len(all_matches),
        "search_text": search_text,
        "match_case": match_case,
        "matches": matches_with_context
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
            - 'table': All tables in the document
            - 'heading': All headings (any level)
            - 'heading1' through 'heading6': Specific heading levels
            - 'title': Document title style elements
            - 'paragraph': Regular paragraphs (non-heading, non-list)
            - 'bullet_list': Bulleted lists
            - 'numbered_list': Numbered/ordered lists
            - 'list': Both bullet and numbered lists
            - 'table_of_contents': Table of contents elements

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
            valid_values=["table", "heading", "heading1-6", "paragraph", "bullet_list", "numbered_list", "list"]
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
            'type': elem['type'],
            'start_index': elem['start_index'],
            'end_index': elem['end_index']
        }
        if 'text' in elem:
            clean_elem['text'] = elem['text']
        if 'level' in elem:
            clean_elem['level'] = elem['level']
        if 'rows' in elem:
            clean_elem['rows'] = elem['rows']
            clean_elem['columns'] = elem.get('columns', 0)
        if 'items' in elem:
            clean_elem['item_count'] = len(elem['items'])
            clean_elem['items'] = [
                {'text': item['text'], 'start_index': item['start_index'], 'end_index': item['end_index']}
                for item in elem['items']
            ]
        clean_elements.append(clean_elem)

    result = {
        'count': len(clean_elements),
        'element_type': element_type,
        'elements': clean_elements
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
            context_description="document position"
        )
        return format_error(error)

    # Get the document
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Get ancestors
    ancestors = get_element_ancestors(doc_data, index)

    result = {
        'index': index,
        'ancestors': ancestors,
        'depth': len(ancestors),
        'innermost_section': ancestors[-1] if ancestors else None
    }

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    if ancestors:
        innermost = ancestors[-1]
        summary = f"Position {index} is within section '{innermost['text']}' (level {innermost['level']}) with {len(ancestors)} containing section(s)"
    else:
        summary = f"Position {index} is not within any heading section"

    return f"{summary}:\n\n{json.dumps(result, indent=2)}\n\nLink: {link}"


@server.tool()
@handle_http_errors("navigate_heading_siblings", is_read_only=True, service_type="docs")
@require_google_service("docs", "docs_read")
async def navigate_heading_siblings(
    service: Any,
    user_google_email: str,
    document_id: str,
    heading: str,
    match_case: bool = False,
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

    logger.debug(f"[navigate_heading_siblings] Doc={document_id}, heading={heading}, match_case={match_case}")

    # Input validation
    validator = ValidationManager()

    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    if not heading or not heading.strip():
        error = DocsErrorBuilder.missing_required_param(
            param_name="heading",
            context_description="for sibling navigation",
            valid_values=["non-empty heading text"]
        )
        return format_error(error)

    # Get the document
    doc_data = await asyncio.to_thread(
        service.documents().get(documentId=document_id).execute
    )

    # Get siblings
    result = get_heading_siblings(doc_data, heading, match_case)

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    if not result['found']:
        # Provide helpful error with available headings
        all_headings = get_all_headings(doc_data)
        heading_list = [h['text'] for h in all_headings] if all_headings else []
        return validator.create_heading_not_found_error(
            heading=heading,
            available_headings=heading_list,
            match_case=match_case
        )

    summary_parts = [f"Heading '{result['heading']['text']}' (level {result['level']})"]
    summary_parts.append(f"Position {result['position_in_siblings']} of {result['siblings_count']} headings at this level")
    if result['previous']:
        summary_parts.append(f"Previous: '{result['previous']['text']}'")
    else:
        summary_parts.append("No previous sibling (first at this level)")
    if result['next']:
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
            if heading['start_index'] <= idx:
                current_section = heading['text']
            else:
                break
        return current_section

    # Traverse document to find all links
    links = []
    body = doc_data.get('body', {})
    content = body.get('content', [])

    def process_elements(elements: list, depth: int = 0) -> None:
        """Recursively process elements to extract links."""
        if depth > 10:  # Prevent infinite recursion
            return

        for element in elements:
            if 'paragraph' in element:
                paragraph = element['paragraph']
                for para_elem in paragraph.get('elements', []):
                    if 'textRun' in para_elem:
                        text_run = para_elem['textRun']
                        text_style = text_run.get('textStyle', {})

                        # Check for link in text style
                        if 'link' in text_style:
                            link_info = text_style['link']
                            url = link_info.get('url', '')

                            # Skip internal document bookmarks
                            if url and not url.startswith('#'):
                                link_text = text_run.get('content', '').strip()
                                start_idx = para_elem.get('startIndex', 0)
                                end_idx = para_elem.get('endIndex', 0)

                                link_entry = {
                                    'text': link_text,
                                    'url': url,
                                    'start_index': start_idx,
                                    'end_index': end_idx,
                                }

                                if include_section_context:
                                    link_entry['section'] = find_section_for_index(start_idx)

                                links.append(link_entry)

            elif 'table' in element:
                # Process table cells
                table = element['table']
                for row in table.get('tableRows', []):
                    for cell in row.get('tableCells', []):
                        process_elements(cell.get('content', []), depth + 1)

    process_elements(content)

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    result = {
        'total_links': len(links),
        'links': links,
        'document_link': link
    }

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
    inline_objects = doc_data.get('inlineObjects', {})

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
            if heading['start_index'] <= idx:
                current_section = heading['text']
            else:
                break
        return current_section

    # Map to track inline object references and their positions
    image_refs = {}  # object_id -> start_index

    body = doc_data.get('body', {})
    content = body.get('content', [])

    def process_elements(elements: list, depth: int = 0) -> None:
        """Recursively process elements to find inline object references."""
        if depth > 10:
            return

        for element in elements:
            if 'paragraph' in element:
                paragraph = element['paragraph']
                for para_elem in paragraph.get('elements', []):
                    if 'inlineObjectElement' in para_elem:
                        obj_elem = para_elem['inlineObjectElement']
                        obj_id = obj_elem.get('inlineObjectId')
                        if obj_id:
                            start_idx = para_elem.get('startIndex', 0)
                            image_refs[obj_id] = start_idx

            elif 'table' in element:
                table = element['table']
                for row in table.get('tableRows', []):
                    for cell in row.get('tableCells', []):
                        process_elements(cell.get('content', []), depth + 1)

    process_elements(content)

    # Build image list from inline objects registry
    images = []
    for obj_id, obj_data in inline_objects.items():
        props = obj_data.get('inlineObjectProperties', {})
        embedded_obj = props.get('embeddedObject', {})

        # Get image properties
        image_props = embedded_obj.get('imageProperties', {})
        content_uri = image_props.get('contentUri', '')
        source_uri = image_props.get('sourceUri', '')

        # Get size
        size = embedded_obj.get('size', {})
        width = size.get('width', {})
        height = size.get('height', {})

        width_pt = width.get('magnitude', 0) if width.get('unit') == 'PT' else width.get('magnitude', 0)
        height_pt = height.get('magnitude', 0) if height.get('unit') == 'PT' else height.get('magnitude', 0)

        start_idx = image_refs.get(obj_id, 0)

        image_entry = {
            'object_id': obj_id,
            'content_uri': content_uri,
            'width_pt': width_pt,
            'height_pt': height_pt,
            'start_index': start_idx,
        }

        if source_uri:
            image_entry['source_uri'] = source_uri

        if include_section_context:
            image_entry['section'] = find_section_for_index(start_idx)

        images.append(image_entry)

    # Sort by position in document
    images.sort(key=lambda x: x['start_index'])

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    result = {
        'total_images': len(images),
        'images': images,
        'document_link': link
    }

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
        'courier new', 'consolas', 'monaco', 'menlo', 'source code pro',
        'fira code', 'jetbrains mono', 'roboto mono', 'ubuntu mono',
        'droid sans mono', 'liberation mono', 'dejavu sans mono',
        'lucida console', 'andale mono', 'courier'
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
            if heading['start_index'] <= idx:
                current_section = heading['text']
            else:
                break
        return current_section

    def is_code_formatted(text_style: dict) -> tuple[bool, str, bool]:
        """Check if text style indicates code formatting.

        Returns:
            Tuple of (is_code, font_family, has_background)
        """
        # Check font family
        font_info = text_style.get('weightedFontFamily', {})
        font_family = font_info.get('fontFamily', '').lower()

        is_monospace = any(mono in font_family for mono in MONOSPACE_FONTS)

        # Check background color
        bg_color = text_style.get('backgroundColor', {})
        has_background = bool(bg_color.get('color', {}))

        return is_monospace, font_info.get('fontFamily', ''), has_background

    # Collect code-formatted text runs
    code_runs = []

    body = doc_data.get('body', {})
    content = body.get('content', [])

    def process_elements(elements: list, depth: int = 0) -> None:
        """Recursively process elements to find code-formatted text."""
        if depth > 10:
            return

        for element in elements:
            if 'paragraph' in element:
                paragraph = element['paragraph']
                for para_elem in paragraph.get('elements', []):
                    if 'textRun' in para_elem:
                        text_run = para_elem['textRun']
                        text_style = text_run.get('textStyle', {})
                        text_content = text_run.get('content', '')

                        is_code, font_family, has_background = is_code_formatted(text_style)

                        if is_code and text_content.strip():
                            code_runs.append({
                                'content': text_content,
                                'font_family': font_family,
                                'start_index': para_elem.get('startIndex', 0),
                                'end_index': para_elem.get('endIndex', 0),
                                'has_background': has_background
                            })

            elif 'table' in element:
                table = element['table']
                for row in table.get('tableRows', []):
                    for cell in row.get('tableCells', []):
                        process_elements(cell.get('content', []), depth + 1)

    process_elements(content)

    # Merge consecutive code runs into blocks
    code_blocks = []
    current_block = None

    for run in sorted(code_runs, key=lambda x: x['start_index']):
        if current_block is None:
            current_block = run.copy()
        elif run['start_index'] <= current_block['end_index'] + 1:
            # Consecutive or overlapping - merge
            current_block['content'] += run['content']
            current_block['end_index'] = max(current_block['end_index'], run['end_index'])
            current_block['has_background'] = current_block['has_background'] or run['has_background']
        else:
            # Gap - save current block and start new one
            if include_section_context:
                current_block['section'] = find_section_for_index(current_block['start_index'])
            code_blocks.append(current_block)
            current_block = run.copy()

    # Don't forget the last block
    if current_block:
        if include_section_context:
            current_block['section'] = find_section_for_index(current_block['start_index'])
        code_blocks.append(current_block)

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    result = {
        'total_code_blocks': len(code_blocks),
        'code_blocks': code_blocks,
        'document_link': link
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
    counts = {
        'headings': 0,
        'paragraphs': 0,
        'tables': 0,
        'lists': 0
    }

    for elem in elements:
        elem_type = elem.get('type', '')
        if elem_type.startswith('heading') or elem_type == 'title':
            counts['headings'] += 1
        elif elem_type == 'paragraph':
            counts['paragraphs'] += 1
        elif elem_type == 'table':
            counts['tables'] += 1
        elif elem_type in ('bullet_list', 'numbered_list'):
            counts['lists'] += 1

    # Calculate total characters
    body = doc_data.get('body', {})
    content = body.get('content', [])
    total_chars = 0
    if content:
        total_chars = content[-1].get('endIndex', 0) - 1  # Subtract 1 for start index

    # Build outline with previews
    outline = build_headings_outline(elements)

    def get_section_preview(heading_start: int, heading_end: int) -> str:
        """Get preview text from section content."""
        preview_parts = []
        word_count = 0

        for elem in elements:
            if elem['start_index'] > heading_end:
                # We're past this heading
                if elem.get('type', '').startswith('heading'):
                    # Hit next heading, stop
                    break

            if elem['start_index'] > heading_start:
                text = elem.get('text', '')
                if text:
                    words = text.split()
                    for word in words:
                        if word_count >= max_preview_words:
                            break
                        preview_parts.append(word)
                        word_count += 1
                    if word_count >= max_preview_words:
                        break

        preview = ' '.join(preview_parts)
        if word_count >= max_preview_words:
            preview += '...'
        return preview

    def add_previews_to_outline(outline_items: list, elements_list: list) -> None:
        """Recursively add preview text to outline items."""
        for i, item in enumerate(outline_items):
            # Find section end (next sibling heading start or document end)
            section_end = total_chars
            if i + 1 < len(outline_items):
                section_end = outline_items[i + 1]['start_index']

            item['preview'] = get_section_preview(item['start_index'], section_end)

            if item.get('children'):
                add_previews_to_outline(item['children'], elements_list)

    add_previews_to_outline(outline, elements)

    link = f"https://docs.google.com/document/d/{document_id}/edit"

    result = {
        'title': doc_data.get('title', ''),
        'total_characters': total_chars,
        'total_headings': counts['headings'],
        'total_paragraphs': counts['paragraphs'],
        'total_tables': counts['tables'],
        'total_lists': counts['lists'],
        'outline': outline,
        'document_link': link
    }

    return json.dumps(result, indent=2)


# Create comment management tools for documents
_comment_tools = create_comment_tools("document", "document_id")

# Extract and register the functions
read_doc_comments = _comment_tools['read_comments']
create_doc_comment = _comment_tools['create_comment']
reply_to_comment = _comment_tools['reply_to_comment']
resolve_comment = _comment_tools['resolve_comment']


# =============================================================================
# Operation History and Undo Tools
# =============================================================================

from gdocs.managers.history_manager import (
    get_history_manager,
    UndoCapability,
)
from gdocs.docs_helpers import extract_text_at_range


@server.tool()
async def get_doc_operation_history(
    document_id: str,
    limit: int = 10,
    include_undone: bool = False,
) -> str:
    """
    Get the operation history for a Google Doc.

    This tool returns a list of recent operations performed on the document
    through the MCP tools. Each operation includes details about what was
    changed and whether it can be undone.

    Note: History is stored in-memory and is lost when the MCP server restarts.
    Operations performed outside this MCP server (e.g., in the Google Docs UI)
    are not tracked.

    Args:
        document_id: ID of the document
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

    Example Response:
        {
            "document_id": "1abc123...",
            "operations": [
                {
                    "id": "op_20250604123456_1",
                    "timestamp": "2025-06-04T12:34:56Z",
                    "operation_type": "insert_text",
                    "start_index": 100,
                    "position_shift": 15,
                    "undo_capability": "full",
                    "undone": false
                }
            ],
            "total_operations": 5,
            "undoable_count": 3
        }
    """
    import json

    logger.debug(f"[get_doc_operation_history] Doc={document_id}, limit={limit}")

    # Validate inputs
    if not document_id or not document_id.strip():
        return json.dumps({
            "success": False,
            "error": "document_id is required"
        }, indent=2)

    limit = min(max(1, limit), 50)  # Clamp between 1 and 50

    manager = get_history_manager()
    operations = manager.get_history(document_id, limit=limit, include_undone=include_undone)

    # Count undoable operations
    undoable_count = sum(
        1 for op in operations
        if not op.get('undone') and op.get('undo_capability') != 'none'
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

    This tool reverses the most recent undoable operation by executing a
    compensating operation (e.g., re-inserting deleted text, deleting
    inserted text, or restoring replaced text).

    Important Limitations:
    - Only operations performed through this MCP server can be undone
    - History is stored in-memory and lost on server restart
    - Undo may fail if the document was modified externally
    - Some operations (find_replace, insert_table) cannot be undone
    - Format undo requires original formatting to be captured

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

    Example Response (success):
        {
            "success": true,
            "message": "Successfully undone: insert_text",
            "operation_id": "op_20250604123456_1",
            "reverse_operation": {
                "type": "delete_text",
                "start_index": 100,
                "end_index": 115
            },
            "document_link": "https://docs.google.com/document/d/.../edit"
        }

    Example Response (failure):
        {
            "success": false,
            "message": "No undoable operations found",
            "error": "All operations have been undone or cannot be undone"
        }
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
        return json.dumps({
            "success": False,
            "message": undo_result.message,
            "error": undo_result.error,
        }, indent=2)

    # Execute the reverse operation
    reverse_op = undo_result.reverse_operation
    op_type = reverse_op.get("type")

    try:
        requests = []

        if op_type == "insert_text":
            requests.append(create_insert_text_request(
                reverse_op["index"],
                reverse_op["text"]
            ))
        elif op_type == "delete_text":
            requests.append(create_delete_range_request(
                reverse_op["start_index"],
                reverse_op["end_index"]
            ))
        elif op_type == "replace_text":
            # Replace = delete + insert
            requests.append(create_delete_range_request(
                reverse_op["start_index"],
                reverse_op["end_index"]
            ))
            requests.append(create_insert_text_request(
                reverse_op["start_index"],
                reverse_op["text"]
            ))
        elif op_type == "format_text":
            requests.append(create_format_text_request(
                reverse_op["start_index"],
                reverse_op["end_index"],
                reverse_op.get("bold"),
                reverse_op.get("italic"),
                reverse_op.get("underline"),
                reverse_op.get("font_size"),
                reverse_op.get("font_family")
            ))
        else:
            return json.dumps({
                "success": False,
                "message": f"Unknown reverse operation type: {op_type}",
                "error": "Cannot execute undo"
            }, indent=2)

        # Execute the batch update
        await asyncio.to_thread(
            service.documents().batchUpdate(
                documentId=document_id,
                body={'requests': requests}
            ).execute
        )

        # Mark the operation as undone
        manager.mark_undone(document_id, undo_result.operation_id)

        return json.dumps({
            "success": True,
            "message": f"Successfully undone operation",
            "operation_id": undo_result.operation_id,
            "reverse_operation": {
                "type": op_type,
                "description": reverse_op.get("description", ""),
            },
            "document_link": f"https://docs.google.com/document/d/{document_id}/edit"
        }, indent=2)

    except Exception as e:
        logger.error(f"Failed to execute undo: {str(e)}")
        return json.dumps({
            "success": False,
            "message": "Failed to execute undo operation",
            "operation_id": undo_result.operation_id,
            "error": str(e)
        }, indent=2)


@server.tool()
async def clear_doc_history(
    document_id: str,
) -> str:
    """
    Clear the operation history for a Google Doc.

    This removes all tracked operations for the specified document,
    making undo unavailable for previous operations.

    Args:
        document_id: ID of the document

    Returns:
        str: JSON containing:
            - success: Whether the history was cleared
            - message: Description of the result
            - document_id: The document ID

    Example Response:
        {
            "success": true,
            "message": "Cleared history for document",
            "document_id": "1abc123..."
        }
    """
    import json

    logger.info(f"[clear_doc_history] Doc={document_id}")

    if not document_id or not document_id.strip():
        return json.dumps({
            "success": False,
            "error": "document_id is required"
        }, indent=2)

    manager = get_history_manager()
    cleared = manager.clear_history(document_id)

    return json.dumps({
        "success": True,
        "message": "Cleared history for document" if cleared else "No history existed for document",
        "document_id": document_id
    }, indent=2)


@server.tool()
async def get_history_stats() -> str:
    """
    Get statistics about tracked operation history across all documents.

    This tool provides an overview of the history tracking system,
    including how many documents and operations are being tracked.

    Returns:
        str: JSON containing:
            - documents_tracked: Number of documents with tracked history
            - total_operations: Total operations across all documents
            - undone_operations: Number of operations that have been undone
            - operations_per_document: Dictionary mapping document IDs to operation counts

    Example Response:
        {
            "documents_tracked": 3,
            "total_operations": 25,
            "undone_operations": 5,
            "operations_per_document": {
                "1abc123...": 10,
                "2def456...": 8,
                "3ghi789...": 7
            }
        }
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

    This tool allows recording operations that need undo support. Call this
    BEFORE performing the operation to capture the original text that will
    be deleted or replaced.

    For most use cases, prefer using modify_doc_text with automatic tracking
    (coming soon). This tool is for advanced scenarios or custom integrations.

    Args:
        user_google_email: User's Google email address
        document_id: ID of the document
        operation_type: Type of operation. Supported values:
            - "insert_text": Inserting new text
            - "delete_text": Deleting existing text
            - "replace_text": Replacing text with new text
            - "format_text": Applying formatting changes
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

    Example - Recording a delete operation:
        record_doc_operation(
            document_id="1abc...",
            operation_type="delete_text",
            start_index=100,
            end_index=150
        )
        # This captures the text at positions 100-150 so it can be restored on undo

    Example - Recording an insert operation:
        record_doc_operation(
            document_id="1abc...",
            operation_type="insert_text",
            start_index=100,
            text="Hello World"
        )
        # Undo will delete the inserted text
    """
    import json

    logger.info(f"[record_doc_operation] Doc={document_id}, Type={operation_type}, Start={start_index}")

    # Validate inputs
    validator = ValidationManager()
    is_valid, structured_error = validator.validate_document_id_structured(document_id)
    if not is_valid:
        return structured_error

    valid_types = ["insert_text", "delete_text", "replace_text", "format_text"]
    if operation_type not in valid_types:
        return json.dumps({
            "success": False,
            "error": f"Invalid operation_type. Must be one of: {', '.join(valid_types)}"
        }, indent=2)

    if start_index < 0:
        return json.dumps({
            "success": False,
            "error": "start_index must be non-negative"
        }, indent=2)

    # Determine undo capability
    undo_capability = UndoCapability.FULL
    undo_notes = None
    deleted_text = None
    original_text = None

    # Capture text if needed for undo
    if capture_deleted_text and operation_type in ["delete_text", "replace_text"]:
        if end_index is None or end_index <= start_index:
            return json.dumps({
                "success": False,
                "error": "end_index is required and must be greater than start_index for delete/replace operations"
            }, indent=2)

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
        undo_notes = "Format undo requires capturing original formatting (not yet supported)"

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
        "message": f"Recorded {operation_type} operation"
    }

    if deleted_text is not None:
        result["deleted_text_preview"] = deleted_text[:100] + "..." if len(deleted_text) > 100 else deleted_text
        result["deleted_text_length"] = len(deleted_text)

    if original_text is not None:
        result["original_text_preview"] = original_text[:100] + "..." if len(original_text) > 100 else original_text
        result["original_text_length"] = len(original_text)

    return json.dumps(result, indent=2)
