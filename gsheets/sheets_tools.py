"""
Google Sheets MCP Tools

This module provides MCP tools for interacting with Google Sheets API.
"""

import logging
import asyncio
import json
import re
from typing import Any, Dict, List, Optional, Union


from auth.service_decorator import require_google_service
from core.server import server
from core.utils import handle_http_errors
from core.comments import create_comment_tools

# Configure module logger
logger = logging.getLogger(__name__)


@server.tool()
@handle_http_errors("list_spreadsheets", is_read_only=True, service_type="sheets")
@require_google_service("drive", "drive_read")
async def list_spreadsheets(
    service,
    user_google_email: str,
    max_results: int = 25,
) -> str:
    """
    Lists spreadsheets from Google Drive that the user has access to.

    Args:
        user_google_email (str): The user's Google email address. Required.
        max_results (int): Maximum number of spreadsheets to return. Defaults to 25.

    Returns:
        str: A formatted list of spreadsheet files (name, ID, modified time).
    """
    logger.info(f"[list_spreadsheets] Invoked. Email: '{user_google_email}'")

    files_response = await asyncio.to_thread(
        service.files()
        .list(
            q="mimeType='application/vnd.google-apps.spreadsheet'",
            pageSize=max_results,
            fields="files(id,name,modifiedTime,webViewLink)",
            orderBy="modifiedTime desc",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute
    )

    files = files_response.get("files", [])
    if not files:
        return f"No spreadsheets found for {user_google_email}."

    spreadsheets_list = [
        f'- "{file["name"]}" (ID: {file["id"]}) | Modified: {file.get("modifiedTime", "Unknown")} | Link: {file.get("webViewLink", "No link")}'
        for file in files
    ]

    text_output = (
        f"Successfully listed {len(files)} spreadsheets for {user_google_email}:\n"
        + "\n".join(spreadsheets_list)
    )

    logger.info(
        f"Successfully listed {len(files)} spreadsheets for {user_google_email}."
    )
    return text_output


@server.tool()
@handle_http_errors("get_spreadsheet_info", is_read_only=True, service_type="sheets")
@require_google_service("sheets", "sheets_read")
async def get_spreadsheet_info(
    service,
    user_google_email: str,
    spreadsheet_id: str,
) -> str:
    """
    Gets information about a specific spreadsheet including its sheets.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet to get info for. Required.

    Returns:
        str: Formatted spreadsheet information including title and sheets list.
    """
    logger.info(
        f"[get_spreadsheet_info] Invoked. Email: '{user_google_email}', Spreadsheet ID: {spreadsheet_id}"
    )

    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )

    title = spreadsheet.get("properties", {}).get("title", "Unknown")
    sheets = spreadsheet.get("sheets", [])

    sheets_info = []
    for sheet in sheets:
        sheet_props = sheet.get("properties", {})
        sheet_name = sheet_props.get("title", "Unknown")
        sheet_id = sheet_props.get("sheetId", "Unknown")
        grid_props = sheet_props.get("gridProperties", {})
        rows = grid_props.get("rowCount", "Unknown")
        cols = grid_props.get("columnCount", "Unknown")

        sheets_info.append(f'  - "{sheet_name}" (ID: {sheet_id}) | Size: {rows}x{cols}')

    text_output = (
        f'Spreadsheet: "{title}" (ID: {spreadsheet_id})\n'
        f"Sheets ({len(sheets)}):\n" + "\n".join(sheets_info)
        if sheets_info
        else "  No sheets found"
    )

    logger.info(
        f"Successfully retrieved info for spreadsheet {spreadsheet_id} for {user_google_email}."
    )
    return text_output


async def _get_sheet_name_by_id(service, spreadsheet_id: str, sheet_id: int) -> str:
    """
    Get the sheet name from a sheet ID.

    Args:
        service: Authenticated Sheets service
        spreadsheet_id: The spreadsheet ID
        sheet_id: Numeric ID of the sheet

    Returns:
        The sheet name (string)

    Raises:
        ValueError if sheet not found
    """
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )

    for sheet in spreadsheet.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("sheetId") == sheet_id:
            return props.get("title")

    available_sheets = [
        f"{s.get('properties', {}).get('title')} (ID: {s.get('properties', {}).get('sheetId')})"
        for s in spreadsheet.get("sheets", [])
    ]
    raise ValueError(
        f"Sheet with ID {sheet_id} not found. Available sheets: {available_sheets}"
    )


async def _build_full_range(
    service,
    spreadsheet_id: str,
    range_name: str,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Build a full range string with sheet reference if needed.

    If the range_name already contains a sheet reference (contains '!'),
    it is returned as-is. Otherwise, if sheet_name or sheet_id is provided,
    the sheet reference is prepended.

    Args:
        service: Authenticated Sheets service
        spreadsheet_id: The spreadsheet ID
        range_name: The cell range (e.g., "A1:D10" or "Sheet1!A1:D10")
        sheet_name: Optional sheet name
        sheet_id: Optional sheet ID (takes precedence over sheet_name)

    Returns:
        Full range string with sheet reference
    """
    # If range already has a sheet reference, use it as-is
    if "!" in range_name:
        return range_name

    # If sheet_id is provided, look up the sheet name
    if sheet_id is not None:
        resolved_sheet_name = await _get_sheet_name_by_id(
            service, spreadsheet_id, sheet_id
        )
    elif sheet_name is not None:
        resolved_sheet_name = sheet_name
    else:
        # No sheet specified, return range as-is (will use first sheet)
        return range_name

    # Quote sheet name if it contains spaces or special characters
    if " " in resolved_sheet_name or any(
        c in resolved_sheet_name for c in ["'", "!", ":", "[", "]"]
    ):
        escaped_name = resolved_sheet_name.replace("'", "''")
        return f"'{escaped_name}'!{range_name}"
    else:
        return f"{resolved_sheet_name}!{range_name}"


@server.tool()
@handle_http_errors("read_sheet_values", is_read_only=True, service_type="sheets")
@require_google_service("sheets", "sheets_read")
async def read_sheet_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str = "A1:Z1000",
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Reads values from a specific range in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to read (e.g., "A1:D10"). Can include sheet name prefix
            like "Sheet1!A1:D10" but this is optional if sheet_name or sheet_id is provided.
            Defaults to "A1:Z1000".
        sheet_name (Optional[str]): Name of the sheet to read from. If provided, the range_name
            only needs the cell range (e.g., "A1:D10" instead of "Sheet1!A1:D10").
            For sheet names with spaces, this avoids needing to escape quotes.
        sheet_id (Optional[int]): Numeric ID of the sheet to read from. Alternative to sheet_name.
            Takes precedence over sheet_name if both are provided.

    Returns:
        str: The formatted values from the specified range.
    """
    logger.info(
        f"[read_sheet_values] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}, Sheet: {sheet_name}, SheetId: {sheet_id}"
    )

    # Build the full range with sheet reference if sheet_name or sheet_id is provided
    full_range = await _build_full_range(
        service, spreadsheet_id, range_name, sheet_name, sheet_id
    )

    result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=full_range)
        .execute
    )

    values = result.get("values", [])
    if not values:
        return f"No data found in range '{range_name}' for {user_google_email}."

    # Format the output as a readable table
    formatted_rows = []
    for i, row in enumerate(values, 1):
        # Pad row with empty strings to show structure
        padded_row = row + [""] * max(0, len(values[0]) - len(row)) if values else row
        formatted_rows.append(f"Row {i:2d}: {padded_row}")

    text_output = (
        f"Successfully read {len(values)} rows from range '{range_name}' in spreadsheet {spreadsheet_id} for {user_google_email}:\n"
        + "\n".join(formatted_rows[:50])  # Limit to first 50 rows for readability
        + (f"\n... and {len(values) - 50} more rows" if len(values) > 50 else "")
    )

    logger.info(f"Successfully read {len(values)} rows for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("modify_sheet_values", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def modify_sheet_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    values: Optional[Union[str, List[List[str]]]] = None,
    value_input_option: str = "USER_ENTERED",
    clear_values: bool = False,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Modifies values in a specific range of a Google Sheet - can write, update, or clear values.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to modify (e.g., "A1:D10"). Can include sheet name prefix
            like "Sheet1!A1:D10" but this is optional if sheet_name or sheet_id is provided. Required.
        values (Optional[Union[str, List[List[str]]]]): 2D array of values to write/update. Can be a JSON string or Python list. Required unless clear_values=True.
        value_input_option (str): How to interpret input values ("RAW" or "USER_ENTERED"). Defaults to "USER_ENTERED".
        clear_values (bool): If True, clears the range instead of writing values. Defaults to False.
        sheet_name (Optional[str]): Name of the sheet to modify. If provided, the range_name
            only needs the cell range (e.g., "A1:D10" instead of "Sheet1!A1:D10").
            For sheet names with spaces, this avoids needing to escape quotes.
        sheet_id (Optional[int]): Numeric ID of the sheet to modify. Alternative to sheet_name.
            Takes precedence over sheet_name if both are provided.

    Returns:
        str: Confirmation message of the successful modification operation.
    """
    operation = "clear" if clear_values else "write"
    logger.info(
        f"[modify_sheet_values] Invoked. Operation: {operation}, Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}, Sheet: {sheet_name}, SheetId: {sheet_id}"
    )

    # Parse values if it's a JSON string (MCP passes parameters as JSON strings)
    if values is not None and isinstance(values, str):
        try:
            parsed_values = json.loads(values)
            if not isinstance(parsed_values, list):
                raise ValueError(
                    f"Values must be a list, got {type(parsed_values).__name__}"
                )
            # Validate it's a list of lists
            for i, row in enumerate(parsed_values):
                if not isinstance(row, list):
                    raise ValueError(
                        f"Row {i} must be a list, got {type(row).__name__}"
                    )
            values = parsed_values
            logger.info(
                f"[modify_sheet_values] Parsed JSON string to Python list with {len(values)} rows"
            )
        except json.JSONDecodeError as e:
            raise Exception(f"Invalid JSON format for values: {e}")
        except ValueError as e:
            raise Exception(f"Invalid values structure: {e}")

    if not clear_values and not values:
        raise Exception(
            "Either 'values' must be provided or 'clear_values' must be True."
        )

    # Build the full range with sheet reference if sheet_name or sheet_id is provided
    full_range = await _build_full_range(
        service, spreadsheet_id, range_name, sheet_name, sheet_id
    )

    if clear_values:
        result = await asyncio.to_thread(
            service.spreadsheets()
            .values()
            .clear(spreadsheetId=spreadsheet_id, range=full_range)
            .execute
        )

        cleared_range = result.get("clearedRange", range_name)
        text_output = f"Successfully cleared range '{cleared_range}' in spreadsheet {spreadsheet_id} for {user_google_email}."
        logger.info(
            f"Successfully cleared range '{cleared_range}' for {user_google_email}."
        )
    else:
        body = {"values": values}

        result = await asyncio.to_thread(
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=full_range,
                valueInputOption=value_input_option,
                body=body,
            )
            .execute
        )

        updated_cells = result.get("updatedCells", 0)
        updated_rows = result.get("updatedRows", 0)
        updated_columns = result.get("updatedColumns", 0)

        text_output = (
            f"Successfully updated range '{range_name}' in spreadsheet {spreadsheet_id} for {user_google_email}. "
            f"Updated: {updated_cells} cells, {updated_rows} rows, {updated_columns} columns."
        )
        logger.info(
            f"Successfully updated {updated_cells} cells for {user_google_email}."
        )

    return text_output


@server.tool()
@handle_http_errors("create_spreadsheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def create_spreadsheet(
    service,
    user_google_email: str,
    title: str,
    sheet_names: Optional[List[str]] = None,
) -> str:
    """
    Creates a new Google Spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        title (str): The title of the new spreadsheet. Required.
        sheet_names (Optional[List[str]]): List of sheet names to create. If not provided, creates one sheet with default name.

    Returns:
        str: Information about the newly created spreadsheet including ID and URL.
    """
    logger.info(
        f"[create_spreadsheet] Invoked. Email: '{user_google_email}', Title: {title}"
    )

    spreadsheet_body = {"properties": {"title": title}}

    if sheet_names:
        spreadsheet_body["sheets"] = [
            {"properties": {"title": sheet_name}} for sheet_name in sheet_names
        ]

    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().create(body=spreadsheet_body).execute
    )

    spreadsheet_id = spreadsheet.get("spreadsheetId")
    spreadsheet_url = spreadsheet.get("spreadsheetUrl")

    text_output = (
        f"Successfully created spreadsheet '{title}' for {user_google_email}. "
        f"ID: {spreadsheet_id} | URL: {spreadsheet_url}"
    )

    logger.info(
        f"Successfully created spreadsheet for {user_google_email}. ID: {spreadsheet_id}"
    )
    return text_output


@server.tool()
@handle_http_errors("create_sheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def create_sheet(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_name: str,
) -> str:
    """
    Creates a new sheet within an existing spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_name (str): The name of the new sheet. Required.

    Returns:
        str: Confirmation message of the successful sheet creation.
    """
    logger.info(
        f"[create_sheet] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Sheet: {sheet_name}"
    )

    request_body = {"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}

    response = await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    sheet_id = response["replies"][0]["addSheet"]["properties"]["sheetId"]

    text_output = f"Successfully created sheet '{sheet_name}' (ID: {sheet_id}) in spreadsheet {spreadsheet_id} for {user_google_email}."

    logger.info(
        f"Successfully created sheet for {user_google_email}. Sheet ID: {sheet_id}"
    )
    return text_output


# Create comment management tools for sheets
_comment_tools = create_comment_tools("spreadsheet", "spreadsheet_id")

# Extract and register the functions
read_sheet_comments = _comment_tools["read_comments"]
create_sheet_comment = _comment_tools["create_comment"]
reply_to_sheet_comment = _comment_tools["reply_to_comment"]
resolve_sheet_comment = _comment_tools["resolve_comment"]


def _parse_cell_reference(cell_ref: str) -> tuple[int, int]:
    """
    Parse a cell reference like 'A1' into (row_index, column_index).

    Args:
        cell_ref: Cell reference in A1 notation (e.g., 'A1', 'B2', 'AA10')

    Returns:
        Tuple of (row_index, column_index) - both 0-indexed
    """
    import re

    match = re.match(r"^([A-Za-z]+)(\d+)$", cell_ref.strip())
    if not match:
        raise ValueError(
            f"Invalid cell reference: '{cell_ref}'. Expected format like 'A1', 'B2', 'AA10'."
        )

    col_letters = match.group(1).upper()
    row_num = int(match.group(2))

    # Convert column letters to index (A=0, B=1, ..., Z=25, AA=26, ...)
    col_index = 0
    for char in col_letters:
        col_index = col_index * 26 + (ord(char) - ord("A") + 1)
    col_index -= 1  # Make 0-indexed

    row_index = row_num - 1  # Make 0-indexed

    return row_index, col_index


def _column_index_to_letter(col_index: int) -> str:
    """
    Convert a 0-indexed column number to column letters (0='A', 25='Z', 26='AA').
    """
    result = ""
    col_index += 1  # Make 1-indexed for calculation
    while col_index > 0:
        col_index -= 1
        result = chr(ord("A") + col_index % 26) + result
        col_index //= 26
    return result


async def _get_sheet_id_by_name(service, spreadsheet_id: str, sheet_name: str) -> int:
    """
    Get the sheet ID from the sheet name.

    Args:
        service: Authenticated Sheets service
        spreadsheet_id: The spreadsheet ID
        sheet_name: Name of the sheet

    Returns:
        The sheet ID (integer)

    Raises:
        ValueError if sheet not found
    """
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )

    for sheet in spreadsheet.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == sheet_name:
            return props.get("sheetId")

    available_sheets = [
        s.get("properties", {}).get("title") for s in spreadsheet.get("sheets", [])
    ]
    raise ValueError(
        f"Sheet '{sheet_name}' not found. Available sheets: {available_sheets}"
    )


@server.tool()
@handle_http_errors("read_cell_notes", is_read_only=True, service_type="sheets")
@require_google_service("sheets", "sheets_read")
async def read_cell_notes(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
) -> str:
    """
    Reads cell notes (yellow popup comments) from a range in a Google Sheet.

    Cell notes are lightweight annotations attached to individual cells that appear
    as yellow popups when you hover over a cell. These are different from threaded
    comments (use read_spreadsheet_comments for those).

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to read notes from (e.g., "Sheet1!A1:D10", "A1:B5"). Required.

    Returns:
        str: Formatted list of cells with notes, including cell reference and note content.
    """
    logger.info(
        f"[read_cell_notes] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
    )

    # Get spreadsheet data including notes
    result = await asyncio.to_thread(
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            ranges=[range_name],
            fields="sheets.data.rowData.values.note,sheets.properties.title",
        )
        .execute
    )

    sheets = result.get("sheets", [])
    if not sheets:
        return f"No data found in range '{range_name}'."

    sheet = sheets[0]
    sheet_name = sheet.get("properties", {}).get("title", "Unknown")
    data = sheet.get("data", [])

    if not data:
        return f"No data found in range '{range_name}'."

    notes_found = []
    grid_data = data[0]
    start_row = grid_data.get("startRow", 0)
    start_col = grid_data.get("startColumn", 0)

    row_data_list = grid_data.get("rowData", [])
    for row_offset, row_data in enumerate(row_data_list):
        values = row_data.get("values", [])
        for col_offset, cell in enumerate(values):
            note = cell.get("note")
            if note:
                row_num = start_row + row_offset + 1
                col_letter = _column_index_to_letter(start_col + col_offset)
                cell_ref = f"{col_letter}{row_num}"
                notes_found.append({"cell": cell_ref, "note": note})

    if not notes_found:
        return f"No cell notes found in range '{range_name}' of sheet '{sheet_name}'."

    output = [
        f"Found {len(notes_found)} cell notes in range '{range_name}' of sheet '{sheet_name}':\n"
    ]
    for item in notes_found:
        output.append(f"  Cell {item['cell']}: {item['note']}")

    logger.info(
        f"Successfully read {len(notes_found)} cell notes for {user_google_email}."
    )
    return "\n".join(output)


@server.tool()
@handle_http_errors("update_cell_note", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def update_cell_note(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    cell: str,
    note: str,
    sheet_name: Optional[str] = None,
) -> str:
    """
    Adds or updates a note on a specific cell in a Google Sheet.

    Cell notes are lightweight yellow popup annotations that appear when hovering
    over a cell. They are simpler than threaded comments - just plain text without
    threading, author tracking, or resolution status.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        cell (str): Cell reference in A1 notation (e.g., 'A1', 'B2', 'AA10'). Required.
        note (str): The note text to add to the cell. Required.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses the first sheet.

    Returns:
        str: Confirmation message of the successful note update.
    """
    logger.info(
        f"[update_cell_note] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Cell: {cell}, Sheet: {sheet_name}"
    )

    # Parse cell reference
    row_index, col_index = _parse_cell_reference(cell)

    # Get sheet ID
    if sheet_name:
        sheet_id = await _get_sheet_id_by_name(service, spreadsheet_id, sheet_name)
    else:
        # Get first sheet ID
        spreadsheet = await asyncio.to_thread(
            service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
        )
        sheets = spreadsheet.get("sheets", [])
        if not sheets:
            raise ValueError("Spreadsheet has no sheets.")
        sheet_id = sheets[0].get("properties", {}).get("sheetId")
        sheet_name = sheets[0].get("properties", {}).get("title", "Sheet1")

    # Build the update request
    request_body = {
        "requests": [
            {
                "updateCells": {
                    "rows": [{"values": [{"note": note}]}],
                    "fields": "note",
                    "start": {
                        "sheetId": sheet_id,
                        "rowIndex": row_index,
                        "columnIndex": col_index,
                    },
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    text_output = (
        f"Successfully added/updated note on cell {cell} in sheet '{sheet_name}' "
        f"of spreadsheet {spreadsheet_id} for {user_google_email}.\n"
        f"Note content: {note}"
    )

    logger.info(f"Successfully updated cell note for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("clear_cell_note", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def clear_cell_note(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    cell: str,
    sheet_name: Optional[str] = None,
) -> str:
    """
    Removes a note from a specific cell in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        cell (str): Cell reference in A1 notation (e.g., 'A1', 'B2', 'AA10'). Required.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses the first sheet.

    Returns:
        str: Confirmation message of the successful note removal.
    """
    logger.info(
        f"[clear_cell_note] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Cell: {cell}, Sheet: {sheet_name}"
    )

    # Parse cell reference
    row_index, col_index = _parse_cell_reference(cell)

    # Get sheet ID
    if sheet_name:
        sheet_id = await _get_sheet_id_by_name(service, spreadsheet_id, sheet_name)
    else:
        # Get first sheet ID
        spreadsheet = await asyncio.to_thread(
            service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
        )
        sheets = spreadsheet.get("sheets", [])
        if not sheets:
            raise ValueError("Spreadsheet has no sheets.")
        sheet_id = sheets[0].get("properties", {}).get("sheetId")
        sheet_name = sheets[0].get("properties", {}).get("title", "Sheet1")

    # Build the update request - setting note to empty string clears it
    request_body = {
        "requests": [
            {
                "updateCells": {
                    "rows": [{"values": [{"note": ""}]}],
                    "fields": "note",
                    "start": {
                        "sheetId": sheet_id,
                        "rowIndex": row_index,
                        "columnIndex": col_index,
                    },
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    text_output = (
        f"Successfully cleared note from cell {cell} in sheet '{sheet_name}' "
        f"of spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully cleared cell note for {user_google_email}.")
    return text_output


def _parse_range_to_grid(range_str: str) -> Dict[str, int]:
    """
    Parse a range string like 'A1:D10' into GridRange coordinates.

    Args:
        range_str: Range in A1 notation (e.g., 'A1:D10', 'B2:C5')

    Returns:
        Dict with startRowIndex, endRowIndex, startColumnIndex, endColumnIndex (all 0-indexed)
    """
    # Handle single cell vs range
    if ":" in range_str:
        start_cell, end_cell = range_str.split(":")
    else:
        start_cell = end_cell = range_str

    start_row, start_col = _parse_cell_reference(start_cell)
    end_row, end_col = _parse_cell_reference(end_cell)

    return {
        "startRowIndex": start_row,
        "endRowIndex": end_row + 1,  # API uses exclusive end
        "startColumnIndex": start_col,
        "endColumnIndex": end_col + 1,  # API uses exclusive end
    }


def _parse_color(color_str: str) -> Dict[str, float]:
    """
    Parse a color string (hex or name) into RGBA dict for Sheets API.

    Args:
        color_str: Color in hex format (#RRGGBB or #RGB) or common color name

    Returns:
        Dict with red, green, blue values (0-1 float range)
    """
    color_names = {
        "black": "#000000",
        "white": "#FFFFFF",
        "red": "#FF0000",
        "green": "#00FF00",
        "blue": "#0000FF",
        "yellow": "#FFFF00",
        "cyan": "#00FFFF",
        "magenta": "#FF00FF",
        "orange": "#FFA500",
        "purple": "#800080",
        "pink": "#FFC0CB",
        "gray": "#808080",
        "grey": "#808080",
        "lightgray": "#D3D3D3",
        "lightgrey": "#D3D3D3",
        "darkgray": "#A9A9A9",
        "darkgrey": "#A9A9A9",
    }

    # Convert color name to hex
    color_lower = color_str.lower().strip()
    if color_lower in color_names:
        color_str = color_names[color_lower]

    # Parse hex color
    hex_match = re.match(r"^#?([0-9A-Fa-f]{6})$", color_str)
    if hex_match:
        hex_val = hex_match.group(1)
        return {
            "red": int(hex_val[0:2], 16) / 255.0,
            "green": int(hex_val[2:4], 16) / 255.0,
            "blue": int(hex_val[4:6], 16) / 255.0,
        }

    # Try short hex (#RGB)
    short_hex_match = re.match(r"^#?([0-9A-Fa-f]{3})$", color_str)
    if short_hex_match:
        hex_val = short_hex_match.group(1)
        return {
            "red": int(hex_val[0] * 2, 16) / 255.0,
            "green": int(hex_val[1] * 2, 16) / 255.0,
            "blue": int(hex_val[2] * 2, 16) / 255.0,
        }

    raise ValueError(
        f"Invalid color format: '{color_str}'. Use hex (#RRGGBB or #RGB) or color name (red, blue, etc.)"
    )


async def _resolve_sheet_id(
    service, spreadsheet_id: str, sheet_name: Optional[str], sheet_id: Optional[int]
) -> int:
    """
    Resolve to a sheet ID, looking up from name if needed.

    Args:
        service: Authenticated Sheets service
        spreadsheet_id: The spreadsheet ID
        sheet_name: Optional sheet name
        sheet_id: Optional sheet ID (takes precedence)

    Returns:
        The resolved sheet ID
    """
    if sheet_id is not None:
        return sheet_id

    if sheet_name is not None:
        return await _get_sheet_id_by_name(service, spreadsheet_id, sheet_name)

    # Get first sheet
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )
    sheets = spreadsheet.get("sheets", [])
    if not sheets:
        raise ValueError("Spreadsheet has no sheets.")
    return sheets[0].get("properties", {}).get("sheetId")


@server.tool()
@handle_http_errors("format_cells", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def format_cells(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    strikethrough: Optional[bool] = None,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
    font_color: Optional[str] = None,
    background_color: Optional[str] = None,
    horizontal_alignment: Optional[str] = None,
    vertical_alignment: Optional[str] = None,
    wrap_strategy: Optional[str] = None,
    number_format_type: Optional[str] = None,
    number_format_pattern: Optional[str] = None,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Applies formatting to a range of cells in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to format (e.g., "A1:D10", "B2:C5"). Required.
        bold (Optional[bool]): Set text bold.
        italic (Optional[bool]): Set text italic.
        underline (Optional[bool]): Set text underline.
        strikethrough (Optional[bool]): Set text strikethrough.
        font_size (Optional[int]): Font size in points (e.g., 10, 12, 14).
        font_family (Optional[str]): Font family name (e.g., "Arial", "Times New Roman", "Courier New").
        font_color (Optional[str]): Text color - hex (#RRGGBB) or name (red, blue, etc.).
        background_color (Optional[str]): Cell background color - hex (#RRGGBB) or name (red, blue, etc.).
        horizontal_alignment (Optional[str]): Horizontal alignment - LEFT, CENTER, or RIGHT.
        vertical_alignment (Optional[str]): Vertical alignment - TOP, MIDDLE, or BOTTOM.
        wrap_strategy (Optional[str]): Text wrapping - OVERFLOW_CELL, CLIP, or WRAP.
        number_format_type (Optional[str]): Number format type - TEXT, NUMBER, CURRENCY, PERCENT, DATE, TIME, DATE_TIME, SCIENTIFIC.
        number_format_pattern (Optional[str]): Custom number format pattern (e.g., "#,##0.00", "$#,##0", "yyyy-mm-dd").
        sheet_name (Optional[str]): Name of the sheet. If not provided with range, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful formatting operation.
    """
    logger.info(
        f"[format_cells] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
    )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(range_name)
    grid_range["sheetId"] = resolved_sheet_id

    # Build cell format
    cell_format: Dict[str, Any] = {}
    fields = []

    # Text format properties
    text_format: Dict[str, Any] = {}
    if bold is not None:
        text_format["bold"] = bold
        fields.append("userEnteredFormat.textFormat.bold")
    if italic is not None:
        text_format["italic"] = italic
        fields.append("userEnteredFormat.textFormat.italic")
    if underline is not None:
        text_format["underline"] = underline
        fields.append("userEnteredFormat.textFormat.underline")
    if strikethrough is not None:
        text_format["strikethrough"] = strikethrough
        fields.append("userEnteredFormat.textFormat.strikethrough")
    if font_size is not None:
        text_format["fontSize"] = font_size
        fields.append("userEnteredFormat.textFormat.fontSize")
    if font_family is not None:
        text_format["fontFamily"] = font_family
        fields.append("userEnteredFormat.textFormat.fontFamily")
    if font_color is not None:
        text_format["foregroundColor"] = _parse_color(font_color)
        fields.append("userEnteredFormat.textFormat.foregroundColor")

    if text_format:
        cell_format["textFormat"] = text_format

    # Background color
    if background_color is not None:
        cell_format["backgroundColor"] = _parse_color(background_color)
        fields.append("userEnteredFormat.backgroundColor")

    # Alignment
    if horizontal_alignment is not None:
        valid_h_align = ["LEFT", "CENTER", "RIGHT"]
        h_align_upper = horizontal_alignment.upper()
        if h_align_upper not in valid_h_align:
            raise ValueError(
                f"Invalid horizontal_alignment: '{horizontal_alignment}'. Must be one of: {valid_h_align}"
            )
        cell_format["horizontalAlignment"] = h_align_upper
        fields.append("userEnteredFormat.horizontalAlignment")

    if vertical_alignment is not None:
        valid_v_align = ["TOP", "MIDDLE", "BOTTOM"]
        v_align_upper = vertical_alignment.upper()
        if v_align_upper not in valid_v_align:
            raise ValueError(
                f"Invalid vertical_alignment: '{vertical_alignment}'. Must be one of: {valid_v_align}"
            )
        cell_format["verticalAlignment"] = v_align_upper
        fields.append("userEnteredFormat.verticalAlignment")

    # Text wrapping
    if wrap_strategy is not None:
        valid_wrap = ["OVERFLOW_CELL", "CLIP", "WRAP"]
        wrap_upper = wrap_strategy.upper()
        if wrap_upper not in valid_wrap:
            raise ValueError(
                f"Invalid wrap_strategy: '{wrap_strategy}'. Must be one of: {valid_wrap}"
            )
        cell_format["wrapStrategy"] = wrap_upper
        fields.append("userEnteredFormat.wrapStrategy")

    # Number format
    if number_format_type is not None or number_format_pattern is not None:
        number_format: Dict[str, Any] = {}
        if number_format_type is not None:
            valid_types = [
                "TEXT",
                "NUMBER",
                "CURRENCY",
                "PERCENT",
                "DATE",
                "TIME",
                "DATE_TIME",
                "SCIENTIFIC",
            ]
            type_upper = number_format_type.upper()
            if type_upper not in valid_types:
                raise ValueError(
                    f"Invalid number_format_type: '{number_format_type}'. Must be one of: {valid_types}"
                )
            number_format["type"] = type_upper
        if number_format_pattern is not None:
            number_format["pattern"] = number_format_pattern
        cell_format["numberFormat"] = number_format
        fields.append("userEnteredFormat.numberFormat")

    if not fields:
        return "No formatting options specified. Please provide at least one formatting option."

    # Build request
    request_body = {
        "requests": [
            {
                "repeatCell": {
                    "range": grid_range,
                    "cell": {"userEnteredFormat": cell_format},
                    "fields": ",".join(fields),
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    # Build summary of applied formatting
    format_summary = []
    if bold is not None:
        format_summary.append(f"bold={bold}")
    if italic is not None:
        format_summary.append(f"italic={italic}")
    if underline is not None:
        format_summary.append(f"underline={underline}")
    if strikethrough is not None:
        format_summary.append(f"strikethrough={strikethrough}")
    if font_size is not None:
        format_summary.append(f"font_size={font_size}")
    if font_family is not None:
        format_summary.append(f"font_family={font_family}")
    if font_color is not None:
        format_summary.append(f"font_color={font_color}")
    if background_color is not None:
        format_summary.append(f"background_color={background_color}")
    if horizontal_alignment is not None:
        format_summary.append(f"horizontal_alignment={horizontal_alignment}")
    if vertical_alignment is not None:
        format_summary.append(f"vertical_alignment={vertical_alignment}")
    if wrap_strategy is not None:
        format_summary.append(f"wrap_strategy={wrap_strategy}")
    if number_format_type is not None:
        format_summary.append(f"number_format_type={number_format_type}")
    if number_format_pattern is not None:
        format_summary.append(f"number_format_pattern={number_format_pattern}")

    text_output = (
        f"Successfully formatted range '{range_name}' in spreadsheet {spreadsheet_id} for {user_google_email}.\n"
        f"Applied formatting: {', '.join(format_summary)}"
    )

    logger.info(f"Successfully formatted cells for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("merge_cells", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def merge_cells(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    merge_type: str = "MERGE_ALL",
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Merges cells in a range in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to merge (e.g., "A1:D4", "B2:C3"). Required.
        merge_type (str): Type of merge - MERGE_ALL (single merged cell), MERGE_COLUMNS (merge columns, keep rows separate), or MERGE_ROWS (merge rows, keep columns separate). Defaults to MERGE_ALL.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful merge operation.
    """
    logger.info(
        f"[merge_cells] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
    )

    # Validate merge type
    valid_merge_types = ["MERGE_ALL", "MERGE_COLUMNS", "MERGE_ROWS"]
    merge_type_upper = merge_type.upper()
    if merge_type_upper not in valid_merge_types:
        raise ValueError(
            f"Invalid merge_type: '{merge_type}'. Must be one of: {valid_merge_types}"
        )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(range_name)
    grid_range["sheetId"] = resolved_sheet_id

    # Build request
    request_body = {
        "requests": [
            {"mergeCells": {"range": grid_range, "mergeType": merge_type_upper}}
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    text_output = (
        f"Successfully merged cells in range '{range_name}' with type '{merge_type_upper}' "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully merged cells for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("unmerge_cells", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def unmerge_cells(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Unmerges any merged cells in a range in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to unmerge (e.g., "A1:D4", "B2:C3"). Required.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful unmerge operation.
    """
    logger.info(
        f"[unmerge_cells] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
    )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(range_name)
    grid_range["sheetId"] = resolved_sheet_id

    # Build request
    request_body = {"requests": [{"unmergeCells": {"range": grid_range}}]}

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    text_output = (
        f"Successfully unmerged cells in range '{range_name}' "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully unmerged cells for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("set_frozen_rows_columns", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def set_frozen_rows_columns(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    frozen_rows: Optional[int] = None,
    frozen_columns: Optional[int] = None,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Sets the number of frozen rows and/or columns in a Google Sheet.

    Frozen rows stay visible at the top when scrolling down.
    Frozen columns stay visible on the left when scrolling right.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        frozen_rows (Optional[int]): Number of rows to freeze from the top. Use 0 to unfreeze all rows.
        frozen_columns (Optional[int]): Number of columns to freeze from the left. Use 0 to unfreeze all columns.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful freeze operation.
    """
    logger.info(
        f"[set_frozen_rows_columns] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Rows: {frozen_rows}, Cols: {frozen_columns}"
    )

    if frozen_rows is None and frozen_columns is None:
        return "No freeze options specified. Please provide frozen_rows and/or frozen_columns."

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Build properties and fields
    grid_properties: Dict[str, Any] = {}
    fields = []

    if frozen_rows is not None:
        if frozen_rows < 0:
            raise ValueError("frozen_rows must be 0 or greater")
        grid_properties["frozenRowCount"] = frozen_rows
        fields.append("gridProperties.frozenRowCount")

    if frozen_columns is not None:
        if frozen_columns < 0:
            raise ValueError("frozen_columns must be 0 or greater")
        grid_properties["frozenColumnCount"] = frozen_columns
        fields.append("gridProperties.frozenColumnCount")

    # Build request
    request_body = {
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": resolved_sheet_id,
                        "gridProperties": grid_properties,
                    },
                    "fields": ",".join(fields),
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    # Build summary
    summary_parts = []
    if frozen_rows is not None:
        summary_parts.append(
            f"{frozen_rows} row(s) frozen" if frozen_rows > 0 else "rows unfrozen"
        )
    if frozen_columns is not None:
        summary_parts.append(
            f"{frozen_columns} column(s) frozen"
            if frozen_columns > 0
            else "columns unfrozen"
        )

    text_output = (
        f"Successfully updated freeze settings in spreadsheet {spreadsheet_id} for {user_google_email}: "
        + ", ".join(summary_parts)
    )

    logger.info(f"Successfully set frozen rows/columns for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("set_column_width", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def set_column_width(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    start_column: str,
    end_column: Optional[str] = None,
    width: int = 100,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Sets the width of one or more columns in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        start_column (str): The starting column letter (e.g., "A", "B", "AA"). Required.
        end_column (Optional[str]): The ending column letter (inclusive). If not provided, only start_column is resized.
        width (int): The width in pixels. Defaults to 100.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful column width update.
    """
    logger.info(
        f"[set_column_width] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Column: {start_column}-{end_column}, Width: {width}"
    )

    if width < 0:
        raise ValueError("width must be 0 or greater")

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Parse column letters to indices
    start_col_ref = f"{start_column}1"
    _, start_col_idx = _parse_cell_reference(start_col_ref)

    if end_column:
        end_col_ref = f"{end_column}1"
        _, end_col_idx = _parse_cell_reference(end_col_ref)
    else:
        end_col_idx = start_col_idx

    # Build request
    request_body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": resolved_sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": start_col_idx,
                        "endIndex": end_col_idx + 1,  # Exclusive end
                    },
                    "properties": {"pixelSize": width},
                    "fields": "pixelSize",
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    col_range = f"{start_column}-{end_column}" if end_column else start_column
    text_output = (
        f"Successfully set column(s) {col_range} width to {width}px "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully set column width for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("set_row_height", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def set_row_height(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    start_row: int,
    end_row: Optional[int] = None,
    height: int = 21,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Sets the height of one or more rows in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        start_row (int): The starting row number (1-indexed, e.g., 1 for the first row). Required.
        end_row (Optional[int]): The ending row number (inclusive, 1-indexed). If not provided, only start_row is resized.
        height (int): The height in pixels. Defaults to 21 (standard row height).
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful row height update.
    """
    logger.info(
        f"[set_row_height] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Row: {start_row}-{end_row}, Height: {height}"
    )

    if height < 0:
        raise ValueError("height must be 0 or greater")
    if start_row < 1:
        raise ValueError("start_row must be 1 or greater (1-indexed)")
    if end_row is not None and end_row < start_row:
        raise ValueError("end_row must be greater than or equal to start_row")

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Convert to 0-indexed
    start_row_idx = start_row - 1
    end_row_idx = (end_row - 1) if end_row else start_row_idx

    # Build request
    request_body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": resolved_sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_row_idx,
                        "endIndex": end_row_idx + 1,  # Exclusive end
                    },
                    "properties": {"pixelSize": height},
                    "fields": "pixelSize",
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    row_range = f"{start_row}-{end_row}" if end_row else str(start_row)
    text_output = (
        f"Successfully set row(s) {row_range} height to {height}px "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully set row height for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("add_conditional_formatting", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def add_conditional_formatting(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    rule_type: str,
    condition_type: Optional[str] = None,
    condition_values: Optional[List[str]] = None,
    background_color: Optional[str] = None,
    font_color: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    min_color: Optional[str] = None,
    mid_color: Optional[str] = None,
    max_color: Optional[str] = None,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Adds a conditional formatting rule to a range in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to apply formatting to (e.g., "A1:D10"). Required.
        rule_type (str): Type of rule - BOOLEAN (condition-based) or GRADIENT (color scale). Required.
        condition_type (Optional[str]): For BOOLEAN rules - type of condition: NUMBER_GREATER, NUMBER_LESS, NUMBER_EQ, NUMBER_BETWEEN, TEXT_CONTAINS, TEXT_NOT_CONTAINS, TEXT_STARTS_WITH, TEXT_ENDS_WITH, BLANK, NOT_BLANK, CUSTOM_FORMULA.
        condition_values (Optional[List[str]]): Values for the condition. For NUMBER_BETWEEN provide [min, max]. For CUSTOM_FORMULA provide the formula.
        background_color (Optional[str]): Background color for BOOLEAN rules - hex (#RRGGBB) or color name.
        font_color (Optional[str]): Font color for BOOLEAN rules - hex (#RRGGBB) or color name.
        bold (Optional[bool]): Apply bold for BOOLEAN rules.
        italic (Optional[bool]): Apply italic for BOOLEAN rules.
        min_color (Optional[str]): Color for minimum value in GRADIENT rules. Defaults to green.
        mid_color (Optional[str]): Color for midpoint in GRADIENT rules (optional).
        max_color (Optional[str]): Color for maximum value in GRADIENT rules. Defaults to red.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful conditional formatting rule creation.
    """
    logger.info(
        f"[add_conditional_formatting] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}, Type: {rule_type}"
    )

    # Parse condition_values if it's a JSON string (MCP passes parameters as JSON strings)
    if condition_values is not None and isinstance(condition_values, str):
        try:
            parsed_values = json.loads(condition_values)
            if not isinstance(parsed_values, list):
                raise ValueError(
                    f"condition_values must be a list, got {type(parsed_values).__name__}"
                )
            condition_values = parsed_values
            logger.info(
                f"[add_conditional_formatting] Parsed JSON string to Python list with {len(condition_values)} values"
            )
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON format for condition_values: {e}")

    # Validate rule type
    rule_type_upper = rule_type.upper()
    if rule_type_upper not in ["BOOLEAN", "GRADIENT"]:
        raise ValueError(
            f"Invalid rule_type: '{rule_type}'. Must be BOOLEAN or GRADIENT."
        )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(range_name)
    grid_range["sheetId"] = resolved_sheet_id

    # Build the rule based on type
    rule: Dict[str, Any] = {"ranges": [grid_range]}

    if rule_type_upper == "BOOLEAN":
        if not condition_type:
            raise ValueError("condition_type is required for BOOLEAN rules")

        # Valid condition types
        valid_conditions = [
            "NUMBER_GREATER",
            "NUMBER_GREATER_THAN_EQ",
            "NUMBER_LESS",
            "NUMBER_LESS_THAN_EQ",
            "NUMBER_EQ",
            "NUMBER_NOT_EQ",
            "NUMBER_BETWEEN",
            "NUMBER_NOT_BETWEEN",
            "TEXT_CONTAINS",
            "TEXT_NOT_CONTAINS",
            "TEXT_STARTS_WITH",
            "TEXT_ENDS_WITH",
            "TEXT_EQ",
            "BLANK",
            "NOT_BLANK",
            "CUSTOM_FORMULA",
        ]
        condition_type_upper = condition_type.upper()
        if condition_type_upper not in valid_conditions:
            raise ValueError(
                f"Invalid condition_type: '{condition_type}'. Must be one of: {valid_conditions}"
            )

        # Build condition
        condition: Dict[str, Any] = {"type": condition_type_upper}
        if condition_values:
            condition["values"] = [{"userEnteredValue": v} for v in condition_values]

        # Build format
        cell_format: Dict[str, Any] = {}
        if background_color:
            cell_format["backgroundColor"] = _parse_color(background_color)
        if font_color or bold is not None or italic is not None:
            text_format: Dict[str, Any] = {}
            if font_color:
                text_format["foregroundColor"] = _parse_color(font_color)
            if bold is not None:
                text_format["bold"] = bold
            if italic is not None:
                text_format["italic"] = italic
            cell_format["textFormat"] = text_format

        rule["booleanRule"] = {"condition": condition, "format": cell_format}

    else:  # GRADIENT
        # Build gradient rule
        gradient_rule: Dict[str, Any] = {
            "minpoint": {
                "type": "MIN",
                "color": _parse_color(min_color or "green"),
            },
            "maxpoint": {
                "type": "MAX",
                "color": _parse_color(max_color or "red"),
            },
        }
        if mid_color:
            gradient_rule["midpoint"] = {
                "type": "PERCENTILE",
                "value": "50",
                "color": _parse_color(mid_color),
            }

        rule["gradientRule"] = gradient_rule

    # Build request
    request_body = {
        "requests": [{"addConditionalFormatRule": {"rule": rule, "index": 0}}]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    text_output = (
        f"Successfully added {rule_type_upper} conditional formatting rule to range '{range_name}' "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully added conditional formatting for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("delete_sheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def delete_sheet(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Deletes a sheet from a spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_name (Optional[str]): Name of the sheet to delete. Either sheet_name or sheet_id is required.
        sheet_id (Optional[int]): Numeric ID of the sheet to delete. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful sheet deletion.
    """
    logger.info(
        f"[delete_sheet] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Sheet: {sheet_name}, SheetId: {sheet_id}"
    )

    if sheet_name is None and sheet_id is None:
        raise ValueError("Either sheet_name or sheet_id must be provided.")

    # Resolve sheet ID and get the sheet name for confirmation message
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Get sheet name for confirmation if we only have ID
    if sheet_name is None:
        resolved_sheet_name = await _get_sheet_name_by_id(
            service, spreadsheet_id, resolved_sheet_id
        )
    else:
        resolved_sheet_name = sheet_name

    # Build the delete request
    request_body = {"requests": [{"deleteSheet": {"sheetId": resolved_sheet_id}}]}

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    text_output = (
        f"Successfully deleted sheet '{resolved_sheet_name}' (ID: {resolved_sheet_id}) "
        f"from spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully deleted sheet for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("rename_sheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def rename_sheet(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    new_name: str,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Renames a sheet in a spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        new_name (str): The new name for the sheet. Required.
        sheet_name (Optional[str]): Current name of the sheet to rename. Either sheet_name or sheet_id is required.
        sheet_id (Optional[int]): Numeric ID of the sheet to rename. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful sheet rename.
    """
    logger.info(
        f"[rename_sheet] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Sheet: {sheet_name}, SheetId: {sheet_id}, NewName: {new_name}"
    )

    if sheet_name is None and sheet_id is None:
        raise ValueError("Either sheet_name or sheet_id must be provided.")

    if not new_name or not new_name.strip():
        raise ValueError("new_name cannot be empty.")

    # Resolve sheet ID and get the old sheet name for confirmation message
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Get old sheet name for confirmation if we only have ID
    if sheet_name is None:
        old_sheet_name = await _get_sheet_name_by_id(
            service, spreadsheet_id, resolved_sheet_id
        )
    else:
        old_sheet_name = sheet_name

    # Build the rename request using updateSheetProperties
    request_body = {
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": resolved_sheet_id, "title": new_name},
                    "fields": "title",
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    text_output = (
        f"Successfully renamed sheet from '{old_sheet_name}' to '{new_name}' "
        f"(ID: {resolved_sheet_id}) in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully renamed sheet for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("set_borders", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def set_borders(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    top: Optional[bool] = None,
    bottom: Optional[bool] = None,
    left: Optional[bool] = None,
    right: Optional[bool] = None,
    inner_horizontal: Optional[bool] = None,
    inner_vertical: Optional[bool] = None,
    border_style: str = "SOLID",
    border_color: Optional[str] = None,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Applies borders to a range of cells in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to apply borders to (e.g., "A1:D10", "B2:C5"). Required.
        top (Optional[bool]): Apply border to the top edge of the range.
        bottom (Optional[bool]): Apply border to the bottom edge of the range.
        left (Optional[bool]): Apply border to the left edge of the range.
        right (Optional[bool]): Apply border to the right edge of the range.
        inner_horizontal (Optional[bool]): Apply horizontal borders between cells inside the range.
        inner_vertical (Optional[bool]): Apply vertical borders between cells inside the range.
        border_style (str): Style of the border - DOTTED, DASHED, SOLID, SOLID_MEDIUM, SOLID_THICK, DOUBLE, or NONE. Defaults to SOLID. Use NONE to remove a border.
        border_color (Optional[str]): Border color - hex (#RRGGBB) or color name (red, blue, etc.). Defaults to black.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful border operation.
    """
    logger.info(
        f"[set_borders] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
    )

    # Validate border style
    valid_styles = [
        "DOTTED",
        "DASHED",
        "SOLID",
        "SOLID_MEDIUM",
        "SOLID_THICK",
        "DOUBLE",
        "NONE",
    ]
    border_style_upper = border_style.upper()
    if border_style_upper not in valid_styles:
        raise ValueError(
            f"Invalid border_style: '{border_style}'. Must be one of: {valid_styles}"
        )

    # Check if at least one border position is specified
    if all(
        v is None for v in [top, bottom, left, right, inner_horizontal, inner_vertical]
    ):
        return "No border positions specified. Please set at least one of: top, bottom, left, right, inner_horizontal, inner_vertical."

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(range_name)
    grid_range["sheetId"] = resolved_sheet_id

    # Build border style object
    def build_border(enabled: Optional[bool]) -> Optional[Dict[str, Any]]:
        if enabled is None:
            return None
        if not enabled:
            # Explicitly set to False means remove the border
            return {"style": "NONE"}
        border: Dict[str, Any] = {"style": border_style_upper}
        if border_color:
            border["color"] = _parse_color(border_color)
        else:
            # Default to black
            border["color"] = {"red": 0, "green": 0, "blue": 0}
        return border

    # Build the updateBorders request
    update_borders_request: Dict[str, Any] = {"range": grid_range}

    if top is not None:
        update_borders_request["top"] = build_border(top)
    if bottom is not None:
        update_borders_request["bottom"] = build_border(bottom)
    if left is not None:
        update_borders_request["left"] = build_border(left)
    if right is not None:
        update_borders_request["right"] = build_border(right)
    if inner_horizontal is not None:
        update_borders_request["innerHorizontal"] = build_border(inner_horizontal)
    if inner_vertical is not None:
        update_borders_request["innerVertical"] = build_border(inner_vertical)

    # Build request
    request_body = {"requests": [{"updateBorders": update_borders_request}]}

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    # Build summary of applied borders
    border_positions = []
    if top is not None:
        border_positions.append(f"top={'on' if top else 'off'}")
    if bottom is not None:
        border_positions.append(f"bottom={'on' if bottom else 'off'}")
    if left is not None:
        border_positions.append(f"left={'on' if left else 'off'}")
    if right is not None:
        border_positions.append(f"right={'on' if right else 'off'}")
    if inner_horizontal is not None:
        border_positions.append(
            f"inner_horizontal={'on' if inner_horizontal else 'off'}"
        )
    if inner_vertical is not None:
        border_positions.append(f"inner_vertical={'on' if inner_vertical else 'off'}")

    style_info = f"style={border_style_upper}"
    if border_color:
        style_info += f", color={border_color}"

    text_output = (
        f"Successfully applied borders to range '{range_name}' in spreadsheet {spreadsheet_id} for {user_google_email}.\n"
        f"Borders: {', '.join(border_positions)} ({style_info})"
    )

    logger.info(f"Successfully set borders for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("sort_range", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def sort_range(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    sort_column: str,
    ascending: bool = True,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Sorts data in a range by a specified column.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to sort (e.g., "A1:D10"). Required.
        sort_column (str): The column letter to sort by (e.g., "A", "B", "AA"). Required.
        ascending (bool): Sort in ascending order if True, descending if False. Defaults to True.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful sort operation.
    """
    logger.info(
        f"[sort_range] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}, Column: {sort_column}"
    )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Parse range to get grid coordinates
    grid_range = _parse_range_to_grid(range_name)
    grid_range["sheetId"] = resolved_sheet_id

    # Parse the sort column letter to get its index
    sort_col_ref = f"{sort_column}1"
    _, sort_col_idx = _parse_cell_reference(sort_col_ref)

    # Validate that sort column is within the range
    if sort_col_idx < grid_range["startColumnIndex"]:
        raise ValueError(
            f"Sort column '{sort_column}' (index {sort_col_idx}) is before the range start column."
        )
    if sort_col_idx >= grid_range["endColumnIndex"]:
        raise ValueError(
            f"Sort column '{sort_column}' (index {sort_col_idx}) is after the range end column."
        )

    # Build the sort request
    request_body = {
        "requests": [
            {
                "sortRange": {
                    "range": grid_range,
                    "sortSpecs": [
                        {
                            "dimensionIndex": sort_col_idx,
                            "sortOrder": "ASCENDING" if ascending else "DESCENDING",
                        }
                    ],
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    sort_order = "ascending" if ascending else "descending"
    text_output = (
        f"Successfully sorted range '{range_name}' by column '{sort_column}' ({sort_order}) "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully sorted range for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("find_and_replace", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def find_and_replace(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    find_text: str,
    replace_text: str,
    range_name: Optional[str] = None,
    match_case: bool = False,
    match_entire_cell: bool = False,
    use_regex: bool = False,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
    all_sheets: bool = False,
) -> str:
    """
    Finds and replaces text in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        find_text (str): The text to find. Required.
        replace_text (str): The text to replace with. Required.
        range_name (Optional[str]): The range to search in (e.g., "A1:D10"). If not provided, searches entire sheet.
        match_case (bool): If True, performs case-sensitive matching. Defaults to False.
        match_entire_cell (bool): If True, only matches cells that contain exactly the find_text. Defaults to False.
        use_regex (bool): If True, treats find_text as a regular expression. Defaults to False.
        sheet_name (Optional[str]): Name of the sheet to search. If not provided and all_sheets=False, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.
        all_sheets (bool): If True, searches all sheets in the spreadsheet. Defaults to False.

    Returns:
        str: Summary of replacements made including count.
    """
    logger.info(
        f"[find_and_replace] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Find: '{find_text}', Replace: '{replace_text}'"
    )

    # Build the find/replace request
    find_replace_request: Dict[str, Any] = {
        "find": find_text,
        "replacement": replace_text,
        "matchCase": match_case,
        "matchEntireCell": match_entire_cell,
        "searchByRegex": use_regex,
    }

    if all_sheets:
        # Search all sheets
        find_replace_request["allSheets"] = True
    elif range_name:
        # Search within a specific range
        resolved_sheet_id = await _resolve_sheet_id(
            service, spreadsheet_id, sheet_name, sheet_id
        )
        grid_range = _parse_range_to_grid(range_name)
        grid_range["sheetId"] = resolved_sheet_id
        find_replace_request["range"] = grid_range
    else:
        # Search entire sheet (single sheet)
        resolved_sheet_id = await _resolve_sheet_id(
            service, spreadsheet_id, sheet_name, sheet_id
        )
        find_replace_request["sheetId"] = resolved_sheet_id

    request_body = {"requests": [{"findReplace": find_replace_request}]}

    response = await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    # Get the result
    replies = response.get("replies", [])
    if replies and "findReplace" in replies[0]:
        result = replies[0]["findReplace"]
        occurrences_changed = result.get("occurrencesChanged", 0)
        values_changed = result.get("valuesChanged", 0)
        rows_changed = result.get("rowsChanged", 0)
        sheets_changed = result.get("sheetsChanged", 0)
    else:
        occurrences_changed = 0
        values_changed = 0
        rows_changed = 0
        sheets_changed = 0

    # Build summary
    search_scope = (
        "all sheets"
        if all_sheets
        else (f"range '{range_name}'" if range_name else "entire sheet")
    )
    text_output = (
        f"Find and replace completed in spreadsheet {spreadsheet_id} for {user_google_email}.\n"
        f"Search scope: {search_scope}\n"
        f"Found: '{find_text}' -> Replaced with: '{replace_text}'\n"
        f"Results: {occurrences_changed} occurrence(s) replaced in {values_changed} cell(s), "
        f"{rows_changed} row(s), {sheets_changed} sheet(s)."
    )

    logger.info(
        f"Successfully completed find/replace: {occurrences_changed} occurrences changed for {user_google_email}."
    )
    return text_output


@server.tool()
@handle_http_errors("clear_conditional_formatting", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def clear_conditional_formatting(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Clears all conditional formatting rules that overlap with a range in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to clear formatting from (e.g., "A1:D10"). Required.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful conditional formatting removal.
    """
    logger.info(
        f"[clear_conditional_formatting] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
    )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(range_name)
    grid_range["sheetId"] = resolved_sheet_id

    # Get existing conditional format rules to find overlapping ones
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields="sheets.conditionalFormats,sheets.properties.sheetId",
        )
        .execute
    )

    # Find rules that overlap with our range
    rules_to_delete = []
    for sheet in spreadsheet.get("sheets", []):
        if sheet.get("properties", {}).get("sheetId") != resolved_sheet_id:
            continue
        for i, rule in enumerate(sheet.get("conditionalFormats", [])):
            for rule_range in rule.get("ranges", []):
                # Check if ranges overlap
                if (
                    rule_range.get("sheetId", 0) == resolved_sheet_id
                    and rule_range.get("startRowIndex", 0) < grid_range["endRowIndex"]
                    and rule_range.get("endRowIndex", float("inf"))
                    > grid_range["startRowIndex"]
                    and rule_range.get("startColumnIndex", 0)
                    < grid_range["endColumnIndex"]
                    and rule_range.get("endColumnIndex", float("inf"))
                    > grid_range["startColumnIndex"]
                ):
                    rules_to_delete.append(i)
                    break

    if not rules_to_delete:
        return (
            f"No conditional formatting rules found overlapping range '{range_name}'."
        )

    # Delete rules in reverse order to maintain indices
    requests = [
        {"deleteConditionalFormatRule": {"sheetId": resolved_sheet_id, "index": i}}
        for i in sorted(rules_to_delete, reverse=True)
    ]

    request_body = {"requests": requests}

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    text_output = (
        f"Successfully cleared {len(rules_to_delete)} conditional formatting rule(s) from range '{range_name}' "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully cleared conditional formatting for {user_google_email}.")
    return text_output
