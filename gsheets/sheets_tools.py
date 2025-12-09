"""
Google Sheets MCP Tools

This module provides MCP tools for interacting with Google Sheets API.
"""

import logging
import asyncio
import json
import re
from typing import Any, Dict, List, Literal, Optional, Union


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
    value_render_option: Literal[
        "FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"
    ] = "FORMATTED_VALUE",
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
        value_render_option (Literal["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"]):
            How values should be represented in the output. Defaults to "FORMATTED_VALUE".
            - FORMATTED_VALUE: Values are formatted according to cell formatting (e.g., "100%" for 1.0 in a percentage cell).
            - UNFORMATTED_VALUE: Values are returned as raw underlying values (e.g., 1.0 instead of "100%").
            - FORMULA: Returns the formula in the cell (e.g., "=A1+B1" instead of the calculated result).

    Returns:
        str: The formatted values from the specified range.
    """
    logger.info(
        f"[read_sheet_values] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}, Sheet: {sheet_name}, SheetId: {sheet_id}, ValueRenderOption: {value_render_option}"
    )

    # Build the full range with sheet reference if sheet_name or sheet_id is provided
    full_range = await _build_full_range(
        service, spreadsheet_id, range_name, sheet_name, sheet_id
    )

    result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=full_range,
            valueRenderOption=value_render_option,
        )
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

    # Build render option context for output message
    render_context = ""
    if value_render_option == "FORMULA":
        render_context = " (showing formulas)"
    elif value_render_option == "UNFORMATTED_VALUE":
        render_context = " (unformatted values)"

    text_output = (
        f"Successfully read {len(values)} rows from range '{range_name}' in spreadsheet {spreadsheet_id} for {user_google_email}{render_context}:\n"
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
@handle_http_errors("append_sheet_values", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def append_sheet_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    values: Union[str, List[List[str]]],
    value_input_option: str = "USER_ENTERED",
    insert_data_option: Literal["INSERT_ROWS", "OVERWRITE"] = "INSERT_ROWS",
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Appends values after existing data in a Google Sheet.

    This tool automatically finds the end of existing data in the specified range
    and appends the new values there. This is ideal for adding new rows to data
    tables without needing to know the exact row number.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to search for a table to append to (e.g., "A:D" or "Sheet1!A:D").
            The API searches this range to find existing data and appends after it.
            Can include sheet name prefix, but this is optional if sheet_name or sheet_id is provided. Required.
        values (Union[str, List[List[str]]]): 2D array of values to append. Can be a JSON string
            (e.g., '[["row1col1", "row1col2"], ["row2col1", "row2col2"]]') or Python list. Required.
        value_input_option (str): How to interpret input values. Defaults to "USER_ENTERED".
            - "USER_ENTERED": Values are parsed as if typed by the user (formulas evaluated, numbers parsed).
            - "RAW": Values are stored exactly as provided.
        insert_data_option (Literal["INSERT_ROWS", "OVERWRITE"]): How to insert data. Defaults to "INSERT_ROWS".
            - "INSERT_ROWS": Insert new rows for the new data.
            - "OVERWRITE": Overwrite existing data in the area where new data is written.
        sheet_name (Optional[str]): Name of the sheet to append to. If provided, the range_name
            only needs the cell range (e.g., "A:D" instead of "Sheet1!A:D").
        sheet_id (Optional[int]): Numeric ID of the sheet to append to. Alternative to sheet_name.
            Takes precedence over sheet_name if both are provided.

    Returns:
        str: Confirmation message including the range where data was actually appended
             and the number of updated cells, rows, and columns.

    Example:
        # Append two rows to a sheet
        append_sheet_values(
            user_google_email="user@example.com",
            spreadsheet_id="abc123",
            range_name="A:C",
            values=[["Alice", "Engineer", "2024"], ["Bob", "Designer", "2023"]]
        )
    """
    logger.info(
        f"[append_sheet_values] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, "
        f"Range: {range_name}, Sheet: {sheet_name}, SheetId: {sheet_id}, InsertDataOption: {insert_data_option}"
    )

    # Parse values if it's a JSON string (MCP passes parameters as JSON strings)
    if isinstance(values, str):
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
                f"[append_sheet_values] Parsed JSON string to Python list with {len(values)} rows"
            )
        except json.JSONDecodeError as e:
            raise Exception(f"Invalid JSON format for values: {e}")
        except ValueError as e:
            raise Exception(f"Invalid values structure: {e}")

    if not values:
        raise Exception("Values cannot be empty. Provide at least one row to append.")

    # Build the full range with sheet reference if sheet_name or sheet_id is provided
    full_range = await _build_full_range(
        service, spreadsheet_id, range_name, sheet_name, sheet_id
    )

    body = {"values": values}

    result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadsheet_id,
            range=full_range,
            valueInputOption=value_input_option,
            insertDataOption=insert_data_option,
            body=body,
        )
        .execute
    )

    # Extract update information from the response
    updates = result.get("updates", {})
    updated_range = updates.get("updatedRange", range_name)
    updated_cells = updates.get("updatedCells", 0)
    updated_rows = updates.get("updatedRows", 0)
    updated_columns = updates.get("updatedColumns", 0)

    text_output = (
        f"Successfully appended data to spreadsheet {spreadsheet_id} for {user_google_email}. "
        f"Range: {updated_range} | "
        f"Updated: {updated_rows} rows, {updated_columns} columns, {updated_cells} cells."
    )

    logger.info(
        f"Successfully appended {updated_rows} rows ({updated_cells} cells) to range '{updated_range}' for {user_google_email}."
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
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Reads cell notes (yellow popup comments) from a range in a Google Sheet.

    Cell notes are lightweight annotations attached to individual cells that appear
    as yellow popups when you hover over a cell. These are different from threaded
    comments (use read_spreadsheet_comments for those).

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to read notes from (e.g., "A1:D10", "A1:B5"). Can include
            sheet name prefix like "Sheet1!A1:D10" but this is optional if sheet_name or
            sheet_id is provided. Required.
        sheet_name (Optional[str]): Name of the sheet to read from. If provided, the range_name
            only needs the cell range (e.g., "A1:D10" instead of "Sheet1!A1:D10").
            For sheet names with spaces, this avoids needing to escape quotes.
        sheet_id (Optional[int]): Numeric ID of the sheet to read from. Alternative to sheet_name.
            Takes precedence over sheet_name if both are provided.

    Returns:
        str: Formatted list of cells with notes, including cell reference and note content.
    """
    logger.info(
        f"[read_cell_notes] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}, Sheet: {sheet_name}, SheetId: {sheet_id}"
    )

    # Build the full range with sheet reference if sheet_name or sheet_id is provided
    full_range = await _build_full_range(
        service, spreadsheet_id, range_name, sheet_name, sheet_id
    )

    # Get spreadsheet data including notes
    result = await asyncio.to_thread(
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            ranges=[full_range],
            fields="sheets.data.rowData.values.note,sheets.properties.title",
        )
        .execute
    )

    sheets = result.get("sheets", [])
    if not sheets:
        return f"No data found in range '{full_range}'."

    sheet = sheets[0]
    resolved_sheet_name = sheet.get("properties", {}).get("title", "Unknown")
    data = sheet.get("data", [])

    if not data:
        return f"No data found in range '{full_range}'."

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
        return f"No cell notes found in range '{full_range}' of sheet '{resolved_sheet_name}'."

    output = [
        f"Found {len(notes_found)} cell notes in range '{full_range}' of sheet '{resolved_sheet_name}':\n"
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
    sheet_id: Optional[int] = None,
) -> str:
    """
    Adds or updates a note on a specific cell in a Google Sheet.

    Cell notes are lightweight yellow popup annotations that appear when hovering
    over a cell. They are simpler than threaded comments - just plain text without
    threading, author tracking, or resolution status.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        cell (str): Cell reference in A1 notation (e.g., 'A1', 'B2', 'AA10'). Can also
            include a sheet prefix (e.g., 'SheetName!A1', "'My Sheet'!B2"). Required.
        note (str): The note text to add to the cell. Required.
        sheet_name (Optional[str]): Name of the sheet. If not provided and cell doesn't
            include a sheet prefix, uses the first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.
            Takes precedence over sheet_name and cell prefix if provided.

    Returns:
        str: Confirmation message of the successful note update.
    """
    logger.info(
        f"[update_cell_note] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Cell: {cell}, Sheet: {sheet_name}, SheetId: {sheet_id}"
    )

    # Check if cell contains sheet prefix (e.g., 'SheetName!A1')
    # Use it only if no explicit sheet_name or sheet_id is provided
    cell_sheet_prefix, clean_cell = _strip_sheet_prefix(cell)
    if cell_sheet_prefix and sheet_name is None and sheet_id is None:
        sheet_name = cell_sheet_prefix
        cell = clean_cell

    # Parse cell reference
    row_index, col_index = _parse_cell_reference(cell)

    # Get sheet ID and name
    resolved_sheet_id: int
    resolved_sheet_name: str
    if sheet_id is not None:
        resolved_sheet_id = sheet_id
        resolved_sheet_name = await _get_sheet_name_by_id(
            service, spreadsheet_id, sheet_id
        )
    elif sheet_name:
        resolved_sheet_id = await _get_sheet_id_by_name(
            service, spreadsheet_id, sheet_name
        )
        resolved_sheet_name = sheet_name
    else:
        # Get first sheet ID and name
        spreadsheet = await asyncio.to_thread(
            service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
        )
        sheets = spreadsheet.get("sheets", [])
        if not sheets:
            raise ValueError("Spreadsheet has no sheets.")
        resolved_sheet_id = sheets[0].get("properties", {}).get("sheetId")
        resolved_sheet_name = sheets[0].get("properties", {}).get("title", "Sheet1")

    # Build the update request
    request_body = {
        "requests": [
            {
                "updateCells": {
                    "rows": [{"values": [{"note": note}]}],
                    "fields": "note",
                    "start": {
                        "sheetId": resolved_sheet_id,
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
        f"Successfully added/updated note on cell {cell} in sheet '{resolved_sheet_name}' "
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
    sheet_id: Optional[int] = None,
) -> str:
    """
    Removes a note from a specific cell in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        cell (str): Cell reference in A1 notation (e.g., 'A1', 'B2', 'AA10'). Can also
            include a sheet prefix (e.g., 'SheetName!A1', "'My Sheet'!B2"). Required.
        sheet_name (Optional[str]): Name of the sheet. If not provided and cell doesn't
            include a sheet prefix, uses the first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.
            Takes precedence over sheet_name and cell prefix if provided.

    Returns:
        str: Confirmation message of the successful note removal.
    """
    logger.info(
        f"[clear_cell_note] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Cell: {cell}, Sheet: {sheet_name}, SheetId: {sheet_id}"
    )

    # Check if cell contains sheet prefix (e.g., 'SheetName!A1')
    # Use it only if no explicit sheet_name or sheet_id is provided
    cell_sheet_prefix, clean_cell = _strip_sheet_prefix(cell)
    if cell_sheet_prefix and sheet_name is None and sheet_id is None:
        sheet_name = cell_sheet_prefix
        cell = clean_cell

    # Parse cell reference
    row_index, col_index = _parse_cell_reference(cell)

    # Get sheet ID and name
    resolved_sheet_id: int
    resolved_sheet_name: str
    if sheet_id is not None:
        resolved_sheet_id = sheet_id
        resolved_sheet_name = await _get_sheet_name_by_id(
            service, spreadsheet_id, sheet_id
        )
    elif sheet_name:
        resolved_sheet_id = await _get_sheet_id_by_name(
            service, spreadsheet_id, sheet_name
        )
        resolved_sheet_name = sheet_name
    else:
        # Get first sheet ID and name
        spreadsheet = await asyncio.to_thread(
            service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
        )
        sheets = spreadsheet.get("sheets", [])
        if not sheets:
            raise ValueError("Spreadsheet has no sheets.")
        resolved_sheet_id = sheets[0].get("properties", {}).get("sheetId")
        resolved_sheet_name = sheets[0].get("properties", {}).get("title", "Sheet1")

    # Build the update request - setting note to empty string clears it
    request_body = {
        "requests": [
            {
                "updateCells": {
                    "rows": [{"values": [{"note": ""}]}],
                    "fields": "note",
                    "start": {
                        "sheetId": resolved_sheet_id,
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
        f"Successfully cleared note from cell {cell} in sheet '{resolved_sheet_name}' "
        f"of spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully cleared cell note for {user_google_email}.")
    return text_output


def _strip_sheet_prefix(range_str: str) -> tuple[Optional[str], str]:
    """
    Strip the sheet name prefix from a range string if present.

    Handles formats like:
    - 'SheetName!A1:D10' -> ('SheetName', 'A1:D10')
    - "'Sheet Name'!A1:D10" -> ('Sheet Name', 'A1:D10')
    - "'Sheet''s Name'!A1:D10" -> ("Sheet's Name", 'A1:D10')
    - 'A1:D10' -> (None, 'A1:D10')

    Args:
        range_str: Range in A1 notation, optionally with sheet prefix

    Returns:
        Tuple of (sheet_name or None, clean_range_without_prefix)
    """
    if "!" not in range_str:
        return None, range_str

    # Split on the last '!' to handle sheet names that might contain '!'
    # However, we need to be careful with quoted sheet names
    if range_str.startswith("'"):
        # Find the closing quote (handling escaped quotes '')
        i = 1
        while i < len(range_str):
            if range_str[i] == "'":
                if i + 1 < len(range_str) and range_str[i + 1] == "'":
                    # Escaped quote, skip both
                    i += 2
                else:
                    # End of quoted name
                    break
            else:
                i += 1

        # After the closing quote, expect '!'
        if i + 1 < len(range_str) and range_str[i + 1] == "!":
            quoted_name = range_str[1:i]  # Strip outer quotes
            # Unescape doubled quotes
            sheet_name = quoted_name.replace("''", "'")
            clean_range = range_str[i + 2 :]  # After the '!'
            return sheet_name, clean_range

    # Simple case: no quotes, split on first '!'
    parts = range_str.split("!", 1)
    if len(parts) == 2:
        return parts[0], parts[1]

    return None, range_str


def _parse_range_to_grid(range_str: str) -> Dict[str, int]:
    """
    Parse a range string like 'A1:D10' into GridRange coordinates.

    Args:
        range_str: Range in A1 notation (e.g., 'A1:D10', 'B2:C5', 'Sheet1!A1:D10')
            If a sheet name prefix is included, it will be stripped before parsing.

    Returns:
        Dict with startRowIndex, endRowIndex, startColumnIndex, endColumnIndex (all 0-indexed)
    """
    # Strip sheet name prefix if present
    _, clean_range = _strip_sheet_prefix(range_str)

    # Handle single cell vs range
    if ":" in clean_range:
        start_cell, end_cell = clean_range.split(":")
    else:
        start_cell = end_cell = clean_range

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

    # Extract sheet name from range_name if present and no explicit sheet_name provided
    extracted_sheet_name, clean_range = _strip_sheet_prefix(range_name)
    effective_sheet_name = (
        sheet_name if sheet_name is not None else extracted_sheet_name
    )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, effective_sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(clean_range)
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

    # Extract sheet name from range_name if present and no explicit sheet_name provided
    extracted_sheet_name, clean_range = _strip_sheet_prefix(range_name)
    effective_sheet_name = (
        sheet_name if sheet_name is not None else extracted_sheet_name
    )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, effective_sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(clean_range)
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

    # Extract sheet name from range_name if present and no explicit sheet_name provided
    extracted_sheet_name, clean_range = _strip_sheet_prefix(range_name)
    effective_sheet_name = (
        sheet_name if sheet_name is not None else extracted_sheet_name
    )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, effective_sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(clean_range)
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

    # Extract sheet name from range_name if present and no explicit sheet_name provided
    extracted_sheet_name, clean_range = _strip_sheet_prefix(range_name)
    effective_sheet_name = (
        sheet_name if sheet_name is not None else extracted_sheet_name
    )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, effective_sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(clean_range)
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

    # Extract sheet name from range_name if present and no explicit sheet_name provided
    extracted_sheet_name, clean_range = _strip_sheet_prefix(range_name)
    effective_sheet_name = (
        sheet_name if sheet_name is not None else extracted_sheet_name
    )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, effective_sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(clean_range)
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

    # Extract sheet name from range_name if present and no explicit sheet_name provided
    extracted_sheet_name, clean_range = _strip_sheet_prefix(range_name)
    effective_sheet_name = (
        sheet_name if sheet_name is not None else extracted_sheet_name
    )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, effective_sheet_name, sheet_id
    )

    # Parse range to get grid coordinates
    grid_range = _parse_range_to_grid(clean_range)
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
        # Extract sheet name from range_name if present and no explicit sheet_name provided
        extracted_sheet_name, clean_range = _strip_sheet_prefix(range_name)
        effective_sheet_name = (
            sheet_name if sheet_name is not None else extracted_sheet_name
        )
        # Search within a specific range
        resolved_sheet_id = await _resolve_sheet_id(
            service, spreadsheet_id, effective_sheet_name, sheet_id
        )
        grid_range = _parse_range_to_grid(clean_range)
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
@handle_http_errors("insert_rows", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def insert_rows(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    start_row: int,
    num_rows: int = 1,
    inherit_from_before: bool = False,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Inserts blank rows at a specified position in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        start_row (int): The row number where new rows will be inserted (1-indexed).
            New rows are inserted BEFORE this row. Required.
        num_rows (int): The number of rows to insert. Defaults to 1.
        inherit_from_before (bool): If True, new rows inherit formatting from the row
            before start_row. If False, inherit from start_row. Defaults to False.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful row insertion.
    """
    logger.info(
        f"[insert_rows] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, StartRow: {start_row}, NumRows: {num_rows}"
    )

    if start_row < 1:
        raise ValueError("start_row must be 1 or greater (1-indexed)")
    if num_rows < 1:
        raise ValueError("num_rows must be 1 or greater")

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Convert to 0-indexed
    start_index = start_row - 1

    # Build the insert request
    request_body = {
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": resolved_sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_index,
                        "endIndex": start_index + num_rows,
                    },
                    "inheritFromBefore": inherit_from_before,
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
        f"Successfully inserted {num_rows} row(s) at row {start_row} "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully inserted rows for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("delete_rows", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def delete_rows(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    start_row: int,
    num_rows: int = 1,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Deletes rows at a specified position in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        start_row (int): The first row number to delete (1-indexed). Required.
        num_rows (int): The number of rows to delete. Defaults to 1.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful row deletion.
    """
    logger.info(
        f"[delete_rows] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, StartRow: {start_row}, NumRows: {num_rows}"
    )

    if start_row < 1:
        raise ValueError("start_row must be 1 or greater (1-indexed)")
    if num_rows < 1:
        raise ValueError("num_rows must be 1 or greater")

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Convert to 0-indexed
    start_index = start_row - 1

    # Build the delete request
    request_body = {
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": resolved_sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_index,
                        "endIndex": start_index + num_rows,
                    }
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    end_row = start_row + num_rows - 1
    row_range = f"{start_row}-{end_row}" if num_rows > 1 else str(start_row)
    text_output = (
        f"Successfully deleted {num_rows} row(s) ({row_range}) "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully deleted rows for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("insert_columns", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def insert_columns(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    start_column: str,
    num_columns: int = 1,
    inherit_from_before: bool = False,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Inserts blank columns at a specified position in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        start_column (str): The column letter where new columns will be inserted (e.g., "A", "B", "AA").
            New columns are inserted BEFORE this column. Required.
        num_columns (int): The number of columns to insert. Defaults to 1.
        inherit_from_before (bool): If True, new columns inherit formatting from the column
            before start_column. If False, inherit from start_column. Defaults to False.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful column insertion.
    """
    logger.info(
        f"[insert_columns] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, StartColumn: {start_column}, NumColumns: {num_columns}"
    )

    if num_columns < 1:
        raise ValueError("num_columns must be 1 or greater")

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Parse column letter to index
    col_ref = f"{start_column}1"
    _, start_col_idx = _parse_cell_reference(col_ref)

    # Build the insert request
    request_body = {
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": resolved_sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": start_col_idx,
                        "endIndex": start_col_idx + num_columns,
                    },
                    "inheritFromBefore": inherit_from_before,
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
        f"Successfully inserted {num_columns} column(s) at column {start_column} "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully inserted columns for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("delete_columns", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def delete_columns(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    start_column: str,
    num_columns: int = 1,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Deletes columns at a specified position in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        start_column (str): The first column letter to delete (e.g., "A", "B", "AA"). Required.
        num_columns (int): The number of columns to delete. Defaults to 1.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful column deletion.
    """
    logger.info(
        f"[delete_columns] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, StartColumn: {start_column}, NumColumns: {num_columns}"
    )

    if num_columns < 1:
        raise ValueError("num_columns must be 1 or greater")

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Parse column letter to index
    col_ref = f"{start_column}1"
    _, start_col_idx = _parse_cell_reference(col_ref)

    # Build the delete request
    request_body = {
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": resolved_sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": start_col_idx,
                        "endIndex": start_col_idx + num_columns,
                    }
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    if num_columns > 1:
        end_col_letter = _column_index_to_letter(start_col_idx + num_columns - 1)
        col_range = f"{start_column}-{end_col_letter}"
    else:
        col_range = start_column

    text_output = (
        f"Successfully deleted {num_columns} column(s) ({col_range}) "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully deleted columns for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("auto_resize_dimension", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def auto_resize_dimension(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    dimension: str,
    start_index: Optional[int] = None,
    end_index: Optional[int] = None,
    start_column: Optional[str] = None,
    end_column: Optional[str] = None,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Auto-resizes rows or columns to fit their content in a Google Sheet.

    For rows, this adjusts the height to fit the tallest cell content.
    For columns, this adjusts the width to fit the widest cell content.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        dimension (str): Either "ROWS" or "COLUMNS". Required.
        start_index (Optional[int]): For ROWS: starting row number (1-indexed).
            For COLUMNS: starting column index (0-indexed). Use start_column instead for columns.
        end_index (Optional[int]): For ROWS: ending row number (inclusive, 1-indexed).
            For COLUMNS: ending column index (exclusive, 0-indexed). Use end_column instead for columns.
        start_column (Optional[str]): For COLUMNS: starting column letter (e.g., "A", "B", "AA").
            More intuitive than start_index for columns.
        end_column (Optional[str]): For COLUMNS: ending column letter (inclusive).
            More intuitive than end_index for columns.
        sheet_name (Optional[str]): Name of the sheet. If not provided, uses first sheet.
        sheet_id (Optional[int]): Numeric ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful auto-resize operation.
    """
    logger.info(
        f"[auto_resize_dimension] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Dimension: {dimension}"
    )

    # Validate dimension
    dimension_upper = dimension.upper()
    if dimension_upper not in ["ROWS", "COLUMNS"]:
        raise ValueError(f"Invalid dimension: '{dimension}'. Must be ROWS or COLUMNS.")

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Determine start and end indices
    if dimension_upper == "COLUMNS":
        # For columns, prefer letter-based specification
        if start_column is not None:
            col_ref = f"{start_column}1"
            _, resolved_start = _parse_cell_reference(col_ref)
        elif start_index is not None:
            resolved_start = start_index
        else:
            resolved_start = 0  # Default to first column

        if end_column is not None:
            col_ref = f"{end_column}1"
            _, end_col_idx = _parse_cell_reference(col_ref)
            resolved_end = end_col_idx + 1  # API uses exclusive end
        elif end_index is not None:
            resolved_end = end_index
        else:
            # Auto-resize all columns - need to get sheet dimensions
            spreadsheet = await asyncio.to_thread(
                service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
            )
            for sheet in spreadsheet.get("sheets", []):
                if sheet.get("properties", {}).get("sheetId") == resolved_sheet_id:
                    resolved_end = (
                        sheet.get("properties", {})
                        .get("gridProperties", {})
                        .get("columnCount", 26)
                    )
                    break
            else:
                resolved_end = 26  # Default if not found
    else:
        # For rows, use 1-indexed start_index and end_index
        if start_index is not None:
            if start_index < 1:
                raise ValueError(
                    "start_index for rows must be 1 or greater (1-indexed)"
                )
            resolved_start = start_index - 1  # Convert to 0-indexed
        else:
            resolved_start = 0  # Default to first row

        if end_index is not None:
            if end_index < start_index if start_index else 1:
                raise ValueError("end_index must be >= start_index")
            resolved_end = (
                end_index  # Already 1-indexed, use as exclusive end in 0-indexed terms
            )
        else:
            # Auto-resize all rows - need to get sheet dimensions
            spreadsheet = await asyncio.to_thread(
                service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
            )
            for sheet in spreadsheet.get("sheets", []):
                if sheet.get("properties", {}).get("sheetId") == resolved_sheet_id:
                    resolved_end = (
                        sheet.get("properties", {})
                        .get("gridProperties", {})
                        .get("rowCount", 1000)
                    )
                    break
            else:
                resolved_end = 1000  # Default if not found

    # Build the auto-resize request
    request_body = {
        "requests": [
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": resolved_sheet_id,
                        "dimension": dimension_upper,
                        "startIndex": resolved_start,
                        "endIndex": resolved_end,
                    }
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
    if dimension_upper == "COLUMNS":
        if start_column and end_column:
            range_desc = f"columns {start_column}-{end_column}"
        elif start_column:
            range_desc = f"columns starting from {start_column}"
        else:
            range_desc = "all columns"
    else:
        if start_index and end_index:
            range_desc = f"rows {start_index}-{end_index}"
        elif start_index:
            range_desc = f"rows starting from {start_index}"
        else:
            range_desc = "all rows"

    text_output = (
        f"Successfully auto-resized {range_desc} to fit content "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully auto-resized dimensions for {user_google_email}.")
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

    # Extract sheet name from range_name if present and no explicit sheet_name provided
    extracted_sheet_name, clean_range = _strip_sheet_prefix(range_name)
    effective_sheet_name = (
        sheet_name if sheet_name is not None else extracted_sheet_name
    )

    # Resolve sheet ID
    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, effective_sheet_name, sheet_id
    )

    # Parse range
    grid_range = _parse_range_to_grid(clean_range)
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


@server.tool()
@handle_http_errors("copy_range", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def copy_range(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    source_range: str,
    destination_range: str,
    source_sheet_name: Optional[str] = None,
    source_sheet_id: Optional[int] = None,
    destination_sheet_name: Optional[str] = None,
    destination_sheet_id: Optional[int] = None,
    paste_type: Literal[
        "PASTE_NORMAL",
        "PASTE_VALUES",
        "PASTE_FORMAT",
        "PASTE_NO_BORDERS",
        "PASTE_FORMULA",
        "PASTE_DATA_VALIDATION",
        "PASTE_CONDITIONAL_FORMATTING",
    ] = "PASTE_NORMAL",
    transpose: bool = False,
) -> str:
    """
    Copies data from one range to another within a spreadsheet.

    This tool copies data from a source range to a destination range. It supports copying
    within the same sheet or between different sheets in the same spreadsheet. You can
    control what content is copied (values, formulas, formatting, etc.) and optionally
    transpose the data.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        source_range (str): The range to copy from (e.g., "A1:D10", "Sheet1!A1:D10"). Required.
        destination_range (str): The range to paste to (e.g., "F1:I10", "Sheet2!A1"). Required.
            Only the starting cell matters - the destination size is determined by the source.
        source_sheet_name (Optional[str]): Name of the source sheet. If not provided, extracted
            from source_range or uses first sheet.
        source_sheet_id (Optional[int]): Numeric ID of the source sheet. Alternative to
            source_sheet_name.
        destination_sheet_name (Optional[str]): Name of the destination sheet. If not provided,
            extracted from destination_range or uses same sheet as source.
        destination_sheet_id (Optional[int]): Numeric ID of the destination sheet. Alternative
            to destination_sheet_name.
        paste_type (str): What to paste. Options:
            - "PASTE_NORMAL": Values, formulas, formats, and merges (default)
            - "PASTE_VALUES": Values only (no formulas or formatting)
            - "PASTE_FORMAT": Formatting and data validation only
            - "PASTE_NO_BORDERS": Like PASTE_NORMAL but excludes borders
            - "PASTE_FORMULA": Formulas only
            - "PASTE_DATA_VALIDATION": Data validation rules only
            - "PASTE_CONDITIONAL_FORMATTING": Conditional formatting rules only
        transpose (bool): If True, rows become columns and columns become rows. Defaults to False.

    Returns:
        str: Confirmation message of the successful copy operation.
    """
    logger.info(
        f"[copy_range] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, "
        f"Source: {source_range}, Destination: {destination_range}, PasteType: {paste_type}"
    )

    # Extract sheet names from ranges if present
    extracted_source_sheet, clean_source_range = _strip_sheet_prefix(source_range)
    extracted_dest_sheet, clean_dest_range = _strip_sheet_prefix(destination_range)

    # Determine effective sheet names
    effective_source_sheet = (
        source_sheet_name if source_sheet_name is not None else extracted_source_sheet
    )
    effective_dest_sheet = (
        destination_sheet_name
        if destination_sheet_name is not None
        else extracted_dest_sheet
    )

    # Resolve source sheet ID
    resolved_source_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, effective_source_sheet, source_sheet_id
    )

    # Resolve destination sheet ID (defaults to source sheet if not specified)
    if (
        destination_sheet_id is None
        and destination_sheet_name is None
        and extracted_dest_sheet is None
    ):
        # No destination sheet specified - use source sheet
        resolved_dest_sheet_id = resolved_source_sheet_id
    else:
        resolved_dest_sheet_id = await _resolve_sheet_id(
            service, spreadsheet_id, effective_dest_sheet, destination_sheet_id
        )

    # Parse source and destination ranges to grid coordinates
    source_grid = _parse_range_to_grid(clean_source_range)
    source_grid["sheetId"] = resolved_source_sheet_id

    dest_grid = _parse_range_to_grid(clean_dest_range)
    dest_grid["sheetId"] = resolved_dest_sheet_id

    # Build the copyPaste request
    request_body = {
        "requests": [
            {
                "copyPaste": {
                    "source": source_grid,
                    "destination": dest_grid,
                    "pasteType": paste_type,
                    "pasteOrientation": "TRANSPOSE" if transpose else "NORMAL",
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    # Build descriptive output message
    transpose_text = " (transposed)" if transpose else ""
    paste_type_text = paste_type.replace("PASTE_", "").replace("_", " ").lower()

    text_output = (
        f"Successfully copied range '{source_range}' to '{destination_range}'{transpose_text} "
        f"with paste type '{paste_type_text}' in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully copied range for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("copy_sheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def copy_sheet(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    source_sheet_name: Optional[str] = None,
    source_sheet_id: Optional[int] = None,
    destination_spreadsheet_id: Optional[str] = None,
    new_sheet_name: Optional[str] = None,
) -> str:
    """
    Copies a sheet within the same spreadsheet or to a different spreadsheet.

    This is useful for:
    - Creating a backup/copy of a sheet before making changes
    - Duplicating a template sheet within the same spreadsheet
    - Copying a sheet to another spreadsheet

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the source spreadsheet. Required.
        source_sheet_name (Optional[str]): Name of the sheet to copy. If neither
            source_sheet_name nor source_sheet_id is provided, copies the first sheet.
        source_sheet_id (Optional[int]): Numeric ID of the sheet to copy. Alternative
            to source_sheet_name.
        destination_spreadsheet_id (Optional[str]): The ID of the destination spreadsheet.
            If not provided, copies within the same spreadsheet.
        new_sheet_name (Optional[str]): Name for the copied sheet. If not provided,
            the API will use the default "Copy of <original>" name.

    Returns:
        str: Confirmation message including the new sheet ID and name.
    """
    logger.info(
        f"[copy_sheet] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, "
        f"SourceSheet: {source_sheet_name}, SourceSheetId: {source_sheet_id}, "
        f"DestSpreadsheet: {destination_spreadsheet_id}, NewName: {new_sheet_name}"
    )

    # Resolve source sheet ID
    resolved_source_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, source_sheet_name, source_sheet_id
    )

    # Get source sheet title for the response message
    source_sheet_title = await _get_sheet_name_by_id(
        service, spreadsheet_id, resolved_source_sheet_id
    )

    # Determine destination spreadsheet
    dest_spreadsheet_id = destination_spreadsheet_id or spreadsheet_id
    is_same_spreadsheet = dest_spreadsheet_id == spreadsheet_id

    # Copy the sheet using the copyTo API
    copy_request_body = {"destinationSpreadsheetId": dest_spreadsheet_id}

    copy_response = await asyncio.to_thread(
        service.spreadsheets()
        .sheets()
        .copyTo(
            spreadsheetId=spreadsheet_id,
            sheetId=resolved_source_sheet_id,
            body=copy_request_body,
        )
        .execute
    )

    new_sheet_id = copy_response.get("sheetId")
    default_new_name = copy_response.get("title", f"Copy of {source_sheet_title}")

    # Rename the copied sheet if a custom name was provided
    final_sheet_name = default_new_name
    if new_sheet_name and new_sheet_name != default_new_name:
        rename_request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": new_sheet_id,
                            "title": new_sheet_name,
                        },
                        "fields": "title",
                    }
                }
            ]
        }

        await asyncio.to_thread(
            service.spreadsheets()
            .batchUpdate(spreadsheetId=dest_spreadsheet_id, body=rename_request_body)
            .execute
        )
        final_sheet_name = new_sheet_name

    # Build response message
    if is_same_spreadsheet:
        text_output = (
            f"Successfully copied sheet '{source_sheet_title}' to '{final_sheet_name}' "
            f"(ID: {new_sheet_id}) in spreadsheet {spreadsheet_id} for {user_google_email}."
        )
    else:
        text_output = (
            f"Successfully copied sheet '{source_sheet_title}' from spreadsheet {spreadsheet_id} "
            f"to '{final_sheet_name}' (ID: {new_sheet_id}) in destination spreadsheet "
            f"{dest_spreadsheet_id} for {user_google_email}."
        )

    logger.info(f"Successfully copied sheet for {user_google_email}.")
    return text_output
