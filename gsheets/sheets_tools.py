"""
Google Sheets MCP Tools

This module provides MCP tools for interacting with Google Sheets API.
"""

import logging
import asyncio
import json
import re
from typing import List, Optional, Union, Dict


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


@server.tool()
@handle_http_errors("read_sheet_values", is_read_only=True, service_type="sheets")
@require_google_service("sheets", "sheets_read")
async def read_sheet_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str = "A1:Z1000",
) -> str:
    """
    Reads values from a specific range in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to read (e.g., "Sheet1!A1:D10", "A1:D10"). Defaults to "A1:Z1000".

    Returns:
        str: The formatted values from the specified range.
    """
    logger.info(
        f"[read_sheet_values] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
    )

    result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=range_name)
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
) -> str:
    """
    Modifies values in a specific range of a Google Sheet - can write, update, or clear values.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to modify (e.g., "Sheet1!A1:D10", "A1:D10"). Required.
        values (Optional[Union[str, List[List[str]]]]): 2D array of values to write/update. Can be a JSON string or Python list. Required unless clear_values=True.
        value_input_option (str): How to interpret input values ("RAW" or "USER_ENTERED"). Defaults to "USER_ENTERED".
        clear_values (bool): If True, clears the range instead of writing values. Defaults to False.

    Returns:
        str: Confirmation message of the successful modification operation.
    """
    operation = "clear" if clear_values else "write"
    logger.info(
        f"[modify_sheet_values] Invoked. Operation: {operation}, Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
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

    if clear_values:
        result = await asyncio.to_thread(
            service.spreadsheets()
            .values()
            .clear(spreadsheetId=spreadsheet_id, range=range_name)
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
                range=range_name,
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


async def _resolve_sheet_id(
    service,
    spreadsheet_id: str,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> int:
    """
    Resolve sheet_name or sheet_id to the actual sheet ID.

    Args:
        service: The Google Sheets API service.
        spreadsheet_id: The ID of the spreadsheet.
        sheet_name: Optional name of the sheet.
        sheet_id: Optional ID of the sheet.

    Returns:
        int: The resolved sheet ID.

    Raises:
        Exception: If neither sheet_name nor sheet_id is provided, or sheet not found.
    """
    if sheet_id is not None:
        return sheet_id

    if sheet_name is None:
        # Get the first sheet by default
        spreadsheet = await asyncio.to_thread(
            service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
        )
        sheets = spreadsheet.get("sheets", [])
        if not sheets:
            raise Exception(f"Spreadsheet {spreadsheet_id} has no sheets.")
        return sheets[0].get("properties", {}).get("sheetId", 0)

    # Resolve sheet_name to sheet_id
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )
    sheets = spreadsheet.get("sheets", [])

    for sheet in sheets:
        props = sheet.get("properties", {})
        if props.get("title") == sheet_name:
            return props.get("sheetId", 0)

    raise Exception(f"Sheet '{sheet_name}' not found in spreadsheet {spreadsheet_id}.")


@server.tool()
@handle_http_errors("insert_rows", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def insert_rows(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    start_row: int,
    num_rows: int = 1,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Inserts blank rows at a specific position in a Google Sheet.

    Existing data at and below the specified position will shift down to make room.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        start_row (int): The 1-indexed row position to insert at. Required.
            For example, start_row=3 inserts rows starting at row 3.
        num_rows (int): Number of rows to insert. Defaults to 1.
        sheet_name (Optional[str]): The name of the sheet. If not provided, uses
            the first sheet or sheet_id if specified.
        sheet_id (Optional[int]): The ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful insertion.

    Example:
        Insert 2 blank rows at row 5 (existing rows 5+ shift to 7+):
        >>> insert_rows(spreadsheet_id="...", start_row=5, num_rows=2)
    """
    logger.info(
        f"[insert_rows] Invoked. Email: '{user_google_email}', "
        f"Spreadsheet: {spreadsheet_id}, start_row: {start_row}, num_rows: {num_rows}"
    )

    if start_row < 1:
        raise Exception("start_row must be >= 1 (1-indexed).")
    if num_rows < 1:
        raise Exception("num_rows must be >= 1.")

    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Convert 1-indexed to 0-indexed for the API
    start_index = start_row - 1

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
                    "inheritFromBefore": start_index > 0,
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    sheet_identifier = sheet_name if sheet_name else f"sheet ID {resolved_sheet_id}"
    text_output = (
        f"Successfully inserted {num_rows} row(s) at row {start_row} "
        f"in '{sheet_identifier}' of spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully inserted {num_rows} row(s) for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("insert_columns", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def insert_columns(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    start_column: int,
    num_columns: int = 1,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Inserts blank columns at a specific position in a Google Sheet.

    Existing data at and to the right of the specified position will shift right.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        start_column (int): The 1-indexed column position to insert at. Required.
            For example, start_column=3 inserts columns starting at column C.
        num_columns (int): Number of columns to insert. Defaults to 1.
        sheet_name (Optional[str]): The name of the sheet. If not provided, uses
            the first sheet or sheet_id if specified.
        sheet_id (Optional[int]): The ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful insertion.

    Example:
        Insert 2 blank columns at column C (existing columns C+ shift to E+):
        >>> insert_columns(spreadsheet_id="...", start_column=3, num_columns=2)
    """
    logger.info(
        f"[insert_columns] Invoked. Email: '{user_google_email}', "
        f"Spreadsheet: {spreadsheet_id}, start_column: {start_column}, num_columns: {num_columns}"
    )

    if start_column < 1:
        raise Exception("start_column must be >= 1 (1-indexed).")
    if num_columns < 1:
        raise Exception("num_columns must be >= 1.")

    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Convert 1-indexed to 0-indexed for the API
    start_index = start_column - 1

    request_body = {
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": resolved_sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": start_index,
                        "endIndex": start_index + num_columns,
                    },
                    "inheritFromBefore": start_index > 0,
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    # Convert column number to letter for user-friendly output
    def column_to_letter(col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord("A")) + result
            col_num //= 26
        return result

    column_letter = column_to_letter(start_column)
    sheet_identifier = sheet_name if sheet_name else f"sheet ID {resolved_sheet_id}"
    text_output = (
        f"Successfully inserted {num_columns} column(s) at column {column_letter} (column {start_column}) "
        f"in '{sheet_identifier}' of spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(
        f"Successfully inserted {num_columns} column(s) for {user_google_email}."
    )
    return text_output


@server.tool()
@handle_http_errors("delete_rows", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def delete_rows(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    start_row: int,
    end_row: int,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Deletes rows from a Google Sheet.

    Removes the specified rows and shifts remaining rows up to fill the gap.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        start_row (int): The 1-indexed first row to delete (inclusive). Required.
        end_row (int): The 1-indexed last row to delete (inclusive). Required.
            For example, start_row=3, end_row=5 deletes rows 3, 4, and 5.
        sheet_name (Optional[str]): The name of the sheet. If not provided, uses
            the first sheet or sheet_id if specified.
        sheet_id (Optional[int]): The ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful deletion.

    Example:
        Delete rows 5 through 7:
        >>> delete_rows(spreadsheet_id="...", start_row=5, end_row=7)
    """
    logger.info(
        f"[delete_rows] Invoked. Email: '{user_google_email}', "
        f"Spreadsheet: {spreadsheet_id}, start_row: {start_row}, end_row: {end_row}"
    )

    if start_row < 1:
        raise Exception("start_row must be >= 1 (1-indexed).")
    if end_row < 1:
        raise Exception("end_row must be >= 1 (1-indexed).")
    if end_row < start_row:
        raise Exception("end_row must be >= start_row.")

    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Convert 1-indexed inclusive to 0-indexed exclusive for the API
    start_index = start_row - 1
    end_index = end_row  # API end_index is exclusive, so end_row (1-indexed) works

    request_body = {
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": resolved_sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_index,
                        "endIndex": end_index,
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

    num_rows = end_row - start_row + 1
    sheet_identifier = sheet_name if sheet_name else f"sheet ID {resolved_sheet_id}"
    text_output = (
        f"Successfully deleted {num_rows} row(s) (rows {start_row}-{end_row}) "
        f"from '{sheet_identifier}' of spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully deleted {num_rows} row(s) for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("delete_columns", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def delete_columns(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    start_column: int,
    end_column: int,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Deletes columns from a Google Sheet.

    Removes the specified columns and shifts remaining columns left to fill the gap.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        start_column (int): The 1-indexed first column to delete (inclusive). Required.
            For example, start_column=3 means column C.
        end_column (int): The 1-indexed last column to delete (inclusive). Required.
            For example, start_column=3, end_column=5 deletes columns C, D, and E.
        sheet_name (Optional[str]): The name of the sheet. If not provided, uses
            the first sheet or sheet_id if specified.
        sheet_id (Optional[int]): The ID of the sheet. Alternative to sheet_name.

    Returns:
        str: Confirmation message of the successful deletion.

    Example:
        Delete columns C through E (columns 3-5):
        >>> delete_columns(spreadsheet_id="...", start_column=3, end_column=5)
    """
    logger.info(
        f"[delete_columns] Invoked. Email: '{user_google_email}', "
        f"Spreadsheet: {spreadsheet_id}, start_column: {start_column}, end_column: {end_column}"
    )

    if start_column < 1:
        raise Exception("start_column must be >= 1 (1-indexed).")
    if end_column < 1:
        raise Exception("end_column must be >= 1 (1-indexed).")
    if end_column < start_column:
        raise Exception("end_column must be >= start_column.")

    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, sheet_name, sheet_id
    )

    # Convert 1-indexed inclusive to 0-indexed exclusive for the API
    start_index = start_column - 1
    end_index = end_column  # API end_index is exclusive

    request_body = {
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": resolved_sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": start_index,
                        "endIndex": end_index,
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

    # Convert column numbers to letters for user-friendly output
    def column_to_letter(col_num: int) -> str:
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord("A")) + result
            col_num //= 26
        return result

    start_letter = column_to_letter(start_column)
    end_letter = column_to_letter(end_column)
    num_columns = end_column - start_column + 1
    sheet_identifier = sheet_name if sheet_name else f"sheet ID {resolved_sheet_id}"
    text_output = (
        f"Successfully deleted {num_columns} column(s) (columns {start_letter}-{end_letter}, "
        f"columns {start_column}-{end_column}) from '{sheet_identifier}' of spreadsheet {spreadsheet_id} "
        f"for {user_google_email}."
    )

    logger.info(
        f"Successfully deleted {num_columns} column(s) for {user_google_email}."
    )
    return text_output


def _column_letter_to_index(col_str: str) -> int:
    """
    Convert a column letter (e.g., 'A', 'B', 'AA') to a 0-indexed column number.

    Args:
        col_str: Column letter(s) like 'A', 'Z', 'AA', 'AZ'.

    Returns:
        int: 0-indexed column number (A=0, B=1, Z=25, AA=26).
    """
    result = 0
    for char in col_str.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result - 1  # Convert to 0-indexed


def _parse_a1_range(range_str: str) -> Dict:
    """
    Parse an A1 notation range string into row and column indices.

    Args:
        range_str: A1 notation range like 'A1:D10', 'Sheet1!A1:D10', 'A:D', '1:10'.

    Returns:
        Dict with keys: sheet_name (optional), start_row, end_row, start_col, end_col.
        Row/col values are 0-indexed. None values mean unbounded.

    Raises:
        ValueError: If the range format is invalid.
    """
    # Remove sheet name if present
    sheet_name = None
    if "!" in range_str:
        sheet_name, range_str = range_str.split("!", 1)
        # Remove quotes from sheet name if present
        sheet_name = sheet_name.strip("'\"")

    # Pattern for A1 notation: optional letters + optional numbers
    cell_pattern = r"([A-Za-z]*)(\d*)"

    if ":" in range_str:
        start_part, end_part = range_str.split(":", 1)
    else:
        start_part = range_str
        end_part = range_str

    start_match = re.fullmatch(cell_pattern, start_part)
    end_match = re.fullmatch(cell_pattern, end_part)

    if not start_match or not end_match:
        raise ValueError(f"Invalid A1 range format: {range_str}")

    start_col_str, start_row_str = start_match.groups()
    end_col_str, end_row_str = end_match.groups()

    # Convert to 0-indexed values
    start_col = _column_letter_to_index(start_col_str) if start_col_str else None
    end_col = (
        _column_letter_to_index(end_col_str) + 1 if end_col_str else None
    )  # +1 for exclusive end
    start_row = (
        int(start_row_str) - 1 if start_row_str else None
    )  # Convert to 0-indexed
    end_row = (
        int(end_row_str) if end_row_str else None
    )  # Already exclusive (1-indexed end = 0-indexed exclusive)

    return {
        "sheet_name": sheet_name,
        "start_row": start_row,
        "end_row": end_row,
        "start_col": start_col,
        "end_col": end_col,
    }


@server.tool()
@handle_http_errors("sort_range", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def sort_range(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    sort_column: int,
    ascending: bool = True,
    has_header_row: bool = False,
    secondary_sort_column: Optional[int] = None,
    secondary_ascending: Optional[bool] = None,
    sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Sorts data in a range of a Google Sheet by one or more columns.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to sort in A1 notation (e.g., "A1:D10", "Sheet1!A1:D10").
            Required. The range should include all columns you want to keep together during sort.
        sort_column (int): The 1-indexed column number to sort by within the range. Required.
            For example, if range is "B2:E10" and sort_column=1, sorts by column B.
        ascending (bool): Sort order. True for A-Z/smallest-first, False for Z-A/largest-first.
            Defaults to True.
        has_header_row (bool): If True, the first row of the range is treated as a header
            and excluded from sorting. Defaults to False.
        secondary_sort_column (Optional[int]): Optional 1-indexed column for secondary sort.
            Used when primary sort column has duplicate values.
        secondary_ascending (Optional[bool]): Sort order for secondary column.
            Defaults to same as 'ascending' if not specified.
        sheet_name (Optional[str]): The name of the sheet. If not provided, uses
            sheet from range_name, or first sheet if neither specified.
        sheet_id (Optional[int]): The ID of the sheet. Alternative to sheet_name.
            Takes precedence over sheet_name.

    Returns:
        str: Confirmation message of the successful sort operation.

    Example:
        Sort cells A1:D10 by column A (ascending), keeping header row in place:
        >>> sort_range(spreadsheet_id="...", range_name="A1:D10", sort_column=1,
        ...            ascending=True, has_header_row=True)
    """
    logger.info(
        f"[sort_range] Invoked. Email: '{user_google_email}', "
        f"Spreadsheet: {spreadsheet_id}, Range: {range_name}, "
        f"sort_column: {sort_column}, ascending: {ascending}"
    )

    if sort_column < 1:
        raise Exception("sort_column must be >= 1 (1-indexed).")

    if secondary_sort_column is not None and secondary_sort_column < 1:
        raise Exception("secondary_sort_column must be >= 1 (1-indexed).")

    # Parse the A1 notation range
    parsed_range = _parse_a1_range(range_name)

    # Determine sheet: priority is sheet_id > sheet_name param > sheet from range > first sheet
    effective_sheet_name = sheet_name or parsed_range.get("sheet_name")

    resolved_sheet_id = await _resolve_sheet_id(
        service, spreadsheet_id, effective_sheet_name, sheet_id
    )

    # Validate we have a proper range
    if parsed_range["start_col"] is None or parsed_range["end_col"] is None:
        raise Exception(
            "Range must specify column bounds (e.g., 'A1:D10'). "
            "Unbounded column ranges are not supported for sorting."
        )

    if parsed_range["start_row"] is None or parsed_range["end_row"] is None:
        raise Exception(
            "Range must specify row bounds (e.g., 'A1:D10'). "
            "Unbounded row ranges are not supported for sorting."
        )

    start_row = parsed_range["start_row"]
    end_row = parsed_range["end_row"]
    start_col = parsed_range["start_col"]
    end_col = parsed_range["end_col"]

    # If has_header_row, adjust start_row to exclude the header
    if has_header_row:
        if end_row - start_row < 2:
            raise Exception(
                "Range must have at least 2 rows when has_header_row=True "
                "(1 header + 1 data row)."
            )
        start_row += 1

    # Validate sort_column is within range
    range_width = end_col - start_col
    if sort_column > range_width:
        raise Exception(
            f"sort_column ({sort_column}) exceeds the range width ({range_width} columns). "
            f"sort_column must be between 1 and {range_width}."
        )

    if secondary_sort_column is not None and secondary_sort_column > range_width:
        raise Exception(
            f"secondary_sort_column ({secondary_sort_column}) exceeds the range width ({range_width} columns). "
            f"secondary_sort_column must be between 1 and {range_width}."
        )

    # Build sort specs - dimensionIndex is the absolute column index (0-based)
    sort_specs = [
        {
            "dimensionIndex": start_col
            + sort_column
            - 1,  # Convert 1-indexed relative to 0-indexed absolute
            "sortOrder": "ASCENDING" if ascending else "DESCENDING",
        }
    ]

    if secondary_sort_column is not None:
        sec_ascending = (
            secondary_ascending if secondary_ascending is not None else ascending
        )
        sort_specs.append(
            {
                "dimensionIndex": start_col + secondary_sort_column - 1,
                "sortOrder": "ASCENDING" if sec_ascending else "DESCENDING",
            }
        )

    request_body = {
        "requests": [
            {
                "sortRange": {
                    "range": {
                        "sheetId": resolved_sheet_id,
                        "startRowIndex": start_row,
                        "endRowIndex": end_row,
                        "startColumnIndex": start_col,
                        "endColumnIndex": end_col,
                    },
                    "sortSpecs": sort_specs,
                }
            }
        ]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    # Build user-friendly message
    sort_order = "ascending" if ascending else "descending"
    header_note = " (excluding header row)" if has_header_row else ""

    sort_description = f"column {sort_column} ({sort_order})"
    if secondary_sort_column is not None:
        sec_order = (
            "ascending"
            if (secondary_ascending if secondary_ascending is not None else ascending)
            else "descending"
        )
        sort_description += f", then column {secondary_sort_column} ({sec_order})"

    sheet_identifier = effective_sheet_name or f"sheet ID {resolved_sheet_id}"
    text_output = (
        f"Successfully sorted range '{range_name}' by {sort_description}{header_note} "
        f"in '{sheet_identifier}' of spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully sorted range '{range_name}' for {user_google_email}.")
    return text_output


# Create comment management tools for sheets
_comment_tools = create_comment_tools("spreadsheet", "spreadsheet_id")

# Extract and register the functions
read_sheet_comments = _comment_tools["read_comments"]
create_sheet_comment = _comment_tools["create_comment"]
reply_to_sheet_comment = _comment_tools["reply_to_comment"]
resolve_sheet_comment = _comment_tools["resolve_comment"]
