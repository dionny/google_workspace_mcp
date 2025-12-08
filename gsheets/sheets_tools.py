"""
Google Sheets MCP Tools

This module provides MCP tools for interacting with Google Sheets API.
"""

import logging
import asyncio
import json
from typing import List, Optional, Union


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
    sheet_names: Optional[Union[str, List[str]]] = None,
) -> str:
    """
    Creates a new Google Spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        title (str): The title of the new spreadsheet. Required.
        sheet_names (Optional[Union[str, List[str]]]): List of sheet names to create. Can be a JSON string or Python list. If not provided, creates one sheet with default name.

    Returns:
        str: Information about the newly created spreadsheet including ID and URL.
    """
    logger.info(
        f"[create_spreadsheet] Invoked. Email: '{user_google_email}', Title: {title}"
    )

    # Parse sheet_names if it's a JSON string (MCP passes parameters as JSON strings)
    if sheet_names is not None and isinstance(sheet_names, str):
        try:
            parsed_sheet_names = json.loads(sheet_names)
            if not isinstance(parsed_sheet_names, list):
                raise ValueError(
                    f"sheet_names must be a list, got {type(parsed_sheet_names).__name__}"
                )
            # Validate each sheet name is a string
            for i, name in enumerate(parsed_sheet_names):
                if not isinstance(name, str):
                    raise ValueError(
                        f"Sheet name at index {i} must be a string, got {type(name).__name__}"
                    )
            sheet_names = parsed_sheet_names
            logger.info(
                f"[create_spreadsheet] Parsed JSON string to Python list with {len(sheet_names)} sheet names"
            )
        except json.JSONDecodeError as e:
            raise Exception(f"Invalid JSON format for sheet_names: {e}")
        except ValueError as e:
            raise Exception(f"Invalid sheet_names structure: {e}")

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


@server.tool()
@handle_http_errors("append_rows", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def append_rows(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    values: Union[str, List[List[str]]],
    value_input_option: str = "USER_ENTERED",
    insert_data_option: str = "INSERT_ROWS",
) -> str:
    """
    Appends rows of data to the end of existing data in a sheet.

    This is the easiest way to add new data to a spreadsheet - no need to
    calculate the next empty row. The API automatically finds the end of
    existing data and appends new rows there.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to append to (e.g., "Sheet1", "Sheet1!A:A", "Sheet1!A1:D1"). Required.
            This defines where to search for existing data and where to append.
            Using just the sheet name (e.g., "Sheet1") will append to the first table found.
        values (Union[str, List[List[str]]]): 2D array of values to append. Can be a JSON string or Python list. Required.
            Each inner list represents a row. Example: [["Name", "Age"], ["Alice", "30"]]
        value_input_option (str): How to interpret input values. Defaults to "USER_ENTERED".
            "USER_ENTERED": Values are parsed as if typed by a user (formulas executed, numbers/dates parsed).
            "RAW": Values are stored exactly as provided.
        insert_data_option (str): How to insert the data. Defaults to "INSERT_ROWS".
            "INSERT_ROWS": Insert new rows for the appended data.
            "OVERWRITE": Overwrite existing data in the range (rarely used for append).

    Returns:
        str: Confirmation message including the range where data was appended.

    Example:
        # Append a single row
        append_rows(spreadsheet_id="abc123", range_name="Sheet1", values=[["New Name", "25", "Seattle"]])

        # Append multiple rows
        append_rows(spreadsheet_id="abc123", range_name="Data!A:C", values=[["Row1Col1", "Row1Col2"], ["Row2Col1", "Row2Col2"]])
    """
    logger.info(
        f"[append_rows] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
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
                f"[append_rows] Parsed JSON string to Python list with {len(values)} rows"
            )
        except json.JSONDecodeError as e:
            raise Exception(f"Invalid JSON format for values: {e}")
        except ValueError as e:
            raise Exception(f"Invalid values structure: {e}")

    if not values:
        raise Exception("'values' must be provided and non-empty.")

    body = {"values": values}

    result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption=value_input_option,
            insertDataOption=insert_data_option,
            body=body,
        )
        .execute
    )

    updates = result.get("updates", {})
    updated_range = updates.get("updatedRange", "unknown range")
    updated_rows = updates.get("updatedRows", 0)
    updated_cells = updates.get("updatedCells", 0)

    text_output = (
        f"Successfully appended {updated_rows} row(s) ({updated_cells} cells) to spreadsheet {spreadsheet_id} for {user_google_email}. "
        f"Data written to range: {updated_range}"
    )

    logger.info(
        f"Successfully appended {updated_rows} row(s) for {user_google_email}. Range: {updated_range}"
    )
    return text_output


@server.tool()
@handle_http_errors("insert_rows", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def insert_rows(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_id: int,
    start_index: int,
    count: int = 1,
    inherit_from_before: bool = True,
) -> str:
    """
    Inserts blank rows at a specific position in a sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_id (int): The ID of the sheet (not the sheet name). Required. Use get_spreadsheet_info to find sheet IDs.
        start_index (int): The row index to insert at (0-based). Required. Rows will be inserted starting at this index.
        count (int): Number of rows to insert. Defaults to 1.
        inherit_from_before (bool): If True, new rows inherit formatting from the row above. Defaults to True.

    Returns:
        str: Confirmation message of the successful row insertion.
    """
    logger.info(
        f"[insert_rows] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Sheet ID: {sheet_id}, Start: {start_index}, Count: {count}"
    )

    request_body = {
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_index,
                        "endIndex": start_index + count,
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
        f"Successfully inserted {count} row(s) at index {start_index} in sheet ID {sheet_id} "
        f"of spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully inserted {count} row(s) for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("delete_rows", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def delete_rows(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_id: int,
    start_index: int,
    end_index: int,
) -> str:
    """
    Deletes rows from a sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_id (int): The ID of the sheet (not the sheet name). Required. Use get_spreadsheet_info to find sheet IDs.
        start_index (int): The starting row index to delete (0-based, inclusive). Required.
        end_index (int): The ending row index (0-based, exclusive). Required. Rows from start_index to end_index-1 will be deleted.

    Returns:
        str: Confirmation message of the successful row deletion.
    """
    logger.info(
        f"[delete_rows] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Sheet ID: {sheet_id}, Start: {start_index}, End: {end_index}"
    )

    if start_index >= end_index:
        raise Exception(
            f"start_index ({start_index}) must be less than end_index ({end_index})"
        )

    request_body = {
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
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

    rows_deleted = end_index - start_index
    text_output = (
        f"Successfully deleted {rows_deleted} row(s) (index {start_index} to {end_index - 1}) "
        f"from sheet ID {sheet_id} of spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully deleted {rows_deleted} row(s) for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("insert_columns", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def insert_columns(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_id: int,
    start_index: int,
    count: int = 1,
    inherit_from_before: bool = True,
) -> str:
    """
    Inserts blank columns at a specific position in a sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_id (int): The ID of the sheet (not the sheet name). Required. Use get_spreadsheet_info to find sheet IDs.
        start_index (int): The column index to insert at (0-based, A=0, B=1, etc.). Required. Columns will be inserted starting at this index.
        count (int): Number of columns to insert. Defaults to 1.
        inherit_from_before (bool): If True, new columns inherit formatting from the column to the left. Defaults to True.

    Returns:
        str: Confirmation message of the successful column insertion.
    """
    logger.info(
        f"[insert_columns] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Sheet ID: {sheet_id}, Start: {start_index}, Count: {count}"
    )

    request_body = {
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": start_index,
                        "endIndex": start_index + count,
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
        f"Successfully inserted {count} column(s) at index {start_index} in sheet ID {sheet_id} "
        f"of spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully inserted {count} column(s) for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("delete_columns", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def delete_columns(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_id: int,
    start_index: int,
    end_index: int,
) -> str:
    """
    Deletes columns from a sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_id (int): The ID of the sheet (not the sheet name). Required. Use get_spreadsheet_info to find sheet IDs.
        start_index (int): The starting column index to delete (0-based, A=0, B=1, etc., inclusive). Required.
        end_index (int): The ending column index (0-based, exclusive). Required. Columns from start_index to end_index-1 will be deleted.

    Returns:
        str: Confirmation message of the successful column deletion.
    """
    logger.info(
        f"[delete_columns] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Sheet ID: {sheet_id}, Start: {start_index}, End: {end_index}"
    )

    if start_index >= end_index:
        raise Exception(
            f"start_index ({start_index}) must be less than end_index ({end_index})"
        )

    request_body = {
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
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

    columns_deleted = end_index - start_index
    text_output = (
        f"Successfully deleted {columns_deleted} column(s) (index {start_index} to {end_index - 1}) "
        f"from sheet ID {sheet_id} of spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(
        f"Successfully deleted {columns_deleted} column(s) for {user_google_email}."
    )
    return text_output


@server.tool()
@handle_http_errors("delete_sheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def delete_sheet(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_id: int,
) -> str:
    """
    Deletes a sheet from a spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_id (int): The ID of the sheet to delete (not the sheet name). Required.
            Use get_spreadsheet_info to find sheet IDs.

    Returns:
        str: Confirmation message of the successful sheet deletion.

    Note:
        A spreadsheet must have at least one sheet. Attempting to delete the last
        remaining sheet will result in an error.
    """
    logger.info(
        f"[delete_sheet] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Sheet ID: {sheet_id}"
    )

    request_body = {"requests": [{"deleteSheet": {"sheetId": sheet_id}}]}

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    text_output = (
        f"Successfully deleted sheet ID {sheet_id} from spreadsheet {spreadsheet_id} "
        f"for {user_google_email}."
    )

    logger.info(f"Successfully deleted sheet ID {sheet_id} for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("rename_sheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def rename_sheet(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_id: int,
    new_name: str,
) -> str:
    """
    Renames a sheet within a spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_id (int): The ID of the sheet to rename (not the sheet name). Required.
            Use get_spreadsheet_info to find sheet IDs.
        new_name (str): The new name for the sheet. Required.

    Returns:
        str: Confirmation message of the successful sheet rename.
    """
    logger.info(
        f"[rename_sheet] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, "
        f"Sheet ID: {sheet_id}, New Name: '{new_name}'"
    )

    request_body = {
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": sheet_id, "title": new_name},
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
        f"Successfully renamed sheet ID {sheet_id} to '{new_name}' in spreadsheet {spreadsheet_id} "
        f"for {user_google_email}."
    )

    logger.info(
        f"Successfully renamed sheet ID {sheet_id} to '{new_name}' for {user_google_email}."
    )
    return text_output


@server.tool()
@handle_http_errors("copy_sheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def copy_sheet(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_id: int,
    new_name: Optional[str] = None,
) -> str:
    """
    Creates a copy of a sheet within the same spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_id (int): The ID of the sheet to copy (not the sheet name). Required.
            Use get_spreadsheet_info to find sheet IDs.
        new_name (Optional[str]): The name for the new sheet. If not provided,
            Google Sheets will generate a name like "Copy of [original name]".

    Returns:
        str: Confirmation message including the new sheet's ID and name.
    """
    logger.info(
        f"[copy_sheet] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, "
        f"Sheet ID: {sheet_id}, New Name: {new_name or '(auto-generated)'}"
    )

    duplicate_request = {"sourceSheetId": sheet_id}
    if new_name:
        duplicate_request["newSheetName"] = new_name

    request_body = {"requests": [{"duplicateSheet": duplicate_request}]}

    response = await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    # Extract the new sheet info from the response
    new_sheet_props = response["replies"][0]["duplicateSheet"]["properties"]
    new_sheet_id = new_sheet_props["sheetId"]
    new_sheet_title = new_sheet_props["title"]

    text_output = (
        f"Successfully copied sheet ID {sheet_id} to new sheet '{new_sheet_title}' (ID: {new_sheet_id}) "
        f"in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(
        f"Successfully copied sheet ID {sheet_id} to new sheet ID {new_sheet_id} for {user_google_email}."
    )
    return text_output


def _parse_hex_color(hex_color: str) -> dict:
    """
    Parse a hex color string to Google Sheets API color format.

    Args:
        hex_color: Hex color string (e.g., '#FF0000' or 'FF0000')

    Returns:
        dict with red, green, blue values as floats (0.0 to 1.0)
    """
    # Remove # prefix if present
    hex_color = hex_color.lstrip("#")

    if len(hex_color) != 6:
        raise ValueError(
            f"Invalid hex color format: {hex_color}. Expected 6 hex digits."
        )

    try:
        r = int(hex_color[0:2], 16) / 255.0
        g = int(hex_color[2:4], 16) / 255.0
        b = int(hex_color[4:6], 16) / 255.0
    except ValueError:
        raise ValueError(
            f"Invalid hex color format: {hex_color}. Must contain valid hex digits."
        )

    return {"red": r, "green": g, "blue": b}


def _parse_range_to_grid_range(range_name: str, sheet_id: int) -> dict:
    """
    Parse an A1 notation range to a GridRange object.

    Args:
        range_name: Range in A1 notation (e.g., 'A1:C10', 'Sheet1!A1:C10')
        sheet_id: The sheet ID to use in the GridRange

    Returns:
        dict representing a GridRange for the Sheets API
    """
    import re

    # Remove sheet name prefix if present (e.g., "Sheet1!A1:C10" -> "A1:C10")
    if "!" in range_name:
        range_name = range_name.split("!")[-1]

    # Parse the range (e.g., "A1:C10" or "A1" or "A:C")
    # Pattern matches: column letters + optional row number
    cell_pattern = r"([A-Za-z]+)(\d*)"

    if ":" in range_name:
        start, end = range_name.split(":")
        start_match = re.match(cell_pattern, start)
        end_match = re.match(cell_pattern, end)

        if not start_match or not end_match:
            raise ValueError(f"Invalid range format: {range_name}")

        start_col = start_match.group(1).upper()
        start_row = start_match.group(2)
        end_col = end_match.group(1).upper()
        end_row = end_match.group(2)
    else:
        # Single cell
        match = re.match(cell_pattern, range_name)
        if not match:
            raise ValueError(f"Invalid range format: {range_name}")
        start_col = end_col = match.group(1).upper()
        start_row = end_row = match.group(2)

    def col_to_index(col: str) -> int:
        """Convert column letter(s) to 0-based index (A=0, B=1, ..., Z=25, AA=26)."""
        result = 0
        for char in col:
            result = result * 26 + (ord(char) - ord("A") + 1)
        return result - 1

    grid_range = {"sheetId": sheet_id}

    # Column indices (always present)
    grid_range["startColumnIndex"] = col_to_index(start_col)
    grid_range["endColumnIndex"] = col_to_index(end_col) + 1

    # Row indices (only if row numbers are specified)
    if start_row:
        grid_range["startRowIndex"] = int(start_row) - 1  # Convert to 0-based
    if end_row:
        grid_range["endRowIndex"] = int(end_row)  # End is exclusive, so no -1

    return grid_range


@server.tool()
@handle_http_errors("format_cells", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def format_cells(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_id: int,
    range_name: str,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    strikethrough: Optional[bool] = None,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
    text_color: Optional[str] = None,
    background_color: Optional[str] = None,
    horizontal_alignment: Optional[str] = None,
    vertical_alignment: Optional[str] = None,
    number_format_type: Optional[str] = None,
    number_format_pattern: Optional[str] = None,
    wrap_strategy: Optional[str] = None,
) -> str:
    """
    Applies formatting to a range of cells in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_id (int): The ID of the sheet (not the sheet name). Required.
            Use get_spreadsheet_info to find sheet IDs.
        range_name (str): The range to format in A1 notation (e.g., 'A1:C10', 'A1'). Required.
            Do not include the sheet name prefix; use sheet_id instead.
        bold (Optional[bool]): Make text bold.
        italic (Optional[bool]): Make text italic.
        underline (Optional[bool]): Underline text.
        strikethrough (Optional[bool]): Strikethrough text.
        font_size (Optional[int]): Font size in points (e.g., 10, 12, 14).
        font_family (Optional[str]): Font family name (e.g., 'Arial', 'Times New Roman').
        text_color (Optional[str]): Text color as hex string (e.g., '#FF0000' for red).
        background_color (Optional[str]): Cell background color as hex string (e.g., '#FFFF00' for yellow).
        horizontal_alignment (Optional[str]): Horizontal alignment: 'LEFT', 'CENTER', or 'RIGHT'.
        vertical_alignment (Optional[str]): Vertical alignment: 'TOP', 'MIDDLE', or 'BOTTOM'.
        number_format_type (Optional[str]): Number format type: 'TEXT', 'NUMBER', 'PERCENT',
            'CURRENCY', 'DATE', 'TIME', 'DATE_TIME', or 'SCIENTIFIC'.
        number_format_pattern (Optional[str]): Custom number format pattern (e.g., '#,##0.00',
            'yyyy-mm-dd', '$#,##0.00'). Used with number_format_type.
        wrap_strategy (Optional[str]): Text wrap strategy: 'OVERFLOW_CELL', 'LEGACY_WRAP',
            'CLIP', or 'WRAP'.

    Returns:
        str: Confirmation message of the successful formatting operation.

    Example:
        # Bold header row with gray background
        format_cells(spreadsheet_id="abc123", sheet_id=0, range_name="A1:Z1",
                     bold=True, background_color="#E0E0E0")

        # Currency formatting
        format_cells(spreadsheet_id="abc123", sheet_id=0, range_name="B2:B100",
                     number_format_type="CURRENCY", number_format_pattern="$#,##0.00")

        # Date formatting with center alignment
        format_cells(spreadsheet_id="abc123", sheet_id=0, range_name="A2:A100",
                     number_format_type="DATE", number_format_pattern="yyyy-mm-dd",
                     horizontal_alignment="CENTER")
    """
    logger.info(
        f"[format_cells] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, "
        f"Sheet ID: {sheet_id}, Range: {range_name}"
    )

    # Build the cell format object
    cell_format = {}
    fields = []

    # Text format options
    text_format = {}
    if bold is not None:
        text_format["bold"] = bold
    if italic is not None:
        text_format["italic"] = italic
    if underline is not None:
        text_format["underline"] = underline
    if strikethrough is not None:
        text_format["strikethrough"] = strikethrough
    if font_size is not None:
        text_format["fontSize"] = font_size
    if font_family is not None:
        text_format["fontFamily"] = font_family
    if text_color is not None:
        text_format["foregroundColor"] = _parse_hex_color(text_color)

    if text_format:
        cell_format["textFormat"] = text_format
        # Build specific fields for text format
        text_fields = []
        if bold is not None:
            text_fields.append("bold")
        if italic is not None:
            text_fields.append("italic")
        if underline is not None:
            text_fields.append("underline")
        if strikethrough is not None:
            text_fields.append("strikethrough")
        if font_size is not None:
            text_fields.append("fontSize")
        if font_family is not None:
            text_fields.append("fontFamily")
        if text_color is not None:
            text_fields.append("foregroundColor")
        fields.extend([f"userEnteredFormat.textFormat.{f}" for f in text_fields])

    # Background color
    if background_color is not None:
        cell_format["backgroundColor"] = _parse_hex_color(background_color)
        fields.append("userEnteredFormat.backgroundColor")

    # Horizontal alignment
    if horizontal_alignment is not None:
        valid_h_alignments = ["LEFT", "CENTER", "RIGHT"]
        if horizontal_alignment.upper() not in valid_h_alignments:
            raise ValueError(
                f"Invalid horizontal_alignment: {horizontal_alignment}. "
                f"Must be one of: {valid_h_alignments}"
            )
        cell_format["horizontalAlignment"] = horizontal_alignment.upper()
        fields.append("userEnteredFormat.horizontalAlignment")

    # Vertical alignment
    if vertical_alignment is not None:
        valid_v_alignments = ["TOP", "MIDDLE", "BOTTOM"]
        if vertical_alignment.upper() not in valid_v_alignments:
            raise ValueError(
                f"Invalid vertical_alignment: {vertical_alignment}. "
                f"Must be one of: {valid_v_alignments}"
            )
        cell_format["verticalAlignment"] = vertical_alignment.upper()
        fields.append("userEnteredFormat.verticalAlignment")

    # Number format
    if number_format_type is not None or number_format_pattern is not None:
        number_format = {}
        if number_format_type is not None:
            valid_types = [
                "TEXT",
                "NUMBER",
                "PERCENT",
                "CURRENCY",
                "DATE",
                "TIME",
                "DATE_TIME",
                "SCIENTIFIC",
            ]
            if number_format_type.upper() not in valid_types:
                raise ValueError(
                    f"Invalid number_format_type: {number_format_type}. "
                    f"Must be one of: {valid_types}"
                )
            number_format["type"] = number_format_type.upper()
        if number_format_pattern is not None:
            number_format["pattern"] = number_format_pattern
        cell_format["numberFormat"] = number_format
        fields.append("userEnteredFormat.numberFormat")

    # Wrap strategy
    if wrap_strategy is not None:
        valid_wrap = ["OVERFLOW_CELL", "LEGACY_WRAP", "CLIP", "WRAP"]
        if wrap_strategy.upper() not in valid_wrap:
            raise ValueError(
                f"Invalid wrap_strategy: {wrap_strategy}. Must be one of: {valid_wrap}"
            )
        cell_format["wrapStrategy"] = wrap_strategy.upper()
        fields.append("userEnteredFormat.wrapStrategy")

    if not cell_format:
        raise ValueError("At least one formatting option must be specified.")

    # Parse the range to a GridRange
    grid_range = _parse_range_to_grid_range(range_name, sheet_id)

    # Build the request
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

    # Build a description of applied formatting
    format_descriptions = []
    if bold is not None:
        format_descriptions.append(f"bold={bold}")
    if italic is not None:
        format_descriptions.append(f"italic={italic}")
    if underline is not None:
        format_descriptions.append(f"underline={underline}")
    if strikethrough is not None:
        format_descriptions.append(f"strikethrough={strikethrough}")
    if font_size is not None:
        format_descriptions.append(f"fontSize={font_size}")
    if font_family is not None:
        format_descriptions.append(f"fontFamily={font_family}")
    if text_color is not None:
        format_descriptions.append(f"textColor={text_color}")
    if background_color is not None:
        format_descriptions.append(f"backgroundColor={background_color}")
    if horizontal_alignment is not None:
        format_descriptions.append(f"horizontalAlignment={horizontal_alignment}")
    if vertical_alignment is not None:
        format_descriptions.append(f"verticalAlignment={vertical_alignment}")
    if number_format_type is not None:
        format_descriptions.append(f"numberFormatType={number_format_type}")
    if number_format_pattern is not None:
        format_descriptions.append(f"numberFormatPattern={number_format_pattern}")
    if wrap_strategy is not None:
        format_descriptions.append(f"wrapStrategy={wrap_strategy}")

    format_summary = ", ".join(format_descriptions)

    text_output = (
        f"Successfully formatted range '{range_name}' in sheet ID {sheet_id} "
        f"of spreadsheet {spreadsheet_id} for {user_google_email}. "
        f"Applied: {format_summary}"
    )

    logger.info(
        f"Successfully formatted range '{range_name}' for {user_google_email}. "
        f"Applied: {format_summary}"
    )
    return text_output


@server.tool()
@handle_http_errors("sort_range", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def sort_range(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_id: int,
    range_name: str,
    sort_specs: Union[str, List[dict]],
) -> str:
    """
    Sorts data in a range of cells in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_id (int): The ID of the sheet (not the sheet name). Required.
            Use get_spreadsheet_info to find sheet IDs.
        range_name (str): The range to sort in A1 notation (e.g., 'A1:D100', 'A2:C50'). Required.
            Do not include the sheet name prefix; use sheet_id instead.
        sort_specs (Union[str, List[dict]]): JSON array or Python list of sort specifications. Required.
            Each specification must contain:
            - column_index (int): The column index to sort by, relative to the range start (0-based).
              For example, if range is 'B1:D10', column_index 0 refers to column B.
            - order (str): Sort order - 'ASCENDING' or 'DESCENDING'.
            Multiple sort specs create multi-level sorting (sort by first, then by second, etc.).

    Returns:
        str: Confirmation message of the successful sort operation.

    Example:
        # Sort by first column ascending
        sort_range(spreadsheet_id="abc123", sheet_id=0, range_name="A1:D100",
                   sort_specs='[{"column_index": 0, "order": "ASCENDING"}]')

        # Multi-column sort: by column A ascending, then by column C descending
        sort_range(spreadsheet_id="abc123", sheet_id=0, range_name="A1:D100",
                   sort_specs='[{"column_index": 0, "order": "ASCENDING"}, {"column_index": 2, "order": "DESCENDING"}]')
    """
    logger.info(
        f"[sort_range] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, "
        f"Sheet ID: {sheet_id}, Range: {range_name}"
    )

    # Parse sort_specs if it's a JSON string
    if isinstance(sort_specs, str):
        try:
            parsed_sort_specs = json.loads(sort_specs)
            if not isinstance(parsed_sort_specs, list):
                raise ValueError(
                    f"sort_specs must be a list, got {type(parsed_sort_specs).__name__}"
                )
            sort_specs = parsed_sort_specs
            logger.info(
                f"[sort_range] Parsed JSON string to Python list with {len(sort_specs)} sort specs"
            )
        except json.JSONDecodeError as e:
            raise Exception(f"Invalid JSON format for sort_specs: {e}")
        except ValueError as e:
            raise Exception(f"Invalid sort_specs structure: {e}")

    if not sort_specs:
        raise Exception("sort_specs must not be empty.")

    # Validate and transform sort specs
    valid_orders = ["ASCENDING", "DESCENDING"]
    api_sort_specs = []

    for i, spec in enumerate(sort_specs):
        if not isinstance(spec, dict):
            raise Exception(
                f"sort_specs[{i}] must be a dict, got {type(spec).__name__}"
            )

        if "column_index" not in spec:
            raise Exception(f"sort_specs[{i}] missing required 'column_index' field")
        if "order" not in spec:
            raise Exception(f"sort_specs[{i}] missing required 'order' field")

        column_index = spec["column_index"]
        order = spec["order"]

        if not isinstance(column_index, int):
            raise Exception(
                f"sort_specs[{i}]['column_index'] must be an int, got {type(column_index).__name__}"
            )
        if column_index < 0:
            raise Exception(
                f"sort_specs[{i}]['column_index'] must be >= 0, got {column_index}"
            )

        if not isinstance(order, str):
            raise Exception(
                f"sort_specs[{i}]['order'] must be a string, got {type(order).__name__}"
            )
        if order.upper() not in valid_orders:
            raise Exception(
                f"sort_specs[{i}]['order'] must be one of {valid_orders}, got '{order}'"
            )

        api_sort_specs.append(
            {"dimensionIndex": column_index, "sortOrder": order.upper()}
        )

    # Parse the range to a GridRange
    grid_range = _parse_range_to_grid_range(range_name, sheet_id)

    # Build the sortRange request
    request_body = {
        "requests": [{"sortRange": {"range": grid_range, "sortSpecs": api_sort_specs}}]
    }

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    # Build description of sort operation
    sort_descriptions = []
    for spec in sort_specs:
        sort_descriptions.append(
            f"column {spec['column_index']} {spec['order'].upper()}"
        )
    sort_summary = ", ".join(sort_descriptions)

    text_output = (
        f"Successfully sorted range '{range_name}' in sheet ID {sheet_id} "
        f"of spreadsheet {spreadsheet_id} for {user_google_email}. "
        f"Sort order: {sort_summary}"
    )

    logger.info(
        f"Successfully sorted range '{range_name}' for {user_google_email}. "
        f"Sort order: {sort_summary}"
    )
    return text_output


# Create comment management tools for sheets
_comment_tools = create_comment_tools("spreadsheet", "spreadsheet_id")

# Extract and register the functions
read_sheet_comments = _comment_tools["read_comments"]
create_sheet_comment = _comment_tools["create_comment"]
reply_to_sheet_comment = _comment_tools["reply_to_comment"]
resolve_sheet_comment = _comment_tools["resolve_comment"]
