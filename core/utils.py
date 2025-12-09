import io
import logging
import os
import zipfile
import xml.etree.ElementTree as ET
import ssl
import asyncio
import functools

from typing import List, Optional

from googleapiclient.errors import HttpError
from .api_enablement import get_api_enablement_message
from auth.google_auth import GoogleAuthenticationError

logger = logging.getLogger(__name__)


class TransientNetworkError(Exception):
    """Custom exception for transient network errors after retries."""

    pass


def check_credentials_directory_permissions(credentials_dir: str = None) -> None:
    """
    Check if the service has appropriate permissions to create and write to the .credentials directory.

    Args:
        credentials_dir: Path to the credentials directory (default: uses get_default_credentials_dir())

    Raises:
        PermissionError: If the service lacks necessary permissions
        OSError: If there are other file system issues
    """
    if credentials_dir is None:
        from auth.google_auth import get_default_credentials_dir

        credentials_dir = get_default_credentials_dir()

    try:
        # Check if directory exists
        if os.path.exists(credentials_dir):
            # Directory exists, check if we can write to it
            test_file = os.path.join(credentials_dir, ".permission_test")
            try:
                with open(test_file, "w") as f:
                    f.write("test")
                os.remove(test_file)
                logger.info(
                    f"Credentials directory permissions check passed: {os.path.abspath(credentials_dir)}"
                )
            except (PermissionError, OSError) as e:
                raise PermissionError(
                    f"Cannot write to existing credentials directory '{os.path.abspath(credentials_dir)}': {e}"
                )
        else:
            # Directory doesn't exist, try to create it and its parent directories
            try:
                os.makedirs(credentials_dir, exist_ok=True)
                # Test writing to the new directory
                test_file = os.path.join(credentials_dir, ".permission_test")
                with open(test_file, "w") as f:
                    f.write("test")
                os.remove(test_file)
                logger.info(
                    f"Created credentials directory with proper permissions: {os.path.abspath(credentials_dir)}"
                )
            except (PermissionError, OSError) as e:
                # Clean up if we created the directory but can't write to it
                try:
                    if os.path.exists(credentials_dir):
                        os.rmdir(credentials_dir)
                except (PermissionError, OSError):
                    pass
                raise PermissionError(
                    f"Cannot create or write to credentials directory '{os.path.abspath(credentials_dir)}': {e}"
                )

    except PermissionError:
        raise
    except Exception as e:
        raise OSError(
            f"Unexpected error checking credentials directory permissions: {e}"
        )


def extract_office_xml_text(file_bytes: bytes, mime_type: str) -> Optional[str]:
    """
    Very light-weight XML scraper for Word, Excel, PowerPoint files.
    Returns plain-text if something readable is found, else None.
    No external deps – just std-lib zipfile + ElementTree.
    """
    shared_strings: List[str] = []
    ns_excel_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            targets: List[str] = []
            # Map MIME → iterable of XML files to inspect
            if (
                mime_type
                == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ):
                targets = ["word/document.xml"]
            elif (
                mime_type
                == "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ):
                targets = [n for n in zf.namelist() if n.startswith("ppt/slides/slide")]
            elif (
                mime_type
                == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ):
                targets = [
                    n
                    for n in zf.namelist()
                    if n.startswith("xl/worksheets/sheet") and "drawing" not in n
                ]
                # Attempt to parse sharedStrings.xml for Excel files
                try:
                    shared_strings_xml = zf.read("xl/sharedStrings.xml")
                    shared_strings_root = ET.fromstring(shared_strings_xml)
                    for si_element in shared_strings_root.findall(
                        f"{{{ns_excel_main}}}si"
                    ):
                        text_parts = []
                        # Find all <t> elements, simple or within <r> runs, and concatenate their text
                        for t_element in si_element.findall(f".//{{{ns_excel_main}}}t"):
                            if t_element.text:
                                text_parts.append(t_element.text)
                        shared_strings.append("".join(text_parts))
                except KeyError:
                    logger.info(
                        "No sharedStrings.xml found in Excel file (this is optional)."
                    )
                except ET.ParseError as e:
                    logger.error(f"Error parsing sharedStrings.xml: {e}")
                except (
                    Exception
                ) as e:  # Catch any other unexpected error during sharedStrings parsing
                    logger.error(
                        f"Unexpected error processing sharedStrings.xml: {e}",
                        exc_info=True,
                    )
            else:
                return None

            pieces: List[str] = []
            for member in targets:
                try:
                    xml_content = zf.read(member)
                    xml_root = ET.fromstring(xml_content)
                    member_texts: List[str] = []

                    if (
                        mime_type
                        == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ):
                        for cell_element in xml_root.findall(
                            f".//{{{ns_excel_main}}}c"
                        ):  # Find all <c> elements
                            value_element = cell_element.find(
                                f"{{{ns_excel_main}}}v"
                            )  # Find <v> under <c>

                            # Skip if cell has no value element or value element has no text
                            if value_element is None or value_element.text is None:
                                continue

                            cell_type = cell_element.get("t")
                            if cell_type == "s":  # Shared string
                                try:
                                    ss_idx = int(value_element.text)
                                    if 0 <= ss_idx < len(shared_strings):
                                        member_texts.append(shared_strings[ss_idx])
                                    else:
                                        logger.warning(
                                            f"Invalid shared string index {ss_idx} in {member}. Max index: {len(shared_strings) - 1}"
                                        )
                                except ValueError:
                                    logger.warning(
                                        f"Non-integer shared string index: '{value_element.text}' in {member}."
                                    )
                            else:  # Direct value (number, boolean, inline string if not 's')
                                member_texts.append(value_element.text)
                    else:  # Word or PowerPoint
                        for elem in xml_root.iter():
                            # For Word: <w:t> where w is "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            # For PowerPoint: <a:t> where a is "http://schemas.openxmlformats.org/drawingml/2006/main"
                            if (
                                elem.tag.endswith("}t") and elem.text
                            ):  # Check for any namespaced tag ending with 't'
                                cleaned_text = elem.text.strip()
                                if (
                                    cleaned_text
                                ):  # Add only if there's non-whitespace text
                                    member_texts.append(cleaned_text)

                    if member_texts:
                        pieces.append(
                            " ".join(member_texts)
                        )  # Join texts from one member with spaces

                except ET.ParseError as e:
                    logger.warning(
                        f"Could not parse XML in member '{member}' for {mime_type} file: {e}"
                    )
                except Exception as e:
                    logger.error(
                        f"Error processing member '{member}' for {mime_type}: {e}",
                        exc_info=True,
                    )
                    # continue processing other members

            if not pieces:  # If no text was extracted at all
                return None

            # Join content from different members (sheets/slides) with double newlines for separation
            text = "\n\n".join(pieces).strip()
            return text or None  # Ensure None is returned if text is empty after strip

    except zipfile.BadZipFile:
        logger.warning(f"File is not a valid ZIP archive (mime_type: {mime_type}).")
        return None
    except (
        ET.ParseError
    ) as e:  # Catch parsing errors at the top level if zipfile itself is XML-like
        logger.error(f"XML parsing error at a high level for {mime_type}: {e}")
        return None
    except Exception as e:
        logger.error(
            f"Failed to extract office XML text for {mime_type}: {e}", exc_info=True
        )
        return None


def _parse_docs_index_error(error_details: str) -> Optional[str]:
    """
    Parse Google Docs API error details to detect index out-of-bounds errors.

    Returns a structured error JSON string if an index error is detected, None otherwise.

    Common patterns:
    - "Index X must be less than the end index of the referenced segment, Y"
    - "The insertion index must be inside the bounds of an existing paragraph"
    """
    import re
    import json

    # Pattern 1: "Index X must be less than the end index of the referenced segment, Y"
    match = re.search(
        r"Index\s+(\d+)\s+must be less than the end index of the referenced segment,?\s*(\d+)?",
        error_details,
        re.IGNORECASE,
    )
    if match:
        index_value = int(match.group(1))
        doc_length = int(match.group(2)) if match.group(2) else None

        error_response = {
            "error": True,
            "code": "INDEX_OUT_OF_BOUNDS",
            "message": f"Index {index_value} exceeds document length"
            + (f" ({doc_length})" if doc_length else ""),
            "reason": f"The requested index {index_value} is beyond the end of the document."
            + (
                f" Document length is {doc_length} characters (valid indices: 1 to {doc_length - 1})."
                if doc_length
                else ""
            ),
            "suggestion": "Use inspect_doc_structure to check document length before editing. "
            "Valid indices are from 1 to document_length - 1.",
            "example": {
                "check_length": "inspect_doc_structure(document_id='...')",
                "valid_usage": f"Use an index less than {doc_length}"
                if doc_length
                else "Use a valid index within document bounds",
            },
            "context": {
                "received": {"index": index_value},
            },
        }
        if doc_length:
            error_response["context"]["document_length"] = doc_length

        return json.dumps(error_response, indent=2)

    # Pattern 2: "The insertion index must be inside the bounds of an existing paragraph"
    if "insertion index must be inside the bounds" in error_details.lower():
        # Try to extract the index value from the error
        idx_match = re.search(r"index[:\s]+(\d+)", error_details, re.IGNORECASE)
        index_value = int(idx_match.group(1)) if idx_match else None

        error_response = {
            "error": True,
            "code": "INDEX_OUT_OF_BOUNDS",
            "message": "Insertion index is outside document bounds"
            + (f" (index: {index_value})" if index_value else ""),
            "reason": "The insertion point is not within a valid text location in the document. "
            "This often happens when trying to insert at an index that doesn't exist yet.",
            "suggestion": "Use inspect_doc_structure to find valid insertion points. "
            "For a new or empty document, use index 1.",
            "example": {
                "check_structure": "inspect_doc_structure(document_id='...')",
                "empty_doc": "For an empty document, use start_index=1",
            },
        }
        if index_value is not None:
            error_response["context"] = {"received": {"index": index_value}}

        return json.dumps(error_response, indent=2)

    return None


def _create_docs_not_found_error(document_id: str) -> str:
    """
    Create a structured error response for document not found (404) errors.

    Args:
        document_id: The document ID that was not found

    Returns:
        A JSON string with structured error details
    """
    import json

    error_response = {
        "error": True,
        "code": "DOCUMENT_NOT_FOUND",
        "message": f"Document '{document_id}' was not found",
        "reason": "The document ID may be incorrect or you may not have access to this document.",
        "suggestion": "Verify the document ID is correct. You can find the ID in the document's URL: docs.google.com/document/d/{document_id}/edit",
        "context": {
            "received": {"document_id": document_id},
            "possible_causes": [
                "Document ID is incorrect",
                "Document was deleted",
                "You don't have permission to access this document",
                "Document ID includes extra characters (quotes, spaces)",
            ],
        },
    }
    return json.dumps(error_response, indent=2)


def handle_http_errors(
    tool_name: str, is_read_only: bool = False, service_type: Optional[str] = None
):
    """
    A decorator to handle Google API HttpErrors and transient SSL errors in a standardized way.

    It wraps a tool function, catches HttpError, logs a detailed error message,
    and raises a generic Exception with a user-friendly message.

    If is_read_only is True, it will also catch ssl.SSLError and retry with
    exponential backoff. After exhausting retries, it raises a TransientNetworkError.

    Args:
        tool_name (str): The name of the tool being decorated (e.g., 'list_calendars').
        is_read_only (bool): If True, the operation is considered safe to retry on
                             transient network errors. Defaults to False.
        service_type (str): Optional. The Google service type (e.g., 'calendar', 'gmail').
    """

    def decorator(func):
        @functools.wraps(func)
        async def wrapper(*args, **kwargs):
            max_retries = 3
            base_delay = 1

            for attempt in range(max_retries):
                try:
                    return await func(*args, **kwargs)
                except ssl.SSLError as e:
                    if is_read_only and attempt < max_retries - 1:
                        delay = base_delay * (2**attempt)
                        logger.warning(
                            f"SSL error in {tool_name} on attempt {attempt + 1}: {e}. Retrying in {delay} seconds..."
                        )
                        await asyncio.sleep(delay)
                    else:
                        logger.error(
                            f"SSL error in {tool_name} on final attempt: {e}. Raising exception."
                        )
                        raise TransientNetworkError(
                            f"A transient SSL error occurred in '{tool_name}' after {max_retries} attempts. "
                            "This is likely a temporary network or certificate issue. Please try again shortly."
                        ) from e
                except HttpError as error:
                    user_google_email = kwargs.get("user_google_email", "N/A")
                    error_details = str(error)

                    # Check if this is an API not enabled error
                    if (
                        error.resp.status == 403
                        and "accessNotConfigured" in error_details
                    ):
                        enablement_msg = get_api_enablement_message(
                            error_details, service_type
                        )

                        if enablement_msg:
                            message = (
                                f"API error in {tool_name}: {enablement_msg}\n\n"
                                f"User: {user_google_email}"
                            )
                        else:
                            message = (
                                f"API error in {tool_name}: {error}. "
                                f"The required API is not enabled for your project. "
                                f"Please check the Google Cloud Console to enable it."
                            )
                    elif error.resp.status in [401, 403]:
                        # Authentication/authorization errors
                        message = (
                            f"API error in {tool_name}: {error}. "
                            f"You might need to re-authenticate for user '{user_google_email}'. "
                            f"LLM: Try 'start_google_auth' with the user's email and the appropriate service_name."
                        )
                    elif error.resp.status == 400 and service_type == "docs":
                        # Check for index out-of-bounds errors in Google Docs API
                        structured_error = _parse_docs_index_error(error_details)
                        if structured_error:
                            logger.error(
                                f"Index error in {tool_name}: {error}", exc_info=True
                            )
                            return structured_error
                        # Fall through to generic error handling if not an index error
                        message = f"API error in {tool_name}: {error}"
                    elif error.resp.status == 404 and service_type == "docs":
                        # Document not found error - return structured error response
                        document_id = kwargs.get("document_id", "unknown")
                        structured_error = _create_docs_not_found_error(document_id)
                        logger.error(
                            f"Document not found in {tool_name}: {error}", exc_info=True
                        )
                        return structured_error
                    else:
                        # Other HTTP errors (400 Bad Request, etc.) - don't suggest re-auth
                        message = f"API error in {tool_name}: {error}"

                    logger.error(f"API error in {tool_name}: {error}", exc_info=True)
                    raise Exception(message) from error
                except TransientNetworkError:
                    # Re-raise without wrapping to preserve the specific error type
                    raise
                except GoogleAuthenticationError:
                    # Re-raise authentication errors without wrapping
                    raise
                except Exception as e:
                    message = f"An unexpected error occurred in {tool_name}: {e}"
                    logger.exception(message)
                    raise Exception(message) from e

        return wrapper

    return decorator
