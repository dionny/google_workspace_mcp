"""
Google Docs Error Handling

This module provides structured, actionable error messages for Google Docs operations.
Errors are designed to be self-documenting and help both humans and AI agents
understand what went wrong and how to fix it.
"""
import json
import logging
from enum import Enum
from dataclasses import dataclass, asdict
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class ErrorCode(str, Enum):
    """Standardized error codes for Google Docs operations."""

    # Validation errors
    INVALID_DOCUMENT_ID = "INVALID_DOCUMENT_ID"
    DOCUMENT_NOT_FOUND = "DOCUMENT_NOT_FOUND"
    PERMISSION_DENIED = "PERMISSION_DENIED"

    # Index errors
    INDEX_OUT_OF_BOUNDS = "INDEX_OUT_OF_BOUNDS"
    INVALID_INDEX_RANGE = "INVALID_INDEX_RANGE"
    INVALID_INDEX_TYPE = "INVALID_INDEX_TYPE"

    # Formatting errors
    FORMATTING_REQUIRES_RANGE = "FORMATTING_REQUIRES_RANGE"
    INVALID_FORMATTING_PARAMS = "INVALID_FORMATTING_PARAMS"
    INVALID_COLOR_FORMAT = "INVALID_COLOR_FORMAT"

    # Search errors
    EMPTY_SEARCH_TEXT = "EMPTY_SEARCH_TEXT"
    SEARCH_TEXT_NOT_FOUND = "SEARCH_TEXT_NOT_FOUND"
    AMBIGUOUS_SEARCH = "AMBIGUOUS_SEARCH"
    INVALID_OCCURRENCE = "INVALID_OCCURRENCE"

    # Heading/Section errors
    HEADING_NOT_FOUND = "HEADING_NOT_FOUND"
    INVALID_SECTION_POSITION = "INVALID_SECTION_POSITION"

    # Table errors
    TABLE_NOT_FOUND = "TABLE_NOT_FOUND"
    INVALID_TABLE_DATA = "INVALID_TABLE_DATA"
    INVALID_TABLE_DIMENSIONS = "INVALID_TABLE_DIMENSIONS"

    # Parameter errors
    MISSING_REQUIRED_PARAM = "MISSING_REQUIRED_PARAM"
    INVALID_PARAM_TYPE = "INVALID_PARAM_TYPE"
    INVALID_PARAM_VALUE = "INVALID_PARAM_VALUE"
    CONFLICTING_PARAMS = "CONFLICTING_PARAMS"

    # Operation errors
    OPERATION_FAILED = "OPERATION_FAILED"
    API_ERROR = "API_ERROR"


@dataclass
class ErrorContext:
    """Additional context for error messages."""
    received: Optional[Dict[str, Any]] = None
    expected: Optional[Dict[str, Any]] = None
    document_length: Optional[int] = None
    available_headings: Optional[List[str]] = None
    similar_found: Optional[List[str]] = None
    occurrences: Optional[List[Dict[str, Any]]] = None
    current_permission: Optional[str] = None
    required_permission: Optional[str] = None
    possible_causes: Optional[List[str]] = None


@dataclass
class StructuredError:
    """
    Structured error response with actionable guidance.

    Attributes:
        error: Always True for error responses
        code: Machine-readable error code from ErrorCode enum
        message: Human-readable error description
        reason: Explanation of why this error occurred
        suggestion: Actionable advice on how to fix the issue
        example: Optional example showing correct usage
        context: Additional context like received values, document length, etc.
        docs_url: Optional URL to relevant documentation
    """
    error: bool = True
    code: str = ""
    message: str = ""
    reason: str = ""
    suggestion: str = ""
    example: Optional[Dict[str, Any]] = None
    context: Optional[ErrorContext] = None
    docs_url: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary, excluding None values."""
        result = {
            "error": self.error,
            "code": self.code,
            "message": self.message,
        }

        if self.reason:
            result["reason"] = self.reason
        if self.suggestion:
            result["suggestion"] = self.suggestion
        if self.example:
            result["example"] = self.example
        if self.context:
            ctx = asdict(self.context)
            ctx = {k: v for k, v in ctx.items() if v is not None}
            if ctx:
                result["context"] = ctx
        if self.docs_url:
            result["docs_url"] = self.docs_url

        return result

    def to_json(self, indent: int = 2) -> str:
        """Convert to JSON string."""
        return json.dumps(self.to_dict(), indent=indent)


class DocsErrorBuilder:
    """
    Builder for creating structured error messages.

    Usage:
        error = DocsErrorBuilder.formatting_requires_range(
            start_index=100, has_text=False
        ).to_json()
    """

    @staticmethod
    def formatting_requires_range(
        start_index: int,
        has_text: bool = False,
        formatting_params: Optional[List[str]] = None
    ) -> StructuredError:
        """Error when formatting is requested without an end_index."""
        params = formatting_params or ["bold", "italic", "underline", "font_size", "font_family"]

        if has_text:
            suggestion = (
                "When inserting text WITH formatting, provide both text and formatting params. "
                "The end_index will be calculated automatically as start_index + len(text)."
            )
            example = {
                "correct_usage": {
                    "description": "Insert and format in one call",
                    "call": f"modify_doc_text(document_id='...', start_index={start_index}, "
                           f"text='IMPORTANT:', bold=True)"
                }
            }
        else:
            suggestion = (
                "To format existing text, provide both start_index and end_index. "
                "To insert formatted text: 1) Insert text first, 2) Format it with the correct range."
            )
            example = {
                "option_1": {
                    "description": "Format existing text",
                    "call": f"modify_doc_text(document_id='...', start_index={start_index}, "
                           f"end_index={start_index + 10}, bold=True)"
                },
                "option_2": {
                    "description": "Insert then format (two calls)",
                    "step1": f"modify_doc_text(document_id='...', start_index={start_index}, text='new text')",
                    "step2": f"modify_doc_text(document_id='...', start_index={start_index}, "
                            f"end_index={start_index + 8}, bold=True)"
                }
            }

        return StructuredError(
            code=ErrorCode.FORMATTING_REQUIRES_RANGE.value,
            message="'end_index' is required when applying formatting without inserting text",
            reason=(
                f"Formatting operations ({', '.join(params)}) modify existing text and need to know "
                "which characters to format. Without an end_index, the system cannot determine the range."
            ),
            suggestion=suggestion,
            example=example,
            context=ErrorContext(
                received={"start_index": start_index, "formatting": params, "text": None if not has_text else "(provided)"}
            )
        )

    @staticmethod
    def index_out_of_bounds(
        index_name: str,
        index_value: int,
        document_length: int
    ) -> StructuredError:
        """Error when an index exceeds document length."""
        return StructuredError(
            code=ErrorCode.INDEX_OUT_OF_BOUNDS.value,
            message=f"{index_name} {index_value} is beyond document length {document_length}",
            reason=(
                f"The document has {document_length} characters (indices 0-{document_length - 1}). "
                f"The requested {index_name} of {index_value} is outside this range."
            ),
            suggestion="Use get_doc_info or get_doc_content to check document length before editing.",
            example={
                "check_length": "get_doc_info(document_id='...')",
                "valid_range": f"Use indices between 1 and {document_length - 1}"
            },
            context=ErrorContext(
                document_length=document_length,
                received={index_name: index_value}
            )
        )

    @staticmethod
    def invalid_index_range(
        start_index: int,
        end_index: int
    ) -> StructuredError:
        """Error when start_index >= end_index."""
        return StructuredError(
            code=ErrorCode.INVALID_INDEX_RANGE.value,
            message=f"start_index ({start_index}) must be less than end_index ({end_index})",
            reason="The start of a range must come before its end.",
            suggestion="Swap the values or correct the range specification.",
            context=ErrorContext(
                received={"start_index": start_index, "end_index": end_index},
                expected={"start_index": min(start_index, end_index), "end_index": max(start_index, end_index)}
            )
        )

    @staticmethod
    def empty_search_text() -> StructuredError:
        """Error when search text is empty."""
        return StructuredError(
            code=ErrorCode.EMPTY_SEARCH_TEXT.value,
            message="Search text cannot be empty",
            reason="An empty string was provided for the search parameter, which would match nothing.",
            suggestion="Provide a non-empty search string to locate text in the document.",
            example={
                "search_mode": "modify_doc_text(document_id='...', search='target text', position='after', text='new text')",
                "find_and_replace": "find_and_replace_doc(document_id='...', find_text='old', replace_text='new')"
            }
        )

    @staticmethod
    def search_text_not_found(
        search_text: str,
        similar_found: Optional[List[str]] = None,
        match_case: bool = True
    ) -> StructuredError:
        """Error when search text is not found in document."""
        suggestions = ["Check spelling and try a shorter, unique phrase"]
        if match_case:
            suggestions.append("Try setting match_case=False for case-insensitive search")

        context = ErrorContext(
            received={"search": search_text, "match_case": match_case}
        )
        if similar_found:
            context.similar_found = similar_found

        return StructuredError(
            code=ErrorCode.SEARCH_TEXT_NOT_FOUND.value,
            message=f"Could not find '{search_text}' in the document",
            reason="The exact text was not found in the document content.",
            suggestion=". ".join(suggestions) + ".",
            example={
                "case_insensitive": f"modify_doc_text(document_id='...', search='{search_text}', "
                                   f"position='after', match_case=False, text='new text')"
            },
            context=context
        )

    @staticmethod
    def ambiguous_search(
        search_text: str,
        occurrences: List[Dict[str, Any]],
        total_count: int
    ) -> StructuredError:
        """Error when multiple occurrences found but no specific one selected."""
        return StructuredError(
            code=ErrorCode.AMBIGUOUS_SEARCH.value,
            message=f"Found {total_count} occurrences of '{search_text}' - specify which one",
            reason="Multiple matches were found and the system needs to know which one to target.",
            suggestion=(
                "Use 'occurrence' parameter (1=first, 2=second, -1=last) "
                "or use more specific search text."
            ),
            example={
                "first_occurrence": f"modify_doc_text(document_id='...', search='{search_text}', "
                                   f"position='replace', occurrence=1, text='replacement')",
                "last_occurrence": f"modify_doc_text(document_id='...', search='{search_text}', "
                                  f"position='replace', occurrence=-1, text='replacement')",
                "specific_occurrence": f"modify_doc_text(document_id='...', search='{search_text}', "
                                      f"position='replace', occurrence=2, text='replacement')"
            },
            context=ErrorContext(
                occurrences=occurrences[:5]  # Limit to first 5
            )
        )

    @staticmethod
    def invalid_occurrence(
        occurrence: int,
        total_found: int,
        search_text: str
    ) -> StructuredError:
        """Error when requested occurrence doesn't exist."""
        return StructuredError(
            code=ErrorCode.INVALID_OCCURRENCE.value,
            message=f"Occurrence {occurrence} requested but only {total_found} found for '{search_text}'",
            reason=f"The document contains {total_found} instance(s) of the search text, but occurrence {occurrence} was requested.",
            suggestion=f"Use occurrence between 1 and {total_found}, or -1 for the last occurrence.",
            context=ErrorContext(
                received={"occurrence": occurrence, "search": search_text},
                expected={"occurrence": f"1 to {total_found} (or -1 for last)"}
            )
        )

    @staticmethod
    def heading_not_found(
        heading: str,
        available_headings: List[str],
        match_case: bool = True
    ) -> StructuredError:
        """Error when a heading is not found in the document."""
        suggestions = ["Check spelling of the heading text"]
        if match_case:
            suggestions.append("Try setting match_case=False for case-insensitive search")

        # Truncate available headings list for display
        display_headings = available_headings[:10]
        if len(available_headings) > 10:
            display_headings.append(f"... and {len(available_headings) - 10} more")

        return StructuredError(
            code=ErrorCode.HEADING_NOT_FOUND.value,
            message=f"Heading '{heading}' not found in document",
            reason="No heading with this exact text was found in the document structure.",
            suggestion=". ".join(suggestions) + ".",
            example={
                "list_headings": "get_doc_info(document_id='...', detail='headings')"
            },
            context=ErrorContext(
                received={"heading": heading, "match_case": match_case},
                available_headings=display_headings
            )
        )

    @staticmethod
    def document_not_found(
        document_id: str
    ) -> StructuredError:
        """Error when a document cannot be found or accessed."""
        return StructuredError(
            code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            message=f"Document with ID '{document_id}' not found or not accessible",
            reason="The document could not be found or you don't have permission to access it.",
            suggestion="Verify the document ID and ensure you have access permissions.",
            context=ErrorContext(
                received={"document_id": document_id},
                possible_causes=[
                    "Document ID is incorrect",
                    "Document was deleted",
                    "You don't have permission to access this document",
                    "Document ID includes extra characters (quotes, spaces)"
                ]
            )
        )

    @staticmethod
    def permission_denied(
        document_id: str,
        current_permission: str = "viewer",
        required_permission: str = "editor"
    ) -> StructuredError:
        """Error when user lacks permission to edit."""
        return StructuredError(
            code=ErrorCode.PERMISSION_DENIED.value,
            message=f"You have {current_permission} access to this document but {required_permission} access is required",
            reason="The operation requires edit permissions that you don't have.",
            suggestion="Request edit access from the document owner.",
            context=ErrorContext(
                received={"document_id": document_id},
                current_permission=current_permission,
                required_permission=required_permission
            )
        )

    @staticmethod
    def invalid_table_data(
        issue: str,
        row_index: Optional[int] = None,
        col_index: Optional[int] = None,
        value: Optional[Any] = None
    ) -> StructuredError:
        """Error when table data format is invalid."""
        context_received = {"issue": issue}
        if row_index is not None:
            context_received["row"] = row_index
        if col_index is not None:
            context_received["column"] = col_index
        if value is not None:
            context_received["value"] = repr(value)

        return StructuredError(
            code=ErrorCode.INVALID_TABLE_DATA.value,
            message=f"Invalid table data: {issue}",
            reason="Table data must be a 2D list of strings with consistent row lengths.",
            suggestion="Ensure all rows have the same number of columns and all cells are strings (use '' for empty cells).",
            example={
                "correct_format": "[['Header1', 'Header2'], ['Data1', 'Data2'], ['', 'Data3']]",
                "create_table": "create_table_with_data(document_id='...', index=..., table_data=[[...], [...]])"
            },
            context=ErrorContext(received=context_received)
        )

    @staticmethod
    def table_not_found(
        table_index: int,
        total_tables: int
    ) -> StructuredError:
        """Error when requested table doesn't exist."""
        return StructuredError(
            code=ErrorCode.TABLE_NOT_FOUND.value,
            message=f"Table index {table_index} not found. Document has {total_tables} table(s)",
            reason=f"Table indices are 0-based. Valid indices are 0 to {total_tables - 1}." if total_tables > 0 else "The document contains no tables.",
            suggestion="Use get_doc_info to see all tables and their indices.",
            example={
                "inspect_tables": "get_doc_info(document_id='...', detail='tables')"
            },
            context=ErrorContext(
                received={"table_index": table_index},
                expected={"valid_indices": f"0 to {total_tables - 1}" if total_tables > 0 else "No tables available"}
            )
        )

    @staticmethod
    def missing_required_param(
        param_name: str,
        context_description: str,
        valid_values: Optional[List[str]] = None
    ) -> StructuredError:
        """Error when a required parameter is missing."""
        suggestion = f"Provide the '{param_name}' parameter"
        if valid_values:
            suggestion += f". Valid values: {', '.join(valid_values)}"

        return StructuredError(
            code=ErrorCode.MISSING_REQUIRED_PARAM.value,
            message=f"'{param_name}' is required {context_description}",
            reason=f"This operation cannot proceed without the '{param_name}' parameter.",
            suggestion=suggestion,
            context=ErrorContext(
                expected={param_name: valid_values[0] if valid_values else "(required)"}
            )
        )

    @staticmethod
    def invalid_param_value(
        param_name: str,
        received_value: Any,
        valid_values: List[str],
        context_description: str = ""
    ) -> StructuredError:
        """Error when a parameter has an invalid value."""
        return StructuredError(
            code=ErrorCode.INVALID_PARAM_VALUE.value,
            message=f"Invalid '{param_name}' value '{received_value}'{'. ' + context_description if context_description else ''}",
            reason=f"The value '{received_value}' is not a valid option for '{param_name}'.",
            suggestion=f"Use one of: {', '.join(valid_values)}",
            context=ErrorContext(
                received={param_name: received_value},
                expected={param_name: valid_values}
            )
        )

    @staticmethod
    def invalid_color_format(
        color_value: str,
        param_name: str = "color"
    ) -> StructuredError:
        """Error when a color value has an invalid format."""
        named_colors = ["red", "green", "blue", "yellow", "orange", "purple", "black", "white", "gray", "grey"]
        return StructuredError(
            code=ErrorCode.INVALID_COLOR_FORMAT.value,
            message=f"Invalid color format for '{param_name}': '{color_value}'",
            reason="Colors must be specified as hex codes (#FF0000, #F00) or named colors.",
            suggestion=f"Use hex format (#RRGGBB or #RGB) or a named color: {', '.join(named_colors)}",
            example={
                "hex_color": "#FF0000",
                "short_hex": "#F00",
                "named_color": "red",
                "usage": f"modify_doc_text(document_id='...', search='text', position='replace', {param_name}='#FF0000', text='colored text')"
            },
            context=ErrorContext(
                received={param_name: color_value},
                expected={"format": "hex (#RRGGBB or #RGB) or named color"}
            )
        )

    @staticmethod
    def conflicting_params(
        params: List[str],
        message: str
    ) -> StructuredError:
        """Error when conflicting parameters are provided."""
        return StructuredError(
            code=ErrorCode.CONFLICTING_PARAMS.value,
            message=message,
            reason=f"The parameters {', '.join(params)} cannot be used together.",
            suggestion="Choose one positioning method and provide only its required parameters.",
            example={
                "heading_mode": "modify_doc_text(document_id='...', heading='Section', section_position='end', text='...')",
                "search_mode": "modify_doc_text(document_id='...', search='text', position='after', text='...')",
                "index_mode": "modify_doc_text(document_id='...', start_index=100, text='...')"
            }
        )

    @staticmethod
    def operation_needs_action(
        operation: str,
        current_state: str,
        required_action: str
    ) -> StructuredError:
        """Error when an operation requires a prerequisite action."""
        return StructuredError(
            code=ErrorCode.OPERATION_FAILED.value,
            message=f"Cannot {operation}: {current_state}",
            reason=f"The operation requires: {required_action}",
            suggestion=required_action
        )

    @staticmethod
    def api_error(
        operation: str,
        error_message: str,
        document_id: Optional[str] = None
    ) -> StructuredError:
        """Error from Google API call."""
        context_data = {"operation": operation}
        if document_id:
            context_data["document_id"] = document_id

        return StructuredError(
            code=ErrorCode.API_ERROR.value,
            message=f"API error during {operation}: {error_message}",
            reason="The Google Docs API returned an error.",
            suggestion="Check the error message for details. Common issues: invalid document ID, insufficient permissions, or rate limiting.",
            context=ErrorContext(
                received=context_data,
                possible_causes=[
                    "Document ID may be incorrect",
                    "You may not have permission to access this document",
                    "The document may have been deleted",
                    "API rate limits may have been exceeded"
                ]
            )
        )

    @staticmethod
    def invalid_image_source(
        image_source: str,
        actual_mime_type: Optional[str] = None,
        error_detail: Optional[str] = None
    ) -> StructuredError:
        """Error when image source is invalid or inaccessible."""
        is_drive_file = not (image_source.startswith('http://') or image_source.startswith('https://'))

        if actual_mime_type and not actual_mime_type.startswith('image/'):
            return StructuredError(
                code=ErrorCode.INVALID_PARAM_VALUE.value,
                message=f"File is not an image (MIME type: {actual_mime_type})",
                reason="The specified file is not a valid image format.",
                suggestion="Provide a valid image file (JPEG, PNG, GIF, etc.) or image URL.",
                context=ErrorContext(
                    received={"image_source": image_source, "mime_type": actual_mime_type},
                    expected={"mime_type": "image/* (e.g., image/jpeg, image/png)"}
                )
            )

        if is_drive_file:
            return StructuredError(
                code=ErrorCode.OPERATION_FAILED.value,
                message=f"Could not access Drive file: {error_detail or 'unknown error'}",
                reason="The Drive file could not be accessed.",
                suggestion="Verify the file ID is correct and you have permission to access it.",
                context=ErrorContext(
                    received={"file_id": image_source},
                    possible_causes=[
                        "File ID is incorrect",
                        "File was deleted or moved",
                        "You don't have permission to access the file",
                        "File is not shared with the service account"
                    ]
                )
            )
        else:
            return StructuredError(
                code=ErrorCode.OPERATION_FAILED.value,
                message=f"Could not access image URL: {error_detail or 'unknown error'}",
                reason="The image URL could not be accessed.",
                suggestion="Verify the URL is publicly accessible and returns a valid image.",
                context=ErrorContext(
                    received={"url": image_source},
                    possible_causes=[
                        "URL is invalid or broken",
                        "Image requires authentication",
                        "Server is unreachable",
                        "URL does not point to an image"
                    ]
                )
            )

    @staticmethod
    def pdf_export_error(
        document_id: str,
        stage: str,
        error_detail: str
    ) -> StructuredError:
        """Error during PDF export operation."""
        suggestions = {
            "access": "Verify the document ID and ensure you have access permissions.",
            "export": "The document may be too large or contain unsupported elements. Try exporting a smaller document.",
            "upload": "Check Drive permissions and storage quota."
        }

        return StructuredError(
            code=ErrorCode.OPERATION_FAILED.value,
            message=f"PDF export failed at {stage} stage: {error_detail}",
            reason=f"The PDF export operation encountered an error during {stage}.",
            suggestion=suggestions.get(stage, "Check the error details and try again."),
            context=ErrorContext(
                received={"document_id": document_id, "stage": stage},
                possible_causes=[
                    "Document may be too large",
                    "Document may contain unsupported elements",
                    "Insufficient Drive storage quota",
                    "Permission issues with target folder"
                ]
            )
        )

    @staticmethod
    def invalid_document_type(
        document_id: str,
        file_name: str,
        actual_mime_type: str,
        expected_mime_type: str = "application/vnd.google-apps.document"
    ) -> StructuredError:
        """Error when file is not the expected Google Docs type."""
        return StructuredError(
            code=ErrorCode.INVALID_PARAM_VALUE.value,
            message=f"File '{file_name}' is not a Google Doc",
            reason=f"Expected a Google Doc but found {actual_mime_type}.",
            suggestion="Only native Google Docs can be used with this operation. Convert the file to Google Docs format first.",
            context=ErrorContext(
                received={"document_id": document_id, "mime_type": actual_mime_type},
                expected={"mime_type": expected_mime_type}
            )
        )

    @staticmethod
    def empty_text_insertion() -> StructuredError:
        """Error when trying to insert empty text without a range to delete."""
        return StructuredError(
            code=ErrorCode.INVALID_PARAM_VALUE.value,
            message="Cannot insert empty text",
            reason="An empty string was provided for text insertion, which would have no effect.",
            suggestion="Provide non-empty text to insert, or use position='replace' with a range (start_index and end_index) and empty text to delete existing content.",
            example={
                "insert_text": "modify_doc_text(document_id='...', location='end', text='new content')",
                "delete_text": "modify_doc_text(document_id='...', start_index=10, end_index=20, text='')",
                "replace_via_search": "modify_doc_text(document_id='...', search='old', position='replace', text='new')"
            },
            context=ErrorContext(
                received={"text": "''"},
                expected={"text": "non-empty string for insertion, or provide end_index > start_index to delete a range"}
            )
        )


def format_error(error: StructuredError) -> str:
    """
    Format a StructuredError for return to the user.

    Returns a JSON string that can be parsed by both humans and AI agents.
    """
    return error.to_json()


def simple_error(code: ErrorCode, message: str, suggestion: str = "") -> str:
    """
    Create a simple error message without full context.

    Useful for quick validation errors where full context isn't needed.
    """
    error = StructuredError(
        code=code.value,
        message=message,
        suggestion=suggestion
    )
    return error.to_json()
