"""
Validation Manager

This module provides centralized validation logic for Google Docs operations,
extracting validation patterns from individual tool functions.
"""
import logging
from typing import Dict, Any, List, Tuple, Optional

from gdocs.errors import (
    DocsErrorBuilder,
    StructuredError,
    ErrorCode,
    format_error,
)

logger = logging.getLogger(__name__)


class ValidationManager:
    """
    Centralized validation manager for Google Docs operations.
    
    Provides consistent validation patterns and error messages across
    all document operations, reducing code duplication and improving
    error message quality.
    """

    def __init__(self):
        """Initialize the validation manager."""
        self.validation_rules = self._setup_validation_rules()

    def _setup_validation_rules(self) -> Dict[str, Any]:
        """Setup validation rules and constraints."""
        return {
            'table_max_rows': 1000,
            'table_max_columns': 20,
            'document_id_pattern': r'^[a-zA-Z0-9-_]+$',
            'max_text_length': 1000000,  # 1MB text limit
            'font_size_range': (1, 400),  # Google Docs font size limits
            'valid_header_footer_types': ["DEFAULT", "FIRST_PAGE", "EVEN_PAGE"],
            'valid_section_types': ["header", "footer"],
            'valid_list_types': ["UNORDERED", "ORDERED"],
            'valid_element_types': ["table", "list", "page_break"]
        }

    def validate_document_id(self, document_id: str) -> Tuple[bool, str]:
        """
        Validate Google Docs document ID format.
        
        Args:
            document_id: Document ID to validate
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        if not document_id:
            return False, "Document ID cannot be empty"

        if not isinstance(document_id, str):
            return False, f"Document ID must be a string, got {type(document_id).__name__}"

        # Basic length check (Google Docs IDs are typically 40+ characters)
        if len(document_id) < 20:
            return False, "Document ID appears too short to be valid"

        return True, ""

    def validate_table_data(self, table_data: List[List[str]]) -> Tuple[bool, str]:
        """
        Comprehensive validation for table data format.
        
        This extracts and centralizes table validation logic from multiple functions.
        
        Args:
            table_data: 2D array of data to validate
            
        Returns:
            Tuple of (is_valid, detailed_error_message)
        """
        if not table_data:
            return False, "Table data cannot be empty. Required format: [['col1', 'col2'], ['row1col1', 'row1col2']]"

        if not isinstance(table_data, list):
            return False, f"Table data must be a list, got {type(table_data).__name__}. Required format: [['col1', 'col2'], ['row1col1', 'row1col2']]"

        # Check if it's a 2D list
        if not all(isinstance(row, list) for row in table_data):
            non_list_rows = [i for i, row in enumerate(table_data) if not isinstance(row, list)]
            return False, f"All rows must be lists. Rows {non_list_rows} are not lists. Required format: [['col1', 'col2'], ['row1col1', 'row1col2']]"

        # Check for empty rows
        if any(len(row) == 0 for row in table_data):
            empty_rows = [i for i, row in enumerate(table_data) if len(row) == 0]
            return False, f"Rows cannot be empty. Empty rows found at indices: {empty_rows}"

        # Check column consistency
        col_counts = [len(row) for row in table_data]
        if len(set(col_counts)) > 1:
            return False, f"All rows must have the same number of columns. Found column counts: {col_counts}. Fix your data structure."

        rows = len(table_data)
        cols = col_counts[0]

        # Check dimension limits
        if rows > self.validation_rules['table_max_rows']:
            return False, f"Too many rows ({rows}). Maximum allowed: {self.validation_rules['table_max_rows']}"

        if cols > self.validation_rules['table_max_columns']:
            return False, f"Too many columns ({cols}). Maximum allowed: {self.validation_rules['table_max_columns']}"

        # Check cell content types
        for row_idx, row in enumerate(table_data):
            for col_idx, cell in enumerate(row):
                if cell is None:
                    return False, f"Cell ({row_idx},{col_idx}) is None. All cells must be strings, use empty string '' for empty cells."

                if not isinstance(cell, str):
                    return False, f"Cell ({row_idx},{col_idx}) is {type(cell).__name__}, not string. All cells must be strings. Value: {repr(cell)}"

        return True, f"Valid table data: {rows}×{cols} table format"

    def validate_text_formatting_params(
        self,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[bool] = None,
        strikethrough: Optional[bool] = None,
        small_caps: Optional[bool] = None,
        subscript: Optional[bool] = None,
        superscript: Optional[bool] = None,
        font_size: Optional[int] = None,
        font_family: Optional[str] = None,
        link: Optional[str] = None,
        foreground_color: Optional[str] = None,
        background_color: Optional[str] = None,
    ) -> Tuple[bool, str]:
        """
        Validate text formatting parameters.

        Args:
            bold: Bold setting
            italic: Italic setting
            underline: Underline setting
            strikethrough: Strikethrough setting
            small_caps: Small caps setting
            subscript: Subscript setting (mutually exclusive with superscript)
            superscript: Superscript setting (mutually exclusive with subscript)
            font_size: Font size in points
            font_family: Font family name
            link: URL for hyperlink (empty string "" removes link)
            foreground_color: Text color (hex or named)
            background_color: Background/highlight color (hex or named)

        Returns:
            Tuple of (is_valid, error_message)
        """
        # Check if at least one formatting option is provided
        formatting_params = [bold, italic, underline, strikethrough, small_caps, subscript, superscript, font_size, font_family, link, foreground_color, background_color]
        if all(param is None for param in formatting_params):
            return False, "At least one formatting parameter must be provided (bold, italic, underline, strikethrough, small_caps, subscript, superscript, font_size, font_family, link, foreground_color, or background_color)"

        # Validate boolean parameters
        for param, name in [(bold, 'bold'), (italic, 'italic'), (underline, 'underline'), (strikethrough, 'strikethrough'), (small_caps, 'small_caps'), (subscript, 'subscript'), (superscript, 'superscript')]:
            if param is not None and not isinstance(param, bool):
                return False, f"{name} parameter must be boolean (True/False), got {type(param).__name__}"

        # Validate that subscript and superscript are not both True
        if subscript and superscript:
            return False, "subscript and superscript are mutually exclusive - only one can be True at a time"

        # Validate font size
        if font_size is not None:
            if not isinstance(font_size, int):
                return False, f"font_size must be an integer, got {type(font_size).__name__}"

            min_size, max_size = self.validation_rules['font_size_range']
            if not (min_size <= font_size <= max_size):
                return False, f"font_size must be between {min_size} and {max_size} points, got {font_size}"

        # Validate font family
        if font_family is not None:
            if not isinstance(font_family, str):
                return False, f"font_family must be a string, got {type(font_family).__name__}"

            if not font_family.strip():
                return False, "font_family cannot be empty"

        # Validate link
        if link is not None:
            if not isinstance(link, str):
                return False, f"link must be a string, got {type(link).__name__}"
            # Empty string is allowed (removes link), but non-empty must look like a URL
            if link and not (link.startswith('http://') or link.startswith('https://') or link.startswith('#')):
                return False, f"link must be a valid URL starting with http://, https://, or # (for internal bookmarks), got '{link}'"

        # Validate colors (full format validation)
        for color, name in [(foreground_color, 'foreground_color'), (background_color, 'background_color')]:
            if color is not None:
                if not isinstance(color, str):
                    return False, f"{name} must be a string, got {type(color).__name__}"
                if not color.strip():
                    return False, f"{name} cannot be empty"
                # Validate color format
                is_valid, error_msg = self._validate_color_format(color, name)
                if not is_valid:
                    return False, error_msg

        return True, ""

    def _validate_color_format(self, color: str, param_name: str = "color") -> Tuple[bool, str]:
        """
        Validate that a color string is in a valid format.

        Args:
            color: Color string to validate
            param_name: Parameter name for error messages

        Returns:
            Tuple of (is_valid, error_message)
        """
        # Valid named colors (must match _parse_color in docs_helpers.py)
        named_colors = {'red', 'green', 'blue', 'yellow', 'orange', 'purple', 'black', 'white', 'gray', 'grey'}

        # Check if it's a named color
        if color.lower() in named_colors:
            return True, ""

        # Check if it's a hex color
        if color.startswith('#'):
            hex_color = color.lstrip('#')
            # Valid formats: #RGB or #RRGGBB
            if len(hex_color) == 3 or len(hex_color) == 6:
                # Verify all characters are valid hex digits
                if all(c in '0123456789abcdefABCDEF' for c in hex_color):
                    return True, ""
            return False, f"Invalid hex color for {param_name}: '{color}'. Use #RGB (e.g., #F00) or #RRGGBB (e.g., #FF0000) format."

        # Not a valid format
        return False, f"Invalid color format for {param_name}: '{color}'. Use hex (#FF0000, #F00) or named colors (red, blue, green, etc.)."

    def validate_index(self, index: int, context: str = "Index") -> Tuple[bool, str]:
        """
        Validate a single document index.
        
        Args:
            index: Index to validate
            context: Context description for error messages
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        if not isinstance(index, int):
            return False, f"{context} must be an integer, got {type(index).__name__}"

        if index < 0:
            return False, f"{context} {index} is negative. You MUST call get_doc_info first to get the proper insertion index."

        return True, ""

    def validate_index_range(
        self,
        start_index: int,
        end_index: Optional[int] = None,
        document_length: Optional[int] = None
    ) -> Tuple[bool, str]:
        """
        Validate document index ranges.
        
        Args:
            start_index: Starting index
            end_index: Ending index (optional)
            document_length: Total document length for bounds checking
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        # Validate start_index
        if not isinstance(start_index, int):
            return False, f"start_index must be an integer, got {type(start_index).__name__}"

        if start_index < 0:
            return False, f"start_index cannot be negative, got {start_index}"

        # Validate end_index if provided
        if end_index is not None:
            if not isinstance(end_index, int):
                return False, f"end_index must be an integer, got {type(end_index).__name__}"

            if end_index <= start_index:
                return False, f"end_index ({end_index}) must be greater than start_index ({start_index})"

        # Validate against document length if provided
        if document_length is not None:
            if start_index >= document_length:
                return False, f"start_index ({start_index}) exceeds document length ({document_length})"

            if end_index is not None and end_index > document_length:
                return False, f"end_index ({end_index}) exceeds document length ({document_length})"

        return True, ""

    def validate_element_insertion_params(
        self,
        element_type: str,
        index: int,
        **kwargs
    ) -> Tuple[bool, str]:
        """
        Validate parameters for element insertion.
        
        Args:
            element_type: Type of element to insert
            index: Insertion index
            **kwargs: Additional parameters specific to element type
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        # Validate element type
        if element_type not in self.validation_rules['valid_element_types']:
            valid_types = ', '.join(self.validation_rules['valid_element_types'])
            return False, f"Invalid element_type '{element_type}'. Must be one of: {valid_types}"

        # Validate index
        if not isinstance(index, int) or index < 0:
            return False, f"index must be a non-negative integer, got {index}"

        # Validate element-specific parameters
        if element_type == "table":
            rows = kwargs.get('rows')
            columns = kwargs.get('columns')

            if not rows or not columns:
                return False, "Table insertion requires 'rows' and 'columns' parameters"

            if not isinstance(rows, int) or not isinstance(columns, int):
                return False, "Table rows and columns must be integers"

            if rows <= 0 or columns <= 0:
                return False, "Table rows and columns must be positive integers"

            if rows > self.validation_rules['table_max_rows']:
                return False, f"Too many rows ({rows}). Maximum: {self.validation_rules['table_max_rows']}"

            if columns > self.validation_rules['table_max_columns']:
                return False, f"Too many columns ({columns}). Maximum: {self.validation_rules['table_max_columns']}"

        elif element_type == "list":
            list_type = kwargs.get('list_type')

            if not list_type:
                return False, "List insertion requires 'list_type' parameter"

            if list_type not in self.validation_rules['valid_list_types']:
                valid_types = ', '.join(self.validation_rules['valid_list_types'])
                return False, f"Invalid list_type '{list_type}'. Must be one of: {valid_types}"

        return True, ""

    def validate_header_footer_params(
        self,
        section_type: str,
        header_footer_type: str = "DEFAULT"
    ) -> Tuple[bool, str]:
        """
        Validate header/footer operation parameters.
        
        Args:
            section_type: Type of section ("header" or "footer")
            header_footer_type: Specific header/footer type
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        if section_type not in self.validation_rules['valid_section_types']:
            valid_types = ', '.join(self.validation_rules['valid_section_types'])
            return False, f"section_type must be one of: {valid_types}, got '{section_type}'"

        if header_footer_type not in self.validation_rules['valid_header_footer_types']:
            valid_types = ', '.join(self.validation_rules['valid_header_footer_types'])
            return False, f"header_footer_type must be one of: {valid_types}, got '{header_footer_type}'"

        return True, ""

    def validate_batch_operations(self, operations: List[Dict[str, Any]]) -> Tuple[bool, str]:
        """
        Validate a list of batch operations.
        
        Args:
            operations: List of operation dictionaries
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        if not operations:
            return False, "Operations list cannot be empty"

        if not isinstance(operations, list):
            return False, f"Operations must be a list, got {type(operations).__name__}"

        # Validate each operation
        for i, op in enumerate(operations):
            if not isinstance(op, dict):
                return False, f"Operation {i+1} must be a dictionary, got {type(op).__name__}"

            if 'type' not in op:
                return False, f"Operation {i+1} missing required 'type' field"

            # Validate operation-specific fields using existing validation logic
            # This would call the validate_operation function from docs_helpers
            # but we're centralizing the logic here

        return True, ""

    def validate_text_content(self, text: str, max_length: Optional[int] = None) -> Tuple[bool, str]:
        """
        Validate text content for insertion.
        
        Args:
            text: Text to validate
            max_length: Maximum allowed length
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        if not isinstance(text, str):
            return False, f"Text must be a string, got {type(text).__name__}"

        max_len = max_length or self.validation_rules['max_text_length']
        if len(text) > max_len:
            return False, f"Text too long ({len(text)} characters). Maximum: {max_len}"

        return True, ""

    def get_validation_summary(self) -> Dict[str, Any]:
        """
        Get a summary of all validation rules and constraints.

        Returns:
            Dictionary containing validation rules
        """
        return {
            'constraints': self.validation_rules.copy(),
            'supported_operations': {
                'table_operations': ['create_table', 'populate_table'],
                'text_operations': ['insert_text', 'format_text', 'find_replace'],
                'element_operations': ['insert_table', 'insert_list', 'insert_page_break'],
                'header_footer_operations': ['update_header', 'update_footer']
            },
            'data_formats': {
                'table_data': "2D list of strings: [['col1', 'col2'], ['row1col1', 'row1col2']]",
                'text_formatting': "Optional boolean/integer parameters for styling",
                'document_indices': "Non-negative integers for position specification"
            }
        }

    # ============================================================
    # Structured Error Methods
    # ============================================================
    # These methods return structured JSON errors for better debugging

    def validate_document_id_structured(self, document_id: str) -> Tuple[bool, Optional[str]]:
        """
        Validate document ID and return structured error if invalid.

        Returns:
            Tuple of (is_valid, structured_error_json or None)
        """
        is_valid, _ = self.validate_document_id(document_id)
        if not is_valid:
            error = DocsErrorBuilder.document_not_found(document_id or "(empty)")
            return False, format_error(error)
        return True, None

    def validate_index_range_structured(
        self,
        start_index: int,
        end_index: Optional[int] = None,
        document_length: Optional[int] = None
    ) -> Tuple[bool, Optional[str]]:
        """
        Validate index range and return structured error if invalid.

        Returns:
            Tuple of (is_valid, structured_error_json or None)
        """
        # Check start_index type
        if not isinstance(start_index, int):
            error = StructuredError(
                code=ErrorCode.INVALID_INDEX_TYPE.value,
                message=f"start_index must be an integer, got {type(start_index).__name__}",
                suggestion="Provide an integer value for start_index"
            )
            return False, format_error(error)

        # Check negative start_index
        if start_index < 0:
            error = StructuredError(
                code=ErrorCode.INVALID_INDEX_RANGE.value,
                message=f"start_index cannot be negative, got {start_index}",
                suggestion="Use get_doc_info to find valid insertion indices"
            )
            return False, format_error(error)

        # Check end_index if provided
        if end_index is not None:
            if not isinstance(end_index, int):
                error = StructuredError(
                    code=ErrorCode.INVALID_INDEX_TYPE.value,
                    message=f"end_index must be an integer, got {type(end_index).__name__}",
                    suggestion="Provide an integer value for end_index"
                )
                return False, format_error(error)

            if end_index <= start_index:
                error = DocsErrorBuilder.invalid_index_range(start_index, end_index)
                return False, format_error(error)

        # Check bounds if document length provided
        if document_length is not None:
            if start_index >= document_length:
                error = DocsErrorBuilder.index_out_of_bounds(
                    "start_index", start_index, document_length
                )
                return False, format_error(error)

            if end_index is not None and end_index > document_length:
                error = DocsErrorBuilder.index_out_of_bounds(
                    "end_index", end_index, document_length
                )
                return False, format_error(error)

        return True, None

    def validate_formatting_range_structured(
        self,
        start_index: int,
        end_index: Optional[int],
        text: Optional[str],
        formatting_params: List[str]
    ) -> Tuple[bool, Optional[str]]:
        """
        Validate that formatting operations have proper range.

        Returns:
            Tuple of (is_valid, structured_error_json or None)
        """
        # Formatting without text and without end_index is invalid
        if end_index is None and text is None:
            error = DocsErrorBuilder.formatting_requires_range(
                start_index=start_index,
                has_text=False,
                formatting_params=formatting_params
            )
            return False, format_error(error)

        return True, None

    def validate_table_data_structured(self, table_data: List[List[str]]) -> Tuple[bool, Optional[str]]:
        """
        Validate table data and return structured error if invalid.

        Returns:
            Tuple of (is_valid, structured_error_json or None)
        """
        if not table_data:
            error = DocsErrorBuilder.invalid_table_data("Table data cannot be empty")
            return False, format_error(error)

        if not isinstance(table_data, list):
            error = DocsErrorBuilder.invalid_table_data(
                f"Table data must be a list, got {type(table_data).__name__}"
            )
            return False, format_error(error)

        # Check if it's a 2D list
        for row_idx, row in enumerate(table_data):
            if not isinstance(row, list):
                error = DocsErrorBuilder.invalid_table_data(
                    f"Row {row_idx} must be a list, got {type(row).__name__}",
                    row_index=row_idx
                )
                return False, format_error(error)

            if len(row) == 0:
                error = DocsErrorBuilder.invalid_table_data(
                    f"Row {row_idx} is empty",
                    row_index=row_idx
                )
                return False, format_error(error)

        # Check column consistency
        col_counts = [len(row) for row in table_data]
        if len(set(col_counts)) > 1:
            error = DocsErrorBuilder.invalid_table_data(
                f"All rows must have same column count. Found: {col_counts}"
            )
            return False, format_error(error)

        # Check cell content
        for row_idx, row in enumerate(table_data):
            for col_idx, cell in enumerate(row):
                if cell is None:
                    error = DocsErrorBuilder.invalid_table_data(
                        "Cell value is None - use empty string '' instead",
                        row_index=row_idx,
                        col_index=col_idx,
                        value=None
                    )
                    return False, format_error(error)

                if not isinstance(cell, str):
                    error = DocsErrorBuilder.invalid_table_data(
                        f"Cell must be string, got {type(cell).__name__}",
                        row_index=row_idx,
                        col_index=col_idx,
                        value=cell
                    )
                    return False, format_error(error)

        return True, None

    def create_empty_search_error(self) -> str:
        """Create a structured error for empty search text."""
        error = DocsErrorBuilder.empty_search_text()
        return format_error(error)

    def create_search_not_found_error(
        self,
        search_text: str,
        match_case: bool = True,
        similar_found: Optional[List[str]] = None
    ) -> str:
        """Create a structured error for search text not found."""
        error = DocsErrorBuilder.search_text_not_found(
            search_text=search_text,
            similar_found=similar_found,
            match_case=match_case
        )
        return format_error(error)

    def create_ambiguous_search_error(
        self,
        search_text: str,
        occurrences: List[Dict[str, Any]],
        total_count: int
    ) -> str:
        """Create a structured error for ambiguous search results."""
        error = DocsErrorBuilder.ambiguous_search(
            search_text=search_text,
            occurrences=occurrences,
            total_count=total_count
        )
        return format_error(error)

    def create_heading_not_found_error(
        self,
        heading: str,
        available_headings: List[str],
        match_case: bool = True
    ) -> str:
        """Create a structured error for heading not found."""
        error = DocsErrorBuilder.heading_not_found(
            heading=heading,
            available_headings=available_headings,
            match_case=match_case
        )
        return format_error(error)

    def create_invalid_occurrence_error(
        self,
        occurrence: int,
        total_found: int,
        search_text: str
    ) -> str:
        """Create a structured error for invalid occurrence number."""
        error = DocsErrorBuilder.invalid_occurrence(
            occurrence=occurrence,
            total_found=total_found,
            search_text=search_text
        )
        return format_error(error)

    def create_table_not_found_error(
        self,
        table_index: int,
        total_tables: int
    ) -> str:
        """Create a structured error for table not found."""
        error = DocsErrorBuilder.table_not_found(
            table_index=table_index,
            total_tables=total_tables
        )
        return format_error(error)

    def create_missing_param_error(
        self,
        param_name: str,
        context: str,
        valid_values: Optional[List[str]] = None
    ) -> str:
        """Create a structured error for missing required parameter."""
        error = DocsErrorBuilder.missing_required_param(
            param_name=param_name,
            context_description=context,
            valid_values=valid_values
        )
        return format_error(error)

    def create_invalid_param_error(
        self,
        param_name: str,
        received: Any,
        valid_values: List[str],
        context: str = ""
    ) -> str:
        """Create a structured error for invalid parameter value."""
        error = DocsErrorBuilder.invalid_param_value(
            param_name=param_name,
            received_value=received,
            valid_values=valid_values,
            context_description=context
        )
        return format_error(error)

    def create_invalid_color_error(
        self,
        color_value: str,
        param_name: str = "color"
    ) -> str:
        """Create a structured error for invalid color format."""
        error = DocsErrorBuilder.invalid_color_format(
            color_value=color_value,
            param_name=param_name
        )
        return format_error(error)

    def create_api_error(
        self,
        operation: str,
        error_message: str,
        document_id: Optional[str] = None
    ) -> str:
        """Create a structured error for API failures."""
        error = DocsErrorBuilder.api_error(
            operation=operation,
            error_message=error_message,
            document_id=document_id
        )
        return format_error(error)

    def create_image_error(
        self,
        image_source: str,
        actual_mime_type: Optional[str] = None,
        error_detail: Optional[str] = None
    ) -> str:
        """Create a structured error for image source issues."""
        error = DocsErrorBuilder.invalid_image_source(
            image_source=image_source,
            actual_mime_type=actual_mime_type,
            error_detail=error_detail
        )
        return format_error(error)

    def create_pdf_export_error(
        self,
        document_id: str,
        stage: str,
        error_detail: str
    ) -> str:
        """Create a structured error for PDF export failures."""
        error = DocsErrorBuilder.pdf_export_error(
            document_id=document_id,
            stage=stage,
            error_detail=error_detail
        )
        return format_error(error)

    def create_invalid_document_type_error(
        self,
        document_id: str,
        file_name: str,
        actual_mime_type: str
    ) -> str:
        """Create a structured error for wrong document type."""
        error = DocsErrorBuilder.invalid_document_type(
            document_id=document_id,
            file_name=file_name,
            actual_mime_type=actual_mime_type
        )
        return format_error(error)

    def validate_index_in_bounds(
        self,
        index: int,
        doc_length: int,
        index_name: str = "index"
    ) -> Tuple[bool, Optional[str]]:
        """
        Validate that an index is within document bounds.

        Args:
            index: The index to validate
            doc_length: The document length (exclusive upper bound)
            index_name: Name of the index parameter for error messages

        Returns:
            Tuple of (is_valid, structured_error_json or None)
        """
        # Check type first
        if not isinstance(index, int):
            error = StructuredError(
                code=ErrorCode.INVALID_INDEX_TYPE.value,
                message=f"{index_name} must be an integer, got {type(index).__name__}",
                suggestion="Provide an integer value for the index"
            )
            return False, format_error(error)

        # Check negative
        if index < 0:
            error = StructuredError(
                code=ErrorCode.INVALID_INDEX_RANGE.value,
                message=f"{index_name} cannot be negative, got {index}",
                suggestion="Use get_doc_info to find valid insertion indices"
            )
            return False, format_error(error)

        # Check upper bound
        if index >= doc_length:
            error = DocsErrorBuilder.index_out_of_bounds(
                index_name, index, doc_length
            )
            return False, format_error(error)

        return True, None

    def validate_mutually_exclusive(
        self,
        params: Dict[str, Any],
        exclusive_groups: List[List[str]]
    ) -> Tuple[bool, Optional[str]]:
        """
        Validate that mutually exclusive parameters are not provided together.

        Args:
            params: Dictionary of parameter names to values
            exclusive_groups: List of groups of mutually exclusive parameter names
                             e.g., [['start_index', 'location'], ['search', 'heading']]

        Returns:
            Tuple of (is_valid, structured_error_json or None)

        Example:
            >>> vm = ValidationManager()
            >>> params = {'start_index': 10, 'location': 'end'}
            >>> is_valid, error = vm.validate_mutually_exclusive(
            ...     params,
            ...     [['start_index', 'location', 'search', 'heading']]
            ... )
        """
        for group in exclusive_groups:
            # Find which params in this group have non-None values
            provided = [p for p in group if params.get(p) is not None]

            if len(provided) > 1:
                error = DocsErrorBuilder.conflicting_params(
                    params=provided,
                    message=f"Cannot use {', '.join(repr(p) for p in provided)} together - these parameters are mutually exclusive"
                )
                return False, format_error(error)

        return True, None

    def create_not_found_error(
        self,
        entity_type: str,
        search_criteria: str,
        available_options: Optional[List[str]] = None,
        suggestion: Optional[str] = None
    ) -> str:
        """Create a structured error for entity not found."""
        error = DocsErrorBuilder.not_found(
            entity_type=entity_type,
            search_criteria=search_criteria,
            available_options=available_options,
            suggestion=suggestion
        )
        return format_error(error)

    def create_invalid_state_error(
        self,
        reason: str,
        current_state: str,
        required_state: str,
        suggestion: Optional[str] = None
    ) -> str:
        """Create a structured error for invalid state."""
        error = DocsErrorBuilder.invalid_state(
            reason=reason,
            current_state=current_state,
            required_state=required_state,
            suggestion=suggestion
        )
        return format_error(error)

    def create_out_of_range_error(
        self,
        param_name: str,
        value: Any,
        min_val: Any,
        max_val: Any,
        suggestion: Optional[str] = None
    ) -> str:
        """Create a structured error for value out of range."""
        error = DocsErrorBuilder.out_of_range(
            param_name=param_name,
            value=value,
            min_val=min_val,
            max_val=max_val,
            suggestion=suggestion
        )
        return format_error(error)
