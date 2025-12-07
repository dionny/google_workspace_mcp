"""
Batch Operation Manager

This module provides high-level batch operation management for Google Docs,
extracting complex validation and request building logic.

Features:
- Atomic batch execution (all operations succeed or all fail)
- Search-based positioning (insert before/after search text)
- Automatic position adjustment for sequential operations
- Per-operation results with position shift tracking
"""
import logging
import asyncio
from typing import Any, Union, Dict, List, Tuple, Optional
from dataclasses import dataclass, asdict

from gdocs.docs_helpers import (
    create_insert_text_request,
    create_delete_range_request,
    create_format_text_request,
    create_clear_formatting_request,
    create_find_replace_request,
    create_insert_table_request,
    create_insert_page_break_request,
    create_paragraph_style_request,
    validate_operation,
    resolve_range,
    RangeResult,
    extract_text_at_range,
    find_all_occurrences_in_document,
    extract_document_text_with_indices,
    SearchPosition,
    interpret_escape_sequences,
    find_paragraph_boundaries,
    find_sentence_boundaries,
    find_line_boundaries,
)
from gdocs.docs_structure import parse_document_structure
from gdocs.managers.history_manager import get_history_manager, UndoCapability

logger = logging.getLogger(__name__)

# Operation type aliases for consistent naming across tools
# Maps both short forms and _text suffix forms to canonical _text suffix
OPERATION_ALIASES = {
    # Short forms -> canonical
    'insert': 'insert_text',
    'delete': 'delete_text',
    'format': 'format_text',
    'replace': 'replace_text',
    # Canonical forms (identity mapping)
    'insert_text': 'insert_text',
    'delete_text': 'delete_text',
    'format_text': 'format_text',
    'replace_text': 'replace_text',
    # Non-aliased operations (identity mapping)
    'insert_table': 'insert_table',
    'insert_page_break': 'insert_page_break',
    'find_replace': 'find_replace',
}

# Valid operation types (all keys in OPERATION_ALIASES)
VALID_OPERATION_TYPES = set(OPERATION_ALIASES.keys())


class VirtualTextTracker:
    """
    Tracks virtual document text state to support chained search operations.

    When executing batch operations, we need to resolve search-based positions
    incrementally. If operation A inserts text "[OP1]", operation B should be
    able to search for "[OP1]" even though it hasn't been inserted yet.

    This class maintains a virtual representation of the document text that
    updates as operations are "applied" (resolved), allowing subsequent
    search operations to find text that earlier operations would insert.

    Example:
        # If document contains "[MARKER]"
        ops = [
            {'type': 'insert', 'search': '[MARKER]', 'position': 'after', 'text': '[OP1]'},
            {'type': 'insert', 'search': '[OP1]', 'position': 'after', 'text': '[OP2]'},
        ]
        # Without VirtualTextTracker, op 2 fails because [OP1] doesn't exist yet
        # With VirtualTextTracker, op 2 succeeds because virtual text is updated after op 1
    """

    def __init__(self, doc_data: Dict[str, Any]):
        """
        Initialize the tracker with document data.

        Args:
            doc_data: Raw document data from Google Docs API
        """
        # Extract text segments with their indices
        text_segments = extract_document_text_with_indices(doc_data)

        # Build a continuous text string and index mapping
        self.text = ""
        self.index_map = []  # Maps position in self.text -> document index

        for segment_text, start_idx, end_idx in text_segments:
            self.text += segment_text
            for i in range(len(segment_text)):
                self.index_map.append(start_idx + i)

        # Track the next available document index for appended content
        # This is 1 past the last character in the document
        if self.index_map:
            self._next_doc_index = self.index_map[-1] + 1
        else:
            self._next_doc_index = 1  # Document index starts at 1

        # Track text inserted by previous operations in this batch
        # List of (inserted_text, doc_start_index, doc_end_index)
        self._recent_inserts: List[Tuple[str, int, int]] = []

        logger.debug(
            f"VirtualTextTracker initialized with {len(self.text)} chars, "
            f"next_index={self._next_doc_index}"
        )

    def search_text(
        self,
        search_text: str,
        position: str,
        occurrence: int = 1,
        match_case: bool = True,
        prefer_recent_insert: bool = True
    ) -> Tuple[bool, Optional[int], Optional[int], str]:
        """
        Search for text in the virtual document state.

        Args:
            search_text: Text to search for
            position: Where to operate relative to found text ("before", "after", "replace")
            occurrence: Which occurrence to target (1=first, -1=last)
            match_case: Whether to match case exactly
            prefer_recent_insert: If True and search text matches text that was inserted
                by a previous operation in this batch, prefer that occurrence. This makes
                it easier to insert text and then format it in the same batch.

        Returns:
            Tuple of (success, start_index, end_index, message)
        """
        if not search_text:
            return (False, None, None, "Search text cannot be empty")

        # Check if search text matches any recently inserted text
        # If so, return that occurrence directly (before searching the full document)
        if prefer_recent_insert and occurrence == 1:
            recent_match = self._find_in_recent_inserts(search_text, match_case)
            if recent_match:
                doc_start, doc_end = recent_match
                logger.debug(
                    f"Found '{search_text}' in recently inserted text at {doc_start}-{doc_end}"
                )
                # Calculate indices based on position
                if position == SearchPosition.BEFORE.value:
                    return (True, doc_start, doc_start, f"Found in recently inserted text at index {doc_start}")
                elif position == SearchPosition.AFTER.value:
                    return (True, doc_end, doc_end, f"Found in recently inserted text at {doc_start}, inserting after index {doc_end}")
                elif position == SearchPosition.REPLACE.value:
                    return (True, doc_start, doc_end, f"Found in recently inserted text at index range {doc_start}-{doc_end}")

        # Find occurrences in virtual text
        search_in = self.text if match_case else self.text.lower()
        search_for = search_text if match_case else search_text.lower()

        occurrences = []
        pos = 0
        while True:
            found = search_in.find(search_for, pos)
            if found == -1:
                break
            occurrences.append(found)
            pos = found + 1

        if not occurrences:
            return (False, None, None, f"Text '{search_text}' not found in document")

        # Select the right occurrence
        if occurrence > 0:
            if occurrence > len(occurrences):
                return (
                    False, None, None,
                    f"Occurrence {occurrence} of '{search_text}' not found. "
                    f"Document contains {len(occurrences)} occurrence(s)."
                )
            target_idx = occurrences[occurrence - 1]
        else:  # Negative = from end
            if abs(occurrence) > len(occurrences):
                return (
                    False, None, None,
                    f"Occurrence {occurrence} of '{search_text}' not found. "
                    f"Document contains {len(occurrences)} occurrence(s)."
                )
            target_idx = occurrences[occurrence]

        # Map virtual text position to document index
        if target_idx >= len(self.index_map):
            # Position is beyond original document (in appended virtual text)
            doc_start = self._next_doc_index + (target_idx - len(self.index_map))
        else:
            doc_start = self.index_map[target_idx]

        end_pos = target_idx + len(search_text)
        if end_pos > len(self.index_map):
            # End is beyond original document
            if target_idx < len(self.index_map):
                # Start in original, end in virtual
                chars_in_original = len(self.index_map) - target_idx
                chars_in_virtual = len(search_text) - chars_in_original
                doc_end = self._next_doc_index + chars_in_virtual
            else:
                # Entirely in virtual text
                doc_end = doc_start + len(search_text)
        else:
            doc_end = self.index_map[end_pos - 1] + 1

        # Calculate indices based on position
        if position == SearchPosition.BEFORE.value:
            return (True, doc_start, doc_start, f"Found at index {doc_start}")
        elif position == SearchPosition.AFTER.value:
            return (True, doc_end, doc_end, f"Found at index {doc_start}, inserting after index {doc_end}")
        elif position == SearchPosition.REPLACE.value:
            return (True, doc_start, doc_end, f"Found at index range {doc_start}-{doc_end}")
        else:
            return (False, None, None, f"Invalid position '{position}'. Use 'before', 'after', or 'replace'.")

    def _find_in_recent_inserts(
        self,
        search_text: str,
        match_case: bool = True
    ) -> Optional[Tuple[int, int]]:
        """
        Check if search text appears in any recently inserted text.

        Args:
            search_text: Text to search for
            match_case: Whether to match case exactly

        Returns:
            Tuple of (doc_start_index, doc_end_index) if found, None otherwise
        """
        for inserted_text, doc_start, doc_end in reversed(self._recent_inserts):
            # Check if search_text is contained in this inserted text
            text_to_search = inserted_text if match_case else inserted_text.lower()
            text_to_find = search_text if match_case else search_text.lower()

            pos = text_to_search.find(text_to_find)
            if pos != -1:
                # Found! Calculate the document indices for this match
                match_start = doc_start + pos
                match_end = match_start + len(search_text)
                return (match_start, match_end)

        return None

    def apply_operation(self, op: Dict[str, Any]) -> None:
        """
        Apply an operation to the virtual text state.

        This updates the virtual text to reflect what the document will look like
        after this operation executes. This allows subsequent search operations
        to find text that was "inserted" by earlier operations.

        Args:
            op: Resolved operation dictionary with indices
        """
        op_type = op.get('type', '')

        if op_type == 'insert_text':
            index = op.get('index', 0)
            text = op.get('text', '')
            self._apply_insert(index, text)

        elif op_type == 'delete_text':
            start = op.get('start_index', 0)
            end = op.get('end_index', start)
            self._apply_delete(start, end)

        elif op_type == 'replace_text':
            start = op.get('start_index', 0)
            end = op.get('end_index', start)
            text = op.get('text', '')
            # Replace = delete + insert
            self._apply_delete(start, end)
            self._apply_insert(start, text)

        # format_text doesn't change text content
        # insert_table, insert_page_break are complex structural changes
        # find_replace affects all occurrences (not tracked for simplicity)

    def _apply_insert(self, doc_index: int, text: str) -> None:
        """Insert text at a document index in the virtual state."""
        if not text:
            return

        # Find position in virtual text that corresponds to doc_index
        virtual_pos = self._doc_index_to_virtual_pos(doc_index)

        # Insert text at position
        self.text = self.text[:virtual_pos] + text + self.text[virtual_pos:]

        # Update index map - insert new indices for the inserted text
        new_indices = [doc_index + i for i in range(len(text))]
        self.index_map = self.index_map[:virtual_pos] + new_indices + self.index_map[virtual_pos:]

        # Shift indices after the insertion point
        for i in range(virtual_pos + len(text), len(self.index_map)):
            self.index_map[i] += len(text)

        # Update next doc index
        self._next_doc_index += len(text)

        # Track this insert for recent-insert preference in search
        # Store the text and its document index range
        self._recent_inserts.append((text, doc_index, doc_index + len(text)))

        logger.debug(f"Virtual insert: {len(text)} chars at doc_index {doc_index}")

    def _apply_delete(self, start_idx: int, end_idx: int) -> None:
        """Delete text in a range from the virtual state."""
        if start_idx >= end_idx:
            return

        # Find positions in virtual text
        start_pos = self._doc_index_to_virtual_pos(start_idx)
        end_pos = self._doc_index_to_virtual_pos(end_idx)

        # Delete text
        self.text = self.text[:start_pos] + self.text[end_pos:]

        # Delete from index map
        del self.index_map[start_pos:end_pos]

        # Shift indices after the deletion point
        deleted_len = end_idx - start_idx
        for i in range(start_pos, len(self.index_map)):
            self.index_map[i] -= deleted_len

        # Update next doc index
        self._next_doc_index -= deleted_len

        logger.debug(f"Virtual delete: {deleted_len} chars at {start_idx}-{end_idx}")

    def _doc_index_to_virtual_pos(self, doc_index: int) -> int:
        """
        Convert a document index to a position in the virtual text string.

        Handles cases where the doc_index is:
        1. In the original document (mapped via index_map)
        2. In the virtual "appended" region (beyond original document)
        """
        # Check if it's beyond the tracked range
        if doc_index >= self._next_doc_index:
            # Position is at or beyond the end of virtual text
            return len(self.text)

        # Binary search for the position
        for i, idx in enumerate(self.index_map):
            if idx >= doc_index:
                return i

        # If not found, it's at the end
        return len(self.text)


def normalize_operation_type(op_type: str) -> str:
    """
    Normalize an operation type to its canonical form.

    Accepts both short forms (insert, delete, format, replace) and
    long forms (insert_text, delete_text, format_text, replace_text).

    Args:
        op_type: Operation type string

    Returns:
        Canonical operation type (always with _text suffix for text operations)
    """
    return OPERATION_ALIASES.get(op_type, op_type)


def normalize_operation(operation: Dict[str, Any]) -> Dict[str, Any]:
    """
    Normalize an operation dictionary, converting operation type aliases
    and interpreting escape sequences in text fields.

    Args:
        operation: Operation dictionary with 'type' field

    Returns:
        Copy of operation with normalized type and interpreted escape sequences
    """
    if 'type' not in operation:
        return operation

    normalized = operation.copy()
    normalized['type'] = normalize_operation_type(operation['type'])

    # Interpret escape sequences in text fields
    if 'text' in normalized and normalized['text'] is not None:
        normalized['text'] = interpret_escape_sequences(normalized['text'])
    if 'replace_text' in normalized and normalized['replace_text'] is not None:
        normalized['replace_text'] = interpret_escape_sequences(normalized['replace_text'])

    return normalized


@dataclass
class BatchOperationResult:
    """Result of a single operation within a batch."""
    index: int  # Operation index in the batch
    type: str  # Operation type
    success: bool
    description: str
    position_shift: int = 0  # How much this operation shifted positions
    affected_range: Optional[Dict[str, int]] = None  # {"start": x, "end": y}
    resolved_index: Optional[int] = None  # Index after search resolution
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary, excluding None values."""
        result = asdict(self)
        return {k: v for k, v in result.items() if v is not None}


@dataclass
class BatchExecutionResult:
    """Complete result of batch execution."""
    success: bool
    operations_completed: int
    total_operations: int
    results: List[BatchOperationResult]
    total_position_shift: int
    message: str
    document_link: Optional[str] = None
    preview: bool = False
    would_modify: bool = False

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON response."""
        result = {
            "success": self.success,
            "operations_completed": self.operations_completed,
            "total_operations": self.total_operations,
            "results": [r.to_dict() for r in self.results],
            "total_position_shift": self.total_position_shift,
            "message": self.message,
            "document_link": self.document_link,
        }
        # Include preview fields only if in preview mode
        if self.preview:
            result["preview"] = True
            result["would_modify"] = self.would_modify
        return result


class BatchOperationManager:
    """
    High-level manager for Google Docs batch operations.

    Handles complex multi-operation requests including:
    - Operation validation and request building
    - Batch execution with proper error handling
    - Operation result processing and reporting
    """

    def __init__(self, service, tab_id: str = None):
        """
        Initialize the batch operation manager.

        Args:
            service: Google Docs API service instance
            tab_id: Optional tab ID for multi-tab documents
        """
        self.service = service
        self.tab_id = tab_id

    async def execute_batch_operations(
        self,
        document_id: str,
        operations: list[dict[str, Any]]
    ) -> tuple[bool, str, dict[str, Any]]:
        """
        Execute multiple document operations in a single atomic batch.

        This method is used by batch_edit_doc tool function.

        Args:
            document_id: ID of the document to update
            operations: List of operation dictionaries

        Returns:
            Tuple of (success, message, metadata) where metadata includes:
            - operations_count: Number of operations executed
            - requests_count: Number of API requests sent
            - replies_count: Number of API replies received
            - operation_summary: First 5 operation descriptions
            - total_position_shift: Cumulative position shift from all operations
            - per_operation_shifts: List of position shifts for each operation
        """
        logger.info(f"Executing batch operations on document {document_id}")
        logger.info(f"Operations count: {len(operations)}")

        if not operations:
            return False, "No operations provided. Please provide at least one operation.", {}

        try:
            # Validate and build requests, tracking position shifts
            requests, operation_descriptions, position_shifts = await self._validate_and_build_requests_with_shifts(operations)

            if not requests:
                return False, "No valid requests could be built from operations", {}

            # Execute the batch
            result = await self._execute_batch_requests(document_id, requests)

            # Calculate cumulative position shift
            total_position_shift = sum(position_shifts)

            # Process results
            metadata = {
                'operations_count': len(operations),
                'requests_count': len(requests),
                'replies_count': len(result.get('replies', [])),
                'operation_summary': operation_descriptions[:5],  # First 5 operations
                'total_position_shift': total_position_shift,
                'per_operation_shifts': position_shifts,
                'document_link': f"https://docs.google.com/document/d/{document_id}/edit"
            }

            summary = self._build_operation_summary(operation_descriptions)

            return True, f"Successfully executed {len(operations)} operations ({summary})", metadata

        except Exception as e:
            logger.error(f"Failed to execute batch operations: {str(e)}")
            return False, f"Batch operation failed: {str(e)}", {}

    async def _validate_and_build_requests(
        self,
        operations: list[dict[str, Any]]
    ) -> tuple[list[dict[str, Any]], list[str]]:
        """
        Validate operations and build API requests.

        Args:
            operations: List of operation dictionaries

        Returns:
            Tuple of (requests, operation_descriptions)
        """
        requests = []
        operation_descriptions = []

        for i, op in enumerate(operations):
            # Normalize operation type aliases (insert -> insert_text, etc.)
            op = normalize_operation(op)

            # Validate operation structure
            is_valid, error_msg = validate_operation(op)
            if not is_valid:
                raise ValueError(f"Operation {i+1}: {error_msg}")

            op_type = op.get('type')

            try:
                # Build request based on operation type
                result = self._build_operation_request(op, op_type)

                # Handle both single request and list of requests
                if isinstance(result[0], list):
                    # Multiple requests (e.g., replace_text)
                    for req in result[0]:
                        requests.append(req)
                    operation_descriptions.append(result[1])
                elif result[0]:
                    # Single request
                    requests.append(result[0])
                    operation_descriptions.append(result[1])

            except KeyError as e:
                raise ValueError(f"Operation {i+1} ({op_type}) missing required field: {e}")
            except Exception as e:
                raise ValueError(f"Operation {i+1} ({op_type}) failed validation: {str(e)}")

        return requests, operation_descriptions

    async def _validate_and_build_requests_with_shifts(
        self,
        operations: list[dict[str, Any]]
    ) -> tuple[list[dict[str, Any]], list[str], list[int]]:
        """
        Validate operations and build API requests, tracking position shifts.

        Args:
            operations: List of operation dictionaries

        Returns:
            Tuple of (requests, operation_descriptions, position_shifts)
        """
        requests = []
        operation_descriptions = []
        position_shifts = []

        for i, op in enumerate(operations):
            # Normalize operation type aliases (insert -> insert_text, etc.)
            op = normalize_operation(op)

            # Validate operation structure
            is_valid, error_msg = validate_operation(op)
            if not is_valid:
                raise ValueError(f"Operation {i+1}: {error_msg}")

            op_type = op.get('type')

            try:
                # Build request based on operation type
                result = self._build_operation_request(op, op_type)

                # Handle both single request and list of requests
                if isinstance(result[0], list):
                    # Multiple requests (e.g., replace_text)
                    for req in result[0]:
                        requests.append(req)
                    operation_descriptions.append(result[1])
                elif result[0]:
                    # Single request
                    requests.append(result[0])
                    operation_descriptions.append(result[1])

                # Calculate position shift for this operation
                shift = self._calculate_operation_shift(op, op_type)
                position_shifts.append(shift)

            except KeyError as e:
                raise ValueError(f"Operation {i+1} ({op_type}) missing required field: {e}")
            except Exception as e:
                raise ValueError(f"Operation {i+1} ({op_type}) failed validation: {str(e)}")

        return requests, operation_descriptions, position_shifts

    def _calculate_operation_shift(self, op: dict[str, Any], op_type: str) -> int:
        """
        Calculate the position shift caused by a single operation.

        Args:
            op: Operation dictionary
            op_type: Operation type

        Returns:
            Position shift (positive = content added, negative = content removed)
        """
        if op_type == 'insert_text':
            return len(op.get('text', ''))

        elif op_type == 'delete_text':
            start = op.get('start_index', 0)
            end = op.get('end_index', start)
            return -(end - start)

        elif op_type == 'replace_text':
            start = op.get('start_index', 0)
            end = op.get('end_index', start)
            old_len = end - start
            new_len = len(op.get('text', ''))
            return new_len - old_len

        elif op_type == 'format_text':
            # Formatting doesn't change positions
            return 0

        elif op_type == 'insert_table':
            # Tables add structure - this is complex but we return 0
            # as an approximation since table structure shift is hard to calculate
            return 0

        elif op_type == 'insert_page_break':
            # Page break adds 1 character
            return 1

        elif op_type == 'find_replace':
            # Find/replace position shift is unpredictable
            # (depends on number of occurrences and length difference)
            # Return 0 as this needs document context to calculate
            return 0

        return 0

    def _build_operation_request(
        self,
        op: dict[str, Any],
        op_type: str
    ) -> Tuple[Union[Dict[str, Any], List[Dict[str, Any]]], str]:
        """
        Build a single operation request.

        Args:
            op: Operation dictionary
            op_type: Operation type

        Returns:
            Tuple of (request, description)
        """
        if op_type == 'insert_text':
            request = create_insert_text_request(op['index'], op['text'], tab_id=self.tab_id)
            description = f"insert text at {op['index']}"

        elif op_type == 'delete_text':
            request = create_delete_range_request(op['start_index'], op['end_index'], tab_id=self.tab_id)
            description = f"delete text {op['start_index']}-{op['end_index']}"

        elif op_type == 'replace_text':
            # Replace is delete + insert (must be done in this order)
            delete_request = create_delete_range_request(op['start_index'], op['end_index'], tab_id=self.tab_id)
            insert_request = create_insert_text_request(op['start_index'], op['text'], tab_id=self.tab_id)
            # Return both requests as a list
            request = [delete_request, insert_request]
            description = f"replace text {op['start_index']}-{op['end_index']} with '{op['text'][:20]}{'...' if len(op['text']) > 20 else ''}'"

        elif op_type == 'format_text':
            # Handle text-level formatting
            request = create_format_text_request(
                op['start_index'], op['end_index'],
                op.get('bold'), op.get('italic'), op.get('underline'),
                op.get('strikethrough'), op.get('small_caps'), op.get('subscript'),
                op.get('superscript'), op.get('font_size'), op.get('font_family'),
                op.get('link'), op.get('foreground_color'), op.get('background_color'),
                tab_id=self.tab_id,
            )

            # Handle paragraph-level formatting
            para_request = create_paragraph_style_request(
                op['start_index'], op['end_index'],
                line_spacing=op.get('line_spacing'),
                heading_style=op.get('heading_style'),
                alignment=op.get('alignment'),
                indent_first_line=op.get('indent_first_line'),
                indent_start=op.get('indent_start'),
                indent_end=op.get('indent_end'),
                space_above=op.get('space_above'),
                space_below=op.get('space_below'),
                tab_id=self.tab_id,
            )

            # Combine text and paragraph requests
            if request and para_request:
                request = [request, para_request]
            elif para_request:
                request = para_request
            elif not request:
                raise ValueError("No formatting options provided")

            # Build format description
            format_changes = []
            # Text-level formatting
            for param, name in [
                ('bold', 'bold'), ('italic', 'italic'), ('underline', 'underline'),
                ('strikethrough', 'strikethrough'), ('small_caps', 'small caps'),
                ('subscript', 'subscript'), ('superscript', 'superscript'),
                ('font_size', 'font size'), ('font_family', 'font family'),
                ('link', 'link'), ('foreground_color', 'foreground color'),
                ('background_color', 'background color')
            ]:
                if op.get(param) is not None:
                    value = f"{op[param]}pt" if param == 'font_size' else op[param]
                    format_changes.append(f"{name}: {value}")
            # Paragraph-level formatting
            for param, name in [
                ('heading_style', 'heading style'), ('alignment', 'alignment'),
                ('line_spacing', 'line spacing'), ('indent_first_line', 'first line indent'),
                ('indent_start', 'left indent'), ('indent_end', 'right indent'),
                ('space_above', 'space above'), ('space_below', 'space below')
            ]:
                if op.get(param) is not None:
                    format_changes.append(f"{name}: {op[param]}")

            description = f"format text {op['start_index']}-{op['end_index']} ({', '.join(format_changes)})"

        elif op_type == 'insert_table':
            # Validate required fields with helpful error messages
            missing = []
            if 'index' not in op:
                missing.append('index')
            if 'rows' not in op:
                missing.append('rows')
            if 'columns' not in op:
                missing.append('columns')
            if missing:
                raise KeyError(
                    f"insert_table operation missing required field(s): {', '.join(missing)}. "
                    f"Example: {{\"type\": \"insert_table\", \"index\": 1, \"rows\": 3, \"columns\": 4}}"
                )
            request = create_insert_table_request(op['index'], op['rows'], op['columns'], tab_id=self.tab_id)
            description = f"insert {op['rows']}x{op['columns']} table at {op['index']}"

        elif op_type == 'insert_page_break':
            request = create_insert_page_break_request(op['index'], tab_id=self.tab_id)
            description = f"insert page break at {op['index']}"

        elif op_type == 'find_replace':
            # Note: find_replace uses tab_ids (list) instead of tab_id (single)
            tab_ids = [self.tab_id] if self.tab_id else None
            request = create_find_replace_request(
                op['find_text'], op['replace_text'], op.get('match_case', False), tab_ids=tab_ids
            )
            description = f"find/replace '{op['find_text']}' → '{op['replace_text']}'"

        else:
            supported_types = [
                'insert_text', 'delete_text', 'replace_text', 'format_text',
                'insert_table', 'insert_page_break', 'find_replace'
            ]
            raise ValueError(f"Unsupported operation type '{op_type}'. Supported: {', '.join(supported_types)}")

        return request, description

    async def _execute_batch_requests(
        self,
        document_id: str,
        requests: list[dict[str, Any]]
    ) -> dict[str, Any]:
        """
        Execute the batch requests against the Google Docs API.
        
        Args:
            document_id: Document ID
            requests: List of API requests
            
        Returns:
            API response
        """
        return await asyncio.to_thread(
            self.service.documents().batchUpdate(
                documentId=document_id,
                body={'requests': requests}
            ).execute
        )

    def _build_operation_summary(self, operation_descriptions: list[str]) -> str:
        """
        Build a concise summary of operations performed.
        
        Args:
            operation_descriptions: List of operation descriptions
            
        Returns:
            Summary string
        """
        if not operation_descriptions:
            return "no operations"

        summary_items = operation_descriptions[:3]  # Show first 3 operations
        summary = ', '.join(summary_items)

        if len(operation_descriptions) > 3:
            remaining = len(operation_descriptions) - 3
            summary += f" and {remaining} more operation{'s' if remaining > 1 else ''}"

        return summary

    def get_supported_operations(self) -> dict[str, Any]:
        """
        Get information about supported batch operations.

        Returns:
            Dictionary with supported operation types and their required parameters
        """
        return {
            'supported_operations': {
                'insert_text': {
                    'aliases': ['insert'],
                    'required': ['index', 'text'],
                    'description': 'Insert text at specified index'
                },
                'delete_text': {
                    'aliases': ['delete'],
                    'required': ['start_index', 'end_index'],
                    'description': 'Delete text in specified range'
                },
                'replace_text': {
                    'aliases': ['replace'],
                    'required': ['start_index', 'end_index', 'text'],
                    'description': 'Replace text in range with new text'
                },
                'format_text': {
                    'aliases': ['format'],
                    'required': ['start_index', 'end_index'],
                    'optional': ['bold', 'italic', 'underline', 'font_size', 'font_family'],
                    'description': 'Apply formatting to text range'
                },
                'insert_table': {
                    'required': ['index', 'rows', 'columns'],
                    'description': 'Insert table at specified index'
                },
                'insert_page_break': {
                    'required': ['index'],
                    'description': 'Insert page break at specified index'
                },
                'find_replace': {
                    'required': ['find_text', 'replace_text'],
                    'optional': ['match_case'],
                    'description': 'Find and replace text throughout document'
                }
            },
            'operation_aliases': {
                'insert': 'insert_text',
                'delete': 'delete_text',
                'replace': 'replace_text',
                'format': 'format_text',
            },
            'example_operations': [
                {"type": "insert", "index": 1, "text": "Hello World"},
                {"type": "format", "start_index": 1, "end_index": 12, "bold": True},
                {"type": "insert_table", "index": 20, "rows": 2, "columns": 3}
            ]
        }

    async def execute_batch_with_search(
        self,
        document_id: str,
        operations: List[Dict[str, Any]],
        auto_adjust_positions: bool = True,
        preview_only: bool = False,
    ) -> BatchExecutionResult:
        """
        Execute batch operations with search-based positioning and position tracking.

        This enhanced method supports:
        - Search-based positioning (e.g., insert before/after "Conclusion")
        - all_occurrences mode (apply operation to ALL matches of search text)
        - Automatic position adjustment for sequential operations
        - Detailed per-operation results with position shifts

        Args:
            document_id: ID of the document to update
            operations: List of operation dictionaries. Each can include:
                - type: Operation type (insert, delete, replace, format)
                - Standard index-based params (index, start_index, end_index)
                - OR search-based params (search, position="before"|"after"|"replace")
                - all_occurrences: If True with search, apply to ALL occurrences (not just first)
                - occurrence: Which occurrence to target (1=first, -1=last) when not using all_occurrences
                - match_case: Whether to match case exactly (default: True)
                - text: Text for insert/replace operations
                - Formatting params for format operations
            auto_adjust_positions: If True, automatically adjust positions for
                subsequent operations based on cumulative position shifts
            preview_only: If True, returns what would be modified without actually
                executing the operations. Useful for validating operations before
                committing changes.

        Returns:
            BatchExecutionResult with per-operation details. In preview mode,
            includes "preview": true and "would_modify": true fields.

        Example operations with search-based positioning:
            [
                {
                    "type": "insert",
                    "search": "Conclusion",
                    "position": "before",
                    "text": "\\n\\nNew section content here.\\n"
                },
                {
                    "type": "format",
                    "search": "Important Note",
                    "position": "replace",
                    "bold": True
                }
            ]

        Example with all_occurrences (format ALL "TODO" markers):
            [
                {
                    "type": "format",
                    "search": "TODO",
                    "all_occurrences": True,
                    "bold": True,
                    "foreground_color": "red"
                },
                {
                    "type": "replace",
                    "search": "DRAFT",
                    "all_occurrences": True,
                    "text": "FINAL"
                }
            ]
        """
        logger.info(f"Executing enhanced batch on document {document_id}")
        logger.info(f"Operations: {len(operations)}, auto_adjust: {auto_adjust_positions}, preview: {preview_only}")

        if not operations:
            return BatchExecutionResult(
                success=False,
                operations_completed=0,
                total_operations=0,
                results=[],
                total_position_shift=0,
                message="No operations provided",
            )

        # First, fetch document to resolve search-based positions
        try:
            doc_data = await asyncio.to_thread(
                self.service.documents().get(documentId=document_id).execute
            )
        except Exception as e:
            return BatchExecutionResult(
                success=False,
                operations_completed=0,
                total_operations=len(operations),
                results=[],
                total_position_shift=0,
                message=f"Failed to fetch document: {str(e)}",
            )

        # Expand all_occurrences operations into individual operations
        # This must happen BEFORE resolution since expanded operations use indices
        expanded_operations = self._expand_all_occurrences_operations(operations, doc_data)
        logger.debug(f"Expanded {len(operations)} operations to {len(expanded_operations)} operations")

        # Resolve all operations and build requests
        resolved_ops, operation_results = await self._resolve_operations_with_search(
            expanded_operations, doc_data, auto_adjust_positions
        )

        # Check if any resolution failed
        failed_results = [r for r in operation_results if not r.success]
        if failed_results:
            return BatchExecutionResult(
                success=False,
                operations_completed=0,
                total_operations=len(operations),
                results=operation_results,
                total_position_shift=0,
                message=f"Failed to resolve {len(failed_results)} operation(s)",
                preview=preview_only,
            )

        # Handle preview mode - return what would change without executing
        if preview_only:
            total_shift = sum(r.position_shift for r in operation_results)

            # Build message with info about expansion if applicable
            if len(expanded_operations) != len(operations):
                message = f"Would execute {len(operations)} operation(s) ({len(expanded_operations)} after all_occurrences expansion)"
            else:
                message = f"Would execute {len(operations)} operation(s)"

            return BatchExecutionResult(
                success=True,
                operations_completed=0,  # No operations actually executed
                total_operations=len(operations),
                results=operation_results,
                total_position_shift=total_shift,
                message=message,
                document_link=f"https://docs.google.com/document/d/{document_id}/edit",
                preview=True,
                would_modify=len(operation_results) > 0,
            )

        # Build API requests from resolved operations
        try:
            requests = self._build_requests_from_resolved(resolved_ops)
        except Exception as e:
            return BatchExecutionResult(
                success=False,
                operations_completed=0,
                total_operations=len(operations),
                results=operation_results,
                total_position_shift=0,
                message=f"Failed to build requests: {str(e)}",
            )

        # Execute the batch
        try:
            await self._execute_batch_requests(document_id, requests)

            # Mark all operations as successful
            for result in operation_results:
                result.success = True

            total_shift = sum(r.position_shift for r in operation_results)

            # Record operations for undo history (automatic tracking)
            try:
                self._record_batch_operations_to_history(
                    document_id, resolved_ops, operation_results, doc_data
                )
            except Exception as e:
                logger.warning(f"Failed to record batch operations for undo history: {e}")

            # Build message with info about expansion if applicable
            if len(expanded_operations) != len(operations):
                message = f"Successfully executed {len(operations)} operation(s) ({len(expanded_operations)} after all_occurrences expansion)"
            else:
                message = f"Successfully executed {len(operations)} operation(s)"

            return BatchExecutionResult(
                success=True,
                operations_completed=len(expanded_operations),
                total_operations=len(operations),
                results=operation_results,
                total_position_shift=total_shift,
                message=message,
                document_link=f"https://docs.google.com/document/d/{document_id}/edit",
            )

        except Exception as e:
            return BatchExecutionResult(
                success=False,
                operations_completed=0,
                total_operations=len(operations),
                results=operation_results,
                total_position_shift=0,
                message=f"Batch execution failed: {str(e)}",
            )

    def _expand_all_occurrences_operations(
        self,
        operations: List[Dict[str, Any]],
        doc_data: Dict[str, Any],
    ) -> List[Dict[str, Any]]:
        """
        Expand operations with all_occurrences=True into individual operations.

        When an operation has all_occurrences=True, it is expanded into separate
        operations for each occurrence of the search text. Operations are ordered
        in reverse document order (last occurrence first) to maintain correct
        indices during sequential execution.

        Args:
            operations: List of operations, some may have all_occurrences=True
            doc_data: Document data for finding occurrences

        Returns:
            Expanded list of operations (all_occurrences operations are split)
        """
        expanded = []

        for op in operations:
            # Only expand search-based operations with all_occurrences=True
            if op.get('search') and op.get('all_occurrences', False):
                search_text = op['search']
                match_case = op.get('match_case', True)

                # Find all occurrences
                occurrences = find_all_occurrences_in_document(
                    doc_data, search_text, match_case
                )

                if not occurrences:
                    # No occurrences - keep original op (will fail with useful error)
                    expanded.append(op)
                    continue

                # Process in reverse order to maintain correct indices
                # (editing from end of doc toward beginning)
                for start_idx, end_idx in reversed(occurrences):
                    # Create a new operation for this occurrence
                    op_copy = op.copy()
                    # Remove search-related fields since we're converting to index-based
                    op_copy.pop('search', None)
                    op_copy.pop('all_occurrences', None)
                    op_copy.pop('occurrence', None)
                    op_copy.pop('position', None)

                    # Apply position based on original position setting
                    position = op.get('position', 'replace')
                    op_type = normalize_operation_type(op.get('type', ''))

                    if op_type in ('insert', 'insert_text'):
                        # For insert, use position to determine index
                        if position == 'before':
                            op_copy['index'] = start_idx
                        elif position == 'after':
                            op_copy['index'] = end_idx
                        else:  # replace - insert at start, will need delete first
                            op_copy['index'] = start_idx
                        op_copy['type'] = 'insert_text'

                    elif op_type in ('delete', 'delete_text'):
                        op_copy['start_index'] = start_idx
                        op_copy['end_index'] = end_idx
                        op_copy['type'] = 'delete_text'

                    elif op_type in ('replace', 'replace_text'):
                        op_copy['start_index'] = start_idx
                        op_copy['end_index'] = end_idx
                        op_copy['type'] = 'replace_text'

                    elif op_type in ('format', 'format_text'):
                        op_copy['start_index'] = start_idx
                        op_copy['end_index'] = end_idx
                        op_copy['type'] = 'format_text'

                    else:
                        # Unsupported operation type for all_occurrences
                        logger.warning(
                            f"all_occurrences not supported for type '{op_type}', "
                            "keeping original operation"
                        )
                        expanded.append(op)
                        break

                    expanded.append(op_copy)
            else:
                # Regular operation - keep as-is
                expanded.append(op)

        return expanded

    async def _resolve_operations_with_search(
        self,
        operations: List[Dict[str, Any]],
        doc_data: Dict[str, Any],
        auto_adjust: bool,
    ) -> Tuple[List[Dict[str, Any]], List[BatchOperationResult]]:
        """
        Resolve search-based positions, range specs, and calculate position shifts.

        Supports three positioning modes:
        1. Direct index-based: {"type": "insert_text", "index": 100, ...}
        2. Search-based: {"type": "insert", "search": "text", "position": "before", ...}
           Optionally with extend: {"search": "text", "position": "after", "extend": "paragraph"}
        3. Range spec: {"type": "delete", "range_spec": {"search": "x", "extend": "paragraph"}, ...}

        The 'extend' parameter in search-based operations allows extending the resolved
        position to semantic boundaries:
        - "paragraph": Extend to paragraph boundaries
        - "sentence": Extend to sentence boundaries
        - "line": Extend to line boundaries

        For example, with extend='paragraph':
        - position='after': Insert after the end of the paragraph containing the search text
        - position='before': Insert before the start of the paragraph containing the search text
        - position='replace': Replace the entire paragraph containing the search text

        Chained operations are supported: if operation A inserts text that operation B
        searches for, the search will find the text because we maintain a virtual document
        state that reflects pending changes.

        Args:
            operations: Raw operation list
            doc_data: Document data for search/range resolution
            auto_adjust: Whether to auto-adjust positions

        Returns:
            Tuple of (resolved_operations, operation_results)
        """
        resolved_ops = []
        results = []
        cumulative_shift = 0

        # Create a virtual text tracker to support chained search operations
        # This allows searching for text that was "inserted" by earlier operations
        virtual_text_tracker = VirtualTextTracker(doc_data)

        for i, op in enumerate(operations):
            op_type = op.get('type', '')
            op_copy = op.copy()

            # Validate operation type before processing
            if not op_type:
                valid_types = sorted(set(OPERATION_ALIASES.values()))  # Canonical types only
                example = {"type": "insert", "search": "target text", "position": "after", "text": "new text"}
                results.append(BatchOperationResult(
                    index=i,
                    type='',
                    success=False,
                    description="Invalid operation",
                    error=f"Missing 'type' field in operation at index {i}. "
                          f"Valid types: {', '.join(valid_types)}. "
                          f"Example: {example}",
                ))
                resolved_ops.append(None)
                continue

            if op_type not in VALID_OPERATION_TYPES:
                valid_types = sorted(set(OPERATION_ALIASES.values()))  # Canonical types only
                results.append(BatchOperationResult(
                    index=i,
                    type=op_type,
                    success=False,
                    description="Invalid operation type",
                    error=f"Unsupported operation type: '{op_type}' at index {i}. Valid types: {', '.join(valid_types)}",
                ))
                resolved_ops.append(None)
                continue

            # Validate required fields for insert operations
            normalized_op_type = normalize_operation_type(op_type)
            if normalized_op_type == 'insert_text':
                # Insert requires 'text' field
                if 'text' not in op:
                    results.append(BatchOperationResult(
                        index=i,
                        type=op_type,
                        success=False,
                        description="Invalid insert operation",
                        error="Insert operation requires 'text' field",
                    ))
                    resolved_ops.append(None)
                    continue

                # Insert requires at least one positioning method
                has_position = any(key in op for key in ['index', 'search', 'location', 'range_spec'])
                if not has_position:
                    results.append(BatchOperationResult(
                        index=i,
                        type=op_type,
                        success=False,
                        description="Invalid insert operation",
                        error="Insert operation requires positioning: 'index', 'search', 'location', or 'range_spec'",
                    ))
                    resolved_ops.append(None)
                    continue

            # Check if this operation uses range_spec (new range-based positioning)
            if 'range_spec' in op:
                range_result = self._resolve_range_spec(op, doc_data, cumulative_shift, auto_adjust)

                if not range_result.success:
                    results.append(BatchOperationResult(
                        index=i,
                        type=op_type,
                        success=False,
                        description="Range resolution failed",
                        error=range_result.message,
                    ))
                    resolved_ops.append(None)
                    continue

                # Convert to index-based operation
                op_copy = self._convert_range_to_index_op(
                    op, op_type, range_result.start_index, range_result.end_index
                )
                resolved_index = range_result.start_index

            # Check if this operation uses location-based positioning
            elif 'location' in op:
                location = op['location']
                if location not in ('start', 'end'):
                    results.append(BatchOperationResult(
                        index=i,
                        type=op_type,
                        success=False,
                        description="Invalid location value",
                        error=f"Invalid location: '{location}'. Valid values: 'start', 'end'",
                    ))
                    resolved_ops.append(None)
                    continue

                # Resolve location to index using document structure
                structure = parse_document_structure(doc_data)
                total_length = structure['total_length']

                if location == 'end':
                    # Use total_length - 1 for safe insertion (last valid index)
                    resolved_index = total_length - 1 if total_length > 1 else 1
                else:  # location == 'start'
                    resolved_index = 1  # After the initial section break at index 0

                # Apply cumulative shift if auto-adjusting
                if auto_adjust and cumulative_shift != 0:
                    resolved_index += cumulative_shift

                # Convert to index-based operation
                op_copy = self._convert_location_to_index_op(op, op_type, resolved_index)

            # Check if this is a search-based operation (legacy mode)
            elif 'search' in op:
                search_text = op['search']
                position = op.get('position', 'replace')
                occurrence = op.get('occurrence', 1)
                match_case = op.get('match_case', True)
                extend = op.get('extend')  # Optional: 'paragraph', 'sentence', 'line'

                # Use virtual text tracker to support chained operations
                # (searching for text inserted by earlier operations in the batch)
                success, start_idx, end_idx, msg = virtual_text_tracker.search_text(
                    search_text, position, occurrence, match_case
                )

                if not success:
                    results.append(BatchOperationResult(
                        index=i,
                        type=op_type,
                        success=False,
                        description=f"Search failed for '{search_text}'",
                        error=msg,
                    ))
                    resolved_ops.append(None)
                    continue

                # If extend is specified, expand the position to boundary
                if extend:
                    extend_lower = extend.lower()
                    # Get the actual position where the search text was found
                    # For 'before' position, start_idx is where text was found
                    # For 'after' position, we need to look back to find original start
                    # For 'replace' position, start_idx is the match start

                    # Find boundaries based on the found position
                    if extend_lower == 'paragraph':
                        boundary_start, boundary_end = find_paragraph_boundaries(doc_data, start_idx)
                    elif extend_lower == 'sentence':
                        boundary_start, boundary_end = find_sentence_boundaries(doc_data, start_idx)
                    elif extend_lower == 'line':
                        boundary_start, boundary_end = find_line_boundaries(doc_data, start_idx)
                    else:
                        results.append(BatchOperationResult(
                            index=i,
                            type=op_type,
                            success=False,
                            description=f"Invalid extend value: '{extend}'",
                            error=f"Invalid extend value: '{extend}'. Valid values: 'paragraph', 'sentence', 'line'",
                        ))
                        resolved_ops.append(None)
                        continue

                    # Adjust indices based on position
                    if position == SearchPosition.BEFORE.value:
                        # Insert before the boundary (e.g., before the paragraph)
                        start_idx = boundary_start
                        end_idx = boundary_start
                    elif position == SearchPosition.AFTER.value:
                        # Insert after the boundary (e.g., after the paragraph)
                        start_idx = boundary_end
                        end_idx = boundary_end
                    elif position == SearchPosition.REPLACE.value:
                        # Replace the entire boundary (e.g., the whole paragraph)
                        start_idx = boundary_start
                        end_idx = boundary_end

                # Note: cumulative_shift is already tracked by virtual_text_tracker
                # We don't need to apply it separately when using the tracker

                # Map to standard index-based operation
                op_copy = self._convert_search_to_index_op(
                    op, op_type, start_idx, end_idx
                )
                resolved_index = start_idx
            else:
                # Index-based operation - normalize type and apply cumulative shift
                resolved_index = None
                # Normalize the operation type (e.g., 'insert' -> 'insert_text')
                if 'type' in op_copy:
                    op_copy['type'] = normalize_operation_type(op_copy['type'])
                if auto_adjust and cumulative_shift != 0:
                    op_copy = self._apply_position_shift(op_copy, cumulative_shift)

            # Calculate position shift for this operation
            shift, affected_range = self._calculate_op_shift(op_copy)

            results.append(BatchOperationResult(
                index=i,
                type=op_type,
                success=True,  # Will be confirmed after execution
                description=self._describe_operation(op_copy),
                position_shift=shift,
                affected_range=affected_range,
                resolved_index=resolved_index,
            ))

            resolved_ops.append(op_copy)

            # Update cumulative shift for next operation
            if auto_adjust:
                cumulative_shift += shift

            # Update virtual text tracker for chained search operations
            # This allows subsequent operations to search for text that was
            # "inserted" by earlier operations in the batch
            virtual_text_tracker.apply_operation(op_copy)

        return resolved_ops, results

    def _resolve_range_spec(
        self,
        op: Dict[str, Any],
        doc_data: Dict[str, Any],
        cumulative_shift: int,
        auto_adjust: bool,
    ) -> RangeResult:
        """
        Resolve a range_spec to start/end indices.

        Range specs support multiple formats:
        1. Search bounds: {"start": {"search": "x"}, "end": {"search": "y"}}
        2. Search with extension: {"search": "x", "extend": "paragraph"}
        3. Search with offsets: {"search": "x", "before_chars": 10, "after_chars": 20}
        4. Section reference: {"section": "Heading Name"}

        Args:
            op: Operation with range_spec field
            doc_data: Document data for resolution
            cumulative_shift: Current cumulative position shift
            auto_adjust: Whether to apply position adjustments

        Returns:
            RangeResult with resolved start/end indices
        """
        range_spec = op.get('range_spec', {})

        # Resolve range using the centralized resolver
        result = resolve_range(doc_data, range_spec)

        if not result.success:
            return result

        # Apply cumulative shift if auto-adjusting
        start_idx = result.start_index
        end_idx = result.end_index

        if auto_adjust and cumulative_shift != 0:
            start_idx += cumulative_shift
            if end_idx is not None:
                end_idx += cumulative_shift

        return RangeResult(
            success=True,
            start_index=start_idx,
            end_index=end_idx,
            message=result.message,
            matched_start=result.matched_start,
            matched_end=result.matched_end,
            extend_type=result.extend_type,
            section_name=result.section_name,
        )

    def _convert_range_to_index_op(
        self,
        op: Dict[str, Any],
        op_type: str,
        start_idx: int,
        end_idx: Optional[int],
    ) -> Dict[str, Any]:
        """Convert a range_spec-based operation to index-based."""
        result = op.copy()

        # Remove range_spec field
        result.pop('range_spec', None)

        # Map based on operation type
        if op_type in ('insert', 'insert_text'):
            result['type'] = 'insert_text'
            result['index'] = start_idx
        elif op_type in ('delete', 'delete_text'):
            result['type'] = 'delete_text'
            result['start_index'] = start_idx
            result['end_index'] = end_idx
        elif op_type in ('replace', 'replace_text'):
            result['type'] = 'replace_text'
            result['start_index'] = start_idx
            result['end_index'] = end_idx
        elif op_type in ('format', 'format_text'):
            result['type'] = 'format_text'
            result['start_index'] = start_idx
            result['end_index'] = end_idx
        elif op_type == 'insert_table':
            result['type'] = 'insert_table'
            result['index'] = start_idx
        elif op_type == 'insert_page_break':
            result['type'] = 'insert_page_break'
            result['index'] = start_idx

        return result

    def _convert_search_to_index_op(
        self,
        op: Dict[str, Any],
        op_type: str,
        start_idx: int,
        end_idx: Optional[int],
    ) -> Dict[str, Any]:
        """Convert a search-based operation to index-based."""
        result = op.copy()

        # Remove search-specific fields
        for key in ['search', 'position', 'occurrence', 'match_case']:
            result.pop(key, None)

        # Map based on operation type
        if op_type in ('insert', 'insert_text'):
            result['type'] = 'insert_text'
            result['index'] = start_idx
        elif op_type in ('delete', 'delete_text'):
            result['type'] = 'delete_text'
            result['start_index'] = start_idx
            result['end_index'] = end_idx
        elif op_type in ('replace', 'replace_text'):
            result['type'] = 'replace_text'
            result['start_index'] = start_idx
            result['end_index'] = end_idx
        elif op_type in ('format', 'format_text'):
            result['type'] = 'format_text'
            result['start_index'] = start_idx
            result['end_index'] = end_idx
        elif op_type == 'insert_table':
            result['type'] = 'insert_table'
            result['index'] = start_idx
        elif op_type == 'insert_page_break':
            result['type'] = 'insert_page_break'
            result['index'] = start_idx

        return result

    def _convert_location_to_index_op(
        self,
        op: Dict[str, Any],
        op_type: str,
        resolved_index: int,
    ) -> Dict[str, Any]:
        """Convert a location-based operation to index-based.

        This handles operations that use location='start' or location='end'
        for convenient document positioning.

        Args:
            op: Original operation dictionary
            op_type: Operation type (insert, format, etc.)
            resolved_index: The resolved document index

        Returns:
            Index-based operation dictionary
        """
        result = op.copy()

        # Remove location-specific field
        result.pop('location', None)

        # Map based on operation type - location-based ops are insert operations
        if op_type in ('insert', 'insert_text'):
            result['type'] = 'insert_text'
            result['index'] = resolved_index
        elif op_type in ('format', 'format_text'):
            # Format at location doesn't make sense without a range
            # but we support it for consistency - uses resolved_index as start
            result['type'] = 'format_text'
            result['start_index'] = resolved_index
            # If no end_index provided, this will need text length
            if 'end_index' not in result:
                text = result.get('text', '')
                result['end_index'] = resolved_index + len(text) if text else resolved_index
        elif op_type == 'insert_table':
            result['type'] = 'insert_table'
            result['index'] = resolved_index
        elif op_type == 'insert_page_break':
            result['type'] = 'insert_page_break'
            result['index'] = resolved_index
        else:
            # For other types, just set the index
            result['index'] = resolved_index

        return result

    def _apply_position_shift(
        self,
        op: Dict[str, Any],
        shift: int,
    ) -> Dict[str, Any]:
        """Apply position shift to an index-based operation."""
        result = op.copy()

        if 'index' in result:
            result['index'] += shift
        if 'start_index' in result:
            result['start_index'] += shift
        if 'end_index' in result:
            result['end_index'] += shift

        return result

    def _calculate_op_shift(
        self,
        op: Dict[str, Any],
    ) -> Tuple[int, Dict[str, int]]:
        """Calculate position shift caused by an operation."""
        op_type = op.get('type', '')

        if op_type == 'insert_text':
            text_len = len(op.get('text', ''))
            start_idx = op.get('index', 0)
            return text_len, {"start": start_idx, "end": start_idx + text_len}

        elif op_type == 'delete_text':
            start_idx = op.get('start_index', 0)
            end_idx = op.get('end_index', start_idx)
            deleted = end_idx - start_idx
            return -deleted, {"start": start_idx, "end": start_idx}

        elif op_type == 'replace_text':
            start_idx = op.get('start_index', 0)
            end_idx = op.get('end_index', start_idx)
            old_len = end_idx - start_idx
            new_len = len(op.get('text', ''))
            shift = new_len - old_len
            return shift, {"start": start_idx, "end": start_idx + new_len}

        elif op_type in ('format_text', 'find_replace'):
            # These don't change positions (find_replace is applied atomically)
            start_idx = op.get('start_index', 0)
            end_idx = op.get('end_index', start_idx)
            return 0, {"start": start_idx, "end": end_idx}

        elif op_type == 'insert_table':
            # Tables add structure but position shift is complex
            # Conservatively estimate based on rows/columns
            start_idx = op.get('index', 0)
            return 0, {"start": start_idx, "end": start_idx}

        elif op_type == 'insert_page_break':
            start_idx = op.get('index', 0)
            return 1, {"start": start_idx, "end": start_idx + 1}

        return 0, {"start": 0, "end": 0}

    def _describe_operation(self, op: Dict[str, Any]) -> str:
        """Generate human-readable description of an operation."""
        op_type = op.get('type', 'unknown')

        if op_type == 'insert_text':
            text = op.get('text', '')
            preview = text[:20] + '...' if len(text) > 20 else text
            return f"insert '{preview}' at {op.get('index', '?')}"

        elif op_type == 'delete_text':
            return f"delete {op.get('start_index', '?')}-{op.get('end_index', '?')}"

        elif op_type == 'replace_text':
            text = op.get('text', '')
            preview = text[:20] + '...' if len(text) > 20 else text
            return f"replace {op.get('start_index', '?')}-{op.get('end_index', '?')} with '{preview}'"

        elif op_type == 'format_text':
            formats = []
            # Text-level formatting
            for key in ['bold', 'italic', 'underline', 'strikethrough', 'small_caps',
                        'subscript', 'superscript', 'font_size', 'font_family',
                        'link', 'foreground_color', 'background_color']:
                if op.get(key) is not None:
                    formats.append(f"{key}={op[key]}")
            # Paragraph-level formatting
            for key in ['heading_style', 'alignment', 'line_spacing',
                        'indent_first_line', 'indent_start', 'indent_end',
                        'space_above', 'space_below']:
                if op.get(key) is not None:
                    formats.append(f"{key}={op[key]}")
            return f"format {op.get('start_index', '?')}-{op.get('end_index', '?')} ({', '.join(formats) or 'none'})"

        elif op_type == 'insert_table':
            return f"insert {op.get('rows', '?')}x{op.get('columns', '?')} table at {op.get('index', '?')}"

        elif op_type == 'insert_page_break':
            return f"insert page break at {op.get('index', '?')}"

        elif op_type == 'find_replace':
            return f"find '{op.get('find_text', '?')}' → '{op.get('replace_text', '?')}'"

        return f"{op_type} operation"

    def _build_requests_from_resolved(
        self,
        resolved_ops: List[Optional[Dict[str, Any]]],
    ) -> List[Dict[str, Any]]:
        """Build Google Docs API requests from resolved operations."""
        requests = []

        for op in resolved_ops:
            if op is None:
                continue

            op_type = op.get('type', '')

            if op_type == 'insert_text':
                requests.append(create_insert_text_request(op['index'], op['text'], tab_id=self.tab_id))
                # Check if insert operation has formatting parameters - apply formatting after insert
                format_req = create_format_text_request(
                    op['index'], op['index'] + len(op['text']),
                    op.get('bold'), op.get('italic'), op.get('underline'),
                    op.get('strikethrough'), op.get('small_caps'), op.get('subscript'),
                    op.get('superscript'), op.get('font_size'), op.get('font_family'),
                    op.get('link'), op.get('foreground_color'), op.get('background_color'),
                    tab_id=self.tab_id,
                )
                # Always clear formatting to prevent inheriting surrounding styles.
                # This ensures plain text insertions stay plain.
                clear_req = create_clear_formatting_request(
                    op['index'], op['index'] + len(op['text']),
                    preserve_links=(op.get('link') is not None),
                    tab_id=self.tab_id,
                )
                requests.append(clear_req)
                if format_req:
                    requests.append(format_req)

            elif op_type == 'delete_text':
                requests.append(create_delete_range_request(
                    op['start_index'], op['end_index'], tab_id=self.tab_id
                ))

            elif op_type == 'replace_text':
                # Replace = delete + insert
                requests.append(create_delete_range_request(
                    op['start_index'], op['end_index'], tab_id=self.tab_id
                ))
                requests.append(create_insert_text_request(
                    op['start_index'], op['text'], tab_id=self.tab_id
                ))
                # Check if replace operation has formatting parameters - apply formatting after insert
                format_req = create_format_text_request(
                    op['start_index'], op['start_index'] + len(op.get('text', '')),
                    op.get('bold'), op.get('italic'), op.get('underline'),
                    op.get('strikethrough'), op.get('small_caps'), op.get('subscript'),
                    op.get('superscript'), op.get('font_size'), op.get('font_family'),
                    op.get('link'), op.get('foreground_color'), op.get('background_color'),
                    tab_id=self.tab_id,
                )
                # Always clear formatting to prevent inheriting surrounding styles.
                # This ensures plain text replacements stay plain.
                clear_req = create_clear_formatting_request(
                    op['start_index'], op['start_index'] + len(op.get('text', '')),
                    preserve_links=(op.get('link') is not None),
                    tab_id=self.tab_id,
                )
                requests.append(clear_req)
                if format_req:
                    requests.append(format_req)

            elif op_type == 'format_text':
                # Handle text-level formatting (bold, italic, font, etc.)
                req = create_format_text_request(
                    op['start_index'], op['end_index'],
                    op.get('bold'), op.get('italic'), op.get('underline'),
                    op.get('strikethrough'), op.get('small_caps'), op.get('subscript'),
                    op.get('superscript'), op.get('font_size'), op.get('font_family'),
                    op.get('link'), op.get('foreground_color'), op.get('background_color'),
                    tab_id=self.tab_id,
                )
                if req:
                    requests.append(req)

                # Handle paragraph-level formatting (heading_style, alignment, etc.)
                para_req = create_paragraph_style_request(
                    op['start_index'], op['end_index'],
                    line_spacing=op.get('line_spacing'),
                    heading_style=op.get('heading_style'),
                    alignment=op.get('alignment'),
                    indent_first_line=op.get('indent_first_line'),
                    indent_start=op.get('indent_start'),
                    indent_end=op.get('indent_end'),
                    space_above=op.get('space_above'),
                    space_below=op.get('space_below'),
                    tab_id=self.tab_id,
                )
                if para_req:
                    requests.append(para_req)

            elif op_type == 'insert_table':
                # Validate required fields with helpful error messages
                missing = []
                if 'index' not in op:
                    missing.append('index')
                if 'rows' not in op:
                    missing.append('rows')
                if 'columns' not in op:
                    missing.append('columns')
                if missing:
                    raise KeyError(
                        f"insert_table operation missing required field(s): {', '.join(missing)}. "
                        f"Example: {{\"type\": \"insert_table\", \"index\": 1, \"rows\": 3, \"columns\": 4}} "
                        f"or {{\"type\": \"insert_table\", \"location\": \"end\", \"rows\": 3, \"columns\": 4}}"
                    )
                requests.append(create_insert_table_request(
                    op['index'], op['rows'], op['columns'], tab_id=self.tab_id
                ))

            elif op_type == 'insert_page_break':
                requests.append(create_insert_page_break_request(op['index'], tab_id=self.tab_id))

            elif op_type == 'find_replace':
                # Note: find_replace uses tab_ids (list) instead of tab_id (single)
                tab_ids = [self.tab_id] if self.tab_id else None
                requests.append(create_find_replace_request(
                    op['find_text'], op['replace_text'], op.get('match_case', False), tab_ids=tab_ids
                ))

        return requests

    def _record_batch_operations_to_history(
        self,
        document_id: str,
        resolved_ops: List[Dict[str, Any]],
        operation_results: List['BatchOperationResult'],
        doc_data: Dict[str, Any],
    ) -> None:
        """
        Record batch operations to history for undo support.

        Args:
            document_id: ID of the document
            resolved_ops: List of resolved operations
            operation_results: List of BatchOperationResult with operation details
            doc_data: Document data (may be stale after operations, but useful for text context)
        """
        history_manager = get_history_manager()

        # Map operation types to history operation types
        op_type_map = {
            'insert_text': 'insert_text',
            'insert': 'insert_text',
            'delete_text': 'delete_text',
            'delete': 'delete_text',
            'replace_text': 'replace_text',
            'replace': 'replace_text',
            'format_text': 'format_text',
            'format': 'format_text',
        }

        for i, (op, result) in enumerate(zip(resolved_ops, operation_results)):
            if op is None or not result.success:
                continue

            op_type = op.get('type', '')
            normalized_type = normalize_operation_type(op_type)
            history_op_type = op_type_map.get(normalized_type, op_type_map.get(op_type))

            if not history_op_type:
                # Skip operations that don't have a mapped history type (e.g., insert_table, find_replace)
                logger.debug(f"Skipping history recording for operation type: {op_type}")
                continue

            start_index = op.get('index') or op.get('start_index', 0)
            end_index = op.get('end_index')
            text = op.get('text', '')
            position_shift = result.position_shift

            # Capture deleted/original text for undo if available from result
            deleted_text = None
            original_text = None
            if result.affected_range:
                # For delete/replace, we'd need to capture text before operation
                # Since we're recording after, this is best-effort from doc_data (which is pre-operation)
                if history_op_type in ['delete_text', 'replace_text'] and end_index:
                    try:
                        extracted = extract_text_at_range(doc_data, start_index, end_index)
                        captured = extracted.get("text", "")
                        if history_op_type == 'delete_text':
                            deleted_text = captured
                        else:
                            original_text = captured
                    except Exception as e:
                        logger.debug(f"Could not capture text for undo: {e}")

            # Determine undo capability
            undo_capability = UndoCapability.FULL
            undo_notes = None
            if history_op_type == 'format_text':
                undo_capability = UndoCapability.NONE
                undo_notes = "Format undo requires capturing original formatting (not yet supported)"

            try:
                history_manager.record_operation(
                    document_id=document_id,
                    operation_type=history_op_type,
                    operation_params={
                        "start_index": start_index,
                        "end_index": end_index,
                        "text": text,
                        "batch_index": i,
                    },
                    start_index=start_index,
                    end_index=end_index,
                    position_shift=position_shift,
                    deleted_text=deleted_text,
                    original_text=original_text,
                    undo_capability=undo_capability,
                    undo_notes=undo_notes,
                )
                logger.debug(f"Recorded batch operation {i} for undo: {history_op_type} at {start_index}")
            except Exception as e:
                logger.warning(f"Failed to record batch operation {i}: {e}")
