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
    create_find_replace_request,
    create_insert_table_request,
    create_insert_page_break_request,
    validate_operation,
    find_text_in_document,
    calculate_search_based_indices,
    OperationType,
    calculate_position_shift,
    resolve_range,
    RangeResult,
)

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
    Normalize an operation dictionary, converting operation type aliases.

    Args:
        operation: Operation dictionary with 'type' field

    Returns:
        Copy of operation with normalized type
    """
    if 'type' not in operation:
        return operation

    normalized = operation.copy()
    normalized['type'] = normalize_operation_type(operation['type'])
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

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON response."""
        return {
            "success": self.success,
            "operations_completed": self.operations_completed,
            "total_operations": self.total_operations,
            "results": [r.to_dict() for r in self.results],
            "total_position_shift": self.total_position_shift,
            "message": self.message,
            "document_link": self.document_link,
        }


class BatchOperationManager:
    """
    High-level manager for Google Docs batch operations.
    
    Handles complex multi-operation requests including:
    - Operation validation and request building
    - Batch execution with proper error handling
    - Operation result processing and reporting
    """
    
    def __init__(self, service):
        """
        Initialize the batch operation manager.
        
        Args:
            service: Google Docs API service instance
        """
        self.service = service
        
    async def execute_batch_operations(
        self,
        document_id: str,
        operations: list[dict[str, Any]]
    ) -> tuple[bool, str, dict[str, Any]]:
        """
        Execute multiple document operations in a single atomic batch.

        This method extracts the complex logic from batch_update_doc tool function.

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
            request = create_insert_text_request(op['index'], op['text'])
            description = f"insert text at {op['index']}"
            
        elif op_type == 'delete_text':
            request = create_delete_range_request(op['start_index'], op['end_index'])
            description = f"delete text {op['start_index']}-{op['end_index']}"
            
        elif op_type == 'replace_text':
            # Replace is delete + insert (must be done in this order)
            delete_request = create_delete_range_request(op['start_index'], op['end_index'])
            insert_request = create_insert_text_request(op['start_index'], op['text'])
            # Return both requests as a list
            request = [delete_request, insert_request]
            description = f"replace text {op['start_index']}-{op['end_index']} with '{op['text'][:20]}{'...' if len(op['text']) > 20 else ''}'"
            
        elif op_type == 'format_text':
            request = create_format_text_request(
                op['start_index'], op['end_index'],
                op.get('bold'), op.get('italic'), op.get('underline'),
                op.get('font_size'), op.get('font_family')
            )
            
            if not request:
                raise ValueError("No formatting options provided")
                
            # Build format description
            format_changes = []
            for param, name in [
                ('bold', 'bold'), ('italic', 'italic'), ('underline', 'underline'),
                ('font_size', 'font size'), ('font_family', 'font family')
            ]:
                if op.get(param) is not None:
                    value = f"{op[param]}pt" if param == 'font_size' else op[param]
                    format_changes.append(f"{name}: {value}")
                    
            description = f"format text {op['start_index']}-{op['end_index']} ({', '.join(format_changes)})"
            
        elif op_type == 'insert_table':
            request = create_insert_table_request(op['index'], op['rows'], op['columns'])
            description = f"insert {op['rows']}x{op['columns']} table at {op['index']}"
            
        elif op_type == 'insert_page_break':
            request = create_insert_page_break_request(op['index'])
            description = f"insert page break at {op['index']}"
            
        elif op_type == 'find_replace':
            request = create_find_replace_request(
                op['find_text'], op['replace_text'], op.get('match_case', False)
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
    ) -> BatchExecutionResult:
        """
        Execute batch operations with search-based positioning and position tracking.

        This enhanced method supports:
        - Search-based positioning (e.g., insert before/after "Conclusion")
        - Automatic position adjustment for sequential operations
        - Detailed per-operation results with position shifts

        Args:
            document_id: ID of the document to update
            operations: List of operation dictionaries. Each can include:
                - type: Operation type (insert, delete, replace, format)
                - Standard index-based params (index, start_index, end_index)
                - OR search-based params (search, position="before"|"after"|"replace")
                - text: Text for insert/replace operations
                - Formatting params for format operations
            auto_adjust_positions: If True, automatically adjust positions for
                subsequent operations based on cumulative position shifts

        Returns:
            BatchExecutionResult with per-operation details

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
        """
        logger.info(f"Executing enhanced batch on document {document_id}")
        logger.info(f"Operations: {len(operations)}, auto_adjust: {auto_adjust_positions}")

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

        # Resolve all operations and build requests
        resolved_ops, operation_results = await self._resolve_operations_with_search(
            operations, doc_data, auto_adjust_positions
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

            return BatchExecutionResult(
                success=True,
                operations_completed=len(operations),
                total_operations=len(operations),
                results=operation_results,
                total_position_shift=total_shift,
                message=f"Successfully executed {len(operations)} operation(s)",
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
        3. Range spec: {"type": "delete", "range_spec": {"search": "x", "extend": "paragraph"}, ...}

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

        for i, op in enumerate(operations):
            op_type = op.get('type', '')
            op_copy = op.copy()

            # Check if this operation uses range_spec (new range-based positioning)
            if 'range_spec' in op:
                range_result = self._resolve_range_spec(op, doc_data, cumulative_shift, auto_adjust)

                if not range_result.success:
                    results.append(BatchOperationResult(
                        index=i,
                        type=op_type,
                        success=False,
                        description=f"Range resolution failed",
                        error=range_result.message,
                    ))
                    resolved_ops.append(None)
                    continue

                # Convert to index-based operation
                op_copy = self._convert_range_to_index_op(
                    op, op_type, range_result.start_index, range_result.end_index
                )
                resolved_index = range_result.start_index

            # Check if this is a search-based operation (legacy mode)
            elif 'search' in op:
                search_text = op['search']
                position = op.get('position', 'replace')
                occurrence = op.get('occurrence', 1)
                match_case = op.get('match_case', True)

                success, start_idx, end_idx, msg = calculate_search_based_indices(
                    doc_data, search_text, position, occurrence, match_case
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

                # Apply cumulative shift if auto-adjusting
                if auto_adjust and cumulative_shift != 0:
                    start_idx += cumulative_shift
                    if end_idx is not None:
                        end_idx += cumulative_shift

                # Map to standard index-based operation
                op_copy = self._convert_search_to_index_op(
                    op, op_type, start_idx, end_idx
                )
                resolved_index = start_idx
            else:
                # Index-based operation - apply cumulative shift
                resolved_index = None
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
            for key in ['bold', 'italic', 'underline', 'font_size', 'font_family']:
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
                requests.append(create_insert_text_request(op['index'], op['text']))

            elif op_type == 'delete_text':
                requests.append(create_delete_range_request(
                    op['start_index'], op['end_index']
                ))

            elif op_type == 'replace_text':
                # Replace = delete + insert
                requests.append(create_delete_range_request(
                    op['start_index'], op['end_index']
                ))
                requests.append(create_insert_text_request(
                    op['start_index'], op['text']
                ))

            elif op_type == 'format_text':
                req = create_format_text_request(
                    op['start_index'], op['end_index'],
                    op.get('bold'), op.get('italic'), op.get('underline'),
                    op.get('font_size'), op.get('font_family')
                )
                if req:
                    requests.append(req)

            elif op_type == 'insert_table':
                requests.append(create_insert_table_request(
                    op['index'], op['rows'], op['columns']
                ))

            elif op_type == 'insert_page_break':
                requests.append(create_insert_page_break_request(op['index']))

            elif op_type == 'find_replace':
                requests.append(create_find_replace_request(
                    op['find_text'], op['replace_text'], op.get('match_case', False)
                ))

        return requests