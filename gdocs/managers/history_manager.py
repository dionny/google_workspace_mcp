"""
Operation History Manager

This module provides operation history tracking and undo capability for Google Docs.

Since the Google Docs API does not support programmatic undo or revision restoration,
this manager tracks operations and generates reverse operations to enable undo.

Design Notes:
- History is tracked per document
- Operations are stored with enough information to generate reverse operations
- Undo works by executing the reverse operation (not by restoring a revision)
- History is stored in-memory (per-process, not persisted)

Limitations:
- History is lost when the MCP server restarts
- Undo may fail if the document was modified externally between operation and undo
- Complex operations (like find_replace) have limited undo support
- Format undo requires storing the original formatting (complex to implement fully)
"""
import logging
import asyncio
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple
from dataclasses import dataclass, field, asdict
from enum import Enum
from collections import defaultdict

logger = logging.getLogger(__name__)


class UndoCapability(str, Enum):
    """Indicates how well an operation can be undone."""
    FULL = "full"  # Can be fully reversed
    PARTIAL = "partial"  # Can be partially reversed (e.g., text but not formatting)
    NONE = "none"  # Cannot be undone (e.g., find_replace with unknown matches)


@dataclass
class OperationSnapshot:
    """
    Snapshot of an operation with information needed for undo.

    Stores both the operation performed and the data needed to reverse it.
    """
    id: str  # Unique operation ID
    document_id: str
    timestamp: datetime
    operation_type: str  # insert_text, delete_text, replace_text, format_text, etc.

    # Operation parameters (what was done)
    operation_params: Dict[str, Any]

    # Data needed for undo
    deleted_text: Optional[str] = None  # Text that was deleted (for undo)
    original_text: Optional[str] = None  # Original text before replace (for undo)
    original_formatting: Optional[Dict[str, Any]] = None  # Original formatting (for undo)

    # Position tracking
    start_index: int = 0
    end_index: Optional[int] = None
    position_shift: int = 0

    # Undo metadata
    undo_capability: UndoCapability = UndoCapability.FULL
    undo_notes: Optional[str] = None
    undone: bool = False
    undone_at: Optional[datetime] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        result = asdict(self)
        result['timestamp'] = self.timestamp.isoformat()
        if self.undone_at:
            result['undone_at'] = self.undone_at.isoformat()
        return result


@dataclass
class UndoResult:
    """Result of an undo operation."""
    success: bool
    message: str
    operation_id: Optional[str] = None
    reverse_operation: Optional[Dict[str, Any]] = None
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary, excluding None values."""
        result = asdict(self)
        return {k: v for k, v in result.items() if v is not None}


@dataclass
class DocumentHistory:
    """History of operations for a single document."""
    document_id: str
    operations: List[OperationSnapshot] = field(default_factory=list)
    max_history_size: int = 50  # Maximum operations to keep per document

    def add_operation(self, operation: OperationSnapshot) -> None:
        """Add an operation to history, trimming if necessary."""
        self.operations.append(operation)
        # Trim old operations if we exceed max size
        if len(self.operations) > self.max_history_size:
            # Keep the most recent operations
            self.operations = self.operations[-self.max_history_size:]

    def get_last_undoable(self) -> Optional[OperationSnapshot]:
        """Get the last operation that can be undone."""
        for op in reversed(self.operations):
            if not op.undone and op.undo_capability != UndoCapability.NONE:
                return op
        return None

    def get_operations(self, limit: int = 10) -> List[OperationSnapshot]:
        """Get recent operations, most recent first."""
        return list(reversed(self.operations[-limit:]))

    def clear(self) -> None:
        """Clear all history for this document."""
        self.operations = []


class HistoryManager:
    """
    Manages operation history and undo capability for Google Docs.

    This class maintains an in-memory history of operations performed through
    the MCP tools, enabling undo functionality by generating and executing
    reverse operations.

    Usage:
        manager = HistoryManager()

        # Before performing an operation, capture the context
        text_to_delete = extract_text_from_doc(doc, start, end)

        # After operation, record it
        manager.record_operation(
            document_id=doc_id,
            operation_type="delete_text",
            operation_params={"start_index": 10, "end_index": 20},
            deleted_text=text_to_delete,
            start_index=10,
            end_index=20,
            position_shift=-10
        )

        # To undo
        undo_op = manager.generate_undo_operation(doc_id)
        # Execute undo_op with the Docs API
        manager.mark_undone(doc_id, undo_op.operation_id)
    """

    def __init__(self, max_history_per_doc: int = 50):
        """
        Initialize the history manager.

        Args:
            max_history_per_doc: Maximum operations to track per document
        """
        self._history: Dict[str, DocumentHistory] = {}
        self._max_history_per_doc = max_history_per_doc
        self._operation_counter = 0

    def _generate_operation_id(self) -> str:
        """Generate a unique operation ID."""
        self._operation_counter += 1
        timestamp = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")
        return f"op_{timestamp}_{self._operation_counter}"

    def _get_or_create_history(self, document_id: str) -> DocumentHistory:
        """Get or create history for a document."""
        if document_id not in self._history:
            self._history[document_id] = DocumentHistory(
                document_id=document_id,
                max_history_size=self._max_history_per_doc
            )
        return self._history[document_id]

    def record_operation(
        self,
        document_id: str,
        operation_type: str,
        operation_params: Dict[str, Any],
        start_index: int,
        end_index: Optional[int] = None,
        position_shift: int = 0,
        deleted_text: Optional[str] = None,
        original_text: Optional[str] = None,
        original_formatting: Optional[Dict[str, Any]] = None,
        undo_capability: UndoCapability = UndoCapability.FULL,
        undo_notes: Optional[str] = None,
    ) -> OperationSnapshot:
        """
        Record an operation in history.

        Args:
            document_id: ID of the document
            operation_type: Type of operation (insert_text, delete_text, etc.)
            operation_params: Parameters passed to the operation
            start_index: Start index of the operation
            end_index: End index (for range operations)
            position_shift: How much positions shifted
            deleted_text: Text that was deleted (for undo of delete/replace)
            original_text: Original text before replace (for undo)
            original_formatting: Original formatting before format change
            undo_capability: How well this operation can be undone
            undo_notes: Notes about undo limitations

        Returns:
            The recorded OperationSnapshot
        """
        history = self._get_or_create_history(document_id)

        snapshot = OperationSnapshot(
            id=self._generate_operation_id(),
            document_id=document_id,
            timestamp=datetime.now(timezone.utc),
            operation_type=operation_type,
            operation_params=operation_params,
            start_index=start_index,
            end_index=end_index,
            position_shift=position_shift,
            deleted_text=deleted_text,
            original_text=original_text,
            original_formatting=original_formatting,
            undo_capability=undo_capability,
            undo_notes=undo_notes,
        )

        history.add_operation(snapshot)
        logger.info(f"Recorded operation {snapshot.id}: {operation_type} on {document_id}")

        return snapshot

    def generate_undo_operation(
        self,
        document_id: str,
    ) -> UndoResult:
        """
        Generate the reverse operation to undo the last operation.

        Args:
            document_id: ID of the document

        Returns:
            UndoResult with the reverse operation to execute
        """
        history = self._history.get(document_id)
        if not history:
            return UndoResult(
                success=False,
                message="No history found for this document",
                error="No operations have been tracked for this document"
            )

        operation = history.get_last_undoable()
        if not operation:
            return UndoResult(
                success=False,
                message="No undoable operations found",
                error="All operations have been undone or cannot be undone"
            )

        if operation.undo_capability == UndoCapability.NONE:
            return UndoResult(
                success=False,
                message=f"Operation {operation.id} cannot be undone",
                operation_id=operation.id,
                error=operation.undo_notes or "This operation type does not support undo"
            )

        # Generate the reverse operation based on type
        reverse_op = self._generate_reverse_operation(operation)

        if reverse_op is None:
            return UndoResult(
                success=False,
                message=f"Could not generate undo for operation {operation.id}",
                operation_id=operation.id,
                error="Missing information required for undo"
            )

        notes = []
        if operation.undo_capability == UndoCapability.PARTIAL:
            notes.append(operation.undo_notes or "Partial undo - some aspects may not be reversed")

        return UndoResult(
            success=True,
            message=f"Generated undo for {operation.operation_type}",
            operation_id=operation.id,
            reverse_operation=reverse_op,
        )

    def _generate_reverse_operation(
        self,
        operation: OperationSnapshot,
    ) -> Optional[Dict[str, Any]]:
        """
        Generate the reverse operation for a given operation.

        Args:
            operation: The operation to reverse

        Returns:
            Dictionary representing the reverse operation, or None if not possible
        """
        op_type = operation.operation_type

        if op_type == "insert_text":
            # Reverse of insert is delete
            # We need to delete the inserted text
            text_len = len(operation.operation_params.get("text", ""))
            start = operation.start_index
            return {
                "type": "delete_text",
                "start_index": start,
                "end_index": start + text_len,
                "description": f"Undo insert: delete {text_len} chars at {start}"
            }

        elif op_type == "delete_text":
            # Reverse of delete is insert (if we have the deleted text)
            if operation.deleted_text is None:
                return None
            return {
                "type": "insert_text",
                "index": operation.start_index,
                "text": operation.deleted_text,
                "description": f"Undo delete: re-insert {len(operation.deleted_text)} chars at {operation.start_index}"
            }

        elif op_type == "replace_text":
            # Reverse of replace is replace back to original
            if operation.original_text is None:
                return None
            new_text_len = len(operation.operation_params.get("text", ""))
            return {
                "type": "replace_text",
                "start_index": operation.start_index,
                "end_index": operation.start_index + new_text_len,
                "text": operation.original_text,
                "description": f"Undo replace: restore original text at {operation.start_index}"
            }

        elif op_type == "format_text":
            # Reverse of format requires original formatting
            # This is complex as we need to know what the original formatting was
            if operation.original_formatting is None:
                # Can't fully undo without original formatting
                return None
            return {
                "type": "format_text",
                "start_index": operation.start_index,
                "end_index": operation.end_index,
                **operation.original_formatting,
                "description": f"Undo format: restore original formatting at {operation.start_index}-{operation.end_index}"
            }

        elif op_type == "insert_table":
            # Reverse of insert table is complex - would need to delete the table
            # Tables add multiple elements, making this difficult
            return None

        elif op_type == "insert_page_break":
            # Reverse of page break is delete one character
            return {
                "type": "delete_text",
                "start_index": operation.start_index,
                "end_index": operation.start_index + 1,
                "description": f"Undo page break: delete at {operation.start_index}"
            }

        elif op_type == "find_replace":
            # Find/replace is very difficult to undo as we'd need to track
            # every match and its original text
            return None

        # Unknown operation type
        logger.warning(f"Unknown operation type for undo: {op_type}")
        return None

    def mark_undone(
        self,
        document_id: str,
        operation_id: str,
    ) -> bool:
        """
        Mark an operation as undone.

        Args:
            document_id: ID of the document
            operation_id: ID of the operation to mark as undone

        Returns:
            True if successful, False otherwise
        """
        history = self._history.get(document_id)
        if not history:
            return False

        for op in history.operations:
            if op.id == operation_id:
                op.undone = True
                op.undone_at = datetime.now(timezone.utc)
                logger.info(f"Marked operation {operation_id} as undone")
                return True

        return False

    def get_history(
        self,
        document_id: str,
        limit: int = 10,
        include_undone: bool = True,
    ) -> List[Dict[str, Any]]:
        """
        Get operation history for a document.

        Args:
            document_id: ID of the document
            limit: Maximum number of operations to return
            include_undone: Whether to include undone operations

        Returns:
            List of operation dictionaries, most recent first
        """
        history = self._history.get(document_id)
        if not history:
            return []

        operations = history.get_operations(limit)

        if not include_undone:
            operations = [op for op in operations if not op.undone]

        return [op.to_dict() for op in operations]

    def clear_history(self, document_id: str) -> bool:
        """
        Clear all history for a document.

        Args:
            document_id: ID of the document

        Returns:
            True if history was cleared, False if no history existed
        """
        if document_id in self._history:
            self._history[document_id].clear()
            logger.info(f"Cleared history for document {document_id}")
            return True
        return False

    def get_stats(self) -> Dict[str, Any]:
        """Get statistics about tracked history."""
        total_ops = sum(len(h.operations) for h in self._history.values())
        undone_ops = sum(
            sum(1 for op in h.operations if op.undone)
            for h in self._history.values()
        )

        return {
            "documents_tracked": len(self._history),
            "total_operations": total_ops,
            "undone_operations": undone_ops,
            "operations_per_document": {
                doc_id: len(h.operations)
                for doc_id, h in self._history.items()
            }
        }


# Global instance for use across the MCP server
_history_manager: Optional[HistoryManager] = None


def get_history_manager() -> HistoryManager:
    """Get the global HistoryManager instance."""
    global _history_manager
    if _history_manager is None:
        _history_manager = HistoryManager()
    return _history_manager


def reset_history_manager() -> None:
    """Reset the global HistoryManager instance (for testing)."""
    global _history_manager
    _history_manager = None
