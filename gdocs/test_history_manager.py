"""
Tests for the HistoryManager and related functionality.
"""
import pytest
from datetime import datetime, timezone

from gdocs.managers.history_manager import (
    HistoryManager,
    OperationSnapshot,
    UndoCapability,
    UndoResult,
    DocumentHistory,
    get_history_manager,
    reset_history_manager,
)


class TestOperationSnapshot:
    """Tests for OperationSnapshot dataclass."""

    def test_create_snapshot(self):
        """Test creating an operation snapshot."""
        snapshot = OperationSnapshot(
            id="op_123",
            document_id="doc_abc",
            timestamp=datetime.now(timezone.utc),
            operation_type="insert_text",
            operation_params={"index": 100, "text": "Hello"},
            start_index=100,
            position_shift=5,
        )

        assert snapshot.id == "op_123"
        assert snapshot.document_id == "doc_abc"
        assert snapshot.operation_type == "insert_text"
        assert snapshot.start_index == 100
        assert snapshot.position_shift == 5
        assert snapshot.undo_capability == UndoCapability.FULL
        assert snapshot.undone is False

    def test_snapshot_to_dict(self):
        """Test converting snapshot to dictionary."""
        timestamp = datetime.now(timezone.utc)
        snapshot = OperationSnapshot(
            id="op_123",
            document_id="doc_abc",
            timestamp=timestamp,
            operation_type="delete_text",
            operation_params={"start_index": 10, "end_index": 20},
            start_index=10,
            end_index=20,
            deleted_text="deleted content",
            position_shift=-10,
        )

        result = snapshot.to_dict()

        assert result["id"] == "op_123"
        assert result["operation_type"] == "delete_text"
        assert result["deleted_text"] == "deleted content"
        assert result["timestamp"] == timestamp.isoformat()


class TestDocumentHistory:
    """Tests for DocumentHistory."""

    def test_add_operation(self):
        """Test adding operations to history."""
        history = DocumentHistory(document_id="doc_abc")

        snapshot = OperationSnapshot(
            id="op_1",
            document_id="doc_abc",
            timestamp=datetime.now(timezone.utc),
            operation_type="insert_text",
            operation_params={},
            start_index=0,
        )

        history.add_operation(snapshot)

        assert len(history.operations) == 1
        assert history.operations[0].id == "op_1"

    def test_max_history_size(self):
        """Test that history is trimmed when exceeding max size."""
        history = DocumentHistory(document_id="doc_abc", max_history_size=5)

        # Add 7 operations
        for i in range(7):
            snapshot = OperationSnapshot(
                id=f"op_{i}",
                document_id="doc_abc",
                timestamp=datetime.now(timezone.utc),
                operation_type="insert_text",
                operation_params={},
                start_index=i,
            )
            history.add_operation(snapshot)

        # Should only keep the last 5
        assert len(history.operations) == 5
        assert history.operations[0].id == "op_2"  # First two trimmed
        assert history.operations[-1].id == "op_6"

    def test_get_last_undoable(self):
        """Test getting the last undoable operation."""
        history = DocumentHistory(document_id="doc_abc")

        # Add an undoable operation
        op1 = OperationSnapshot(
            id="op_1",
            document_id="doc_abc",
            timestamp=datetime.now(timezone.utc),
            operation_type="insert_text",
            operation_params={},
            start_index=0,
            undo_capability=UndoCapability.FULL,
        )
        history.add_operation(op1)

        # Add an already-undone operation
        op2 = OperationSnapshot(
            id="op_2",
            document_id="doc_abc",
            timestamp=datetime.now(timezone.utc),
            operation_type="delete_text",
            operation_params={},
            start_index=10,
            undone=True,
        )
        history.add_operation(op2)

        last_undoable = history.get_last_undoable()

        assert last_undoable is not None
        assert last_undoable.id == "op_1"  # op_2 is undone, so op_1 is returned

    def test_get_last_undoable_none_capability(self):
        """Test that operations with NONE capability are skipped."""
        history = DocumentHistory(document_id="doc_abc")

        # Add operation that can't be undone
        op1 = OperationSnapshot(
            id="op_1",
            document_id="doc_abc",
            timestamp=datetime.now(timezone.utc),
            operation_type="find_replace",
            operation_params={},
            start_index=0,
            undo_capability=UndoCapability.NONE,
        )
        history.add_operation(op1)

        last_undoable = history.get_last_undoable()

        assert last_undoable is None


class TestHistoryManager:
    """Tests for HistoryManager."""

    def setup_method(self):
        """Reset the global manager before each test."""
        reset_history_manager()

    def test_record_operation(self):
        """Test recording an operation."""
        manager = HistoryManager()

        snapshot = manager.record_operation(
            document_id="doc_abc",
            operation_type="insert_text",
            operation_params={"index": 100, "text": "Hello"},
            start_index=100,
            position_shift=5,
        )

        assert snapshot is not None
        assert snapshot.operation_type == "insert_text"
        assert snapshot.start_index == 100
        assert snapshot.position_shift == 5

    def test_record_delete_with_text(self):
        """Test recording a delete operation with captured text."""
        manager = HistoryManager()

        snapshot = manager.record_operation(
            document_id="doc_abc",
            operation_type="delete_text",
            operation_params={"start_index": 10, "end_index": 20},
            start_index=10,
            end_index=20,
            position_shift=-10,
            deleted_text="deleted content",
        )

        assert snapshot.deleted_text == "deleted content"
        assert snapshot.undo_capability == UndoCapability.FULL

    def test_generate_undo_insert(self):
        """Test generating undo for insert operation."""
        manager = HistoryManager()

        manager.record_operation(
            document_id="doc_abc",
            operation_type="insert_text",
            operation_params={"index": 100, "text": "Hello World"},
            start_index=100,
            position_shift=11,
        )

        result = manager.generate_undo_operation("doc_abc")

        assert result.success is True
        assert result.reverse_operation is not None
        assert result.reverse_operation["type"] == "delete_text"
        assert result.reverse_operation["start_index"] == 100
        assert result.reverse_operation["end_index"] == 111  # 100 + len("Hello World")

    def test_generate_undo_delete(self):
        """Test generating undo for delete operation."""
        manager = HistoryManager()

        manager.record_operation(
            document_id="doc_abc",
            operation_type="delete_text",
            operation_params={"start_index": 50, "end_index": 60},
            start_index=50,
            end_index=60,
            position_shift=-10,
            deleted_text="removed text",
        )

        result = manager.generate_undo_operation("doc_abc")

        assert result.success is True
        assert result.reverse_operation["type"] == "insert_text"
        assert result.reverse_operation["index"] == 50
        assert result.reverse_operation["text"] == "removed text"

    def test_generate_undo_replace(self):
        """Test generating undo for replace operation."""
        manager = HistoryManager()

        manager.record_operation(
            document_id="doc_abc",
            operation_type="replace_text",
            operation_params={"start_index": 20, "end_index": 30, "text": "new text"},
            start_index=20,
            end_index=30,
            position_shift=-2,  # 8 chars new - 10 chars old
            original_text="old content",
        )

        result = manager.generate_undo_operation("doc_abc")

        assert result.success is True
        assert result.reverse_operation["type"] == "replace_text"
        assert result.reverse_operation["start_index"] == 20
        assert result.reverse_operation["end_index"] == 28  # 20 + len("new text")
        assert result.reverse_operation["text"] == "old content"

    def test_generate_undo_no_history(self):
        """Test generating undo when no history exists."""
        manager = HistoryManager()

        result = manager.generate_undo_operation("nonexistent_doc")

        assert result.success is False
        assert "No history found" in result.message

    def test_generate_undo_missing_deleted_text(self):
        """Test that undo fails for delete without captured text."""
        manager = HistoryManager()

        manager.record_operation(
            document_id="doc_abc",
            operation_type="delete_text",
            operation_params={"start_index": 10, "end_index": 20},
            start_index=10,
            end_index=20,
            position_shift=-10,
            # Note: deleted_text is not provided
        )

        result = manager.generate_undo_operation("doc_abc")

        assert result.success is False
        assert "Missing information" in result.error

    def test_mark_undone(self):
        """Test marking an operation as undone."""
        manager = HistoryManager()

        snapshot = manager.record_operation(
            document_id="doc_abc",
            operation_type="insert_text",
            operation_params={},
            start_index=0,
        )

        success = manager.mark_undone("doc_abc", snapshot.id)

        assert success is True

        # Check that the operation is now undone
        history = manager.get_history("doc_abc")
        assert history[0]["undone"] is True

    def test_get_history(self):
        """Test getting operation history."""
        manager = HistoryManager()

        # Record multiple operations
        for i in range(5):
            manager.record_operation(
                document_id="doc_abc",
                operation_type="insert_text",
                operation_params={},
                start_index=i * 10,
                position_shift=5,
            )

        history = manager.get_history("doc_abc", limit=3)

        assert len(history) == 3
        # Most recent first
        assert history[0]["start_index"] == 40  # Last recorded

    def test_get_history_exclude_undone(self):
        """Test getting history excluding undone operations."""
        manager = HistoryManager()

        # Record and undo first operation
        snapshot1 = manager.record_operation(
            document_id="doc_abc",
            operation_type="insert_text",
            operation_params={},
            start_index=0,
        )
        manager.mark_undone("doc_abc", snapshot1.id)

        # Record second operation
        manager.record_operation(
            document_id="doc_abc",
            operation_type="insert_text",
            operation_params={},
            start_index=10,
        )

        history = manager.get_history("doc_abc", include_undone=False)

        assert len(history) == 1
        assert history[0]["start_index"] == 10

    def test_clear_history(self):
        """Test clearing document history."""
        manager = HistoryManager()

        manager.record_operation(
            document_id="doc_abc",
            operation_type="insert_text",
            operation_params={},
            start_index=0,
        )

        cleared = manager.clear_history("doc_abc")

        assert cleared is True
        assert len(manager.get_history("doc_abc")) == 0

    def test_clear_history_nonexistent(self):
        """Test clearing history for nonexistent document."""
        manager = HistoryManager()

        cleared = manager.clear_history("nonexistent")

        assert cleared is False

    def test_get_stats(self):
        """Test getting manager statistics."""
        manager = HistoryManager()

        # Record operations for multiple documents
        for i in range(3):
            manager.record_operation(
                document_id="doc_1",
                operation_type="insert_text",
                operation_params={},
                start_index=i,
            )

        for i in range(2):
            manager.record_operation(
                document_id="doc_2",
                operation_type="delete_text",
                operation_params={},
                start_index=i,
                deleted_text="text",
            )

        # Undo one operation
        history = manager.get_history("doc_1")
        manager.mark_undone("doc_1", history[0]["id"])

        stats = manager.get_stats()

        assert stats["documents_tracked"] == 2
        assert stats["total_operations"] == 5
        assert stats["undone_operations"] == 1
        assert stats["operations_per_document"]["doc_1"] == 3
        assert stats["operations_per_document"]["doc_2"] == 2


class TestGlobalHistoryManager:
    """Tests for global history manager singleton."""

    def setup_method(self):
        """Reset the global manager before each test."""
        reset_history_manager()

    def test_get_history_manager_singleton(self):
        """Test that get_history_manager returns same instance."""
        manager1 = get_history_manager()
        manager2 = get_history_manager()

        assert manager1 is manager2

    def test_reset_history_manager(self):
        """Test resetting the global manager."""
        manager1 = get_history_manager()
        manager1.record_operation(
            document_id="doc_abc",
            operation_type="insert_text",
            operation_params={},
            start_index=0,
        )

        reset_history_manager()
        manager2 = get_history_manager()

        assert manager1 is not manager2
        assert len(manager2.get_history("doc_abc")) == 0


class TestUndoEdgeCases:
    """Tests for edge cases in undo functionality."""

    def setup_method(self):
        """Reset the global manager before each test."""
        reset_history_manager()

    def test_undo_page_break(self):
        """Test undo for page break insertion."""
        manager = HistoryManager()

        manager.record_operation(
            document_id="doc_abc",
            operation_type="insert_page_break",
            operation_params={"index": 50},
            start_index=50,
            position_shift=1,
        )

        result = manager.generate_undo_operation("doc_abc")

        assert result.success is True
        assert result.reverse_operation["type"] == "delete_text"
        assert result.reverse_operation["start_index"] == 50
        assert result.reverse_operation["end_index"] == 51

    def test_undo_format_without_original(self):
        """Test that format undo fails without original formatting."""
        manager = HistoryManager()

        manager.record_operation(
            document_id="doc_abc",
            operation_type="format_text",
            operation_params={"bold": True},
            start_index=10,
            end_index=20,
            position_shift=0,
            # Note: original_formatting not provided
        )

        result = manager.generate_undo_operation("doc_abc")

        assert result.success is False

    def test_undo_format_with_original(self):
        """Test format undo with original formatting captured."""
        manager = HistoryManager()

        manager.record_operation(
            document_id="doc_abc",
            operation_type="format_text",
            operation_params={"bold": True},
            start_index=10,
            end_index=20,
            position_shift=0,
            original_formatting={"bold": False, "italic": True},
        )

        result = manager.generate_undo_operation("doc_abc")

        assert result.success is True
        assert result.reverse_operation["type"] == "format_text"
        assert result.reverse_operation["bold"] is False
        assert result.reverse_operation["italic"] is True

    def test_multiple_undo_in_sequence(self):
        """Test undoing multiple operations in sequence."""
        manager = HistoryManager()

        # Record 3 operations
        snapshot1 = manager.record_operation(
            document_id="doc_abc",
            operation_type="insert_text",
            operation_params={"text": "First"},
            start_index=0,
            position_shift=5,
        )

        snapshot2 = manager.record_operation(
            document_id="doc_abc",
            operation_type="insert_text",
            operation_params={"text": "Second"},
            start_index=5,
            position_shift=6,
        )

        snapshot3 = manager.record_operation(
            document_id="doc_abc",
            operation_type="insert_text",
            operation_params={"text": "Third"},
            start_index=11,
            position_shift=5,
        )

        # Undo third operation
        result1 = manager.generate_undo_operation("doc_abc")
        assert result1.success is True
        assert result1.operation_id == snapshot3.id
        manager.mark_undone("doc_abc", snapshot3.id)

        # Undo second operation
        result2 = manager.generate_undo_operation("doc_abc")
        assert result2.success is True
        assert result2.operation_id == snapshot2.id
        manager.mark_undone("doc_abc", snapshot2.id)

        # Undo first operation
        result3 = manager.generate_undo_operation("doc_abc")
        assert result3.success is True
        assert result3.operation_id == snapshot1.id
        manager.mark_undone("doc_abc", snapshot1.id)

        # No more operations to undo
        result4 = manager.generate_undo_operation("doc_abc")
        assert result4.success is False


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
