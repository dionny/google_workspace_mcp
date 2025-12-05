"""
Unit tests for batch_modify_doc with search-based positioning and position tracking.

These tests verify:
- Search-based positioning (before/after/replace)
- Position auto-adjustment for sequential operations
- Per-operation result tracking
- Request building from resolved operations
"""
import pytest
from unittest.mock import MagicMock, AsyncMock, patch
from gdocs.managers.batch_operation_manager import (
    BatchOperationManager,
    BatchOperationResult,
    BatchExecutionResult,
)


class TestBatchOperationResult:
    """Tests for BatchOperationResult dataclass."""

    def test_to_dict_excludes_none(self):
        """Test that to_dict excludes None values."""
        result = BatchOperationResult(
            index=0,
            type="insert_text",
            success=True,
            description="insert 'Hello' at 100",
            position_shift=5,
            affected_range={"start": 100, "end": 105},
            resolved_index=None,  # Should be excluded
            error=None,  # Should be excluded
        )
        result_dict = result.to_dict()

        assert "resolved_index" not in result_dict
        assert "error" not in result_dict
        assert result_dict["index"] == 0
        assert result_dict["position_shift"] == 5

    def test_to_dict_includes_all_present_values(self):
        """Test that to_dict includes all non-None values."""
        result = BatchOperationResult(
            index=1,
            type="format_text",
            success=False,
            description="format failed",
            position_shift=0,
            affected_range={"start": 50, "end": 60},
            resolved_index=50,
            error="Search text not found",
        )
        result_dict = result.to_dict()

        assert result_dict["index"] == 1
        assert result_dict["resolved_index"] == 50
        assert result_dict["error"] == "Search text not found"


class TestBatchExecutionResult:
    """Tests for BatchExecutionResult dataclass."""

    def test_to_dict_formats_correctly(self):
        """Test that to_dict produces the expected format."""
        op_result = BatchOperationResult(
            index=0,
            type="insert_text",
            success=True,
            description="insert at 100",
            position_shift=10,
        )
        result = BatchExecutionResult(
            success=True,
            operations_completed=1,
            total_operations=1,
            results=[op_result],
            total_position_shift=10,
            message="Success",
            document_link="https://docs.google.com/document/d/123/edit",
        )
        result_dict = result.to_dict()

        assert result_dict["success"] is True
        assert result_dict["operations_completed"] == 1
        assert len(result_dict["results"]) == 1
        assert result_dict["total_position_shift"] == 10
        assert "document_link" in result_dict


class TestBatchOperationManagerCalculations:
    """Tests for position shift calculations."""

    def setup_method(self):
        """Set up test fixtures."""
        self.manager = BatchOperationManager(MagicMock())

    def test_calculate_insert_shift(self):
        """Test position shift calculation for insert operations."""
        op = {"type": "insert_text", "index": 100, "text": "Hello World"}
        shift, affected = self.manager._calculate_op_shift(op)

        assert shift == 11  # len("Hello World")
        assert affected["start"] == 100
        assert affected["end"] == 111

    def test_calculate_delete_shift(self):
        """Test position shift calculation for delete operations."""
        op = {"type": "delete_text", "start_index": 50, "end_index": 60}
        shift, affected = self.manager._calculate_op_shift(op)

        assert shift == -10  # Deleted 10 characters
        assert affected["start"] == 50
        assert affected["end"] == 50

    def test_calculate_replace_shift_grows(self):
        """Test position shift when replacement is longer."""
        op = {
            "type": "replace_text",
            "start_index": 100,
            "end_index": 105,  # 5 chars
            "text": "Hello World",  # 11 chars
        }
        shift, affected = self.manager._calculate_op_shift(op)

        assert shift == 6  # 11 - 5 = +6
        assert affected["start"] == 100
        assert affected["end"] == 111

    def test_calculate_replace_shift_shrinks(self):
        """Test position shift when replacement is shorter."""
        op = {
            "type": "replace_text",
            "start_index": 100,
            "end_index": 120,  # 20 chars
            "text": "Hi",  # 2 chars
        }
        shift, affected = self.manager._calculate_op_shift(op)

        assert shift == -18  # 2 - 20 = -18
        assert affected["start"] == 100
        assert affected["end"] == 102

    def test_calculate_format_no_shift(self):
        """Test that format operations don't shift positions."""
        op = {
            "type": "format_text",
            "start_index": 50,
            "end_index": 100,
            "bold": True,
        }
        shift, affected = self.manager._calculate_op_shift(op)

        assert shift == 0
        assert affected["start"] == 50
        assert affected["end"] == 100

    def test_calculate_page_break_shift(self):
        """Test that page break adds 1 character."""
        op = {"type": "insert_page_break", "index": 75}
        shift, affected = self.manager._calculate_op_shift(op)

        assert shift == 1


class TestBatchOperationManagerPositionAdjustment:
    """Tests for position auto-adjustment."""

    def setup_method(self):
        """Set up test fixtures."""
        self.manager = BatchOperationManager(MagicMock())

    def test_apply_position_shift_to_index(self):
        """Test applying shift to index-based operation."""
        op = {"type": "insert_text", "index": 100, "text": "test"}
        shifted = self.manager._apply_position_shift(op, 50)

        assert shifted["index"] == 150
        assert op["index"] == 100  # Original unchanged

    def test_apply_position_shift_to_range(self):
        """Test applying shift to range-based operation."""
        op = {
            "type": "delete_text",
            "start_index": 100,
            "end_index": 150,
        }
        shifted = self.manager._apply_position_shift(op, -25)

        assert shifted["start_index"] == 75
        assert shifted["end_index"] == 125

    def test_apply_negative_shift(self):
        """Test applying negative shift (from deletion)."""
        op = {"type": "insert_text", "index": 200, "text": "after delete"}
        shifted = self.manager._apply_position_shift(op, -30)

        assert shifted["index"] == 170


class TestBatchOperationManagerSearchConversion:
    """Tests for search-to-index conversion."""

    def setup_method(self):
        """Set up test fixtures."""
        self.manager = BatchOperationManager(MagicMock())

    def test_convert_insert_search_to_index(self):
        """Test converting search-based insert to index-based."""
        op = {
            "type": "insert",
            "search": "Conclusion",
            "position": "before",
            "text": "New text",
        }
        result = self.manager._convert_search_to_index_op(op, "insert", 500, 510)

        assert result["type"] == "insert_text"
        assert result["index"] == 500
        assert result["text"] == "New text"
        assert "search" not in result
        assert "position" not in result

    def test_convert_format_search_to_index(self):
        """Test converting search-based format to index-based."""
        op = {
            "type": "format",
            "search": "Important",
            "position": "replace",
            "bold": True,
            "font_size": 14,
        }
        result = self.manager._convert_search_to_index_op(op, "format", 200, 209)

        assert result["type"] == "format_text"
        assert result["start_index"] == 200
        assert result["end_index"] == 209
        assert result["bold"] is True
        assert result["font_size"] == 14

    def test_convert_delete_search_to_index(self):
        """Test converting search-based delete to index-based."""
        op = {
            "type": "delete",
            "search": "old text",
            "position": "replace",
        }
        result = self.manager._convert_search_to_index_op(op, "delete", 100, 108)

        assert result["type"] == "delete_text"
        assert result["start_index"] == 100
        assert result["end_index"] == 108


class TestBatchOperationManagerDescriptions:
    """Tests for operation description generation."""

    def setup_method(self):
        """Set up test fixtures."""
        self.manager = BatchOperationManager(MagicMock())

    def test_describe_insert(self):
        """Test description for insert operation."""
        op = {"type": "insert_text", "index": 100, "text": "Hello"}
        desc = self.manager._describe_operation(op)

        assert "insert" in desc.lower()
        assert "Hello" in desc
        assert "100" in desc

    def test_describe_insert_truncates_long_text(self):
        """Test that long text is truncated in description."""
        op = {
            "type": "insert_text",
            "index": 50,
            "text": "This is a very long piece of text that should be truncated",
        }
        desc = self.manager._describe_operation(op)

        assert "..." in desc
        assert len(desc) < 100

    def test_describe_format(self):
        """Test description for format operation."""
        op = {
            "type": "format_text",
            "start_index": 10,
            "end_index": 20,
            "bold": True,
            "italic": True,
        }
        desc = self.manager._describe_operation(op)

        assert "format" in desc.lower()
        assert "bold" in desc.lower()
        assert "italic" in desc.lower()

    def test_describe_find_replace(self):
        """Test description for find/replace operation."""
        op = {
            "type": "find_replace",
            "find_text": "old",
            "replace_text": "new",
        }
        desc = self.manager._describe_operation(op)

        assert "old" in desc
        assert "new" in desc


class TestBatchOperationManagerRequestBuilding:
    """Tests for building API requests from resolved operations."""

    def setup_method(self):
        """Set up test fixtures."""
        self.manager = BatchOperationManager(MagicMock())

    def test_build_insert_request(self):
        """Test building insert text request."""
        ops = [{"type": "insert_text", "index": 100, "text": "Hello"}]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 1
        assert "insertText" in requests[0]
        assert requests[0]["insertText"]["location"]["index"] == 100

    def test_build_delete_request(self):
        """Test building delete request."""
        ops = [{"type": "delete_text", "start_index": 50, "end_index": 100}]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 1
        assert "deleteContentRange" in requests[0]

    def test_build_replace_request_creates_two(self):
        """Test that replace creates delete + insert requests."""
        ops = [
            {
                "type": "replace_text",
                "start_index": 100,
                "end_index": 110,
                "text": "new text",
            }
        ]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 2
        assert "deleteContentRange" in requests[0]
        assert "insertText" in requests[1]

    def test_build_format_request(self):
        """Test building format request."""
        ops = [
            {
                "type": "format_text",
                "start_index": 0,
                "end_index": 50,
                "bold": True,
            }
        ]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 1
        assert "updateTextStyle" in requests[0]

    def test_build_skips_none_operations(self):
        """Test that None operations are skipped."""
        ops = [
            {"type": "insert_text", "index": 100, "text": "Hello"},
            None,
            {"type": "insert_text", "index": 200, "text": "World"},
        ]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 2

    def test_build_table_request(self):
        """Test building table insert request."""
        ops = [{"type": "insert_table", "index": 50, "rows": 3, "columns": 4}]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 1
        assert "insertTable" in requests[0]
        assert requests[0]["insertTable"]["rows"] == 3
        assert requests[0]["insertTable"]["columns"] == 4


class TestBatchOperationManagerIntegration:
    """Integration tests for batch operations with search."""

    @pytest.fixture
    def mock_service(self):
        """Create a mock Google Docs service."""
        service = MagicMock()
        service.documents.return_value.get.return_value.execute = MagicMock(
            return_value={
                "body": {
                    "content": [
                        {
                            "paragraph": {
                                "elements": [
                                    {
                                        "textRun": {
                                            "content": "Hello World. This is a test document. Conclusion section here.",
                                        },
                                        "startIndex": 1,
                                        "endIndex": 62,
                                    }
                                ]
                            }
                        }
                    ]
                }
            }
        )
        service.documents.return_value.batchUpdate.return_value.execute = MagicMock(
            return_value={"replies": [{}]}
        )
        return service

    @pytest.mark.asyncio
    async def test_execute_batch_with_index_operations(self, mock_service):
        """Test batch execution with index-based operations."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert_text", "index": 1, "text": "PREFIX: "},
            {"type": "format_text", "start_index": 1, "end_index": 9, "bold": True},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 2
        assert len(result.results) == 2
        # First operation inserted 8 chars, second formatted (no shift)
        assert result.results[0].position_shift == 8
        assert result.results[1].position_shift == 0

    @pytest.mark.asyncio
    async def test_execute_batch_tracks_cumulative_shift(self, mock_service):
        """Test that auto_adjust_positions correctly applies cumulative shifts."""
        manager = BatchOperationManager(mock_service)

        # Two inserts at the same original position
        # With auto-adjust, the second should shift by first insert's length
        operations = [
            {"type": "insert_text", "index": 100, "text": "AAAA"},  # +4
            {"type": "insert_text", "index": 100, "text": "BBBB"},  # Should become 104
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        # Total shift is 8 (4 + 4)
        assert result.total_position_shift == 8

    @pytest.mark.asyncio
    async def test_execute_batch_empty_operations(self, mock_service):
        """Test batch execution with empty operations list."""
        manager = BatchOperationManager(mock_service)

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            [],
            auto_adjust_positions=True,
        )

        assert result.success is False
        assert "No operations" in result.message

    @pytest.mark.asyncio
    async def test_execute_batch_document_fetch_error(self):
        """Test handling of document fetch errors."""
        service = MagicMock()
        service.documents.return_value.get.return_value.execute = MagicMock(
            side_effect=Exception("Document not found")
        )
        manager = BatchOperationManager(service)

        result = await manager.execute_batch_with_search(
            "invalid-doc-id",
            [{"type": "insert_text", "index": 1, "text": "test"}],
        )

        assert result.success is False
        assert "Failed to fetch" in result.message

    @pytest.mark.asyncio
    async def test_execute_batch_with_search_based_insert(self, mock_service):
        """Test batch execution with search-based insert operations."""
        manager = BatchOperationManager(mock_service)

        # Search for "Conclusion" and insert before it
        operations = [
            {
                "type": "insert",
                "search": "Conclusion",
                "position": "before",
                "text": "NEW SECTION\n\n",
            },
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 1
        assert result.results[0].position_shift == 13  # len("NEW SECTION\n\n")
        assert result.results[0].resolved_index is not None

    @pytest.mark.asyncio
    async def test_execute_batch_search_not_found(self, mock_service):
        """Test batch execution when search text is not found."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "insert",
                "search": "NonExistentText",
                "position": "before",
                "text": "test",
            },
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is False
        assert "Failed to resolve" in result.message
        assert result.results[0].error is not None

    @pytest.mark.asyncio
    async def test_execute_batch_search_then_index_with_auto_adjust(self, mock_service):
        """Test mixed search and index operations with auto-adjustment."""
        manager = BatchOperationManager(mock_service)

        # First insert by search, then follow up with index-based operation
        # The index operation should be adjusted by the first insert's shift
        operations = [
            {
                "type": "insert",
                "search": "Hello",
                "position": "after",
                "text": " WORLD",  # 6 chars inserted after "Hello"
            },
            {
                "type": "insert_text",
                "index": 50,  # With auto_adjust, should become 56
                "text": "TEST",
            },
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 2
        assert result.results[0].position_shift == 6
        assert result.results[1].position_shift == 4
        assert result.total_position_shift == 10

    @pytest.mark.asyncio
    async def test_execute_batch_without_auto_adjust(self, mock_service):
        """Test that auto_adjust=False does not modify positions."""
        manager = BatchOperationManager(mock_service)

        # Two inserts at the same position - without auto-adjust
        # they should BOTH try to insert at 100 (which may cause issues in real API)
        operations = [
            {"type": "insert_text", "index": 100, "text": "FIRST"},
            {"type": "insert_text", "index": 100, "text": "SECOND"},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=False,
        )

        assert result.success is True
        # Both operations report their individual shifts
        assert result.results[0].position_shift == 5
        assert result.results[1].position_shift == 6

    @pytest.mark.asyncio
    async def test_execute_batch_delete_then_insert_adjusts_correctly(self, mock_service):
        """Test that delete shifts subsequent insert positions correctly."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "delete_text", "start_index": 50, "end_index": 60},  # -10 chars
            {"type": "insert_text", "index": 100, "text": "NEW"},  # Should become 90
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.results[0].position_shift == -10
        assert result.results[1].position_shift == 3
        assert result.total_position_shift == -7  # -10 + 3

    @pytest.mark.asyncio
    async def test_execute_batch_api_error(self, mock_service):
        """Test handling of API errors during batch execution."""
        mock_service.documents.return_value.batchUpdate.return_value.execute = MagicMock(
            side_effect=Exception("API Error: Rate limit exceeded")
        )
        manager = BatchOperationManager(mock_service)

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            [{"type": "insert_text", "index": 1, "text": "test"}],
        )

        assert result.success is False
        assert "Batch execution failed" in result.message


class TestOperationAliasNormalization:
    """Tests for operation type alias normalization."""

    def test_normalize_operation_type_short_forms(self):
        """Test that short operation type forms are normalized to canonical forms."""
        from gdocs.managers.batch_operation_manager import normalize_operation_type

        assert normalize_operation_type('insert') == 'insert_text'
        assert normalize_operation_type('delete') == 'delete_text'
        assert normalize_operation_type('format') == 'format_text'
        assert normalize_operation_type('replace') == 'replace_text'

    def test_normalize_operation_type_canonical_forms(self):
        """Test that canonical forms remain unchanged."""
        from gdocs.managers.batch_operation_manager import normalize_operation_type

        assert normalize_operation_type('insert_text') == 'insert_text'
        assert normalize_operation_type('delete_text') == 'delete_text'
        assert normalize_operation_type('format_text') == 'format_text'
        assert normalize_operation_type('replace_text') == 'replace_text'

    def test_normalize_operation_type_non_aliased(self):
        """Test that non-aliased operations pass through unchanged."""
        from gdocs.managers.batch_operation_manager import normalize_operation_type

        assert normalize_operation_type('insert_table') == 'insert_table'
        assert normalize_operation_type('insert_page_break') == 'insert_page_break'
        assert normalize_operation_type('find_replace') == 'find_replace'

    def test_normalize_operation_type_unknown(self):
        """Test that unknown types pass through unchanged."""
        from gdocs.managers.batch_operation_manager import normalize_operation_type

        assert normalize_operation_type('unknown_type') == 'unknown_type'

    def test_normalize_operation_dict(self):
        """Test normalizing an operation dictionary."""
        from gdocs.managers.batch_operation_manager import normalize_operation

        op = {'type': 'insert', 'index': 100, 'text': 'Hello'}
        normalized = normalize_operation(op)

        assert normalized['type'] == 'insert_text'
        assert normalized['index'] == 100
        assert normalized['text'] == 'Hello'
        # Original should be unchanged
        assert op['type'] == 'insert'

    def test_normalize_operation_missing_type(self):
        """Test that operations without type field pass through unchanged."""
        from gdocs.managers.batch_operation_manager import normalize_operation

        op = {'index': 100, 'text': 'Hello'}
        normalized = normalize_operation(op)

        assert normalized == op

    @pytest.mark.asyncio
    async def test_execute_batch_with_short_type_names(self):
        """Test that batch execution accepts short operation type names."""
        service = MagicMock()
        service.documents.return_value.get.return_value.execute = MagicMock(
            return_value={
                "body": {
                    "content": [
                        {
                            "paragraph": {
                                "elements": [
                                    {
                                        "textRun": {"content": "Test document content."},
                                        "startIndex": 1,
                                        "endIndex": 23,
                                    }
                                ]
                            }
                        }
                    ]
                }
            }
        )
        service.documents.return_value.batchUpdate.return_value.execute = MagicMock(
            return_value={"replies": [{}]}
        )
        manager = BatchOperationManager(service)

        # Use short type names instead of canonical names
        operations = [
            {"type": "insert", "index": 1, "text": "PREFIX: "},
            {"type": "format", "start_index": 1, "end_index": 9, "bold": True},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 2

    @pytest.mark.asyncio
    async def test_execute_batch_operations_with_short_type_names(self):
        """Test execute_batch_operations method accepts short operation type names."""
        service = MagicMock()
        service.documents.return_value.batchUpdate.return_value.execute = MagicMock(
            return_value={"replies": [{}, {}]}
        )
        manager = BatchOperationManager(service)

        # Use short type names
        operations = [
            {"type": "insert", "index": 1, "text": "Hello"},
            {"type": "delete", "start_index": 50, "end_index": 60},
        ]

        success, message, metadata = await manager.execute_batch_operations(
            "test-doc-id",
            operations,
        )

        assert success is True
        assert metadata['operations_count'] == 2
