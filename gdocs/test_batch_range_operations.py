"""
Unit tests for batch operations with range-based positioning.

These tests verify:
- Range spec resolution within batch operations
- Position auto-adjustment with range specs
- Various range spec formats (search bounds, extend, offsets, section)
- Error handling for range resolution failures
"""
import pytest
from unittest.mock import MagicMock, AsyncMock
from gdocs.managers.batch_operation_manager import (
    BatchOperationManager,
    BatchOperationResult,
    BatchExecutionResult,
)
from gdocs.docs_helpers import RangeResult


class TestBatchOperationManagerRangeResolution:
    """Tests for range_spec resolution in batch operations."""

    def setup_method(self):
        """Set up test fixtures."""
        self.manager = BatchOperationManager(MagicMock())
        # Create sample document data for testing
        self.doc_data = {
            "body": {
                "content": [
                    {
                        "startIndex": 1,
                        "endIndex": 150,
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "Introduction\nThis is the intro paragraph. It has multiple sentences. Here is more text.\n",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 89,
                                }
                            ]
                        }
                    },
                    {
                        "startIndex": 89,
                        "endIndex": 150,
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "Conclusion\nThis is the conclusion paragraph.\n",
                                    },
                                    "startIndex": 89,
                                    "endIndex": 134,
                                }
                            ]
                        }
                    }
                ]
            }
        }

    def test_resolve_range_spec_search_with_extend(self):
        """Test resolving range_spec with search + extend."""
        op = {
            "type": "delete",
            "range_spec": {
                "search": "intro",
                "extend": "paragraph",
            }
        }
        result = self.manager._resolve_range_spec(op, self.doc_data, 0, False)

        assert result.success is True
        assert result.start_index is not None
        assert result.end_index is not None
        assert result.extend_type == "paragraph"

    def test_resolve_range_spec_search_with_offsets(self):
        """Test resolving range_spec with search + character offsets."""
        op = {
            "type": "format",
            "range_spec": {
                "search": "intro",
                "before_chars": 5,
                "after_chars": 10,
            }
        }
        result = self.manager._resolve_range_spec(op, self.doc_data, 0, False)

        assert result.success is True
        assert result.start_index is not None
        assert result.end_index is not None

    def test_resolve_range_spec_search_bounds(self):
        """Test resolving range_spec with start/end search bounds."""
        op = {
            "type": "format",
            "range_spec": {
                "start": {"search": "This is the intro"},
                "end": {"search": "more text"},
            }
        }
        result = self.manager._resolve_range_spec(op, self.doc_data, 0, False)

        assert result.success is True
        assert result.start_index is not None
        assert result.end_index is not None
        assert result.start_index < result.end_index

    def test_resolve_range_spec_with_cumulative_shift(self):
        """Test that cumulative shift is applied to resolved range."""
        op = {
            "type": "format",
            "range_spec": {
                "search": "intro",
                "extend": "paragraph",
            }
        }
        result_no_shift = self.manager._resolve_range_spec(op, self.doc_data, 0, True)
        result_with_shift = self.manager._resolve_range_spec(op, self.doc_data, 50, True)

        assert result_with_shift.success is True
        assert result_with_shift.start_index == result_no_shift.start_index + 50
        assert result_with_shift.end_index == result_no_shift.end_index + 50

    def test_resolve_range_spec_no_shift_when_auto_adjust_false(self):
        """Test that cumulative shift is not applied when auto_adjust=False."""
        op = {
            "type": "format",
            "range_spec": {
                "search": "intro",
                "extend": "paragraph",
            }
        }
        result_no_adjust = self.manager._resolve_range_spec(op, self.doc_data, 50, False)
        result_baseline = self.manager._resolve_range_spec(op, self.doc_data, 0, False)

        assert result_no_adjust.start_index == result_baseline.start_index

    def test_resolve_range_spec_not_found(self):
        """Test error handling when search text not found."""
        op = {
            "type": "delete",
            "range_spec": {
                "search": "NonexistentText",
                "extend": "paragraph",
            }
        }
        result = self.manager._resolve_range_spec(op, self.doc_data, 0, False)

        assert result.success is False
        assert "not found" in result.message.lower()


class TestBatchOperationManagerRangeConversion:
    """Tests for converting range_spec operations to index-based."""

    def setup_method(self):
        """Set up test fixtures."""
        self.manager = BatchOperationManager(MagicMock())

    def test_convert_range_to_insert(self):
        """Test converting range-based insert to index-based."""
        op = {
            "type": "insert",
            "range_spec": {"search": "text", "extend": "paragraph"},
            "text": "New content",
        }
        result = self.manager._convert_range_to_index_op(op, "insert", 100, 150)

        assert result["type"] == "insert_text"
        assert result["index"] == 100
        assert result["text"] == "New content"
        assert "range_spec" not in result

    def test_convert_range_to_delete(self):
        """Test converting range-based delete to index-based."""
        op = {
            "type": "delete",
            "range_spec": {"search": "text", "extend": "sentence"},
        }
        result = self.manager._convert_range_to_index_op(op, "delete", 50, 75)

        assert result["type"] == "delete_text"
        assert result["start_index"] == 50
        assert result["end_index"] == 75
        assert "range_spec" not in result

    def test_convert_range_to_replace(self):
        """Test converting range-based replace to index-based."""
        op = {
            "type": "replace",
            "range_spec": {"search": "old text"},
            "text": "new text",
        }
        result = self.manager._convert_range_to_index_op(op, "replace", 100, 108)

        assert result["type"] == "replace_text"
        assert result["start_index"] == 100
        assert result["end_index"] == 108
        assert result["text"] == "new text"

    def test_convert_range_to_format(self):
        """Test converting range-based format to index-based."""
        op = {
            "type": "format",
            "range_spec": {"section": "Introduction"},
            "bold": True,
            "font_size": 14,
        }
        result = self.manager._convert_range_to_index_op(op, "format", 1, 89)

        assert result["type"] == "format_text"
        assert result["start_index"] == 1
        assert result["end_index"] == 89
        assert result["bold"] is True
        assert result["font_size"] == 14


class TestBatchOperationManagerRangeIntegration:
    """Integration tests for batch operations with range_spec."""

    @pytest.fixture
    def mock_service(self):
        """Create a mock Google Docs service."""
        service = MagicMock()
        service.documents.return_value.get.return_value.execute = MagicMock(
            return_value={
                "body": {
                    "content": [
                        {
                            "startIndex": 1,
                            "endIndex": 100,
                            "paragraph": {
                                "elements": [
                                    {
                                        "textRun": {
                                            "content": "Introduction\nThis is the intro paragraph with some text.\n",
                                        },
                                        "startIndex": 1,
                                        "endIndex": 58,
                                    }
                                ]
                            }
                        },
                        {
                            "startIndex": 58,
                            "endIndex": 120,
                            "paragraph": {
                                "elements": [
                                    {
                                        "textRun": {
                                            "content": "Conclusion\nThis is the conclusion.\n",
                                        },
                                        "startIndex": 58,
                                        "endIndex": 93,
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
    async def test_execute_batch_with_range_spec_extend(self, mock_service):
        """Test batch execution with range_spec using extend."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "format",
                "range_spec": {
                    "search": "intro",
                    "extend": "paragraph",
                },
                "bold": True,
            },
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 1
        assert result.results[0].resolved_index is not None

    @pytest.mark.asyncio
    async def test_execute_batch_with_range_spec_offsets(self, mock_service):
        """Test batch execution with range_spec using character offsets."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "format",
                "range_spec": {
                    "search": "paragraph",
                    "before_chars": 5,
                    "after_chars": 10,
                },
                "italic": True,
            },
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 1

    @pytest.mark.asyncio
    async def test_execute_batch_with_range_spec_bounds(self, mock_service):
        """Test batch execution with range_spec using start/end bounds."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "format",
                "range_spec": {
                    "start": {"search": "This is the intro"},
                    "end": {"search": "some text"},
                },
                "underline": True,
            },
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 1

    @pytest.mark.asyncio
    async def test_execute_batch_range_spec_with_auto_adjust(self, mock_service):
        """Test that range_spec operations work with auto position adjustment."""
        manager = BatchOperationManager(mock_service)

        operations = [
            # First: insert text at beginning
            {"type": "insert_text", "index": 1, "text": "PREFIX: "},  # +8 chars
            # Second: format using range_spec (should be adjusted by +8)
            {
                "type": "format",
                "range_spec": {
                    "search": "intro",
                    "extend": "line",
                },
                "bold": True,
            },
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 2
        # First operation shifts by 8
        assert result.results[0].position_shift == 8

    @pytest.mark.asyncio
    async def test_execute_batch_range_spec_then_delete_adjusts_following(self, mock_service):
        """Test that delete with range_spec properly shifts following operations."""
        manager = BatchOperationManager(mock_service)

        operations = [
            # Delete a range using range_spec
            {
                "type": "delete",
                "range_spec": {
                    "search": "intro paragraph",
                    "extend": "sentence",
                },
            },
            # Insert at a later position (should be adjusted)
            {"type": "insert_text", "index": 100, "text": "NEW"},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 2
        # First operation should have negative shift (deletion)
        assert result.results[0].position_shift < 0

    @pytest.mark.asyncio
    async def test_execute_batch_range_spec_not_found_fails_batch(self, mock_service):
        """Test that batch fails when range_spec cannot be resolved."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "delete",
                "range_spec": {
                    "search": "NonExistentText",
                    "extend": "paragraph",
                },
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
    async def test_execute_batch_mixed_range_and_search_operations(self, mock_service):
        """Test batch with mixed range_spec and search-based operations."""
        manager = BatchOperationManager(mock_service)

        operations = [
            # Range spec operation
            {
                "type": "format",
                "range_spec": {
                    "search": "intro",
                    "extend": "paragraph",
                },
                "bold": True,
            },
            # Legacy search-based operation
            {
                "type": "insert",
                "search": "Conclusion",
                "position": "before",
                "text": "\n\nNew Section\n\n",
            },
            # Direct index-based operation
            {"type": "insert_text", "index": 1, "text": ">> "},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 3

    @pytest.mark.asyncio
    async def test_execute_batch_range_spec_replace_calculates_shift(self, mock_service):
        """Test that replace with range_spec correctly calculates position shift."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "replace",
                "range_spec": {
                    "search": "intro paragraph",
                    "before_chars": 0,
                    "after_chars": 0,  # Select exactly the search text
                },
                "text": "overview",  # Replacing longer text with shorter
            },
            {"type": "insert_text", "index": 100, "text": "X"},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        # Replace should have negative shift (shorter replacement)
        first_shift = result.results[0].position_shift
        assert first_shift != 0  # There should be some shift


class TestBatchOperationManagerRangeEdgeCases:
    """Edge case tests for range_spec operations."""

    @pytest.fixture
    def mock_service(self):
        """Create a mock Google Docs service."""
        service = MagicMock()
        service.documents.return_value.get.return_value.execute = MagicMock(
            return_value={
                "body": {
                    "content": [
                        {
                            "startIndex": 1,
                            "endIndex": 50,
                            "paragraph": {
                                "elements": [
                                    {
                                        "textRun": {
                                            "content": "Short document with minimal content.\n",
                                        },
                                        "startIndex": 1,
                                        "endIndex": 38,
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
    async def test_range_spec_empty_object_fails(self, mock_service):
        """Test that empty range_spec fails gracefully."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "format",
                "range_spec": {},
                "bold": True,
            },
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is False
        assert result.results[0].error is not None

    @pytest.mark.asyncio
    async def test_range_spec_invalid_extend_type(self, mock_service):
        """Test that invalid extend type fails gracefully."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "format",
                "range_spec": {
                    "search": "document",
                    "extend": "invalid_type",
                },
                "bold": True,
            },
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is False
        assert "invalid" in result.results[0].error.lower()

    @pytest.mark.asyncio
    async def test_range_spec_preserves_other_operation_fields(self, mock_service):
        """Test that range_spec resolution preserves other operation fields."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "format",
                "range_spec": {
                    "search": "document",
                    "extend": "paragraph",
                },
                "bold": True,
                "italic": True,
                "font_size": 16,
                "font_family": "Arial",
            },
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        # The operation should have been executed with all formatting options
        assert result.operations_completed == 1
