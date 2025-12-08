"""
Unit tests for batch_edit_doc with search-based positioning and position tracking.

These tests verify:
- Search-based positioning (before/after/replace)
- Position auto-adjustment for sequential operations
- Per-operation result tracking
- Request building from resolved operations
- Recent insert preference for format operations
"""
import pytest
from unittest.mock import MagicMock
from gdocs.managers.batch_operation_manager import (
    BatchOperationManager,
    BatchOperationResult,
    BatchExecutionResult,
    VirtualTextTracker,
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
            resolved=True,
            executed=True,
        )
        result_dict = result.to_dict()

        assert "resolved_index" not in result_dict
        assert "error" not in result_dict
        assert result_dict["index"] == 0
        assert result_dict["position_shift"] == 5
        assert result_dict["resolved"] is True
        assert result_dict["executed"] is True

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
            resolved=False,
            executed=False,
        )
        result_dict = result.to_dict()

        assert result_dict["index"] == 1
        assert result_dict["resolved_index"] == 50
        assert result_dict["error"] == "Search text not found"
        assert result_dict["resolved"] is False
        assert result_dict["executed"] is False

    def test_resolved_not_executed_state(self):
        """Test that resolved but not executed state is represented correctly.

        This tests the fix for google_workspace_mcp-c96a where operations that
        were resolved but not executed (e.g., due to another operation failing)
        were misleadingly showing success=True.
        """
        result = BatchOperationResult(
            index=1,
            type="insert_text",
            success=False,  # Not successful yet - needs execution
            description="insert at 100",
            position_shift=5,
            resolved_index=100,
            resolved=True,  # Resolution succeeded
            executed=False,  # But not executed yet
        )
        result_dict = result.to_dict()

        # The key fix: success should be False for unexecuted operations
        assert result_dict["success"] is False
        assert result_dict["resolved"] is True
        assert result_dict["executed"] is False
        assert result_dict["resolved_index"] == 100


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
            resolved=True,
            executed=True,
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
        # Verify operation result includes resolved and executed fields
        assert result_dict["results"][0]["resolved"] is True
        assert result_dict["results"][0]["executed"] is True


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

    def test_convert_insert_table_search_to_index(self):
        """Test converting search-based insert_table to index-based.

        Fix for google_workspace_mcp-9b57: insert_table with search-based positioning
        should correctly set the index field.
        """
        op = {
            "type": "insert_table",
            "search": "Conclusion",
            "position": "after",
            "rows": 3,
            "columns": 4,
        }
        result = self.manager._convert_search_to_index_op(op, "insert_table", 500, 510)

        assert result["type"] == "insert_table"
        assert result["index"] == 500
        assert result["rows"] == 3
        assert result["columns"] == 4
        assert "search" not in result
        assert "position" not in result

    def test_convert_insert_page_break_search_to_index(self):
        """Test converting search-based insert_page_break to index-based.

        Fix for google_workspace_mcp-9b57: insert_page_break with search-based positioning
        should correctly set the index field.
        """
        op = {
            "type": "insert_page_break",
            "search": "Chapter 2",
            "position": "before",
        }
        result = self.manager._convert_search_to_index_op(op, "insert_page_break", 300, 309)

        assert result["type"] == "insert_page_break"
        assert result["index"] == 300
        assert "search" not in result
        assert "position" not in result


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

    def test_describe_format_with_colors(self):
        """Test description for format operation with foreground_color and background_color."""
        op = {
            "type": "format_text",
            "start_index": 10,
            "end_index": 20,
            "foreground_color": "#FF0000",
            "background_color": "#00FF00",
        }
        desc = self.manager._describe_operation(op)

        assert "format" in desc.lower()
        assert "foreground_color" in desc
        assert "#FF0000" in desc
        assert "background_color" in desc
        assert "#00FF00" in desc
        assert "none" not in desc.lower()

    def test_describe_format_with_all_options(self):
        """Test description includes all formatting options."""
        op = {
            "type": "format_text",
            "start_index": 10,
            "end_index": 20,
            "strikethrough": True,
            "small_caps": True,
            "subscript": True,
            "superscript": False,
            "link": "https://example.com",
        }
        desc = self.manager._describe_operation(op)

        assert "strikethrough" in desc
        assert "small_caps" in desc
        assert "subscript" in desc
        assert "superscript" in desc
        assert "link" in desc

    def test_describe_format_with_paragraph_options(self):
        """Test description includes paragraph-level formatting options.

        Fix for google_workspace_mcp-a332: batch_edit_doc format operation
        description should include heading_style and other paragraph options.
        """
        op = {
            "type": "format_text",
            "start_index": 10,
            "end_index": 20,
            "heading_style": "HEADING_2",
            "alignment": "CENTER",
            "line_spacing": 150,
        }
        desc = self.manager._describe_operation(op)

        assert "heading_style" in desc
        assert "HEADING_2" in desc
        assert "alignment" in desc
        assert "CENTER" in desc
        assert "line_spacing" in desc
        assert "150" in desc

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
        """Test building insert text request with clear formatting."""
        ops = [{"type": "insert_text", "index": 100, "text": "Hello"}]
        requests = self.manager._build_requests_from_resolved(ops)

        # Insert creates 2 requests: insert + clear formatting (to prevent inheritance)
        assert len(requests) == 2
        assert "insertText" in requests[0]
        assert requests[0]["insertText"]["location"]["index"] == 100
        # Second request clears formatting to prevent inheriting surrounding styles
        assert "updateTextStyle" in requests[1]

    def test_build_delete_request(self):
        """Test building delete request."""
        ops = [{"type": "delete_text", "start_index": 50, "end_index": 100}]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 1
        assert "deleteContentRange" in requests[0]

    def test_build_replace_request_creates_three(self):
        """Test that replace creates delete + insert + clear formatting requests."""
        ops = [
            {
                "type": "replace_text",
                "start_index": 100,
                "end_index": 110,
                "text": "new text",
            }
        ]
        requests = self.manager._build_requests_from_resolved(ops)

        # Replace creates 3 requests: delete + insert + clear formatting
        assert len(requests) == 3
        assert "deleteContentRange" in requests[0]
        assert "insertText" in requests[1]
        # Third request clears formatting to prevent inheriting surrounding styles
        assert "updateTextStyle" in requests[2]

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

    def test_build_format_request_with_heading_style(self):
        """Test building format request with heading_style.

        Fix for google_workspace_mcp-a332: batch_edit_doc format operation should
        apply heading_style by creating an updateParagraphStyle request.
        """
        ops = [
            {
                "type": "format_text",
                "start_index": 0,
                "end_index": 50,
                "heading_style": "HEADING_2",
            }
        ]
        requests = self.manager._build_requests_from_resolved(ops)

        # Should create 1 request: updateParagraphStyle for heading
        assert len(requests) == 1
        assert "updateParagraphStyle" in requests[0]
        assert requests[0]["updateParagraphStyle"]["paragraphStyle"]["namedStyleType"] == "HEADING_2"

    def test_build_format_request_with_heading_style_and_bold(self):
        """Test building format request with both heading_style and text formatting.

        Fix for google_workspace_mcp-a332: batch_edit_doc format operation should
        handle both text-level and paragraph-level formatting together.
        """
        ops = [
            {
                "type": "format_text",
                "start_index": 0,
                "end_index": 50,
                "bold": True,
                "heading_style": "HEADING_2",
            }
        ]
        requests = self.manager._build_requests_from_resolved(ops)

        # Should create 2 requests: updateTextStyle for bold + updateParagraphStyle for heading
        assert len(requests) == 2

        # Check for text style request
        text_style_reqs = [r for r in requests if "updateTextStyle" in r]
        assert len(text_style_reqs) == 1
        assert text_style_reqs[0]["updateTextStyle"]["textStyle"]["bold"] is True

        # Check for paragraph style request
        para_style_reqs = [r for r in requests if "updateParagraphStyle" in r]
        assert len(para_style_reqs) == 1
        assert para_style_reqs[0]["updateParagraphStyle"]["paragraphStyle"]["namedStyleType"] == "HEADING_2"

    def test_build_format_request_with_alignment(self):
        """Test building format request with paragraph alignment."""
        ops = [
            {
                "type": "format_text",
                "start_index": 0,
                "end_index": 50,
                "alignment": "CENTER",
            }
        ]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 1
        assert "updateParagraphStyle" in requests[0]
        assert requests[0]["updateParagraphStyle"]["paragraphStyle"]["alignment"] == "CENTER"

    def test_build_format_request_with_line_spacing(self):
        """Test building format request with line spacing."""
        ops = [
            {
                "type": "format_text",
                "start_index": 0,
                "end_index": 50,
                "line_spacing": 150,  # 1.5x spacing
            }
        ]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 1
        assert "updateParagraphStyle" in requests[0]
        assert requests[0]["updateParagraphStyle"]["paragraphStyle"]["lineSpacing"] == 150

    def test_build_skips_none_operations(self):
        """Test that None operations are skipped."""
        ops = [
            {"type": "insert_text", "index": 100, "text": "Hello"},
            None,
            {"type": "insert_text", "index": 200, "text": "World"},
        ]
        requests = self.manager._build_requests_from_resolved(ops)

        # 2 inserts * 2 requests each (insert + clear formatting) = 4 requests
        assert len(requests) == 4

    def test_build_table_request(self):
        """Test building table insert request."""
        ops = [{"type": "insert_table", "index": 50, "rows": 3, "columns": 4}]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 1
        assert "insertTable" in requests[0]
        assert requests[0]["insertTable"]["rows"] == 3
        assert requests[0]["insertTable"]["columns"] == 4


class TestInsertWithFormatting:
    """Tests for insert operations with formatting parameters.

    This tests the fix for google_workspace_mcp-0b22: batch_edit_doc insert
    operations should apply formatting parameters like bold, italic, etc.
    """

    def setup_method(self):
        """Set up test fixtures."""
        self.manager = BatchOperationManager(MagicMock())

    def test_build_insert_with_bold_creates_format_request(self):
        """Test that insert with bold=True creates insert, clear formatting, and format requests.

        Fix for google_workspace_mcp-332b: Clear formatting before applying styles to prevent
        inheriting surrounding text formatting.
        """
        ops = [{"type": "insert_text", "index": 100, "text": "Bold Text", "bold": True}]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 3  # insert + clear formatting + apply formatting
        assert "insertText" in requests[0]
        assert "updateTextStyle" in requests[1]  # clear formatting request
        assert "updateTextStyle" in requests[2]  # apply formatting request
        # Verify the clear formatting request resets styles
        clear_req = requests[1]["updateTextStyle"]
        assert clear_req["textStyle"]["bold"] is False
        assert clear_req["textStyle"]["italic"] is False
        # Verify the format request targets the correct range and sets bold
        format_req = requests[2]["updateTextStyle"]
        assert format_req["range"]["startIndex"] == 100
        assert format_req["range"]["endIndex"] == 109  # 100 + len("Bold Text")
        assert format_req["textStyle"]["bold"] is True

    def test_build_insert_with_italic_creates_format_request(self):
        """Test that insert with italic=True creates insert, clear, and format requests."""
        ops = [{"type": "insert_text", "index": 50, "text": "Italic", "italic": True}]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 3  # insert + clear formatting + apply formatting
        assert "insertText" in requests[0]
        assert "updateTextStyle" in requests[1]  # clear formatting
        assert "updateTextStyle" in requests[2]  # apply formatting
        format_req = requests[2]["updateTextStyle"]
        assert format_req["textStyle"]["italic"] is True

    def test_build_insert_with_multiple_formats(self):
        """Test insert with multiple formatting options."""
        ops = [{
            "type": "insert_text",
            "index": 1,
            "text": "Formatted",
            "bold": True,
            "italic": True,
            "underline": True,
        }]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 3  # insert + clear formatting + apply formatting
        format_req = requests[2]["updateTextStyle"]
        assert format_req["textStyle"]["bold"] is True
        assert format_req["textStyle"]["italic"] is True
        assert format_req["textStyle"]["underline"] is True

    def test_build_insert_with_font_size(self):
        """Test insert with font_size formatting."""
        ops = [{
            "type": "insert_text",
            "index": 1,
            "text": "Big Text",
            "font_size": 24,
        }]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 3  # insert + clear formatting + apply formatting
        format_req = requests[2]["updateTextStyle"]
        assert format_req["textStyle"]["fontSize"]["magnitude"] == 24

    def test_build_insert_with_foreground_color(self):
        """Test insert with foreground_color (text color) formatting."""
        ops = [{
            "type": "insert_text",
            "index": 1,
            "text": "Red Text",
            "foreground_color": "#FF0000",
        }]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 3  # insert + clear formatting + apply formatting
        format_req = requests[2]["updateTextStyle"]
        assert "foregroundColor" in format_req["textStyle"]

    def test_build_insert_without_formatting_clears_inherited_styles(self):
        """Test that insert without formatting creates insert + clear formatting requests.

        This prevents inserted text from inheriting surrounding formatting.
        """
        ops = [{"type": "insert_text", "index": 100, "text": "Plain Text"}]
        requests = self.manager._build_requests_from_resolved(ops)

        # Insert without explicit formatting still clears inherited styles
        assert len(requests) == 2
        assert "insertText" in requests[0]
        # Second request clears formatting to ensure plain text stays plain
        assert "updateTextStyle" in requests[1]

    def test_build_replace_with_bold_creates_format_request(self):
        """Test that replace with bold=True creates delete, insert, clear, and format requests."""
        ops = [{
            "type": "replace_text",
            "start_index": 100,
            "end_index": 110,
            "text": "Bold Replacement",
            "bold": True,
        }]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 4  # delete + insert + clear formatting + apply formatting
        assert "deleteContentRange" in requests[0]
        assert "insertText" in requests[1]
        assert "updateTextStyle" in requests[2]  # clear formatting
        assert "updateTextStyle" in requests[3]  # apply formatting
        # Verify format range matches the replacement text
        format_req = requests[3]["updateTextStyle"]
        assert format_req["range"]["startIndex"] == 100
        assert format_req["range"]["endIndex"] == 116  # 100 + len("Bold Replacement")

    def test_build_replace_without_formatting_clears_inherited_styles(self):
        """Test that replace without formatting creates delete, insert, and clear formatting requests.

        This prevents replaced text from inheriting surrounding formatting.
        """
        ops = [{
            "type": "replace_text",
            "start_index": 50,
            "end_index": 60,
            "text": "Plain",
        }]
        requests = self.manager._build_requests_from_resolved(ops)

        # Replace without explicit formatting still clears inherited styles
        assert len(requests) == 3
        assert "deleteContentRange" in requests[0]
        assert "insertText" in requests[1]
        # Third request clears formatting to ensure plain text stays plain
        assert "updateTextStyle" in requests[2]

    def test_build_insert_with_link(self):
        """Test insert with hyperlink formatting."""
        ops = [{
            "type": "insert_text",
            "index": 1,
            "text": "Click here",
            "link": "https://example.com",
        }]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 3  # insert + clear formatting + apply formatting
        format_req = requests[2]["updateTextStyle"]
        assert format_req["textStyle"]["link"]["url"] == "https://example.com"

    def test_build_insert_with_strikethrough(self):
        """Test insert with strikethrough formatting."""
        ops = [{
            "type": "insert_text",
            "index": 1,
            "text": "Crossed out",
            "strikethrough": True,
        }]
        requests = self.manager._build_requests_from_resolved(ops)

        assert len(requests) == 3  # insert + clear formatting + apply formatting
        format_req = requests[2]["updateTextStyle"]
        assert format_req["textStyle"]["strikethrough"] is True


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

    @pytest.mark.asyncio
    async def test_execute_batch_invalid_operation_type_rejected(self, mock_service):
        """Test that invalid operation types are rejected with clear error message."""
        manager = BatchOperationManager(mock_service)

        # Try an invalid operation type
        operations = [
            {"type": "fake_op", "text": "bad"},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
        )

        assert result.success is False
        assert result.operations_completed == 0
        assert len(result.results) == 1
        assert result.results[0].success is False
        assert "Unsupported operation type: 'fake_op'" in result.results[0].error
        assert "Valid types:" in result.results[0].error

    @pytest.mark.asyncio
    async def test_execute_batch_missing_type_field_rejected(self, mock_service):
        """Test that operations missing the type field are rejected."""
        manager = BatchOperationManager(mock_service)

        # Try an operation without a type field
        operations = [
            {"index": 100, "text": "no type field"},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
        )

        assert result.success is False
        assert result.operations_completed == 0
        assert len(result.results) == 1
        assert result.results[0].success is False
        assert "Missing 'type' field" in result.results[0].error

    @pytest.mark.asyncio
    async def test_execute_batch_mixed_valid_invalid_types(self, mock_service):
        """Test that a batch with mix of valid and invalid types fails on invalid ones."""
        manager = BatchOperationManager(mock_service)

        # First operation is valid, second is invalid
        operations = [
            {"type": "insert_text", "index": 1, "text": "valid"},
            {"type": "completely_fake_op", "text": "invalid"},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
        )

        # The batch should fail because of the invalid operation
        assert result.success is False
        assert len(result.results) == 2
        # First operation resolved OK (marked success initially, but batch failed)
        assert result.results[1].success is False
        assert "Unsupported operation type: 'completely_fake_op'" in result.results[1].error


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


class TestBatchAllOccurrences:
    """Tests for all_occurrences expansion in batch operations."""

    @pytest.fixture
    def mock_service_with_multiple_matches(self):
        """Create a mock service with a document containing multiple matches."""
        service = MagicMock()
        # Document: "TODO: First task. TODO: Second task. TODO: Third task."
        # Positions: TODO at indices 1, 19, 38 (1-indexed for Google Docs)
        service.documents.return_value.get.return_value.execute = MagicMock(
            return_value={
                "body": {
                    "content": [
                        {
                            "paragraph": {
                                "elements": [
                                    {
                                        "textRun": {
                                            "content": "TODO: First task. TODO: Second task. TODO: Third task.\n",
                                        },
                                        "startIndex": 1,
                                        "endIndex": 56,
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

    def test_expand_all_occurrences_format(self, mock_service_with_multiple_matches):
        """Test that format operations with all_occurrences expands to multiple operations."""
        manager = BatchOperationManager(mock_service_with_multiple_matches)

        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "TODO: First task. TODO: Second task. TODO: Third task.\n",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 56,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        operations = [
            {
                "type": "format",
                "search": "TODO",
                "all_occurrences": True,
                "bold": True,
            }
        ]

        expanded = manager._expand_all_occurrences_operations(operations, doc_data)

        # Should expand to 3 operations (one for each TODO)
        assert len(expanded) == 3

        # Each should be format_text with start_index/end_index
        for op in expanded:
            assert op["type"] == "format_text"
            assert "start_index" in op
            assert "end_index" in op
            assert op["bold"] is True
            assert "search" not in op
            assert "all_occurrences" not in op

    def test_expand_all_occurrences_preserves_order_reversed(self, mock_service_with_multiple_matches):
        """Test that expanded operations are in reverse document order (for correct index handling)."""
        manager = BatchOperationManager(mock_service_with_multiple_matches)

        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "TODO: First task. TODO: Second task. TODO: Third task.\n",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 56,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        operations = [
            {
                "type": "format",
                "search": "TODO",
                "all_occurrences": True,
                "bold": True,
            }
        ]

        expanded = manager._expand_all_occurrences_operations(operations, doc_data)

        # Operations should be in reverse order (last occurrence first)
        # so that editing from end to start doesn't invalidate indices
        indices = [op["start_index"] for op in expanded]
        assert indices == sorted(indices, reverse=True)

    def test_expand_all_occurrences_replace(self, mock_service_with_multiple_matches):
        """Test that replace operations with all_occurrences expands correctly."""
        manager = BatchOperationManager(mock_service_with_multiple_matches)

        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "TODO: First task. TODO: Second task. TODO: Third task.\n",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 56,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        operations = [
            {
                "type": "replace",
                "search": "TODO",
                "all_occurrences": True,
                "text": "DONE",
            }
        ]

        expanded = manager._expand_all_occurrences_operations(operations, doc_data)

        assert len(expanded) == 3
        for op in expanded:
            assert op["type"] == "replace_text"
            assert op["text"] == "DONE"

    def test_expand_all_occurrences_no_matches(self, mock_service_with_multiple_matches):
        """Test that all_occurrences with no matches keeps original operation (for error reporting)."""
        manager = BatchOperationManager(mock_service_with_multiple_matches)

        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "No matches here.\n",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 18,
                                }
                            ]
                        }
                    }
                ]
            }
        }

        operations = [
            {
                "type": "format",
                "search": "NONEXISTENT",
                "all_occurrences": True,
                "bold": True,
            }
        ]

        expanded = manager._expand_all_occurrences_operations(operations, doc_data)

        # Should keep original operation (which will fail with useful error later)
        assert len(expanded) == 1
        assert expanded[0]["search"] == "NONEXISTENT"

    def test_expand_preserves_non_all_occurrences_operations(self, mock_service_with_multiple_matches):
        """Test that operations without all_occurrences are not expanded."""
        manager = BatchOperationManager(mock_service_with_multiple_matches)

        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "TODO: First task. TODO: Second task.\n",
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

        operations = [
            {"type": "insert_text", "index": 1, "text": "PREFIX "},
            {
                "type": "format",
                "search": "TODO",
                "position": "replace",  # Single occurrence
                "bold": True,
            },
        ]

        expanded = manager._expand_all_occurrences_operations(operations, doc_data)

        # Neither should be expanded
        assert len(expanded) == 2
        assert expanded[0]["type"] == "insert_text"
        assert expanded[1]["search"] == "TODO"

    @pytest.mark.asyncio
    async def test_execute_batch_with_all_occurrences(self, mock_service_with_multiple_matches):
        """Test full batch execution with all_occurrences operations."""
        manager = BatchOperationManager(mock_service_with_multiple_matches)

        operations = [
            {
                "type": "format",
                "search": "TODO",
                "all_occurrences": True,
                "bold": True,
            }
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        # Original operation count is 1
        assert result.total_operations == 1
        # But 3 operations were executed (one per TODO)
        assert result.operations_completed == 3
        # Results should have 3 entries
        assert len(result.results) == 3
        # Message should mention expansion
        assert "all_occurrences" in result.message

    @pytest.mark.asyncio
    async def test_execute_batch_mixed_all_occurrences_and_regular(self, mock_service_with_multiple_matches):
        """Test batch with both all_occurrences and regular operations."""
        manager = BatchOperationManager(mock_service_with_multiple_matches)

        operations = [
            {"type": "insert_text", "index": 1, "text": "PREFIX: "},
            {
                "type": "format",
                "search": "TODO",
                "all_occurrences": True,
                "bold": True,
            },
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.total_operations == 2  # Original count
        assert result.operations_completed == 4  # 1 insert + 3 formats

    def test_expand_all_occurrences_delete(self, mock_service_with_multiple_matches):
        """Test that delete operations with all_occurrences expands correctly."""
        manager = BatchOperationManager(mock_service_with_multiple_matches)

        doc_data = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "TODO: First task. TODO: Second task.\n",
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

        operations = [
            {
                "type": "delete",
                "search": "TODO",
                "all_occurrences": True,
            }
        ]

        expanded = manager._expand_all_occurrences_operations(operations, doc_data)

        assert len(expanded) == 2
        for op in expanded:
            assert op["type"] == "delete_text"
            assert "start_index" in op
            assert "end_index" in op


class TestBatchPreviewMode:
    """Tests for preview mode in batch operations."""

    @pytest.mark.asyncio
    async def test_preview_mode_returns_preview_fields(self):
        """Test that preview mode returns preview-specific fields."""
        mock_service = MagicMock()
        mock_service.documents().get().execute = MagicMock(
            return_value={
                "body": {
                    "content": [
                        {
                            "paragraph": {
                                "elements": [
                                    {
                                        "textRun": {"content": "Hello World\n"},
                                        "startIndex": 1,
                                        "endIndex": 13,
                                    }
                                ]
                            }
                        }
                    ]
                }
            }
        )
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert_text", "index": 1, "text": "New text "},
        ]

        result = await manager.execute_batch_with_search(
            "doc123", operations, preview_only=True
        )

        assert result.success is True
        assert result.preview is True
        assert result.would_modify is True
        assert result.operations_completed == 0  # Nothing actually executed
        assert result.total_operations == 1
        assert len(result.results) == 1
        assert "Would execute" in result.message

    @pytest.mark.asyncio
    async def test_preview_mode_does_not_execute(self):
        """Test that preview mode does not call the API to execute operations."""
        mock_service = MagicMock()
        mock_documents = MagicMock()
        mock_service.documents.return_value = mock_documents
        mock_documents.get.return_value.execute.return_value = {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {"content": "Hello World\n"},
                                    "startIndex": 1,
                                    "endIndex": 13,
                                }
                            ]
                        }
                    }
                ]
            }
        }
        mock_documents.batchUpdate.return_value.execute.return_value = {"replies": []}
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert_text", "index": 1, "text": "New text "},
        ]

        result = await manager.execute_batch_with_search(
            "doc123", operations, preview_only=True
        )

        # Verify batchUpdate was NOT called
        mock_documents.batchUpdate.assert_not_called()
        assert result.success is True
        assert result.preview is True

    @pytest.mark.asyncio
    async def test_preview_mode_calculates_position_shifts(self):
        """Test that preview mode correctly calculates cumulative position shifts."""
        mock_service = MagicMock()
        mock_service.documents().get().execute = MagicMock(
            return_value={
                "body": {
                    "content": [
                        {
                            "paragraph": {
                                "elements": [
                                    {
                                        "textRun": {"content": "Hello World\n"},
                                        "startIndex": 1,
                                        "endIndex": 13,
                                    }
                                ]
                            }
                        }
                    ]
                }
            }
        )
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert_text", "index": 1, "text": "AAAA"},  # +4
            {"type": "insert_text", "index": 10, "text": "BBBBB"},  # +5
        ]

        result = await manager.execute_batch_with_search(
            "doc123", operations, preview_only=True
        )

        assert result.total_position_shift == 9  # 4 + 5
        assert result.results[0].position_shift == 4
        assert result.results[1].position_shift == 5

    @pytest.mark.asyncio
    async def test_preview_mode_with_search_based_operations(self):
        """Test preview mode with search-based positioning."""
        mock_service = MagicMock()
        mock_service.documents().get().execute = MagicMock(
            return_value={
                "body": {
                    "content": [
                        {
                            "paragraph": {
                                "elements": [
                                    {
                                        "textRun": {"content": "Find this text\n"},
                                        "startIndex": 1,
                                        "endIndex": 16,
                                    }
                                ]
                            }
                        }
                    ]
                }
            }
        )
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "insert",
                "search": "this",
                "position": "before",
                "text": "NEW ",
            }
        ]

        result = await manager.execute_batch_with_search(
            "doc123", operations, preview_only=True
        )

        assert result.success is True
        assert result.preview is True
        assert result.would_modify is True
        assert result.results[0].resolved_index is not None

    @pytest.mark.asyncio
    async def test_preview_mode_search_not_found_returns_failure(self):
        """Test that preview mode returns failure when search text not found."""
        mock_service = MagicMock()
        mock_service.documents().get().execute = MagicMock(
            return_value={
                "body": {
                    "content": [
                        {
                            "paragraph": {
                                "elements": [
                                    {
                                        "textRun": {"content": "Hello World\n"},
                                        "startIndex": 1,
                                        "endIndex": 13,
                                    }
                                ]
                            }
                        }
                    ]
                }
            }
        )
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "insert",
                "search": "NOT FOUND TEXT",
                "position": "before",
                "text": "NEW ",
            }
        ]

        result = await manager.execute_batch_with_search(
            "doc123", operations, preview_only=True
        )

        assert result.success is False
        assert result.preview is True
        assert "Failed to resolve" in result.message

    @pytest.mark.asyncio
    async def test_preview_mode_with_all_occurrences(self):
        """Test preview mode with all_occurrences operations."""
        mock_service = MagicMock()
        mock_service.documents().get().execute = MagicMock(
            return_value={
                "body": {
                    "content": [
                        {
                            "paragraph": {
                                "elements": [
                                    {
                                        "textRun": {
                                            "content": "TODO: First. TODO: Second.\n"
                                        },
                                        "startIndex": 1,
                                        "endIndex": 28,
                                    }
                                ]
                            }
                        }
                    ]
                }
            }
        )
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "format",
                "search": "TODO",
                "all_occurrences": True,
                "bold": True,
            }
        ]

        result = await manager.execute_batch_with_search(
            "doc123", operations, preview_only=True
        )

        assert result.success is True
        assert result.preview is True
        # Should have 2 operations after expansion (one for each TODO)
        assert len(result.results) == 2
        assert "all_occurrences expansion" in result.message

    def test_to_dict_includes_preview_fields_when_preview_true(self):
        """Test that to_dict includes preview fields when in preview mode."""
        result = BatchExecutionResult(
            success=True,
            operations_completed=0,
            total_operations=1,
            results=[],
            total_position_shift=5,
            message="Would execute",
            document_link="https://docs.google.com/document/d/123/edit",
            preview=True,
            would_modify=True,
        )
        result_dict = result.to_dict()

        assert result_dict["preview"] is True
        assert result_dict["would_modify"] is True

    def test_to_dict_excludes_preview_fields_when_not_preview(self):
        """Test that to_dict excludes preview fields when not in preview mode."""
        result = BatchExecutionResult(
            success=True,
            operations_completed=1,
            total_operations=1,
            results=[],
            total_position_shift=5,
            message="Success",
            document_link="https://docs.google.com/document/d/123/edit",
            preview=False,
            would_modify=False,
        )
        result_dict = result.to_dict()

        assert "preview" not in result_dict
        assert "would_modify" not in result_dict


class TestVirtualTextTracker:
    """Tests for VirtualTextTracker class that supports chained search operations."""

    def _make_doc_data(self, text: str, start_index: int = 1) -> dict:
        """Create a mock document data structure with the given text."""
        return {
            "body": {
                "content": [
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": text,
                                    },
                                    "startIndex": start_index,
                                    "endIndex": start_index + len(text),
                                }
                            ]
                        }
                    }
                ]
            }
        }

    def test_init_extracts_text(self):
        """Test that tracker initializes correctly from document data."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("Hello World")
        tracker = VirtualTextTracker(doc_data)

        assert tracker.text == "Hello World"
        assert len(tracker.index_map) == 11

    def test_search_text_finds_existing(self):
        """Test searching for text that exists in document."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("Hello World")
        tracker = VirtualTextTracker(doc_data)

        success, start, end, msg = tracker.search_text("World", "replace")
        assert success is True
        assert start == 7  # 1-based index: "World" starts at index 7
        assert end == 12   # End is exclusive

    def test_search_text_not_found(self):
        """Test searching for text that doesn't exist."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("Hello World")
        tracker = VirtualTextTracker(doc_data)

        success, start, end, msg = tracker.search_text("Goodbye", "replace")
        assert success is False
        assert "not found" in msg

    def test_search_text_before_position(self):
        """Test search with 'before' position."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("Hello World")
        tracker = VirtualTextTracker(doc_data)

        success, start, end, msg = tracker.search_text("World", "before")
        assert success is True
        assert start == 7
        assert end == 7  # Insert point is before "World"

    def test_search_text_after_position(self):
        """Test search with 'after' position."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("Hello World")
        tracker = VirtualTextTracker(doc_data)

        success, start, end, msg = tracker.search_text("Hello", "after")
        assert success is True
        assert start == 6  # Insert point is after "Hello" (index 6)
        assert end == 6

    def test_search_text_case_insensitive(self):
        """Test case-insensitive search."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("Hello World")
        tracker = VirtualTextTracker(doc_data)

        success, start, end, msg = tracker.search_text("hello", "replace", match_case=False)
        assert success is True
        assert start == 1

    def test_apply_insert_updates_virtual_text(self):
        """Test that applying an insert updates the virtual text."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("Hello World")
        tracker = VirtualTextTracker(doc_data)

        # Insert "[NEW]" after "Hello"
        tracker.apply_operation({
            "type": "insert_text",
            "index": 6,  # After "Hello"
            "text": "[NEW]"
        })

        assert "[NEW]" in tracker.text
        assert tracker.text == "Hello[NEW] World"

    def test_apply_delete_updates_virtual_text(self):
        """Test that applying a delete updates the virtual text."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("Hello World")
        tracker = VirtualTextTracker(doc_data)

        # Delete "Hello "
        tracker.apply_operation({
            "type": "delete_text",
            "start_index": 1,
            "end_index": 7
        })

        assert tracker.text == "World"

    def test_apply_replace_updates_virtual_text(self):
        """Test that applying a replace updates the virtual text."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("Hello World")
        tracker = VirtualTextTracker(doc_data)

        # Replace "World" with "Universe"
        tracker.apply_operation({
            "type": "replace_text",
            "start_index": 7,
            "end_index": 12,
            "text": "Universe"
        })

        assert tracker.text == "Hello Universe"

    def test_chained_inserts_are_searchable(self):
        """Test the key feature: text inserted by one operation can be searched by the next."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("[MARKER]")
        tracker = VirtualTextTracker(doc_data)

        # First: insert [OP1] after [MARKER]
        tracker.apply_operation({
            "type": "insert_text",
            "index": 9,  # After [MARKER]
            "text": "[OP1]"
        })

        # Now search for [OP1] - should find it in the virtual text
        success, start, end, msg = tracker.search_text("[OP1]", "after")
        assert success is True, f"Failed to find [OP1]: {msg}"

        # Apply second insert after [OP1]
        tracker.apply_operation({
            "type": "insert_text",
            "index": end,
            "text": "[OP2]"
        })

        # Should also be able to find [OP2]
        success2, start2, end2, msg2 = tracker.search_text("[OP2]", "replace")
        assert success2 is True, f"Failed to find [OP2]: {msg2}"

    def test_multiple_occurrences(self):
        """Test finding specific occurrences of repeated text."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("apple banana apple cherry apple")
        tracker = VirtualTextTracker(doc_data)

        # First occurrence
        success, start, end, msg = tracker.search_text("apple", "replace", occurrence=1)
        assert success is True
        assert start == 1

        # Second occurrence
        success, start, end, msg = tracker.search_text("apple", "replace", occurrence=2)
        assert success is True
        assert start == 14  # After "apple banana " (1-based: 1 + len("apple banana "))

        # Last occurrence (occurrence=-1)
        success, start, end, msg = tracker.search_text("apple", "replace", occurrence=-1)
        assert success is True
        assert start == 27  # Last "apple" (1 + len("apple banana apple cherry "))

    def test_occurrence_out_of_bounds(self):
        """Test requesting an occurrence that doesn't exist."""
        from gdocs.managers.batch_operation_manager import VirtualTextTracker

        doc_data = self._make_doc_data("apple banana")
        tracker = VirtualTextTracker(doc_data)

        success, start, end, msg = tracker.search_text("apple", "replace", occurrence=2)
        assert success is False
        assert "1 occurrence" in msg


class TestChainedBatchOperations:
    """Integration tests for chained batch operations (the main bug fix)."""

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
                                            "content": "[MARKER] End of document.",
                                        },
                                        "startIndex": 1,
                                        "endIndex": 26,
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
    async def test_chained_search_operations(self, mock_service):
        """
        Test the main bug fix: chained operations where later ops search for
        text inserted by earlier ops.

        This was the original bug from google_workspace_mcp-9f5f:
        Operation 2 searches for '[OP1]' which is inserted by operation 1.
        Without the fix, it would fail with 'Text [OP1] not found in document'.
        """
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert", "search": "[MARKER]", "position": "after", "text": "[OP1]"},
            {"type": "insert", "search": "[OP1]", "position": "after", "text": "[OP2]"},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True, f"Expected success but got: {result.message}"
        assert result.operations_completed == 2
        assert len(result.results) == 2
        assert all(r.success for r in result.results)

    @pytest.mark.asyncio
    async def test_three_level_chain(self, mock_service):
        """Test three levels of chained insertions."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert", "search": "[MARKER]", "position": "after", "text": "[A]"},
            {"type": "insert", "search": "[A]", "position": "after", "text": "[B]"},
            {"type": "insert", "search": "[B]", "position": "after", "text": "[C]"},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 3

    @pytest.mark.asyncio
    async def test_chain_with_replace(self, mock_service):
        """Test chained operations including a replace."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert", "search": "[MARKER]", "position": "after", "text": "[TEMP]"},
            {"type": "replace", "search": "[TEMP]", "position": "replace", "text": "[FINAL]"},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 2

    @pytest.mark.asyncio
    async def test_chain_preview_mode(self, mock_service):
        """Test that chained operations work in preview mode too."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert", "search": "[MARKER]", "position": "after", "text": "[OP1]"},
            {"type": "insert", "search": "[OP1]", "position": "after", "text": "[OP2]"},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
            preview_only=True,
        )

        assert result.success is True
        assert result.preview is True
        assert len(result.results) == 2
        # Both operations should have resolved successfully
        assert all(r.resolved_index is not None for r in result.results)


class TestLocationBasedPositioning:
    """Tests for location-based positioning in batch operations."""

    @pytest.fixture
    def mock_service(self):
        """Create a mock Google Docs service."""
        service = MagicMock()
        mock_doc = {
            "body": {
                "content": [
                    {"startIndex": 0, "endIndex": 1, "sectionBreak": {}},
                    {
                        "startIndex": 1,
                        "endIndex": 50,
                        "paragraph": {
                            "elements": [
                                {
                                    "startIndex": 1,
                                    "endIndex": 50,
                                    "textRun": {"content": "Hello World. This is test content.\n"},
                                }
                            ]
                        },
                    },
                ]
            }
        }
        service.documents().get().execute.return_value = mock_doc
        service.documents().batchUpdate().execute.return_value = {"replies": []}
        return service

    @pytest.mark.asyncio
    async def test_location_end_resolves_to_last_index(self, mock_service):
        """Test that location='end' resolves to the last valid document index."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert", "location": "end", "text": "[APPENDED]"}
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 1
        # location='end' should resolve to total_length - 1 = 50 - 1 = 49
        assert result.results[0].resolved_index == 49

    @pytest.mark.asyncio
    async def test_location_start_resolves_to_index_1(self, mock_service):
        """Test that location='start' resolves to index 1 (after section break)."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert", "location": "start", "text": "[PREPENDED]"}
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 1
        # location='start' should resolve to index 1
        assert result.results[0].resolved_index == 1

    @pytest.mark.asyncio
    async def test_invalid_location_returns_error(self, mock_service):
        """Test that invalid location values return an error."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert", "location": "middle", "text": "[INVALID]"}
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is False
        assert "Invalid location" in result.results[0].error

    @pytest.mark.asyncio
    async def test_location_with_preview_mode(self, mock_service):
        """Test that location-based operations work in preview mode."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert", "location": "end", "text": "[TEST]"}
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
            preview_only=True,
        )

        assert result.success is True
        assert result.preview is True
        assert result.results[0].resolved_index == 49

    @pytest.mark.asyncio
    async def test_location_with_auto_adjust(self, mock_service):
        """Test that multiple location operations with auto-adjust work correctly."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert", "location": "end", "text": "[FIRST]"},  # 7 chars
            {"type": "insert", "location": "end", "text": "[SECOND]"},  # Should adjust
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 2
        # First insert at 49
        assert result.results[0].resolved_index == 49
        # Second insert should be at 49 + 7 (length of "[FIRST]") = 56
        assert result.results[1].resolved_index == 56

    @pytest.mark.asyncio
    async def test_location_insert_table(self, mock_service):
        """Test that location works with insert_table operations."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert_table", "location": "end", "rows": 2, "columns": 2}
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 1
        assert result.results[0].resolved_index == 49

    @pytest.mark.asyncio
    async def test_location_insert_page_break(self, mock_service):
        """Test that location works with insert_page_break operations."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {"type": "insert_page_break", "location": "end"}
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 1
        assert result.results[0].resolved_index == 49


class TestRecentInsertPreference:
    """Tests for the recent insert preference feature.

    When inserting text and then searching for it in the same batch, the search
    should prefer the recently inserted text over earlier occurrences in the document.
    This fixes the bug where batch_edit_doc format operation formatted wrong location.
    """

    def _create_mock_doc_with_existing_text(self, text: str, start_index: int = 1):
        """Create mock document data with the given text."""
        return {
            "body": {
                "content": [
                    {
                        "sectionBreak": {"sectionStyle": {}},
                        "startIndex": 0,
                        "endIndex": 1,
                    },
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {"content": text},
                                    "startIndex": start_index,
                                    "endIndex": start_index + len(text),
                                }
                            ]
                        },
                    },
                ]
            }
        }

    def test_virtual_tracker_tracks_recent_inserts(self):
        """Test that VirtualTextTracker tracks recently inserted text."""
        doc_data = self._create_mock_doc_with_existing_text("Hello World\n")
        tracker = VirtualTextTracker(doc_data)

        # Initially no recent inserts
        assert tracker._recent_inserts == []

        # Apply an insert operation
        tracker.apply_operation({
            "type": "insert_text",
            "index": 12,  # After "Hello World"
            "text": "[MARKER]"
        })

        # Should have tracked the insert
        assert len(tracker._recent_inserts) == 1
        assert tracker._recent_inserts[0] == ("[MARKER]", 12, 20)

    def test_search_prefers_recent_insert_over_existing(self):
        """Test that search prefers recently inserted text over existing occurrences."""
        # Document already contains "TEST" at the beginning
        doc_data = self._create_mock_doc_with_existing_text("TEST existing content\n")
        tracker = VirtualTextTracker(doc_data)

        # Insert "TEST" again at a later position
        tracker.apply_operation({
            "type": "insert_text",
            "index": 22,
            "text": "NEW TEST"
        })

        # Search should find "TEST" in the recently inserted text, not the original
        success, start, end, msg = tracker.search_text("TEST", "replace")

        assert success is True
        # Should find the TEST in "NEW TEST" (at position 26, not at 1)
        assert start == 26  # "NEW " is 4 chars, so TEST starts at 22+4=26
        assert end == 30
        assert "recently inserted" in msg

    def test_search_falls_back_to_existing_when_not_in_recent(self):
        """Test that search falls back to existing text when search text not in recent inserts."""
        # Document contains "EXISTING" text
        doc_data = self._create_mock_doc_with_existing_text("EXISTING content here\n")
        tracker = VirtualTextTracker(doc_data)

        # Insert different text
        tracker.apply_operation({
            "type": "insert_text",
            "index": 22,
            "text": "[MARKER]"
        })

        # Search for "EXISTING" - should find in original document
        success, start, end, msg = tracker.search_text("EXISTING", "replace")

        assert success is True
        assert start == 1  # Found at original position
        assert "recently inserted" not in msg

    def test_search_respects_occurrence_parameter(self):
        """Test that explicit occurrence parameter bypasses recent insert preference."""
        doc_data = self._create_mock_doc_with_existing_text("TEST first TEST second\n")
        tracker = VirtualTextTracker(doc_data)

        # Insert another "TEST" at the end
        tracker.apply_operation({
            "type": "insert_text",
            "index": 24,
            "text": " TEST third"
        })

        # Search with occurrence=2 should find the second occurrence, not recent insert
        success, start, end, msg = tracker.search_text("TEST", "replace", occurrence=2)

        assert success is True
        assert start == 12  # Second "TEST" in "TEST first TEST second"
        assert "recently inserted" not in msg

    def test_search_recent_insert_case_insensitive(self):
        """Test that recent insert search respects case sensitivity."""
        doc_data = self._create_mock_doc_with_existing_text("test existing\n")
        tracker = VirtualTextTracker(doc_data)

        tracker.apply_operation({
            "type": "insert_text",
            "index": 15,
            "text": "NEW TEST"
        })

        # Case-insensitive search should find recent insert
        success, start, end, msg = tracker.search_text("test", "replace", match_case=False)

        assert success is True
        # Should find "TEST" in recently inserted "NEW TEST"
        assert start == 19  # "NEW " is 4 chars
        assert "recently inserted" in msg

    @pytest.fixture
    def mock_service_with_existing_marker(self):
        """Create mock service with document containing existing 'BATCH EDIT TEST' text."""
        service = MagicMock()
        mock_doc = {
            "body": {
                "content": [
                    {
                        "sectionBreak": {"sectionStyle": {}},
                        "startIndex": 0,
                        "endIndex": 1,
                    },
                    {
                        "paragraph": {
                            "elements": [
                                {
                                    "textRun": {
                                        "content": "Some content with BATCH EDIT TEST already here.\n",
                                    },
                                    "startIndex": 1,
                                    "endIndex": 49,
                                }
                            ]
                        },
                        "startIndex": 1,
                        "endIndex": 49,
                    },
                ]
            },
            "title": "Test Doc"
        }
        service.documents().get().execute.return_value = mock_doc
        service.documents().batchUpdate().execute.return_value = {"replies": []}
        return service

    @pytest.mark.asyncio
    async def test_insert_then_format_targets_inserted_text(self, mock_service_with_existing_marker):
        """Test the main bug fix: insert text then format should target inserted text.

        This reproduces the bug from google_workspace_mcp-0aac:
        When using batch_edit_doc with a format operation that searches for text,
        the format should apply to the NEWLY inserted text, not an earlier occurrence.
        """
        manager = BatchOperationManager(mock_service_with_existing_marker)

        operations = [
            {"type": "insert", "location": "end", "text": "\n\n=== BATCH EDIT TEST ===\n"},
            {"type": "format", "search": "BATCH EDIT TEST", "bold": True},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 2

        # The format operation should target the recently inserted text
        format_result = result.results[1]
        # The inserted text starts at index 48 (end of doc), + 4 for "\n\n=== "
        # So "BATCH EDIT TEST" starts at 52
        assert format_result.resolved_index >= 48  # Should be in the inserted text, not at 18

    @pytest.mark.asyncio
    async def test_insert_then_format_preview_mode(self, mock_service_with_existing_marker):
        """Test insert+format in preview mode shows correct targeting."""
        manager = BatchOperationManager(mock_service_with_existing_marker)

        operations = [
            {"type": "insert", "location": "end", "text": "[NEW MARKER]"},
            {"type": "format", "search": "MARKER", "bold": True},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
            preview_only=True,
        )

        assert result.success is True
        assert result.preview is True

        # Format should target the "MARKER" in the newly inserted "[NEW MARKER]"
        format_result = result.results[1]
        # Inserted text starts at 48, "[NEW " is 5 chars, so "MARKER" at 53
        assert format_result.resolved_index >= 48

    @pytest.mark.asyncio
    async def test_multiple_inserts_then_format(self, mock_service_with_existing_marker):
        """Test that format finds text in the most recent insert."""
        manager = BatchOperationManager(mock_service_with_existing_marker)

        operations = [
            {"type": "insert", "location": "end", "text": "First MARKER insert.\n"},
            {"type": "insert", "location": "end", "text": "Second MARKER insert.\n"},
            {"type": "format", "search": "MARKER", "bold": True},
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 3

        # Format should target "MARKER" in the most recent (second) insert
        format_result = result.results[2]
        # Second insert starts after first (48 + 21 = 69), "Second " is 7 chars, so MARKER at 76
        assert format_result.resolved_index >= 69


class TestBatchOperationInsertValidation:
    """Tests for validation of insert operations in batch_edit_doc."""

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
                                            "content": "Test document content.",
                                        },
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
        return service

    @pytest.mark.asyncio
    async def test_insert_without_text_fails_validation(self, mock_service):
        """Test that insert operation without 'text' field fails with clear error."""
        manager = BatchOperationManager(mock_service)

        # Insert with position but no text
        operations = [{"type": "insert", "index": 1}]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is False
        assert len(result.results) == 1
        assert result.results[0].success is False
        assert "text" in result.results[0].error.lower()

    @pytest.mark.asyncio
    async def test_insert_without_position_fails_validation(self, mock_service):
        """Test that insert operation without any positioning fails with clear error."""
        manager = BatchOperationManager(mock_service)

        # Insert with text but no position (no index, search, location, or range_spec)
        operations = [{"type": "insert", "text": "Hello World"}]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is False
        assert len(result.results) == 1
        assert result.results[0].success is False
        assert "position" in result.results[0].error.lower() or "index" in result.results[0].error.lower()

    @pytest.mark.asyncio
    async def test_insert_without_text_or_position_fails_validation(self, mock_service):
        """Test that insert operation without text AND without position fails."""
        manager = BatchOperationManager(mock_service)

        # Insert with only type - no text, no position
        operations = [{"type": "insert"}]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is False
        assert len(result.results) == 1
        assert result.results[0].success is False
        # Should fail on first validation (missing text)
        assert "text" in result.results[0].error.lower()

    @pytest.mark.asyncio
    async def test_insert_with_index_and_text_succeeds(self, mock_service):
        """Test that valid insert with index and text succeeds."""
        manager = BatchOperationManager(mock_service)

        operations = [{"type": "insert", "index": 1, "text": "Hello"}]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 1

    @pytest.mark.asyncio
    async def test_insert_with_location_and_text_succeeds(self, mock_service):
        """Test that valid insert with location and text succeeds."""
        manager = BatchOperationManager(mock_service)

        operations = [{"type": "insert", "location": "end", "text": "Appended text"}]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 1

    @pytest.mark.asyncio
    async def test_insert_with_search_and_text_succeeds(self, mock_service):
        """Test that valid insert with search positioning and text succeeds."""
        manager = BatchOperationManager(mock_service)

        operations = [
            {
                "type": "insert",
                "search": "document",
                "position": "after",
                "text": " (modified)",
            }
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is True
        assert result.operations_completed == 1

    @pytest.mark.asyncio
    async def test_insert_text_alias_also_validates(self, mock_service):
        """Test that 'insert_text' type also validates correctly."""
        manager = BatchOperationManager(mock_service)

        # Using canonical 'insert_text' instead of alias 'insert'
        operations = [{"type": "insert_text"}]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        assert result.success is False
        assert result.results[0].success is False
        assert "text" in result.results[0].error.lower()

    @pytest.mark.asyncio
    async def test_batch_with_one_invalid_insert_fails_all(self, mock_service):
        """Test that batch with one invalid insert operation fails before API call."""
        manager = BatchOperationManager(mock_service)

        # First valid, second invalid
        operations = [
            {"type": "insert", "index": 1, "text": "Valid"},
            {"type": "insert"},  # Invalid - no text or position
        ]

        result = await manager.execute_batch_with_search(
            "test-doc-id",
            operations,
            auto_adjust_positions=True,
        )

        # Should fail because resolution fails for operation 1
        assert result.success is False
        # API should NOT have been called (batchUpdate)
        mock_service.documents.return_value.batchUpdate.return_value.execute.assert_not_called()
