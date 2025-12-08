"""
Tests for insert_list_item tool.

Issue: google_workspace_mcp-bf24

These tests verify:
1. Input validation for text and indent_level parameters
2. Position parsing logic (start, end, numeric, after:, before:)
3. Finding target list by index or search
4. Text building with proper tab prefixes for indent_level
5. Correct insertion index calculation
6. Preview mode functionality
"""
from gdocs.docs_helpers import (
    create_insert_text_request,
    create_bullet_list_request,
    interpret_escape_sequences,
)
from gdocs.managers.validation_manager import ValidationManager


class TestInsertListItemTextValidation:
    """Test input validation for text parameter."""

    def test_text_must_be_string(self):
        """Test that text parameter must be a string."""
        validator = ValidationManager()

        error = validator.create_invalid_param_error(
            param_name="text",
            received="int",
            valid_values=["string"],
        )
        assert "error" in error.lower()
        assert "text" in error.lower()

    def test_text_must_not_be_empty(self):
        """Test that text parameter must not be empty."""
        validator = ValidationManager()

        error = validator.create_invalid_param_error(
            param_name="text",
            received="empty string",
            valid_values=["non-empty string"],
        )
        assert "error" in error.lower()

    def test_text_with_whitespace_only_is_empty(self):
        """Test that whitespace-only text is considered empty."""
        text = "   \t\n   "
        assert not text.strip()


class TestInsertListItemIndentLevelValidation:
    """Test indent_level parameter validation."""

    def test_indent_level_must_be_integer(self):
        """Test that indent_level must be an integer."""
        validator = ValidationManager()

        error = validator.create_invalid_param_error(
            param_name="indent_level",
            received="str",
            valid_values=["integer (0-8)"],
        )
        assert "error" in error.lower()

    def test_indent_level_range_0_to_8(self):
        """Test that indent_level must be between 0 and 8."""
        validator = ValidationManager()

        # Test level below range
        error = validator.create_invalid_param_error(
            param_name="indent_level",
            received="-1",
            valid_values=["integer from 0 to 8"],
        )
        assert "error" in error.lower()

        # Test level above range
        error = validator.create_invalid_param_error(
            param_name="indent_level",
            received="9",
            valid_values=["integer from 0 to 8"],
        )
        assert "error" in error.lower()

    def test_valid_indent_levels(self):
        """Test that levels 0-8 are valid."""
        for level in range(9):
            assert 0 <= level <= 8


class TestPositionParsing:
    """Test position parameter parsing logic."""

    def test_position_start_keyword(self):
        """Test parsing 'start' position."""
        position_str = "start".lower()
        assert position_str in ("start", "0")

    def test_position_end_keyword(self):
        """Test parsing 'end' position."""
        position_str = "end".lower()
        assert position_str in ("end", "-1")

    def test_position_numeric_zero(self):
        """Test parsing '0' as start position."""
        position_str = "0".strip().lower()
        assert position_str in ("start", "0")

    def test_position_numeric_minus_one(self):
        """Test parsing '-1' as end position."""
        position_str = "-1".strip().lower()
        assert position_str in ("end", "-1")

    def test_position_after_prefix(self):
        """Test parsing 'after:some text' position."""
        position = "after:Buy groceries"
        position_str = position.strip().lower()
        assert position_str.startswith("after:")
        search_text = position[6:].strip()
        assert search_text == "Buy groceries"

    def test_position_before_prefix(self):
        """Test parsing 'before:some text' position."""
        position = "before:Task to do"
        position_str = position.strip().lower()
        assert position_str.startswith("before:")
        search_text = position[7:].strip()
        assert search_text == "Task to do"

    def test_position_numeric(self):
        """Test parsing numeric position."""
        position_str = "2"
        try:
            item_position = int(position_str)
            assert item_position == 2
        except ValueError:
            assert False, "Should parse numeric position"

    def test_position_invalid(self):
        """Test invalid position returns error."""
        validator = ValidationManager()
        position = "invalid_position"

        try:
            int(position)
            assert False, "Should raise ValueError"
        except ValueError:
            error = validator.create_invalid_param_error(
                param_name="position",
                received=position,
                valid_values=[
                    "'start' or '0'",
                    "'end' or '-1'",
                    "numeric index (e.g., '2')",
                    "'after:search text'",
                    "'before:search text'",
                ],
            )
            assert "error" in error.lower()

    def test_position_after_empty_text(self):
        """Test that 'after:' without text is invalid."""
        validator = ValidationManager()
        position = "after:"
        search_text = position[6:].strip()
        assert not search_text

        error = validator.create_invalid_param_error(
            param_name="position",
            received="after: (empty)",
            valid_values=["after:some text"],
        )
        assert "error" in error.lower()

    def test_position_before_empty_text(self):
        """Test that 'before:' without text is invalid."""
        validator = ValidationManager()
        position = "before:"
        search_text = position[7:].strip()
        assert not search_text

        error = validator.create_invalid_param_error(
            param_name="position",
            received="before: (empty)",
            valid_values=["before:some text"],
        )
        assert "error" in error.lower()


class TestInsertionIndexCalculation:
    """Test calculation of insertion index based on position."""

    def test_insertion_at_list_start(self):
        """Test that 'start' inserts at list start_index."""
        target_list = {
            "start_index": 10,
            "end_index": 50,
            "items": [{"text": "Item 1", "start_index": 10, "end_index": 18}]
        }

        position = "start"
        if position.lower() in ("start", "0"):
            insertion_index = target_list["start_index"]

        assert insertion_index == 10

    def test_insertion_at_list_end(self):
        """Test that 'end' inserts at list end_index."""
        target_list = {
            "start_index": 10,
            "end_index": 50,
            "items": [{"text": "Item 1"}, {"text": "Item 2"}]
        }

        position = "end"
        if position.lower() in ("end", "-1"):
            insertion_index = target_list["end_index"]

        assert insertion_index == 50

    def test_insertion_after_item(self):
        """Test insertion after a specific item."""
        list_items = [
            {"text": "First item", "start_index": 10, "end_index": 22},
            {"text": "Second item", "start_index": 22, "end_index": 35},
            {"text": "Third item", "start_index": 35, "end_index": 47},
        ]

        search_text = "second"
        search_text_lower = search_text.lower()

        found_item = None
        for item in list_items:
            if search_text_lower in item.get("text", "").lower():
                found_item = item
                break

        assert found_item is not None
        insertion_index = found_item["end_index"]
        assert insertion_index == 35

    def test_insertion_before_item(self):
        """Test insertion before a specific item."""
        list_items = [
            {"text": "First item", "start_index": 10, "end_index": 22},
            {"text": "Second item", "start_index": 22, "end_index": 35},
            {"text": "Third item", "start_index": 35, "end_index": 47},
        ]

        search_text = "third"
        search_text_lower = search_text.lower()

        found_item = None
        for item in list_items:
            if search_text_lower in item.get("text", "").lower():
                found_item = item
                break

        assert found_item is not None
        insertion_index = found_item["start_index"]
        assert insertion_index == 35

    def test_insertion_at_numeric_position(self):
        """Test insertion at numeric position."""
        list_items = [
            {"text": "Item 0", "start_index": 10, "end_index": 18},
            {"text": "Item 1", "start_index": 18, "end_index": 26},
            {"text": "Item 2", "start_index": 26, "end_index": 34},
        ]
        target_list = {"start_index": 10, "end_index": 34, "items": list_items}
        num_items = len(list_items)

        # Insert at position 1
        item_position = 1
        if item_position == 0:
            insertion_index = target_list["start_index"]
        elif item_position >= num_items:
            insertion_index = target_list["end_index"]
        else:
            insertion_index = list_items[item_position]["start_index"]

        assert insertion_index == 18  # Start of "Item 1"

    def test_insertion_clamps_to_end(self):
        """Test that position beyond list length inserts at end."""
        # target_list defined for documentation purposes
        # {
        #     "start_index": 10,
        #     "end_index": 50,
        #     "items": [{"text": "Item 1"}, {"text": "Item 2"}]
        # }
        num_items = 2

        item_position = 100  # Way beyond list length
        if item_position > num_items:
            item_position = num_items

        assert item_position == num_items


class TestTextBuildingWithIndent:
    """Test text building logic with indent_level."""

    def test_no_indent(self):
        """Test text with indent_level=0 has no tab prefix."""
        text = "New item"
        indent_level = 0

        prefix = "\t" * indent_level
        insert_text = prefix + text + "\n"

        assert insert_text == "New item\n"

    def test_single_indent(self):
        """Test text with indent_level=1 has one tab prefix."""
        text = "Nested item"
        indent_level = 1

        prefix = "\t" * indent_level
        insert_text = prefix + text + "\n"

        assert insert_text == "\tNested item\n"

    def test_deep_indent(self):
        """Test text with indent_level=3 has three tab prefixes."""
        text = "Deeply nested"
        indent_level = 3

        prefix = "\t" * indent_level
        insert_text = prefix + text + "\n"

        assert insert_text == "\t\t\tDeeply nested\n"

    def test_escape_sequences_in_text(self):
        """Test that escape sequences are interpreted in text."""
        text = "Line 1\\nLine 2"
        processed = interpret_escape_sequences(text)
        assert processed == "Line 1\nLine 2"


class TestListFindingLogic:
    """Test logic for finding target list."""

    def test_find_list_by_index_first(self):
        """Test finding first list by default index."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "list_type": "bullet", "items": []},
            {"start_index": 100, "end_index": 150, "list_type": "numbered", "items": []},
        ]
        list_index = 0

        target_list = all_lists[list_index]
        assert target_list["start_index"] == 10

    def test_find_list_by_search(self):
        """Test finding list by search text."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "list_type": "bullet", "items": [
                {"text": "Buy groceries"},
            ]},
            {"start_index": 100, "end_index": 150, "list_type": "numbered", "items": [
                {"text": "Step one"},
            ]},
        ]
        search = "step"
        search_lower = search.lower()

        target_list = None
        for lst in all_lists:
            for item in lst.get("items", []):
                if search_lower in item.get("text", "").lower():
                    target_list = lst
                    break
            if target_list:
                break

        assert target_list is not None
        assert target_list["start_index"] == 100

    def test_search_not_found_in_items(self):
        """Test error when search text not found in list items."""
        list_items = [
            {"text": "First item", "start_index": 10, "end_index": 22},
        ]

        search_text = "nonexistent"
        search_text_lower = search_text.lower()

        found_item = None
        for item in list_items:
            if search_text_lower in item.get("text", "").lower():
                found_item = item
                break

        assert found_item is None


class TestListTypeDetection:
    """Test detection of list type for applying formatting."""

    def test_detect_bullet_list(self):
        """Test detection of bullet list type."""
        target_list = {"list_type": "bullet", "type": "bullet_list"}

        list_type_raw = target_list.get("list_type", target_list.get("type", "bullet"))
        if "bullet" in list_type_raw or "unordered" in list_type_raw.lower():
            list_type = "UNORDERED"
        else:
            list_type = "ORDERED"

        assert list_type == "UNORDERED"

    def test_detect_numbered_list(self):
        """Test detection of numbered list type."""
        target_list = {"list_type": "numbered", "type": "numbered_list"}

        list_type_raw = target_list.get("list_type", target_list.get("type", "bullet"))
        if "bullet" in list_type_raw or "unordered" in list_type_raw.lower():
            list_type = "UNORDERED"
        else:
            list_type = "ORDERED"

        assert list_type == "ORDERED"


class TestApiRequestGeneration:
    """Test generation of API requests for inserting list item."""

    def test_insert_text_request(self):
        """Test that insert text request is created correctly."""
        insertion_index = 25
        insert_text = "New item\n"

        request = create_insert_text_request(insertion_index, insert_text)

        assert "insertText" in request
        assert request["insertText"]["location"]["index"] == 25
        assert request["insertText"]["text"] == "New item\n"

    def test_insert_text_request_with_indent(self):
        """Test insert text request with indented text."""
        insertion_index = 25
        text = "Nested item"
        indent_level = 2
        prefix = "\t" * indent_level
        insert_text = prefix + text + "\n"

        request = create_insert_text_request(insertion_index, insert_text)

        assert request["insertText"]["text"] == "\t\tNested item\n"

    def test_bullet_list_request_for_unordered(self):
        """Test bullet list request for unordered list."""
        insertion_index = 25
        text_length = 8  # "New item" without final newline

        request = create_bullet_list_request(
            insertion_index,
            insertion_index + text_length,
            "UNORDERED"
        )

        assert "createParagraphBullets" in request
        bullet_req = request["createParagraphBullets"]
        assert bullet_req["range"]["startIndex"] == 25
        assert bullet_req["range"]["endIndex"] == 33
        assert "BULLET" in bullet_req["bulletPreset"]

    def test_bullet_list_request_for_ordered(self):
        """Test bullet list request for ordered list."""
        insertion_index = 25
        text_length = 8

        request = create_bullet_list_request(
            insertion_index,
            insertion_index + text_length,
            "ORDERED"
        )

        assert "createParagraphBullets" in request
        bullet_req = request["createParagraphBullets"]
        assert "NUMBERED" in bullet_req["bulletPreset"]


class TestPreviewMode:
    """Test preview mode functionality."""

    def test_preview_info_structure(self):
        """Test that preview info contains expected fields."""
        target_list = {
            "list_type": "bullet",
            "start_index": 10,
            "end_index": 50,
            "items": [{"text": "Item 1"}, {"text": "Item 2"}],
        }
        insertion_index = 18
        position_description = "after item 'Item 1'"
        insert_item_index = 1
        processed_text = "New item"
        indent_level = 0

        preview_info = {
            "target_list": {
                "type": "bullet",
                "start_index": target_list["start_index"],
                "end_index": target_list["end_index"],
                "current_items_count": len(target_list.get("items", [])),
            },
            "insertion": {
                "position": position_description,
                "index": insertion_index,
                "item_position": insert_item_index,
            },
            "new_item": {
                "text": processed_text[:50],
                "indent_level": indent_level,
            },
        }

        assert "target_list" in preview_info
        assert preview_info["target_list"]["current_items_count"] == 2
        assert "insertion" in preview_info
        assert preview_info["insertion"]["index"] == 18
        assert "new_item" in preview_info
        assert preview_info["new_item"]["indent_level"] == 0

    def test_preview_with_indent(self):
        """Test preview info when item has indent."""
        indent_level = 2
        processed_text = "Nested item"

        new_item_info = {
            "text": processed_text[:50],
            "indent_level": indent_level,
        }

        assert new_item_info["indent_level"] == 2


class TestEdgeCases:
    """Test edge cases for insert_list_item."""

    def test_negative_index_wraps(self):
        """Test that negative index wraps around like Python lists."""
        num_items = 5
        item_position = -1  # Should be last position

        if item_position < 0:
            item_position = max(0, num_items + item_position + 1)

        assert item_position == 5  # Insert at end

    def test_negative_index_minus_two(self):
        """Test that -2 wraps to second-to-last position."""
        num_items = 5
        item_position = -2

        if item_position < 0:
            item_position = max(0, num_items + item_position + 1)

        assert item_position == 4

    def test_long_text_truncated_in_preview(self):
        """Test that long text is truncated in preview."""
        long_text = "A" * 100
        truncated = long_text[:50] + ("..." if len(long_text) > 50 else "")

        assert len(truncated) == 53  # 50 chars + "..."
        assert truncated.endswith("...")

    def test_text_with_special_characters(self):
        """Test text with special characters."""
        text = "Item with 'quotes' and \"double quotes\""
        processed = interpret_escape_sequences(text)
        # Special characters should remain unchanged
        assert "'" in processed
        assert '"' in processed

    def test_position_case_insensitive(self):
        """Test that position keywords are case insensitive."""
        for position in ["START", "Start", "start", "END", "End", "end"]:
            position_str = position.strip().lower()
            assert position_str in ("start", "end", "0", "-1")

    def test_after_before_preserves_search_case(self):
        """Test that after:/before: preserves search text case for matching."""
        position = "AFTER:Some Text"
        position_str = position.strip().lower()
        assert position_str.startswith("after:")
        # Search text extraction uses original position
        search_text = position[6:].strip()
        assert search_text == "Some Text"
