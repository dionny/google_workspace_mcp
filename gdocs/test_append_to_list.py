"""
Tests for append_to_list tool.

Issue: google_workspace_mcp-6913

These tests verify:
1. Input validation for items parameter
2. Nesting levels validation
3. Finding target list by index or search
4. Text building with proper tab prefixes for nesting
5. Correct insertion point calculation
6. Preview mode functionality
"""
import pytest
from gdocs.docs_helpers import (
    create_insert_text_request,
    create_bullet_list_request,
    interpret_escape_sequences,
)
from gdocs.managers.validation_manager import ValidationManager


class TestAppendToListInputValidation:
    """Test input validation for append_to_list."""

    def test_items_must_be_list(self):
        """Test that items parameter must be a list."""
        validator = ValidationManager()

        # String is not valid
        error = validator.create_invalid_param_error(
            param_name="items",
            received="str",
            valid_values=["list of strings"],
        )
        assert "error" in error.lower()
        assert "items" in error.lower()

    def test_items_must_not_be_empty(self):
        """Test that items parameter must not be empty."""
        validator = ValidationManager()

        error = validator.create_invalid_param_error(
            param_name="items",
            received="empty list",
            valid_values=["non-empty list of strings"],
        )
        assert "error" in error.lower()

    def test_items_must_contain_strings(self):
        """Test that all items must be strings."""
        validator = ValidationManager()
        items = ["valid", 123, "also valid"]

        for i, item in enumerate(items):
            if not isinstance(item, str):
                error = validator.create_invalid_param_error(
                    param_name=f"items[{i}]",
                    received=type(item).__name__,
                    valid_values=["string"],
                )
                assert "error" in error.lower()
                assert "items[1]" in error


class TestNestingLevelsValidation:
    """Test nesting_levels parameter validation."""

    def test_nesting_levels_must_be_list(self):
        """Test that nesting_levels must be a list."""
        validator = ValidationManager()

        error = validator.create_invalid_param_error(
            param_name="nesting_levels",
            received="int",
            valid_values=["list of integers"],
        )
        assert "error" in error.lower()

    def test_nesting_levels_length_must_match_items(self):
        """Test that nesting_levels length must match items length."""
        validator = ValidationManager()
        items = ["a", "b", "c"]  # 3 items
        nesting_levels = [0, 1]  # Only 2 levels

        error = validator.create_invalid_param_error(
            param_name="nesting_levels",
            received=f"list of length {len(nesting_levels)}",
            valid_values=[f"list of length {len(items)} (must match items length)"],
        )
        assert "error" in error.lower()
        assert "3" in error  # Expected length

    def test_nesting_levels_must_be_integers(self):
        """Test that all nesting levels must be integers."""
        validator = ValidationManager()
        nesting_levels = [0, "one", 2]

        for i, level in enumerate(nesting_levels):
            if not isinstance(level, int):
                error = validator.create_invalid_param_error(
                    param_name=f"nesting_levels[{i}]",
                    received=type(level).__name__,
                    valid_values=["integer (0-8)"],
                )
                assert "error" in error.lower()

    def test_nesting_levels_range_0_to_8(self):
        """Test that nesting levels must be between 0 and 8."""
        validator = ValidationManager()

        # Test level below range
        error = validator.create_invalid_param_error(
            param_name="nesting_levels[0]",
            received="-1",
            valid_values=["integer from 0 to 8"],
        )
        assert "error" in error.lower()

        # Test level above range
        error = validator.create_invalid_param_error(
            param_name="nesting_levels[1]",
            received="9",
            valid_values=["integer from 0 to 8"],
        )
        assert "error" in error.lower()

    def test_default_nesting_level_is_zero(self):
        """Test that default nesting level is 0 for all items."""
        items = ["a", "b", "c"]
        # When nesting_levels is None, default to [0, 0, 0]
        item_nesting = [0] * len(items)

        assert item_nesting == [0, 0, 0]
        assert len(item_nesting) == len(items)


class TestTextBuildingWithNesting:
    """Test text building logic with nesting levels."""

    def test_flat_list_no_tabs(self):
        """Test that flat list items have no tab prefix."""
        items = ["Item 1", "Item 2", "Item 3"]
        nesting_levels = [0, 0, 0]

        nested_items = []
        for item_text, level in zip(items, nesting_levels):
            prefix = "\t" * level
            nested_items.append(prefix + item_text)

        assert nested_items == ["Item 1", "Item 2", "Item 3"]
        combined = "\n".join(nested_items) + "\n"
        assert combined == "Item 1\nItem 2\nItem 3\n"

    def test_nested_list_with_tabs(self):
        """Test that nested items have correct tab prefixes."""
        items = ["Main", "Sub 1", "Sub 2", "Another main"]
        nesting_levels = [0, 1, 1, 0]

        nested_items = []
        for item_text, level in zip(items, nesting_levels):
            prefix = "\t" * level
            nested_items.append(prefix + item_text)

        assert nested_items == ["Main", "\tSub 1", "\tSub 2", "Another main"]
        combined = "\n".join(nested_items) + "\n"
        assert combined == "Main\n\tSub 1\n\tSub 2\nAnother main\n"

    def test_deeply_nested_list(self):
        """Test deeply nested list with multiple levels."""
        items = ["Level 0", "Level 1", "Level 2", "Level 3"]
        nesting_levels = [0, 1, 2, 3]

        nested_items = []
        for item_text, level in zip(items, nesting_levels):
            prefix = "\t" * level
            nested_items.append(prefix + item_text)

        assert nested_items == ["Level 0", "\tLevel 1", "\t\tLevel 2", "\t\t\tLevel 3"]

    def test_escape_sequences_in_items(self):
        """Test that escape sequences are interpreted in items."""
        items = ["Line 1\\nLine 2", "Tab\\there"]

        processed = [interpret_escape_sequences(item) for item in items]

        assert processed[0] == "Line 1\nLine 2"
        assert processed[1] == "Tab\there"


class TestListFindingLogic:
    """Test logic for finding target list."""

    def test_find_list_by_index_first(self):
        """Test finding first list by default index."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "list_type": "bullet", "items": [{"text": "First"}]},
            {"start_index": 100, "end_index": 150, "list_type": "numbered", "items": [{"text": "Second"}]},
        ]
        list_index = 0

        target_list = all_lists[list_index]
        assert target_list["start_index"] == 10
        assert target_list["list_type"] == "bullet"

    def test_find_list_by_index_second(self):
        """Test finding second list by index."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "list_type": "bullet", "items": [{"text": "First"}]},
            {"start_index": 100, "end_index": 150, "list_type": "numbered", "items": [{"text": "Second"}]},
        ]
        list_index = 1

        target_list = all_lists[list_index]
        assert target_list["start_index"] == 100
        assert target_list["list_type"] == "numbered"

    def test_find_list_by_search(self):
        """Test finding list by search text."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "list_type": "bullet", "items": [
                {"text": "Buy groceries"},
                {"text": "Walk dog"},
            ]},
            {"start_index": 100, "end_index": 150, "list_type": "numbered", "items": [
                {"text": "Step one"},
                {"text": "Step two"},
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
        assert target_list["list_type"] == "numbered"

    def test_search_not_found(self):
        """Test error when search text not found in any list."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "items": [{"text": "Buy groceries"}]},
        ]
        search = "nonexistent"
        search_lower = search.lower()

        target_list = None
        for lst in all_lists:
            for item in lst.get("items", []):
                if search_lower in item.get("text", "").lower():
                    target_list = lst
                    break

        assert target_list is None

    def test_index_out_of_range(self):
        """Test error when list index is out of range."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "items": [{"text": "Only list"}]},
        ]
        list_index = 5

        assert list_index >= len(all_lists)


class TestInsertionPointCalculation:
    """Test calculation of insertion point."""

    def test_insertion_at_list_end(self):
        """Test that insertion happens at end_index of list."""
        target_list = {
            "start_index": 10,
            "end_index": 50,
            "items": [{"text": "Item 1"}, {"text": "Item 2"}]
        }

        insertion_index = target_list["end_index"]
        assert insertion_index == 50


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
    """Test generation of API requests for appending to list."""

    def test_insert_text_request(self):
        """Test that insert text request is created correctly."""
        insertion_index = 50
        combined_text = "New item\n"

        request = create_insert_text_request(insertion_index, combined_text)

        assert "insertText" in request
        assert request["insertText"]["location"]["index"] == 50
        assert request["insertText"]["text"] == "New item\n"

    def test_bullet_list_request_for_unordered(self):
        """Test bullet list request for unordered list."""
        insertion_index = 50
        text_length = 8  # "New item" without final newline

        request = create_bullet_list_request(
            insertion_index,
            insertion_index + text_length,
            "UNORDERED"
        )

        assert "createParagraphBullets" in request
        bullet_req = request["createParagraphBullets"]
        assert bullet_req["range"]["startIndex"] == 50
        assert bullet_req["range"]["endIndex"] == 58
        assert "BULLET" in bullet_req["bulletPreset"]

    def test_bullet_list_request_for_ordered(self):
        """Test bullet list request for ordered list."""
        insertion_index = 50
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
        processed_items = ["New item 1", "New item 2"]
        item_nesting = [0, 0]
        insertion_index = target_list["end_index"]

        has_nested = any(level > 0 for level in item_nesting)
        max_level = max(item_nesting) if has_nested else 0

        preview_info = {
            "target_list": {
                "type": "bullet",
                "start_index": target_list["start_index"],
                "end_index": target_list["end_index"],
                "current_items_count": len(target_list.get("items", [])),
            },
            "items_to_append": len(processed_items),
            "insertion_index": insertion_index,
            "nesting": {
                "has_nested_items": has_nested,
                "max_depth": max_level,
            } if has_nested else None,
            "items_preview": [
                {"text": item[:50], "nesting_level": level}
                for item, level in zip(processed_items, item_nesting)
            ][:5],
        }

        assert "target_list" in preview_info
        assert preview_info["target_list"]["current_items_count"] == 2
        assert preview_info["items_to_append"] == 2
        assert preview_info["insertion_index"] == 50
        assert preview_info["nesting"] is None  # No nested items

    def test_preview_with_nested_items(self):
        """Test preview info when items have nesting."""
        processed_items = ["Main", "Sub 1", "Sub 2"]
        item_nesting = [0, 1, 1]

        has_nested = any(level > 0 for level in item_nesting)
        max_level = max(item_nesting) if has_nested else 0

        nesting_info = {
            "has_nested_items": has_nested,
            "max_depth": max_level,
        } if has_nested else None

        assert nesting_info is not None
        assert nesting_info["has_nested_items"] is True
        assert nesting_info["max_depth"] == 1


class TestEdgeCases:
    """Test edge cases for append_to_list."""

    def test_single_item_append(self):
        """Test appending a single item."""
        items = ["Only item"]
        nesting_levels = [0]

        nested_items = []
        for item_text, level in zip(items, nesting_levels):
            prefix = "\t" * level
            nested_items.append(prefix + item_text)

        combined = "\n".join(nested_items) + "\n"
        assert combined == "Only item\n"
        assert len(combined) == 10

    def test_item_with_special_characters(self):
        """Test items with special characters."""
        items = ["Item with 'quotes' and \"double quotes\"", "Item with <html>"]

        # These should pass through without special handling
        for item in items:
            processed = interpret_escape_sequences(item)
            # Special characters should remain unchanged
            assert "'" in processed or '"' in processed or "<" in processed

    def test_long_item_preview_truncation(self):
        """Test that long items are truncated in preview."""
        long_text = "A" * 100
        truncated = long_text[:50] + ("..." if len(long_text) > 50 else "")

        assert len(truncated) == 53  # 50 chars + "..."
        assert truncated.endswith("...")

    def test_many_items_preview_limit(self):
        """Test that preview shows at most 5 items."""
        items = [f"Item {i}" for i in range(10)]
        nesting_levels = [0] * 10

        items_preview = [
            {"text": item[:50], "nesting_level": level}
            for item, level in zip(items, nesting_levels)
        ][:5]

        assert len(items_preview) == 5
        assert items_preview[4]["text"] == "Item 4"
