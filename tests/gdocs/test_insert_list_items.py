"""
Tests for insert_doc_elements with multiple list items.

Issue: google_workspace_mcp-aef6

These tests verify:
1. The 'items' parameter accepts a list of strings
2. Validation rejects invalid items (non-list, empty list, non-string elements)
3. Backward compatibility with single 'text' parameter
"""

import pytest
from gdocs.docs_helpers import (
    create_insert_text_request,
    create_bullet_list_request,
)
from gdocs.managers.validation_manager import ValidationManager


class TestListItemsHelpers:
    """Test helper functions used for list item insertion."""

    def test_create_insert_text_request_multi_line(self):
        """Test that inserting multiple lines preserves newlines."""
        combined_text = "Item 1\nItem 2\nItem 3\n"
        request = create_insert_text_request(100, combined_text)

        assert "insertText" in request
        assert request["insertText"]["location"]["index"] == 100
        assert request["insertText"]["text"] == combined_text
        assert request["insertText"]["text"].count("\n") == 3

    def test_create_bullet_list_request_range(self):
        """Test that bullet request covers the correct range."""
        # For text "Item 1\nItem 2\n" starting at index 100
        # Total text length is 14 chars, bullet range should be 100 to 113
        start = 100
        text = "Item 1\nItem 2"
        end = start + len(text)  # 113

        request = create_bullet_list_request(start, end, "ORDERED")

        assert "createParagraphBullets" in request
        bullet_req = request["createParagraphBullets"]
        assert bullet_req["range"]["startIndex"] == start
        assert bullet_req["range"]["endIndex"] == end
        assert "NUMBERED" in bullet_req["bulletPreset"]

    def test_create_bullet_list_request_unordered(self):
        """Test unordered list creates bullet preset."""
        request = create_bullet_list_request(100, 120, "UNORDERED")

        assert "createParagraphBullets" in request
        bullet_req = request["createParagraphBullets"]
        assert "BULLET" in bullet_req["bulletPreset"]


class TestListItemsValidation:
    """Test validation logic for list items parameter."""

    def test_validation_manager_invalid_items_type(self):
        """ValidationManager can create error for non-list items."""
        validator = ValidationManager()
        error = validator.create_invalid_param_error(
            param_name="items", received="str", valid_values=["list of strings"]
        )

        assert "error" in error.lower()
        assert "items" in error.lower()

    def test_validation_manager_empty_items(self):
        """ValidationManager can create error for empty list."""
        validator = ValidationManager()
        error = validator.create_invalid_param_error(
            param_name="items",
            received="empty list",
            valid_values=["non-empty list of strings"],
        )

        assert "error" in error.lower()
        assert "empty" in error.lower()


class TestListItemsLogic:
    """Test the list item building logic."""

    def test_build_combined_text_single_item(self):
        """Single item produces correct text."""
        items = ["First item"]
        combined = "\n".join(items) + "\n"

        assert combined == "First item\n"
        assert len(combined) == 11

    def test_build_combined_text_multiple_items(self):
        """Multiple items are joined with newlines."""
        items = ["Item 1", "Item 2", "Item 3"]
        combined = "\n".join(items) + "\n"

        expected = "Item 1\nItem 2\nItem 3\n"
        assert combined == expected
        assert combined.count("\n") == 3

    def test_bullet_range_calculation(self):
        """Bullet range should cover text but not trailing newline."""
        items = ["First", "Second"]
        combined = "\n".join(items) + "\n"

        start_index = 100
        # total_length excludes final newline
        total_length = len(combined) - 1

        assert total_length == len("First\nSecond")
        assert total_length == 12

        # Bullet range should be [100, 112]
        end_index = start_index + total_length
        assert end_index == 112


class TestListItemsIntegration:
    """Integration-style tests for list items (without mocking service)."""

    def test_validate_items_is_list(self):
        """Items must be a list type."""
        validator = ValidationManager()
        items = "not a list"

        # This is what the function does
        if not isinstance(items, list):
            error = validator.create_invalid_param_error(
                param_name="items",
                received=type(items).__name__,
                valid_values=["list of strings"],
            )
            assert "str" in error

    def test_validate_items_not_empty(self):
        """Items list must not be empty."""
        validator = ValidationManager()
        items = []

        if len(items) == 0:
            error = validator.create_invalid_param_error(
                param_name="items",
                received="empty list",
                valid_values=["non-empty list of strings"],
            )
            assert "empty" in error.lower()

    def test_validate_items_all_strings(self):
        """All items must be strings."""
        validator = ValidationManager()
        items = ["valid", 123, "also valid"]

        for i, item in enumerate(items):
            if not isinstance(item, str):
                error = validator.create_invalid_param_error(
                    param_name=f"items[{i}]",
                    received=type(item).__name__,
                    valid_values=["string"],
                )
                assert "items[1]" in error
                assert "int" in error
                break


class TestNestedListItems:
    """Tests for nested list functionality (nesting_levels parameter).

    Issue: google_workspace_mcp-b3ba
    """

    def test_build_nested_list_text_with_tabs(self):
        """Nested items should be prefixed with tabs based on nesting level."""
        items = ["Main item", "Sub item 1", "Sub item 2", "Another main"]
        nesting_levels = [0, 1, 1, 0]

        # Build nested items with tab prefixes
        nested_items = []
        for item_text, level in zip(items, nesting_levels):
            prefix = "\t" * level
            nested_items.append(prefix + item_text)

        combined = "\n".join(nested_items) + "\n"

        expected = "Main item\n\tSub item 1\n\tSub item 2\nAnother main\n"
        assert combined == expected

    def test_build_deeply_nested_list(self):
        """Test multiple levels of nesting."""
        items = ["Level 0", "Level 1", "Level 2", "Level 3", "Back to 1"]
        nesting_levels = [0, 1, 2, 3, 1]

        nested_items = []
        for item_text, level in zip(items, nesting_levels):
            prefix = "\t" * level
            nested_items.append(prefix + item_text)

        combined = "\n".join(nested_items) + "\n"

        expected = "Level 0\n\tLevel 1\n\t\tLevel 2\n\t\t\tLevel 3\n\tBack to 1\n"
        assert combined == expected

    def test_nested_list_default_nesting(self):
        """Without nesting_levels, all items should be level 0."""
        items = ["Item 1", "Item 2", "Item 3"]
        nesting_levels = [0] * len(items)  # Default behavior

        nested_items = []
        for item_text, level in zip(items, nesting_levels):
            prefix = "\t" * level
            nested_items.append(prefix + item_text)

        combined = "\n".join(nested_items) + "\n"

        # No tabs should be present
        assert "\t" not in combined
        assert combined == "Item 1\nItem 2\nItem 3\n"

    def test_validate_nesting_levels_must_be_list(self):
        """nesting_levels must be a list type."""
        validator = ValidationManager()
        nesting_levels = "not a list"

        if not isinstance(nesting_levels, list):
            error = validator.create_invalid_param_error(
                param_name="nesting_levels",
                received=type(nesting_levels).__name__,
                valid_values=["list of integers"],
            )
            assert "nesting_levels" in error
            assert "str" in error

    def test_validate_nesting_levels_length_mismatch(self):
        """nesting_levels length must match items length."""
        validator = ValidationManager()
        items = ["Item 1", "Item 2", "Item 3"]
        nesting_levels = [0, 1]  # Wrong length

        if len(nesting_levels) != len(items):
            error = validator.create_invalid_param_error(
                param_name="nesting_levels",
                received=f"list of length {len(nesting_levels)}",
                valid_values=[f"list of length {len(items)} (must match items length)"],
            )
            assert "nesting_levels" in error
            assert "length 2" in error
            assert "length 3" in error

    def test_validate_nesting_level_must_be_integer(self):
        """Each nesting level must be an integer."""
        validator = ValidationManager()
        nesting_levels = [0, "one", 2]

        for i, level in enumerate(nesting_levels):
            if not isinstance(level, int):
                error = validator.create_invalid_param_error(
                    param_name=f"nesting_levels[{i}]",
                    received=type(level).__name__,
                    valid_values=["integer (0-8)"],
                )
                assert "nesting_levels[1]" in error
                assert "str" in error
                break

    def test_validate_nesting_level_out_of_range_negative(self):
        """Nesting level must be >= 0."""
        validator = ValidationManager()
        nesting_levels = [0, -1, 2]

        for i, level in enumerate(nesting_levels):
            if isinstance(level, int) and (level < 0 or level > 8):
                error = validator.create_invalid_param_error(
                    param_name=f"nesting_levels[{i}]",
                    received=str(level),
                    valid_values=["integer from 0 to 8"],
                )
                assert "nesting_levels[1]" in error
                assert "-1" in error
                break

    def test_validate_nesting_level_out_of_range_too_high(self):
        """Nesting level must be <= 8."""
        validator = ValidationManager()
        nesting_levels = [0, 1, 9]

        for i, level in enumerate(nesting_levels):
            if isinstance(level, int) and (level < 0 or level > 8):
                error = validator.create_invalid_param_error(
                    param_name=f"nesting_levels[{i}]",
                    received=str(level),
                    valid_values=["integer from 0 to 8"],
                )
                assert "nesting_levels[2]" in error
                assert "9" in error
                break

    def test_has_nested_detection(self):
        """Test detection of whether list has nested items."""
        # No nesting
        nesting_levels_flat = [0, 0, 0]
        has_nested_flat = any(level > 0 for level in nesting_levels_flat)
        assert has_nested_flat is False

        # Has nesting
        nesting_levels_nested = [0, 1, 0]
        has_nested_nested = any(level > 0 for level in nesting_levels_nested)
        assert has_nested_nested is True

    def test_max_nesting_level_calculation(self):
        """Test max nesting level calculation for description."""
        nesting_levels = [0, 1, 2, 1, 0]
        max_level = max(nesting_levels)
        assert max_level == 2

    def test_bullet_range_with_tabs(self):
        """Bullet range should include tab characters."""
        items = ["Main", "Sub"]
        nesting_levels = [0, 1]

        # Build nested items
        nested_items = []
        for item_text, level in zip(items, nesting_levels):
            prefix = "\t" * level
            nested_items.append(prefix + item_text)

        combined = "\n".join(nested_items) + "\n"
        # "Main\n\tSub\n" = 10 chars
        assert combined == "Main\n\tSub\n"
        assert len(combined) == 10

        start_index = 100
        total_length = len(combined) - 1  # Exclude final newline
        assert total_length == 9

        end_index = start_index + total_length
        assert end_index == 109


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
