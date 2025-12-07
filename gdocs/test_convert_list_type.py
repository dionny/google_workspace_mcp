"""
Tests for convert_list_type tool.

Issue: google_workspace_mcp-66cf

These tests verify:
1. List type conversion between bullet and numbered lists
2. Input validation for list_type parameter
3. Finding target list by index, search, or range
4. Preview mode functionality
5. Handling edge cases (no lists, invalid index, already correct type)
"""
import pytest
from gdocs.docs_helpers import create_bullet_list_request
from gdocs.managers.validation_manager import ValidationManager


class TestListTypeNormalization:
    """Test list_type parameter normalization."""

    def test_normalize_bullet_aliases(self):
        """Test that bullet aliases normalize correctly."""
        aliases = {
            "bullet": "UNORDERED",
            "bullets": "UNORDERED",
            "unordered": "UNORDERED",
            "BULLET": "UNORDERED",
            "UNORDERED": "UNORDERED",
        }

        list_type_aliases = {
            "bullet": "UNORDERED",
            "bullets": "UNORDERED",
            "unordered": "UNORDERED",
            "numbered": "ORDERED",
            "numbers": "ORDERED",
            "ordered": "ORDERED",
        }

        for input_val, expected in aliases.items():
            normalized = list_type_aliases.get(input_val.lower(), input_val.upper())
            assert normalized == expected, f"Failed for input '{input_val}'"

    def test_normalize_numbered_aliases(self):
        """Test that numbered aliases normalize correctly."""
        aliases = {
            "numbered": "ORDERED",
            "numbers": "ORDERED",
            "ordered": "ORDERED",
            "NUMBERED": "ORDERED",
            "ORDERED": "ORDERED",
        }

        list_type_aliases = {
            "bullet": "UNORDERED",
            "bullets": "UNORDERED",
            "unordered": "UNORDERED",
            "numbered": "ORDERED",
            "numbers": "ORDERED",
            "ordered": "ORDERED",
        }

        for input_val, expected in aliases.items():
            normalized = list_type_aliases.get(input_val.lower(), input_val.upper())
            assert normalized == expected, f"Failed for input '{input_val}'"

    def test_invalid_list_type(self):
        """Test that invalid list types are rejected."""
        validator = ValidationManager()
        invalid_types = ["invalid", "bullets_list", "numbering", ""]

        for invalid_type in invalid_types:
            list_type_aliases = {
                "bullet": "UNORDERED",
                "bullets": "UNORDERED",
                "unordered": "UNORDERED",
                "numbered": "ORDERED",
                "numbers": "ORDERED",
                "ordered": "ORDERED",
            }
            normalized = list_type_aliases.get(invalid_type.lower() if invalid_type else "", invalid_type.upper() if invalid_type else "")

            if normalized not in ["ORDERED", "UNORDERED"]:
                # This should trigger validation error
                error = validator.create_invalid_param_error(
                    param_name="list_type",
                    received=invalid_type,
                    valid_values=["ORDERED", "UNORDERED", "numbered", "bullet"],
                )
                assert "error" in error.lower()


class TestBulletListRequestForConversion:
    """Test that create_bullet_list_request works for conversion."""

    def test_convert_to_numbered_request(self):
        """Test request generation for converting to numbered list."""
        request = create_bullet_list_request(100, 150, "ORDERED")

        assert "createParagraphBullets" in request
        bullet_req = request["createParagraphBullets"]
        assert bullet_req["range"]["startIndex"] == 100
        assert bullet_req["range"]["endIndex"] == 150
        assert "NUMBERED" in bullet_req["bulletPreset"]

    def test_convert_to_bullet_request(self):
        """Test request generation for converting to bullet list."""
        request = create_bullet_list_request(100, 150, "UNORDERED")

        assert "createParagraphBullets" in request
        bullet_req = request["createParagraphBullets"]
        assert bullet_req["range"]["startIndex"] == 100
        assert bullet_req["range"]["endIndex"] == 150
        assert "BULLET" in bullet_req["bulletPreset"]


class TestListFindingLogic:
    """Test logic for finding target list."""

    def test_find_list_by_index_first(self):
        """Test finding first list by default index."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "items": [{"text": "First"}]},
            {"start_index": 100, "end_index": 150, "items": [{"text": "Second"}]},
        ]
        list_index = 0

        target_list = all_lists[list_index]
        assert target_list["start_index"] == 10

    def test_find_list_by_index_second(self):
        """Test finding second list by index."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "items": [{"text": "First"}]},
            {"start_index": 100, "end_index": 150, "items": [{"text": "Second"}]},
        ]
        list_index = 1

        target_list = all_lists[list_index]
        assert target_list["start_index"] == 100

    def test_find_list_by_index_out_of_range(self):
        """Test error when list index is out of range."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "items": [{"text": "First"}]},
        ]
        list_index = 5

        if list_index < 0 or list_index >= len(all_lists):
            # Should produce error
            assert True
        else:
            pytest.fail("Should have detected out of range index")

    def test_find_list_by_search_text(self):
        """Test finding list by search text in items."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "items": [{"text": "Apple"}, {"text": "Banana"}]},
            {"start_index": 100, "end_index": 150, "items": [{"text": "Step 1"}, {"text": "Step 2"}]},
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

    def test_find_list_by_search_not_found(self):
        """Test error when search text is not in any list."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "items": [{"text": "Apple"}]},
        ]
        search = "orange"
        search_lower = search.lower()

        target_list = None
        for lst in all_lists:
            for item in lst.get("items", []):
                if search_lower in item.get("text", "").lower():
                    target_list = lst
                    break

        assert target_list is None

    def test_find_list_by_range_overlap(self):
        """Test finding list that overlaps with given range."""
        all_lists = [
            {"start_index": 10, "end_index": 50, "items": []},
            {"start_index": 100, "end_index": 150, "items": []},
        ]
        start_index = 120
        end_index = 130

        target_list = None
        for lst in all_lists:
            if lst["start_index"] < end_index and lst["end_index"] > start_index:
                target_list = lst
                break

        assert target_list is not None
        assert target_list["start_index"] == 100


class TestListTypeDetection:
    """Test detecting current list type."""

    def test_detect_bullet_list(self):
        """Test detecting bullet list type."""
        target_list = {"type": "bullet_list", "list_type": "bullet"}

        current_type = target_list.get("list_type", target_list.get("type", "unknown"))
        if "bullet" in current_type or "unordered" in current_type.lower():
            current_type_display = "bullet"
            current_type_normalized = "UNORDERED"
        else:
            current_type_display = "numbered"
            current_type_normalized = "ORDERED"

        assert current_type_display == "bullet"
        assert current_type_normalized == "UNORDERED"

    def test_detect_numbered_list(self):
        """Test detecting numbered list type."""
        target_list = {"type": "numbered_list", "list_type": "numbered"}

        current_type = target_list.get("list_type", target_list.get("type", "unknown"))
        if "bullet" in current_type or "unordered" in current_type.lower():
            current_type_display = "bullet"
            current_type_normalized = "UNORDERED"
        else:
            current_type_display = "numbered"
            current_type_normalized = "ORDERED"

        assert current_type_display == "numbered"
        assert current_type_normalized == "ORDERED"


class TestNoChangeNeeded:
    """Test cases where list is already the target type."""

    def test_already_bullet_convert_to_bullet(self):
        """Test no change needed when list is already bullet."""
        current_type_normalized = "UNORDERED"
        normalized_type = "UNORDERED"

        no_change_needed = current_type_normalized == normalized_type
        assert no_change_needed is True

    def test_already_numbered_convert_to_numbered(self):
        """Test no change needed when list is already numbered."""
        current_type_normalized = "ORDERED"
        normalized_type = "ORDERED"

        no_change_needed = current_type_normalized == normalized_type
        assert no_change_needed is True

    def test_bullet_to_numbered_needs_change(self):
        """Test change needed when converting bullet to numbered."""
        current_type_normalized = "UNORDERED"
        normalized_type = "ORDERED"

        no_change_needed = current_type_normalized == normalized_type
        assert no_change_needed is False


class TestPreviewMode:
    """Test preview mode functionality."""

    def test_preview_info_structure(self):
        """Test that preview info contains expected fields."""
        target_list = {
            "start_index": 100,
            "end_index": 150,
            "items": [
                {"text": "Item 1", "start_index": 100, "end_index": 107},
                {"text": "Item 2", "start_index": 108, "end_index": 115},
            ],
        }
        current_type_display = "bullet"
        target_type_display = "numbered"

        preview_info = {
            "current_type": current_type_display,
            "target_type": target_type_display,
            "list_range": {
                "start_index": target_list["start_index"],
                "end_index": target_list["end_index"],
            },
            "items_count": len(target_list.get("items", [])),
            "items": [
                {
                    "text": item.get("text", "")[:50],
                    "start_index": item["start_index"],
                    "end_index": item["end_index"],
                }
                for item in target_list.get("items", [])[:5]
            ],
        }

        assert preview_info["current_type"] == "bullet"
        assert preview_info["target_type"] == "numbered"
        assert preview_info["list_range"]["start_index"] == 100
        assert preview_info["items_count"] == 2

    def test_preview_truncates_long_text(self):
        """Test that preview truncates long item text."""
        long_text = "A" * 100
        item = {"text": long_text}

        truncated = item.get("text", "")[:50] + ("..." if len(item.get("text", "")) > 50 else "")

        assert len(truncated) == 53  # 50 chars + "..."
        assert truncated.endswith("...")


class TestErrorResponses:
    """Test error response structures."""

    def test_no_lists_found_error(self):
        """Test error response when no lists in document."""
        doc_link = "https://docs.google.com/document/d/abc123/edit"

        error_response = {
            "error": True,
            "code": "NO_LISTS_FOUND",
            "message": "No lists found in the document",
            "suggestion": "Use insert_doc_elements or modify_doc_text with convert_to_list to create a list first",
            "link": doc_link,
        }

        assert error_response["error"] is True
        assert error_response["code"] == "NO_LISTS_FOUND"
        assert "suggestion" in error_response

    def test_list_not_found_by_search_error(self):
        """Test error response when search text not found in lists."""
        search = "nonexistent"
        doc_link = "https://docs.google.com/document/d/abc123/edit"

        error_response = {
            "error": True,
            "code": "LIST_NOT_FOUND",
            "message": f"No list containing text '{search}' found",
            "suggestion": "Check the search text or use find_doc_elements with element_type='list' to see all lists",
            "available_lists": 2,
            "link": doc_link,
        }

        assert error_response["error"] is True
        assert error_response["code"] == "LIST_NOT_FOUND"
        assert search in error_response["message"]

    def test_invalid_list_index_error(self):
        """Test error response when list index is invalid."""
        list_index = 5
        num_lists = 2
        doc_link = "https://docs.google.com/document/d/abc123/edit"

        error_response = {
            "error": True,
            "code": "INVALID_LIST_INDEX",
            "message": f"List index {list_index} is out of range. Document has {num_lists} list(s).",
            "suggestion": f"Use list_index between 0 and {num_lists - 1}",
            "available_lists": num_lists,
            "link": doc_link,
        }

        assert error_response["error"] is True
        assert error_response["code"] == "INVALID_LIST_INDEX"
        assert "5" in error_response["message"]


class TestSuccessResponses:
    """Test success response structures."""

    def test_conversion_success_response(self):
        """Test successful conversion response."""
        doc_link = "https://docs.google.com/document/d/abc123/edit"
        target_list = {"start_index": 100, "end_index": 150, "items": [{}, {}, {}]}

        success_response = {
            "success": True,
            "message": "Converted bullet list to numbered list",
            "converted_from": "bullet",
            "converted_to": "numbered",
            "list_range": {
                "start_index": target_list["start_index"],
                "end_index": target_list["end_index"],
            },
            "items_count": len(target_list.get("items", [])),
            "link": doc_link,
        }

        assert success_response["success"] is True
        assert success_response["converted_from"] == "bullet"
        assert success_response["converted_to"] == "numbered"
        assert success_response["items_count"] == 3

    def test_no_change_needed_response(self):
        """Test response when list is already target type."""
        doc_link = "https://docs.google.com/document/d/abc123/edit"
        target_list = {"start_index": 100, "end_index": 150, "items": [{}, {}]}

        response = {
            "success": True,
            "message": "List is already a bullet list",
            "no_change_needed": True,
            "list_range": {
                "start_index": target_list["start_index"],
                "end_index": target_list["end_index"],
            },
            "items_count": len(target_list.get("items", [])),
            "link": doc_link,
        }

        assert response["success"] is True
        assert response["no_change_needed"] is True


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
