"""
Unit tests for horizontal_rule insertion functionality.

Tests verify proper creation of horizontal rule requests using the table-based workaround.
NOTE: Google Docs API does not have a native insertHorizontalRule request, so we use
a 1x1 table with styled borders to simulate a horizontal line.
"""
from gdocs.docs_helpers import create_insert_horizontal_rule_requests


class TestCreateInsertHorizontalRuleRequests:
    """Tests for create_insert_horizontal_rule_requests helper function."""

    def test_returns_list_of_two_requests(self):
        """Horizontal rule should return two requests (insertTable + updateTableCellStyle)."""
        requests = create_insert_horizontal_rule_requests(index=1)

        assert isinstance(requests, list)
        assert len(requests) == 2

    def test_first_request_is_insert_table(self):
        """First request should insert a 1x1 table."""
        requests = create_insert_horizontal_rule_requests(index=1)

        first_request = requests[0]
        assert "insertTable" in first_request
        assert first_request["insertTable"]["location"]["index"] == 1
        assert first_request["insertTable"]["rows"] == 1
        assert first_request["insertTable"]["columns"] == 1

    def test_second_request_is_update_table_cell_style(self):
        """Second request should style the table cell borders."""
        requests = create_insert_horizontal_rule_requests(index=1)

        second_request = requests[1]
        assert "updateTableCellStyle" in second_request
        style_request = second_request["updateTableCellStyle"]

        # Should target the table at index + 1
        assert style_request["tableStartLocation"]["index"] == 2

        # Should have border styles
        cell_style = style_request["tableCellStyle"]
        assert "borderTop" in cell_style
        assert "borderBottom" in cell_style
        assert "borderLeft" in cell_style
        assert "borderRight" in cell_style

    def test_top_border_is_visible(self):
        """Top border should be visible (this creates the horizontal line effect)."""
        requests = create_insert_horizontal_rule_requests(index=1)

        cell_style = requests[1]["updateTableCellStyle"]["tableCellStyle"]
        border_top = cell_style["borderTop"]

        # Visible border should have width > 0 and dark color
        assert border_top["width"]["magnitude"] == 1
        assert border_top["color"]["color"]["rgbColor"]["red"] == 0
        assert border_top["color"]["color"]["rgbColor"]["green"] == 0
        assert border_top["color"]["color"]["rgbColor"]["blue"] == 0

    def test_other_borders_are_invisible(self):
        """Bottom, left, right borders should be invisible."""
        requests = create_insert_horizontal_rule_requests(index=1)

        cell_style = requests[1]["updateTableCellStyle"]["tableCellStyle"]

        for border_name in ["borderBottom", "borderLeft", "borderRight"]:
            border = cell_style[border_name]
            # Invisible borders should have width 0 and white color
            assert border["width"]["magnitude"] == 0
            assert border["color"]["color"]["rgbColor"]["red"] == 1
            assert border["color"]["color"]["rgbColor"]["green"] == 1
            assert border["color"]["color"]["rgbColor"]["blue"] == 1

    def test_request_with_different_index(self):
        """Horizontal rule requests should use provided index."""
        requests = create_insert_horizontal_rule_requests(index=100)

        # First request should insert at index 100
        assert requests[0]["insertTable"]["location"]["index"] == 100
        # Second request should target table at index 101
        assert requests[1]["updateTableCellStyle"]["tableStartLocation"]["index"] == 101

    def test_padding_is_zero(self):
        """Table cell should have zero padding for compact appearance."""
        requests = create_insert_horizontal_rule_requests(index=1)

        cell_style = requests[1]["updateTableCellStyle"]["tableCellStyle"]

        for padding_name in ["paddingTop", "paddingBottom", "paddingLeft", "paddingRight"]:
            assert cell_style[padding_name]["magnitude"] == 0
            assert cell_style[padding_name]["unit"] == "PT"

    def test_fields_mask_includes_all_styled_properties(self):
        """Fields mask should include all properties being styled."""
        requests = create_insert_horizontal_rule_requests(index=1)

        fields = requests[1]["updateTableCellStyle"]["fields"]

        expected_fields = [
            "borderTop", "borderBottom", "borderLeft", "borderRight",
            "paddingTop", "paddingBottom", "paddingLeft", "paddingRight"
        ]
        for field in expected_fields:
            assert field in fields
