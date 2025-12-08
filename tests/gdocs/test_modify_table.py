"""
Unit tests for modify_table tool and table modification helper functions.

These tests verify:
- Table modification request builders (insert/delete row/column)
- Input validation for modify_table operations
- Operation execution logic
"""
import pytest
from gdocs.docs_helpers import (
    create_insert_table_row_request,
    create_delete_table_row_request,
    create_insert_table_column_request,
    create_delete_table_column_request,
    create_delete_range_request,
    create_merge_table_cells_request,
    create_unmerge_table_cells_request,
    create_update_table_cell_style_request,
    create_update_table_column_properties_request,
)


class TestTableRowRequests:
    """Tests for table row request builders."""

    def test_insert_table_row_below(self):
        """Test creating an insert row request below a row."""
        request = create_insert_table_row_request(
            table_start_index=10,
            row_index=2,
            insert_below=True,
            column_index=0
        )

        assert 'insertTableRow' in request
        insert_req = request['insertTableRow']
        assert insert_req['insertBelow'] is True
        assert insert_req['tableCellLocation']['tableStartLocation']['index'] == 10
        assert insert_req['tableCellLocation']['rowIndex'] == 2
        assert insert_req['tableCellLocation']['columnIndex'] == 0

    def test_insert_table_row_above(self):
        """Test creating an insert row request above a row."""
        request = create_insert_table_row_request(
            table_start_index=10,
            row_index=0,
            insert_below=False
        )

        assert request['insertTableRow']['insertBelow'] is False
        assert request['insertTableRow']['tableCellLocation']['rowIndex'] == 0

    def test_delete_table_row_request(self):
        """Test creating a delete row request."""
        request = create_delete_table_row_request(
            table_start_index=15,
            row_index=3,
            column_index=0
        )

        assert 'deleteTableRow' in request
        delete_req = request['deleteTableRow']
        assert delete_req['tableCellLocation']['tableStartLocation']['index'] == 15
        assert delete_req['tableCellLocation']['rowIndex'] == 3


class TestTableColumnRequests:
    """Tests for table column request builders."""

    def test_insert_table_column_right(self):
        """Test creating an insert column request to the right."""
        request = create_insert_table_column_request(
            table_start_index=10,
            row_index=0,
            column_index=1,
            insert_right=True
        )

        assert 'insertTableColumn' in request
        insert_req = request['insertTableColumn']
        assert insert_req['insertRight'] is True
        assert insert_req['tableCellLocation']['tableStartLocation']['index'] == 10
        assert insert_req['tableCellLocation']['columnIndex'] == 1

    def test_insert_table_column_left(self):
        """Test creating an insert column request to the left."""
        request = create_insert_table_column_request(
            table_start_index=10,
            row_index=0,
            column_index=2,
            insert_right=False
        )

        assert request['insertTableColumn']['insertRight'] is False
        assert request['insertTableColumn']['tableCellLocation']['columnIndex'] == 2

    def test_delete_table_column_request(self):
        """Test creating a delete column request."""
        request = create_delete_table_column_request(
            table_start_index=20,
            row_index=0,
            column_index=1
        )

        assert 'deleteTableColumn' in request
        delete_req = request['deleteTableColumn']
        assert delete_req['tableCellLocation']['tableStartLocation']['index'] == 20
        assert delete_req['tableCellLocation']['columnIndex'] == 1


class TestRequestStructure:
    """Tests for request structure correctness."""

    def test_insert_row_request_structure_matches_api(self):
        """Verify insert row request matches Google Docs API format."""
        request = create_insert_table_row_request(
            table_start_index=5,
            row_index=0,
            insert_below=True
        )

        # Verify full structure
        assert 'insertTableRow' in request
        assert 'tableCellLocation' in request['insertTableRow']
        assert 'tableStartLocation' in request['insertTableRow']['tableCellLocation']
        assert 'index' in request['insertTableRow']['tableCellLocation']['tableStartLocation']
        assert 'rowIndex' in request['insertTableRow']['tableCellLocation']
        assert 'columnIndex' in request['insertTableRow']['tableCellLocation']
        assert 'insertBelow' in request['insertTableRow']

    def test_delete_row_request_structure_matches_api(self):
        """Verify delete row request matches Google Docs API format."""
        request = create_delete_table_row_request(
            table_start_index=5,
            row_index=1
        )

        assert 'deleteTableRow' in request
        assert 'tableCellLocation' in request['deleteTableRow']
        assert 'tableStartLocation' in request['deleteTableRow']['tableCellLocation']
        assert 'index' in request['deleteTableRow']['tableCellLocation']['tableStartLocation']
        assert 'rowIndex' in request['deleteTableRow']['tableCellLocation']

    def test_insert_column_request_structure_matches_api(self):
        """Verify insert column request matches Google Docs API format."""
        request = create_insert_table_column_request(
            table_start_index=5,
            row_index=0,
            column_index=0,
            insert_right=True
        )

        assert 'insertTableColumn' in request
        assert 'tableCellLocation' in request['insertTableColumn']
        assert 'insertRight' in request['insertTableColumn']

    def test_delete_column_request_structure_matches_api(self):
        """Verify delete column request matches Google Docs API format."""
        request = create_delete_table_column_request(
            table_start_index=5,
            row_index=0,
            column_index=2
        )

        assert 'deleteTableColumn' in request
        assert 'tableCellLocation' in request['deleteTableColumn']


class TestDefaultValues:
    """Tests for default parameter values."""

    def test_insert_row_default_column_index(self):
        """Test that insert_row defaults column_index to 0."""
        request = create_insert_table_row_request(
            table_start_index=10,
            row_index=1,
            insert_below=True
        )

        assert request['insertTableRow']['tableCellLocation']['columnIndex'] == 0

    def test_delete_row_default_column_index(self):
        """Test that delete_row defaults column_index to 0."""
        request = create_delete_table_row_request(
            table_start_index=10,
            row_index=1
        )

        assert request['deleteTableRow']['tableCellLocation']['columnIndex'] == 0

    def test_insert_row_default_insert_below(self):
        """Test that insert_below defaults to True."""
        request = create_insert_table_row_request(
            table_start_index=10,
            row_index=1
        )

        assert request['insertTableRow']['insertBelow'] is True

    def test_insert_column_default_insert_right(self):
        """Test that insert_right defaults to True."""
        request = create_insert_table_column_request(
            table_start_index=10,
            row_index=0,
            column_index=0
        )

        assert request['insertTableColumn']['insertRight'] is True


class TestDeleteRangeRequest:
    """Tests for delete range request builder (used by delete_table)."""

    def test_delete_range_request_structure(self):
        """Test creating a delete range request for table deletion."""
        request = create_delete_range_request(
            start_index=10,
            end_index=100
        )

        assert 'deleteContentRange' in request
        delete_req = request['deleteContentRange']
        assert 'range' in delete_req
        assert delete_req['range']['startIndex'] == 10
        assert delete_req['range']['endIndex'] == 100

    def test_delete_range_request_with_table_bounds(self):
        """Test delete range request with typical table boundaries."""
        # Simulate a table at indices 50-150
        request = create_delete_range_request(
            start_index=50,
            end_index=150
        )

        assert request['deleteContentRange']['range']['startIndex'] == 50
        assert request['deleteContentRange']['range']['endIndex'] == 150


class TestEdgeCases:
    """Tests for edge cases and boundary conditions."""

    def test_insert_row_at_first_row(self):
        """Test inserting row at position 0."""
        request = create_insert_table_row_request(
            table_start_index=10,
            row_index=0,
            insert_below=False
        )

        assert request['insertTableRow']['tableCellLocation']['rowIndex'] == 0
        assert request['insertTableRow']['insertBelow'] is False

    def test_insert_column_at_first_column(self):
        """Test inserting column at position 0."""
        request = create_insert_table_column_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            insert_right=False
        )

        assert request['insertTableColumn']['tableCellLocation']['columnIndex'] == 0
        assert request['insertTableColumn']['insertRight'] is False

    def test_large_table_indices(self):
        """Test with large table index values."""
        request = create_insert_table_row_request(
            table_start_index=99999,
            row_index=500,
            insert_below=True,
            column_index=100
        )

        assert request['insertTableRow']['tableCellLocation']['tableStartLocation']['index'] == 99999
        assert request['insertTableRow']['tableCellLocation']['rowIndex'] == 500
        assert request['insertTableRow']['tableCellLocation']['columnIndex'] == 100


class TestMergeTableCellsRequest:
    """Tests for merge table cells request builder."""

    def test_merge_cells_basic(self):
        """Test creating a basic merge cells request."""
        request = create_merge_table_cells_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            row_span=1,
            column_span=2
        )

        assert 'mergeTableCells' in request
        merge_req = request['mergeTableCells']
        assert 'tableRange' in merge_req
        assert merge_req['tableRange']['tableCellLocation']['tableStartLocation']['index'] == 10
        assert merge_req['tableRange']['tableCellLocation']['rowIndex'] == 0
        assert merge_req['tableRange']['tableCellLocation']['columnIndex'] == 0
        assert merge_req['tableRange']['rowSpan'] == 1
        assert merge_req['tableRange']['columnSpan'] == 2

    def test_merge_cells_multi_row(self):
        """Test merging cells across multiple rows."""
        request = create_merge_table_cells_request(
            table_start_index=15,
            row_index=1,
            column_index=2,
            row_span=3,
            column_span=1
        )

        table_range = request['mergeTableCells']['tableRange']
        assert table_range['rowSpan'] == 3
        assert table_range['columnSpan'] == 1
        assert table_range['tableCellLocation']['rowIndex'] == 1
        assert table_range['tableCellLocation']['columnIndex'] == 2

    def test_merge_cells_full_row(self):
        """Test merging all cells in a row to create a header."""
        request = create_merge_table_cells_request(
            table_start_index=5,
            row_index=0,
            column_index=0,
            row_span=1,
            column_span=5
        )

        table_range = request['mergeTableCells']['tableRange']
        assert table_range['rowSpan'] == 1
        assert table_range['columnSpan'] == 5

    def test_merge_cells_request_structure_matches_api(self):
        """Verify merge cells request matches Google Docs API format."""
        request = create_merge_table_cells_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            row_span=2,
            column_span=2
        )

        # Verify full structure
        assert 'mergeTableCells' in request
        assert 'tableRange' in request['mergeTableCells']
        table_range = request['mergeTableCells']['tableRange']
        assert 'tableCellLocation' in table_range
        assert 'tableStartLocation' in table_range['tableCellLocation']
        assert 'index' in table_range['tableCellLocation']['tableStartLocation']
        assert 'rowIndex' in table_range['tableCellLocation']
        assert 'columnIndex' in table_range['tableCellLocation']
        assert 'rowSpan' in table_range
        assert 'columnSpan' in table_range


class TestUnmergeTableCellsRequest:
    """Tests for unmerge table cells request builder."""

    def test_unmerge_cells_basic(self):
        """Test creating a basic unmerge cells request."""
        request = create_unmerge_table_cells_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            row_span=1,
            column_span=2
        )

        assert 'unmergeTableCells' in request
        unmerge_req = request['unmergeTableCells']
        assert 'tableRange' in unmerge_req
        assert unmerge_req['tableRange']['tableCellLocation']['tableStartLocation']['index'] == 10
        assert unmerge_req['tableRange']['tableCellLocation']['rowIndex'] == 0
        assert unmerge_req['tableRange']['tableCellLocation']['columnIndex'] == 0
        assert unmerge_req['tableRange']['rowSpan'] == 1
        assert unmerge_req['tableRange']['columnSpan'] == 2

    def test_unmerge_cells_multi_row(self):
        """Test unmerging cells across multiple rows."""
        request = create_unmerge_table_cells_request(
            table_start_index=15,
            row_index=1,
            column_index=0,
            row_span=3,
            column_span=2
        )

        table_range = request['unmergeTableCells']['tableRange']
        assert table_range['rowSpan'] == 3
        assert table_range['columnSpan'] == 2
        assert table_range['tableCellLocation']['rowIndex'] == 1
        assert table_range['tableCellLocation']['columnIndex'] == 0

    def test_unmerge_cells_request_structure_matches_api(self):
        """Verify unmerge cells request matches Google Docs API format."""
        request = create_unmerge_table_cells_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            row_span=2,
            column_span=2
        )

        # Verify full structure
        assert 'unmergeTableCells' in request
        assert 'tableRange' in request['unmergeTableCells']
        table_range = request['unmergeTableCells']['tableRange']
        assert 'tableCellLocation' in table_range
        assert 'tableStartLocation' in table_range['tableCellLocation']
        assert 'index' in table_range['tableCellLocation']['tableStartLocation']
        assert 'rowIndex' in table_range['tableCellLocation']
        assert 'columnIndex' in table_range['tableCellLocation']
        assert 'rowSpan' in table_range
        assert 'columnSpan' in table_range


class TestMergeCellsEdgeCases:
    """Tests for edge cases in merge/unmerge operations."""

    def test_merge_single_cell(self):
        """Test merge request with single cell (no actual merge)."""
        request = create_merge_table_cells_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            row_span=1,
            column_span=1
        )

        table_range = request['mergeTableCells']['tableRange']
        assert table_range['rowSpan'] == 1
        assert table_range['columnSpan'] == 1

    def test_merge_at_non_zero_position(self):
        """Test merge starting at non-zero position."""
        request = create_merge_table_cells_request(
            table_start_index=100,
            row_index=5,
            column_index=3,
            row_span=2,
            column_span=3
        )

        table_range = request['mergeTableCells']['tableRange']
        assert table_range['tableCellLocation']['rowIndex'] == 5
        assert table_range['tableCellLocation']['columnIndex'] == 3
        assert table_range['rowSpan'] == 2
        assert table_range['columnSpan'] == 3

    def test_large_span_values(self):
        """Test with large span values."""
        request = create_merge_table_cells_request(
            table_start_index=50,
            row_index=0,
            column_index=0,
            row_span=100,
            column_span=50
        )

        table_range = request['mergeTableCells']['tableRange']
        assert table_range['rowSpan'] == 100
        assert table_range['columnSpan'] == 50


class TestUpdateTableCellStyleRequest:
    """Tests for table cell style request builder (format_cell operation)."""

    def test_background_color_hex(self):
        """Test creating a request with hex background color."""
        request = create_update_table_cell_style_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            background_color="#FF0000"
        )

        assert 'updateTableCellStyle' in request
        style_req = request['updateTableCellStyle']
        assert 'tableCellStyle' in style_req
        assert 'backgroundColor' in style_req['tableCellStyle']
        bg = style_req['tableCellStyle']['backgroundColor']
        assert bg['color']['rgbColor']['red'] == 1.0
        assert bg['color']['rgbColor']['green'] == 0.0
        assert bg['color']['rgbColor']['blue'] == 0.0
        assert 'backgroundColor' in style_req['fields']

    def test_background_color_named(self):
        """Test creating a request with named background color."""
        request = create_update_table_cell_style_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            background_color="yellow"
        )

        bg = request['updateTableCellStyle']['tableCellStyle']['backgroundColor']
        assert bg['color']['rgbColor']['red'] == 1.0
        assert bg['color']['rgbColor']['green'] == 1.0
        assert bg['color']['rgbColor']['blue'] == 0.0

    def test_all_borders_with_defaults(self):
        """Test creating a request with default border settings for all borders."""
        request = create_update_table_cell_style_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            border_color="red",
            border_width=2,
            border_dash_style="SOLID"
        )

        style = request['updateTableCellStyle']['tableCellStyle']
        fields = request['updateTableCellStyle']['fields']

        # All four borders should be present
        for border_name in ['borderTop', 'borderBottom', 'borderLeft', 'borderRight']:
            assert border_name in style
            border = style[border_name]
            assert border['color']['color']['rgbColor']['red'] == 1.0
            assert border['width']['magnitude'] == 2
            assert border['width']['unit'] == 'PT'
            assert border['dashStyle'] == 'SOLID'
            assert border_name in fields

    def test_individual_border_override(self):
        """Test creating a request with specific border overrides."""
        request = create_update_table_cell_style_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            border_top={"color": "blue", "width": 3, "dash_style": "DASH"}
        )

        style = request['updateTableCellStyle']['tableCellStyle']
        fields = request['updateTableCellStyle']['fields']

        assert 'borderTop' in style
        border = style['borderTop']
        assert border['color']['color']['rgbColor']['blue'] == 1.0
        assert border['width']['magnitude'] == 3
        assert border['dashStyle'] == 'DASH'
        assert 'borderTop' in fields

        # Other borders should not be present
        assert 'borderBottom' not in style
        assert 'borderLeft' not in style
        assert 'borderRight' not in style

    def test_padding(self):
        """Test creating a request with padding settings."""
        request = create_update_table_cell_style_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            padding_top=5,
            padding_bottom=10,
            padding_left=8,
            padding_right=8
        )

        style = request['updateTableCellStyle']['tableCellStyle']
        fields = request['updateTableCellStyle']['fields']

        assert style['paddingTop']['magnitude'] == 5
        assert style['paddingTop']['unit'] == 'PT'
        assert style['paddingBottom']['magnitude'] == 10
        assert style['paddingLeft']['magnitude'] == 8
        assert style['paddingRight']['magnitude'] == 8

        for padding_field in ['paddingTop', 'paddingBottom', 'paddingLeft', 'paddingRight']:
            assert padding_field in fields

    def test_content_alignment(self):
        """Test creating a request with content alignment."""
        request = create_update_table_cell_style_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            content_alignment="MIDDLE"
        )

        style = request['updateTableCellStyle']['tableCellStyle']
        assert style['contentAlignment'] == 'MIDDLE'
        assert 'contentAlignment' in request['updateTableCellStyle']['fields']

    def test_table_range_structure(self):
        """Test that table range is properly structured."""
        request = create_update_table_cell_style_request(
            table_start_index=100,
            row_index=2,
            column_index=3,
            row_span=2,
            column_span=4,
            background_color="gray"
        )

        table_range = request['updateTableCellStyle']['tableRange']
        assert table_range['tableCellLocation']['tableStartLocation']['index'] == 100
        assert table_range['tableCellLocation']['rowIndex'] == 2
        assert table_range['tableCellLocation']['columnIndex'] == 3
        assert table_range['rowSpan'] == 2
        assert table_range['columnSpan'] == 4

    def test_combined_formatting(self):
        """Test creating a request with multiple formatting options."""
        request = create_update_table_cell_style_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            background_color="#E0E0E0",
            border_color="black",
            border_width=1,
            padding_top=5,
            padding_bottom=5,
            content_alignment="TOP"
        )

        style = request['updateTableCellStyle']['tableCellStyle']
        fields = request['updateTableCellStyle']['fields']

        # Background should be set
        assert 'backgroundColor' in style
        # Borders should be set
        assert 'borderTop' in style
        assert 'borderBottom' in style
        # Padding should be set
        assert 'paddingTop' in style
        assert 'paddingBottom' in style
        # Alignment should be set
        assert style['contentAlignment'] == 'TOP'

        # Verify fields contains all set properties
        assert 'backgroundColor' in fields
        assert 'contentAlignment' in fields

    def test_default_row_column_span(self):
        """Test that default row/column span is 1."""
        request = create_update_table_cell_style_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            background_color="white"
        )

        table_range = request['updateTableCellStyle']['tableRange']
        assert table_range['rowSpan'] == 1
        assert table_range['columnSpan'] == 1

    def test_fields_no_duplicates(self):
        """Test that fields list has no duplicates."""
        request = create_update_table_cell_style_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            border_color="red",
            border_width=1,
            border_dash_style="SOLID"
        )

        fields = request['updateTableCellStyle']['fields'].split(',')
        assert len(fields) == len(set(fields)), "Fields list should have no duplicates"

    def test_short_hex_color(self):
        """Test that short hex colors (#FFF) are expanded correctly."""
        request = create_update_table_cell_style_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            background_color="#F00"  # Short for #FF0000
        )

        bg = request['updateTableCellStyle']['tableCellStyle']['backgroundColor']
        assert bg['color']['rgbColor']['red'] == 1.0
        assert bg['color']['rgbColor']['green'] == 0.0
        assert bg['color']['rgbColor']['blue'] == 0.0

    def test_zero_padding(self):
        """Test that zero padding values are correctly set."""
        request = create_update_table_cell_style_request(
            table_start_index=10,
            row_index=0,
            column_index=0,
            padding_top=0,
            padding_bottom=0
        )

        style = request['updateTableCellStyle']['tableCellStyle']
        assert style['paddingTop']['magnitude'] == 0
        assert style['paddingBottom']['magnitude'] == 0
        assert 'paddingTop' in request['updateTableCellStyle']['fields']


class TestUpdateTableColumnPropertiesRequest:
    """Tests for table column properties request builder (resize_column operation)."""

    def test_resize_single_column(self):
        """Test creating a request to resize a single column."""
        request = create_update_table_column_properties_request(
            table_start_index=10,
            column_indices=[0],
            width=100.0
        )

        assert 'updateTableColumnProperties' in request
        update_req = request['updateTableColumnProperties']
        assert update_req['tableStartLocation']['index'] == 10
        assert update_req['columnIndices'] == [0]
        assert update_req['tableColumnProperties']['width']['magnitude'] == 100.0
        assert update_req['tableColumnProperties']['width']['unit'] == 'PT'
        assert update_req['tableColumnProperties']['widthType'] == 'FIXED_WIDTH'
        assert update_req['fields'] == '*'

    def test_resize_multiple_columns(self):
        """Test creating a request to resize multiple columns."""
        request = create_update_table_column_properties_request(
            table_start_index=15,
            column_indices=[0, 2, 4],
            width=75.5
        )

        update_req = request['updateTableColumnProperties']
        assert update_req['columnIndices'] == [0, 2, 4]
        assert update_req['tableColumnProperties']['width']['magnitude'] == 75.5

    def test_resize_with_evenly_distributed(self):
        """Test creating a request with EVENLY_DISTRIBUTED width type."""
        request = create_update_table_column_properties_request(
            table_start_index=10,
            column_indices=[1],
            width=50.0,
            width_type="EVENLY_DISTRIBUTED"
        )

        update_req = request['updateTableColumnProperties']
        assert update_req['tableColumnProperties']['widthType'] == 'EVENLY_DISTRIBUTED'

    def test_resize_all_columns_empty_list(self):
        """Test that empty column indices list updates all columns."""
        request = create_update_table_column_properties_request(
            table_start_index=10,
            column_indices=[],
            width=100.0
        )

        update_req = request['updateTableColumnProperties']
        assert update_req['columnIndices'] == []

    def test_request_structure_matches_api(self):
        """Verify the request structure matches Google Docs API format."""
        request = create_update_table_column_properties_request(
            table_start_index=20,
            column_indices=[0],
            width=150.0
        )

        # Verify full structure
        assert 'updateTableColumnProperties' in request
        update_req = request['updateTableColumnProperties']
        assert 'tableStartLocation' in update_req
        assert 'index' in update_req['tableStartLocation']
        assert 'columnIndices' in update_req
        assert 'tableColumnProperties' in update_req
        assert 'widthType' in update_req['tableColumnProperties']
        assert 'width' in update_req['tableColumnProperties']
        assert 'magnitude' in update_req['tableColumnProperties']['width']
        assert 'unit' in update_req['tableColumnProperties']['width']
        assert 'fields' in update_req

    def test_minimum_width(self):
        """Test with minimum allowed width (5 points)."""
        request = create_update_table_column_properties_request(
            table_start_index=10,
            column_indices=[0],
            width=5.0
        )

        update_req = request['updateTableColumnProperties']
        assert update_req['tableColumnProperties']['width']['magnitude'] == 5.0

    def test_large_width_value(self):
        """Test with a large width value."""
        request = create_update_table_column_properties_request(
            table_start_index=10,
            column_indices=[0],
            width=500.0
        )

        update_req = request['updateTableColumnProperties']
        assert update_req['tableColumnProperties']['width']['magnitude'] == 500.0

    def test_integer_width_converted_to_float(self):
        """Test that integer width values work correctly."""
        request = create_update_table_column_properties_request(
            table_start_index=10,
            column_indices=[0],
            width=100  # Integer instead of float
        )

        update_req = request['updateTableColumnProperties']
        assert update_req['tableColumnProperties']['width']['magnitude'] == 100


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
