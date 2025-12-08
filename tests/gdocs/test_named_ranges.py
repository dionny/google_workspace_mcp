"""
Unit tests for Named Range functionality.

Tests verify proper handling of named range creation, listing, and deletion requests.
"""
from gdocs.docs_helpers import (
    create_named_range_request,
    create_delete_named_range_request,
)
import pytest


class TestCreateNamedRangeRequest:
    """Tests for create_named_range_request helper function."""

    def test_basic_named_range_request(self):
        """Named range request should have correct structure."""
        request = create_named_range_request(
            name="test_range",
            start_index=10,
            end_index=50
        )

        assert "createNamedRange" in request
        assert request["createNamedRange"]["name"] == "test_range"
        assert "range" in request["createNamedRange"]
        assert request["createNamedRange"]["range"]["startIndex"] == 10
        assert request["createNamedRange"]["range"]["endIndex"] == 50

    def test_named_range_at_start(self):
        """Named range at document start (index 1)."""
        request = create_named_range_request(
            name="start_marker",
            start_index=1,
            end_index=1
        )

        assert request["createNamedRange"]["range"]["startIndex"] == 1
        assert request["createNamedRange"]["range"]["endIndex"] == 1

    def test_named_range_with_segment_id(self):
        """Named range with segment ID for headers/footers."""
        request = create_named_range_request(
            name="header_section",
            start_index=0,
            end_index=20,
            segment_id="kix.header123"
        )

        assert request["createNamedRange"]["range"]["segmentId"] == "kix.header123"

    def test_named_range_with_tab_id(self):
        """Named range with tab ID for multi-tab documents."""
        request = create_named_range_request(
            name="tab_section",
            start_index=1,
            end_index=100,
            tab_id="tab.abc123"
        )

        assert request["createNamedRange"]["range"]["tabId"] == "tab.abc123"

    def test_named_range_with_all_options(self):
        """Named range with all optional parameters."""
        request = create_named_range_request(
            name="full_range",
            start_index=50,
            end_index=150,
            segment_id="kix.segment",
            tab_id="tab.xyz"
        )

        range_obj = request["createNamedRange"]["range"]
        assert range_obj["startIndex"] == 50
        assert range_obj["endIndex"] == 150
        assert range_obj["segmentId"] == "kix.segment"
        assert range_obj["tabId"] == "tab.xyz"

    def test_named_range_structure(self):
        """Verify full request structure matches Google Docs API spec."""
        request = create_named_range_request(
            name="structure_test",
            start_index=100,
            end_index=200
        )

        # Should have exactly one key at top level
        assert len(request) == 1
        assert "createNamedRange" in request

        # Should have name and range
        create_request = request["createNamedRange"]
        assert "name" in create_request
        assert "range" in create_request

        # Range should have startIndex and endIndex
        range_obj = create_request["range"]
        assert "startIndex" in range_obj
        assert "endIndex" in range_obj

    def test_named_range_name_with_special_characters(self):
        """Named range names can contain special characters."""
        request = create_named_range_request(
            name="section_v2.0_alpha-test",
            start_index=1,
            end_index=10
        )

        assert request["createNamedRange"]["name"] == "section_v2.0_alpha-test"

    def test_named_range_name_with_unicode(self):
        """Named range names can contain unicode characters."""
        request = create_named_range_request(
            name="sección_principal",
            start_index=1,
            end_index=10
        )

        assert request["createNamedRange"]["name"] == "sección_principal"


class TestDeleteNamedRangeRequest:
    """Tests for create_delete_named_range_request helper function."""

    def test_delete_by_id(self):
        """Delete request by named range ID should have correct structure."""
        request = create_delete_named_range_request(
            named_range_id="kix.abc123def456"
        )

        assert "deleteNamedRange" in request
        assert request["deleteNamedRange"]["namedRangeId"] == "kix.abc123def456"
        assert "name" not in request["deleteNamedRange"]

    def test_delete_by_name(self):
        """Delete request by name should have correct structure."""
        request = create_delete_named_range_request(
            name="old_marker"
        )

        assert "deleteNamedRange" in request
        assert request["deleteNamedRange"]["name"] == "old_marker"
        assert "namedRangeId" not in request["deleteNamedRange"]

    def test_delete_requires_identifier(self):
        """Delete request should raise error if neither ID nor name provided."""
        with pytest.raises(ValueError) as exc_info:
            create_delete_named_range_request()

        assert "Either named_range_id or name must be provided" in str(exc_info.value)

    def test_delete_with_tabs_criteria(self):
        """Delete request can include tabs criteria."""
        request = create_delete_named_range_request(
            name="multi_tab_range",
            tabs_criteria={"tabIds": ["tab1", "tab2"]}
        )

        assert request["deleteNamedRange"]["name"] == "multi_tab_range"
        assert request["deleteNamedRange"]["tabsCriteria"] == {"tabIds": ["tab1", "tab2"]}

    def test_delete_by_id_structure(self):
        """Verify delete by ID request structure matches Google Docs API spec."""
        request = create_delete_named_range_request(
            named_range_id="kix.test123"
        )

        # Should have exactly one key at top level
        assert len(request) == 1
        assert "deleteNamedRange" in request

        # Should have exactly namedRangeId
        delete_request = request["deleteNamedRange"]
        assert len(delete_request) == 1
        assert "namedRangeId" in delete_request

    def test_delete_by_name_structure(self):
        """Verify delete by name request structure matches Google Docs API spec."""
        request = create_delete_named_range_request(
            name="test_name"
        )

        # Should have exactly one key at top level
        assert len(request) == 1
        assert "deleteNamedRange" in request

        # Should have exactly name
        delete_request = request["deleteNamedRange"]
        assert len(delete_request) == 1
        assert "name" in delete_request

    def test_id_takes_precedence_over_name(self):
        """When both ID and name are provided, ID should be used."""
        request = create_delete_named_range_request(
            named_range_id="kix.priority",
            name="should_be_ignored"
        )

        assert request["deleteNamedRange"]["namedRangeId"] == "kix.priority"
        assert "name" not in request["deleteNamedRange"]
