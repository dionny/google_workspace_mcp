"""
Tests for sheet range building functionality in Google Sheets.

These tests verify the _build_full_range helper function that builds
full range strings with sheet references for consistent API.
"""

import pytest
from unittest.mock import MagicMock

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from gsheets.sheets_tools import _build_full_range, _get_sheet_name_by_id


class TestBuildFullRangeSync:
    """Tests for _build_full_range function - synchronous parts."""

    @pytest.mark.asyncio
    async def test_range_with_existing_sheet_reference(self):
        """Test that range with ! is returned as-is."""
        service = MagicMock()
        result = await _build_full_range(
            service, "spreadsheet123", "Sheet1!A1:D10", None, None
        )
        assert result == "Sheet1!A1:D10"

    @pytest.mark.asyncio
    async def test_range_with_quoted_sheet_reference(self):
        """Test that range with quoted sheet name is returned as-is."""
        service = MagicMock()
        result = await _build_full_range(
            service, "spreadsheet123", "'My Sheet'!A1:D10", None, None
        )
        assert result == "'My Sheet'!A1:D10"

    @pytest.mark.asyncio
    async def test_range_without_sheet_reference(self):
        """Test that range without sheet reference is returned as-is when no sheet specified."""
        service = MagicMock()
        result = await _build_full_range(
            service, "spreadsheet123", "A1:D10", None, None
        )
        assert result == "A1:D10"

    @pytest.mark.asyncio
    async def test_range_with_simple_sheet_name(self):
        """Test adding simple sheet name to range."""
        service = MagicMock()
        result = await _build_full_range(
            service, "spreadsheet123", "A1:D10", "Sheet1", None
        )
        assert result == "Sheet1!A1:D10"

    @pytest.mark.asyncio
    async def test_range_with_sheet_name_containing_space(self):
        """Test that sheet names with spaces are properly quoted."""
        service = MagicMock()
        result = await _build_full_range(
            service, "spreadsheet123", "A1:D10", "My Sheet", None
        )
        assert result == "'My Sheet'!A1:D10"

    @pytest.mark.asyncio
    async def test_range_with_sheet_name_containing_apostrophe(self):
        """Test that sheet names with apostrophes are properly escaped."""
        service = MagicMock()
        result = await _build_full_range(
            service, "spreadsheet123", "A1:D10", "John's Data", None
        )
        assert result == "'John''s Data'!A1:D10"

    @pytest.mark.asyncio
    async def test_range_with_sheet_name_containing_colon(self):
        """Test that sheet names with colons are properly quoted."""
        service = MagicMock()
        result = await _build_full_range(
            service, "spreadsheet123", "A1:D10", "Data:2024", None
        )
        assert result == "'Data:2024'!A1:D10"

    @pytest.mark.asyncio
    async def test_range_with_sheet_name_containing_exclamation(self):
        """Test that sheet names with exclamation marks are properly quoted."""
        service = MagicMock()
        result = await _build_full_range(
            service, "spreadsheet123", "A1:D10", "Important!", None
        )
        assert result == "'Important!'!A1:D10"

    @pytest.mark.asyncio
    async def test_sheet_id_takes_precedence_over_sheet_name(self):
        """Test that sheet_id takes precedence when both are provided."""
        # Mock the service to return spreadsheet info
        mock_spreadsheet_data = {
            "sheets": [
                {"properties": {"sheetId": 0, "title": "Sheet1"}},
                {"properties": {"sheetId": 123, "title": "DataSheet"}},
            ]
        }

        service = MagicMock()
        service.spreadsheets().get().execute = MagicMock(
            return_value=mock_spreadsheet_data
        )

        result = await _build_full_range(
            service, "spreadsheet123", "A1:D10", "Sheet1", 123
        )
        # Should use DataSheet (from sheet_id 123), not Sheet1
        assert result == "DataSheet!A1:D10"


class TestGetSheetNameById:
    """Tests for _get_sheet_name_by_id function."""

    @pytest.mark.asyncio
    async def test_get_sheet_name_by_id_found(self):
        """Test finding sheet name by ID."""
        mock_spreadsheet_data = {
            "sheets": [
                {"properties": {"sheetId": 0, "title": "Sheet1"}},
                {"properties": {"sheetId": 123, "title": "DataSheet"}},
                {"properties": {"sheetId": 456, "title": "Summary"}},
            ]
        }

        service = MagicMock()
        service.spreadsheets().get().execute = MagicMock(
            return_value=mock_spreadsheet_data
        )

        result = await _get_sheet_name_by_id(service, "spreadsheet123", 123)
        assert result == "DataSheet"

    @pytest.mark.asyncio
    async def test_get_sheet_name_by_id_not_found(self):
        """Test error when sheet ID not found."""
        mock_spreadsheet_data = {
            "sheets": [
                {"properties": {"sheetId": 0, "title": "Sheet1"}},
            ]
        }

        service = MagicMock()
        service.spreadsheets().get().execute = MagicMock(
            return_value=mock_spreadsheet_data
        )

        with pytest.raises(ValueError, match="Sheet with ID 999 not found"):
            await _get_sheet_name_by_id(service, "spreadsheet123", 999)

    @pytest.mark.asyncio
    async def test_get_sheet_name_by_id_first_sheet(self):
        """Test finding first sheet (ID 0)."""
        mock_spreadsheet_data = {
            "sheets": [
                {"properties": {"sheetId": 0, "title": "FirstSheet"}},
            ]
        }

        service = MagicMock()
        service.spreadsheets().get().execute = MagicMock(
            return_value=mock_spreadsheet_data
        )

        result = await _get_sheet_name_by_id(service, "spreadsheet123", 0)
        assert result == "FirstSheet"
