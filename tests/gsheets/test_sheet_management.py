"""
Tests for sheet management tools in Google Sheets.

These tests verify:
1. _resolve_sheet_id - Helper function for resolving sheet ID
2. _get_sheet_id_by_name - Helper function for looking up sheet ID by name

Note: The actual delete_sheet and rename_sheet tools are wrapped by decorators
and cannot be called directly in unit tests. Use tools_cli.py for integration testing.
"""

import pytest
from unittest.mock import MagicMock

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from gsheets.sheets_tools import _resolve_sheet_id, _get_sheet_id_by_name


class TestResolveSheetId:
    """Tests for _resolve_sheet_id helper function."""

    @pytest.mark.asyncio
    async def test_resolve_with_sheet_id(self):
        """Test that sheet_id is returned when provided."""
        service = MagicMock()
        result = await _resolve_sheet_id(service, "spreadsheet123", None, 456)
        assert result == 456
        # Service should not be called
        service.spreadsheets.assert_not_called()

    @pytest.mark.asyncio
    async def test_resolve_with_sheet_name(self):
        """Test that sheet ID is looked up when sheet_name is provided."""
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

        result = await _resolve_sheet_id(service, "spreadsheet123", "DataSheet", None)
        assert result == 123

    @pytest.mark.asyncio
    async def test_resolve_with_neither_gets_first_sheet(self):
        """Test that first sheet ID is returned when neither name nor ID provided."""
        mock_spreadsheet_data = {
            "sheets": [
                {"properties": {"sheetId": 789, "title": "FirstSheet"}},
                {"properties": {"sheetId": 123, "title": "SecondSheet"}},
            ]
        }

        service = MagicMock()
        service.spreadsheets().get().execute = MagicMock(
            return_value=mock_spreadsheet_data
        )

        result = await _resolve_sheet_id(service, "spreadsheet123", None, None)
        assert result == 789

    @pytest.mark.asyncio
    async def test_resolve_sheet_name_not_found(self):
        """Test error when sheet name not found."""
        mock_spreadsheet_data = {
            "sheets": [
                {"properties": {"sheetId": 0, "title": "Sheet1"}},
            ]
        }

        service = MagicMock()
        service.spreadsheets().get().execute = MagicMock(
            return_value=mock_spreadsheet_data
        )

        with pytest.raises(ValueError, match="Sheet 'NonExistent' not found"):
            await _resolve_sheet_id(service, "spreadsheet123", "NonExistent", None)

    @pytest.mark.asyncio
    async def test_resolve_no_sheets_error(self):
        """Test error when spreadsheet has no sheets."""
        mock_spreadsheet_data = {"sheets": []}

        service = MagicMock()
        service.spreadsheets().get().execute = MagicMock(
            return_value=mock_spreadsheet_data
        )

        with pytest.raises(ValueError, match="Spreadsheet has no sheets"):
            await _resolve_sheet_id(service, "spreadsheet123", None, None)


class TestGetSheetIdByName:
    """Tests for _get_sheet_id_by_name helper function."""

    @pytest.mark.asyncio
    async def test_get_sheet_id_by_name_found(self):
        """Test finding sheet ID by name."""
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

        result = await _get_sheet_id_by_name(service, "spreadsheet123", "DataSheet")
        assert result == 123

    @pytest.mark.asyncio
    async def test_get_sheet_id_by_name_not_found(self):
        """Test error when sheet name not found."""
        mock_spreadsheet_data = {
            "sheets": [
                {"properties": {"sheetId": 0, "title": "Sheet1"}},
            ]
        }

        service = MagicMock()
        service.spreadsheets().get().execute = MagicMock(
            return_value=mock_spreadsheet_data
        )

        with pytest.raises(ValueError, match="Sheet 'NonExistent' not found"):
            await _get_sheet_id_by_name(service, "spreadsheet123", "NonExistent")

    @pytest.mark.asyncio
    async def test_get_sheet_id_by_name_first_sheet(self):
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

        result = await _get_sheet_id_by_name(service, "spreadsheet123", "FirstSheet")
        assert result == 0
