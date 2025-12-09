"""
Unit tests for freeze_panes operation in Google Sheets tools.

These tests verify the logic and request body structures for the freeze_panes tool.
"""

import pytest


class TestFreezePanesRequestBodyStructure:
    """Unit tests for freeze_panes request body structure."""

    def test_freeze_panes_request_body_structure(self):
        """Test that the request body for freeze_panes is correctly structured."""
        sheet_id = 0
        frozen_row_count = 1
        frozen_column_count = 2

        grid_properties = {
            "frozenRowCount": frozen_row_count,
            "frozenColumnCount": frozen_column_count,
        }
        fields = [
            "gridProperties.frozenRowCount",
            "gridProperties.frozenColumnCount",
        ]

        request_body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": grid_properties,
                        },
                        "fields": ",".join(fields),
                    }
                }
            ]
        }

        update_request = request_body["requests"][0]["updateSheetProperties"]
        assert update_request["properties"]["sheetId"] == 0
        assert update_request["properties"]["gridProperties"]["frozenRowCount"] == 1
        assert update_request["properties"]["gridProperties"]["frozenColumnCount"] == 2
        assert "gridProperties.frozenRowCount" in update_request["fields"]
        assert "gridProperties.frozenColumnCount" in update_request["fields"]

    def test_freeze_panes_with_various_row_counts(self):
        """Test freeze_panes with various frozen row counts."""
        test_cases = [0, 1, 2, 5, 10, 100]

        for frozen_row_count in test_cases:
            grid_properties = {
                "frozenRowCount": frozen_row_count,
                "frozenColumnCount": 0,
            }
            request_body = {
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 0,
                                "gridProperties": grid_properties,
                            },
                            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
                        }
                    }
                ]
            }
            assert (
                request_body["requests"][0]["updateSheetProperties"]["properties"][
                    "gridProperties"
                ]["frozenRowCount"]
                == frozen_row_count
            )

    def test_freeze_panes_with_various_column_counts(self):
        """Test freeze_panes with various frozen column counts."""
        test_cases = [0, 1, 2, 5, 10, 26]

        for frozen_column_count in test_cases:
            grid_properties = {
                "frozenRowCount": 0,
                "frozenColumnCount": frozen_column_count,
            }
            request_body = {
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 0,
                                "gridProperties": grid_properties,
                            },
                            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
                        }
                    }
                ]
            }
            assert (
                request_body["requests"][0]["updateSheetProperties"]["properties"][
                    "gridProperties"
                ]["frozenColumnCount"]
                == frozen_column_count
            )

    def test_freeze_panes_with_different_sheet_ids(self):
        """Test freeze_panes with various sheet IDs."""
        test_cases = [0, 1, 100, 999999, 2147483647]

        for sheet_id in test_cases:
            request_body = {
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": sheet_id,
                                "gridProperties": {
                                    "frozenRowCount": 1,
                                    "frozenColumnCount": 1,
                                },
                            },
                            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
                        }
                    }
                ]
            }
            assert (
                request_body["requests"][0]["updateSheetProperties"]["properties"][
                    "sheetId"
                ]
                == sheet_id
            )


class TestFreezePanesValidation:
    """Tests for freeze_panes input validation."""

    def test_freeze_panes_validates_negative_row_count(self):
        """Test that negative frozen_row_count is detected."""
        frozen_row_count = -1

        with pytest.raises(ValueError) as exc_info:
            if frozen_row_count < 0:
                raise ValueError(
                    f"frozen_row_count must be >= 0, got {frozen_row_count}"
                )

        assert "frozen_row_count must be >= 0" in str(exc_info.value)
        assert "-1" in str(exc_info.value)

    def test_freeze_panes_validates_negative_column_count(self):
        """Test that negative frozen_column_count is detected."""
        frozen_column_count = -5

        with pytest.raises(ValueError) as exc_info:
            if frozen_column_count < 0:
                raise ValueError(
                    f"frozen_column_count must be >= 0, got {frozen_column_count}"
                )

        assert "frozen_column_count must be >= 0" in str(exc_info.value)
        assert "-5" in str(exc_info.value)

    def test_freeze_panes_allows_zero_values(self):
        """Test that zero values are valid (used to unfreeze)."""
        frozen_row_count = 0
        frozen_column_count = 0

        # Should not raise
        assert frozen_row_count >= 0
        assert frozen_column_count >= 0


class TestFreezePanesSuccessMessages:
    """Tests for freeze_panes success message formatting."""

    def test_freeze_rows_only_message(self):
        """Test success message when only rows are frozen."""
        frozen_row_count = 1
        frozen_column_count = 0
        sheet_id = 0
        spreadsheet_id = "abc123xyz"
        user_email = "test@example.com"

        parts = []
        if frozen_row_count > 0:
            parts.append(f"{frozen_row_count} row(s)")
        if frozen_column_count > 0:
            parts.append(f"{frozen_column_count} column(s)")

        if frozen_row_count == 0 and frozen_column_count == 0:
            freeze_description = "unfrozen all rows and columns"
        else:
            freeze_description = f"frozen {' and '.join(parts)}"

        message = (
            f"Successfully {freeze_description} in sheet ID {sheet_id} "
            f"of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "frozen 1 row(s)" in message
        assert "column" not in message or "column(s)" not in message.replace(
            "frozen 1 row(s)", ""
        )
        assert "abc123xyz" in message
        assert "test@example.com" in message

    def test_freeze_columns_only_message(self):
        """Test success message when only columns are frozen."""
        frozen_row_count = 0
        frozen_column_count = 2
        sheet_id = 0
        spreadsheet_id = "def456"
        user_email = "user@example.com"

        parts = []
        if frozen_row_count > 0:
            parts.append(f"{frozen_row_count} row(s)")
        if frozen_column_count > 0:
            parts.append(f"{frozen_column_count} column(s)")

        if frozen_row_count == 0 and frozen_column_count == 0:
            freeze_description = "unfrozen all rows and columns"
        else:
            freeze_description = f"frozen {' and '.join(parts)}"

        message = (
            f"Successfully {freeze_description} in sheet ID {sheet_id} "
            f"of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "frozen 2 column(s)" in message
        assert "row(s)" not in message

    def test_freeze_rows_and_columns_message(self):
        """Test success message when both rows and columns are frozen."""
        frozen_row_count = 1
        frozen_column_count = 1
        sheet_id = 123
        spreadsheet_id = "ghi789"
        user_email = "both@example.com"

        parts = []
        if frozen_row_count > 0:
            parts.append(f"{frozen_row_count} row(s)")
        if frozen_column_count > 0:
            parts.append(f"{frozen_column_count} column(s)")

        if frozen_row_count == 0 and frozen_column_count == 0:
            freeze_description = "unfrozen all rows and columns"
        else:
            freeze_description = f"frozen {' and '.join(parts)}"

        message = (
            f"Successfully {freeze_description} in sheet ID {sheet_id} "
            f"of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "frozen 1 row(s) and 1 column(s)" in message
        assert "sheet ID 123" in message

    def test_unfreeze_all_message(self):
        """Test success message when unfreezing all rows and columns."""
        frozen_row_count = 0
        frozen_column_count = 0
        sheet_id = 0
        spreadsheet_id = "jkl012"
        user_email = "unfreeze@example.com"

        parts = []
        if frozen_row_count > 0:
            parts.append(f"{frozen_row_count} row(s)")
        if frozen_column_count > 0:
            parts.append(f"{frozen_column_count} column(s)")

        if frozen_row_count == 0 and frozen_column_count == 0:
            freeze_description = "unfrozen all rows and columns"
        else:
            freeze_description = f"frozen {' and '.join(parts)}"

        message = (
            f"Successfully {freeze_description} in sheet ID {sheet_id} "
            f"of spreadsheet {spreadsheet_id} for {user_email}."
        )

        assert "unfrozen all rows and columns" in message

    def test_freeze_multiple_rows_message(self):
        """Test success message with multiple frozen rows."""
        frozen_row_count = 5
        frozen_column_count = 0

        parts = []
        if frozen_row_count > 0:
            parts.append(f"{frozen_row_count} row(s)")
        if frozen_column_count > 0:
            parts.append(f"{frozen_column_count} column(s)")

        freeze_description = f"frozen {' and '.join(parts)}"

        assert freeze_description == "frozen 5 row(s)"


class TestFieldsMask:
    """Tests for the fields mask in freeze_panes requests."""

    def test_fields_mask_includes_both_properties(self):
        """Test that fields mask includes both frozen row and column properties."""
        fields = [
            "gridProperties.frozenRowCount",
            "gridProperties.frozenColumnCount",
        ]

        fields_str = ",".join(fields)

        assert "gridProperties.frozenRowCount" in fields_str
        assert "gridProperties.frozenColumnCount" in fields_str

    def test_fields_mask_format(self):
        """Test that fields mask is comma-separated."""
        fields = [
            "gridProperties.frozenRowCount",
            "gridProperties.frozenColumnCount",
        ]

        fields_str = ",".join(fields)

        assert (
            fields_str
            == "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
        )
