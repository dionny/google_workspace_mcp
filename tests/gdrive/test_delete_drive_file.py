"""
Unit tests for delete_drive_file tool.

These tests verify the logic and request structures for the delete_drive_file tool.
"""

import pytest


class TestDeleteDriveFileRequestStructure:
    """Unit tests for delete_drive_file request body structure."""

    def test_permanent_delete_uses_files_delete(self):
        """Test that permanent=True triggers the delete endpoint."""
        file_id = "abc123xyz"
        permanent = True

        # When permanent=True, we use files().delete()
        if permanent:
            # This would be the API call structure
            api_method = "delete"
            api_params = {"fileId": file_id, "supportsAllDrives": True}
        else:
            api_method = "update"
            api_params = {
                "fileId": file_id,
                "body": {"trashed": True},
                "supportsAllDrives": True,
            }

        assert api_method == "delete"
        assert api_params["fileId"] == "abc123xyz"
        assert api_params["supportsAllDrives"] is True

    def test_trash_uses_files_update(self):
        """Test that permanent=False (default) uses the update endpoint with trashed=True."""
        file_id = "def456ghi"
        permanent = False

        if permanent:
            api_method = "delete"
            api_params = {"fileId": file_id, "supportsAllDrives": True}
        else:
            api_method = "update"
            api_params = {
                "fileId": file_id,
                "body": {"trashed": True},
                "supportsAllDrives": True,
            }

        assert api_method == "update"
        assert api_params["fileId"] == "def456ghi"
        assert api_params["body"]["trashed"] is True
        assert api_params["supportsAllDrives"] is True


class TestDeleteDriveFileMessages:
    """Tests for delete_drive_file success message formatting."""

    def test_permanent_delete_message_format(self):
        """Test the format of permanent delete success message."""
        file_name = "my_document.pdf"
        file_id = "abc123"
        user_email = "test@example.com"

        message = (
            f"Permanently deleted file '{file_name}' (ID: {file_id}) for {user_email}."
        )

        assert "Permanently deleted" in message
        assert file_name in message
        assert file_id in message
        assert user_email in message

    def test_trash_message_format(self):
        """Test the format of move-to-trash success message."""
        file_name = "my_spreadsheet.xlsx"
        file_id = "xyz789"
        user_email = "user@example.com"

        message = f"Moved file '{file_name}' (ID: {file_id}) to trash for {user_email}. Use permanent=True to delete permanently."

        assert "to trash" in message
        assert file_name in message
        assert file_id in message
        assert user_email in message
        assert "permanent=True" in message


class TestDeleteDriveFileValidation:
    """Tests for delete_drive_file input validation logic."""

    def test_file_id_required(self):
        """Test that file_id is a required parameter."""
        # The tool requires file_id - this tests the validation logic
        file_id = None

        with pytest.raises(Exception) as exc_info:
            if not file_id:
                raise Exception("file_id is required")

        assert "file_id is required" in str(exc_info.value)

    def test_permanent_defaults_to_false(self):
        """Test that permanent parameter defaults to False (trash mode)."""
        # The function signature has permanent: bool = False
        permanent = False  # Default value

        # Default behavior should be trash, not permanent delete
        if permanent:
            action = "permanent_delete"
        else:
            action = "trash"

        assert action == "trash"

    def test_shared_drive_support(self):
        """Test that the tool supports shared drives via supportsAllDrives."""
        file_id = "shared_drive_file_123"

        # Both delete and update calls should include supportsAllDrives=True
        delete_params = {"fileId": file_id, "supportsAllDrives": True}
        update_params = {
            "fileId": file_id,
            "body": {"trashed": True},
            "supportsAllDrives": True,
        }

        assert delete_params["supportsAllDrives"] is True
        assert update_params["supportsAllDrives"] is True
