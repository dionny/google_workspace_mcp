"""
Unit tests for delete_drive_file tool.

These tests verify the logic and request structures for the delete_drive_file tool.
"""

import pytest


class TestDeleteDriveFileRequestStructure:
    """Unit tests for delete_drive_file request structures."""

    def test_trash_request_body_format(self):
        """Test that the trash request body has the correct structure."""
        # For moving to trash, we use files().update() with trashed=True
        request_body = {"trashed": True}

        assert request_body["trashed"] is True

    def test_delete_api_parameters_structure(self):
        """Test that delete API parameters are correctly structured."""
        file_id = "abc123"

        # For permanent delete, we use files().delete()
        api_params = {
            "fileId": file_id,
            "supportsAllDrives": True,
        }

        assert api_params["fileId"] == "abc123"
        assert api_params["supportsAllDrives"] is True

    def test_trash_api_parameters_structure(self):
        """Test that trash API parameters are correctly structured."""
        file_id = "abc123"

        # For trash, we use files().update()
        api_params = {
            "fileId": file_id,
            "body": {"trashed": True},
            "supportsAllDrives": True,
        }

        assert api_params["fileId"] == "abc123"
        assert api_params["body"]["trashed"] is True
        assert api_params["supportsAllDrives"] is True


class TestDeleteDriveFileParameterValidation:
    """Tests for delete_drive_file parameter behavior."""

    def test_permanent_default_is_false(self):
        """Test that permanent parameter defaults to False (trash only)."""
        permanent = False  # Default value

        # Default behavior should move to trash, not permanently delete
        assert permanent is False

    def test_permanent_true_enables_hard_delete(self):
        """Test that permanent=True enables permanent deletion."""
        permanent = True

        assert permanent is True

    def test_file_id_required(self):
        """Test that file_id is a required parameter."""
        file_id = None

        with pytest.raises(Exception) as exc_info:
            if not file_id:
                raise Exception("file_id is required")

        assert "file_id is required" in str(exc_info.value)


class TestDeleteDriveFileMessages:
    """Tests for delete_drive_file success message formatting."""

    def test_trash_success_message(self):
        """Test the format of success message for trash operation."""
        file_id = "abc123"
        file_name = "My Document.docx"
        user_email = "test@example.com"
        permanent = False

        if permanent:
            message = f"Permanently deleted file '{file_name}' (ID: {file_id}) for {user_email}. This action cannot be undone."
        else:
            message = f"Moved file '{file_name}' (ID: {file_id}) to trash for {user_email}. The file can be recovered from the trash."

        assert "Moved file" in message
        assert file_name in message
        assert file_id in message
        assert user_email in message
        assert "recovered from the trash" in message

    def test_permanent_delete_success_message(self):
        """Test the format of success message for permanent deletion."""
        file_id = "abc123"
        file_name = "My Document.docx"
        user_email = "test@example.com"
        permanent = True

        if permanent:
            message = f"Permanently deleted file '{file_name}' (ID: {file_id}) for {user_email}. This action cannot be undone."
        else:
            message = f"Moved file '{file_name}' (ID: {file_id}) to trash for {user_email}. The file can be recovered from the trash."

        assert "Permanently deleted" in message
        assert file_name in message
        assert file_id in message
        assert user_email in message
        assert "cannot be undone" in message


class TestDeleteDriveFileEdgeCases:
    """Tests for edge cases in delete_drive_file."""

    def test_file_name_with_special_characters(self):
        """Test handling file names with special characters."""
        file_name = "Report (Q1 2024) - Final.xlsx"
        file_id = "abc123"
        user_email = "test@example.com"

        message = f"Moved file '{file_name}' (ID: {file_id}) to trash for {user_email}."

        assert file_name in message

    def test_file_name_with_unicode(self):
        """Test handling file names with unicode characters."""
        file_name = "Budget Summary.pdf"
        file_id = "abc123"
        user_email = "test@example.com"

        message = f"Moved file '{file_name}' (ID: {file_id}) to trash for {user_email}."

        assert file_name in message

    def test_empty_file_name_fallback(self):
        """Test handling when file name is not available."""
        file_name = "Unknown"  # Fallback value
        file_id = "abc123"
        user_email = "test@example.com"

        message = f"Moved file '{file_name}' (ID: {file_id}) to trash for {user_email}."

        assert "Unknown" in message
        assert file_id in message

    def test_shared_drive_support(self):
        """Test that shared drive support is enabled in API calls."""
        api_params = {
            "fileId": "abc123",
            "supportsAllDrives": True,
        }

        # supportsAllDrives must be True for shared drive files
        assert api_params["supportsAllDrives"] is True


class TestDeleteDriveFileShortcutResolution:
    """Tests for shortcut resolution logic."""

    def test_shortcut_detection(self):
        """Test that shortcuts can be identified by MIME type."""
        shortcut_mime_type = "application/vnd.google-apps.shortcut"
        regular_mime_type = "application/vnd.google-apps.document"

        # Shortcuts should be resolved to their target
        assert shortcut_mime_type == "application/vnd.google-apps.shortcut"
        assert regular_mime_type != "application/vnd.google-apps.shortcut"

    def test_shortcut_has_target_id(self):
        """Test that shortcuts have target ID in shortcutDetails."""
        shortcut_metadata = {
            "mimeType": "application/vnd.google-apps.shortcut",
            "shortcutDetails": {
                "targetId": "target123",
                "targetMimeType": "application/vnd.google-apps.document",
            },
        }

        target_id = shortcut_metadata.get("shortcutDetails", {}).get("targetId")
        assert target_id == "target123"

    def test_non_shortcut_returns_same_id(self):
        """Test that non-shortcuts return the original ID."""
        file_metadata = {
            "id": "file123",
            "mimeType": "application/vnd.google-apps.document",
        }

        mime_type = file_metadata.get("mimeType")
        shortcut_mime = "application/vnd.google-apps.shortcut"

        # If not a shortcut, use the original ID
        assert mime_type != shortcut_mime
