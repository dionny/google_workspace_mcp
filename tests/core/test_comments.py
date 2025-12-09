"""
Tests for core comments functionality.

These tests verify:
1. Comment reading includes anchor and quotedFileContent fields
2. Comment creation supports optional anchor parameter
3. Spreadsheet-specific docstrings clarify the difference between Drive comments and cell notes
"""

import pytest
from unittest.mock import MagicMock

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from core.comments import (
    _read_comments_impl,
    _create_comment_impl,
)


class TestReadCommentsImpl:
    """Tests for _read_comments_impl function."""

    @pytest.mark.asyncio
    async def test_read_comments_includes_anchor_field(self):
        """Test that reading comments includes the anchor field."""
        mock_service = MagicMock()
        mock_comments = MagicMock()
        mock_service.comments.return_value = mock_comments
        mock_list = MagicMock()
        mock_comments.list.return_value = mock_list
        mock_list.execute.return_value = {
            "comments": [
                {
                    "id": "test_id_1",
                    "content": "Test comment",
                    "author": {"displayName": "Test User"},
                    "createdTime": "2024-01-01T00:00:00Z",
                    "resolved": False,
                    "anchor": '{"type":"workbook-range","uid":0,"range":"A1"}',
                    "quotedFileContent": {"value": "Cell content"},
                    "replies": [],
                }
            ]
        }

        result = await _read_comments_impl(mock_service, "spreadsheet", "test_file_id")

        assert "Anchor:" in result
        assert '{"type":"workbook-range","uid":0,"range":"A1"}' in result
        assert "Quoted content: Cell content" in result

    @pytest.mark.asyncio
    async def test_read_comments_without_anchor(self):
        """Test that reading comments works when no anchor is present."""
        mock_service = MagicMock()
        mock_comments = MagicMock()
        mock_service.comments.return_value = mock_comments
        mock_list = MagicMock()
        mock_comments.list.return_value = mock_list
        mock_list.execute.return_value = {
            "comments": [
                {
                    "id": "test_id_1",
                    "content": "Test comment without anchor",
                    "author": {"displayName": "Test User"},
                    "createdTime": "2024-01-01T00:00:00Z",
                    "resolved": False,
                    "replies": [],
                }
            ]
        }

        result = await _read_comments_impl(mock_service, "spreadsheet", "test_file_id")

        assert "Comment ID: test_id_1" in result
        assert "Test comment without anchor" in result
        assert "Anchor:" not in result

    @pytest.mark.asyncio
    async def test_read_comments_no_comments(self):
        """Test that reading no comments returns appropriate message."""
        mock_service = MagicMock()
        mock_comments = MagicMock()
        mock_service.comments.return_value = mock_comments
        mock_list = MagicMock()
        mock_comments.list.return_value = mock_list
        mock_list.execute.return_value = {"comments": []}

        result = await _read_comments_impl(mock_service, "spreadsheet", "test_file_id")

        assert "No comments found" in result


class TestCreateCommentImpl:
    """Tests for _create_comment_impl function."""

    @pytest.mark.asyncio
    async def test_create_comment_without_anchor(self):
        """Test creating a comment without an anchor."""
        mock_service = MagicMock()
        mock_comments = MagicMock()
        mock_service.comments.return_value = mock_comments
        mock_create = MagicMock()
        mock_comments.create.return_value = mock_create
        mock_create.execute.return_value = {
            "id": "new_comment_id",
            "content": "Test content",
            "author": {"displayName": "Test Author"},
            "createdTime": "2024-01-01T00:00:00Z",
        }

        result = await _create_comment_impl(
            mock_service, "spreadsheet", "test_file_id", "Test content"
        )

        assert "Comment created successfully" in result
        assert "new_comment_id" in result
        # Verify the body didn't include anchor
        call_args = mock_comments.create.call_args
        body = call_args.kwargs.get("body") or call_args[1].get("body")
        assert "anchor" not in body

    @pytest.mark.asyncio
    async def test_create_comment_with_anchor(self):
        """Test creating a comment with an anchor."""
        mock_service = MagicMock()
        mock_comments = MagicMock()
        mock_service.comments.return_value = mock_comments
        mock_create = MagicMock()
        mock_comments.create.return_value = mock_create
        mock_create.execute.return_value = {
            "id": "new_comment_id",
            "content": "Test content",
            "author": {"displayName": "Test Author"},
            "createdTime": "2024-01-01T00:00:00Z",
            "anchor": '{"type":"workbook-range","uid":0,"range":"A1"}',
        }

        anchor = '{"type":"workbook-range","uid":0,"range":"A1"}'
        result = await _create_comment_impl(
            mock_service, "spreadsheet", "test_file_id", "Test content", anchor=anchor
        )

        assert "Comment created successfully" in result
        assert "new_comment_id" in result
        assert "Anchor:" in result
        # Verify the body included anchor
        call_args = mock_comments.create.call_args
        body = call_args.kwargs.get("body") or call_args[1].get("body")
        assert body.get("anchor") == anchor

    @pytest.mark.asyncio
    async def test_create_comment_anchor_none(self):
        """Test creating a comment with anchor explicitly set to None."""
        mock_service = MagicMock()
        mock_comments = MagicMock()
        mock_service.comments.return_value = mock_comments
        mock_create = MagicMock()
        mock_comments.create.return_value = mock_create
        mock_create.execute.return_value = {
            "id": "new_comment_id",
            "content": "Test content",
            "author": {"displayName": "Test Author"},
            "createdTime": "2024-01-01T00:00:00Z",
        }

        result = await _create_comment_impl(
            mock_service, "spreadsheet", "test_file_id", "Test content", anchor=None
        )

        assert "Comment created successfully" in result
        # Verify the body didn't include anchor
        call_args = mock_comments.create.call_args
        body = call_args.kwargs.get("body") or call_args[1].get("body")
        assert "anchor" not in body
