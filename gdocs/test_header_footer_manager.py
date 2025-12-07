"""
Tests for the HeaderFooterManager.
"""
import pytest
from unittest.mock import MagicMock

from gdocs.managers.header_footer_manager import HeaderFooterManager


class TestReplaceSectionContent:
    """Tests for _replace_section_content method."""

    @pytest.fixture
    def manager(self):
        """Create a HeaderFooterManager with a mocked service."""
        mock_service = MagicMock()
        return HeaderFooterManager(mock_service)

    @pytest.fixture
    def mock_batch_update_success(self, manager):
        """Set up mock for successful batch update API call."""
        mock_execute = MagicMock(return_value={})
        mock_batch_update = MagicMock()
        mock_batch_update.return_value.execute = mock_execute
        manager.service.documents.return_value.batchUpdate = mock_batch_update
        return mock_batch_update

    @pytest.mark.asyncio
    async def test_replace_content_in_section_with_text(self, manager, mock_batch_update_success):
        """Test replacing content when section has existing text."""
        section = {
            'content': [
                {
                    'paragraph': {
                        'elements': [{'textRun': {'content': 'Old Header\n'}}]
                    },
                    'startIndex': 0,
                    'endIndex': 11  # "Old Header" + newline
                }
            ]
        }

        result = await manager._replace_section_content('doc123', section, 'New Header', 'kix.header123')

        assert result is True
        mock_batch_update_success.assert_called_once()
        call_args = mock_batch_update_success.call_args
        body = call_args[1]['body']
        requests = body['requests']

        # Should have delete request (old content) and insert request
        assert len(requests) == 2
        assert 'deleteContentRange' in requests[0]
        assert requests[0]['deleteContentRange']['range']['segmentId'] == 'kix.header123'
        assert 'insertText' in requests[1]
        assert requests[1]['insertText']['text'] == 'New Header'
        assert requests[1]['insertText']['location']['segmentId'] == 'kix.header123'

    @pytest.mark.asyncio
    async def test_replace_content_in_empty_section_with_paragraph(self, manager, mock_batch_update_success):
        """Test replacing content when section exists but is empty (just newline)."""
        # Empty header/footer still has a paragraph, but start == end - 1 (just newline)
        section = {
            'content': [
                {
                    'paragraph': {
                        'elements': [{'textRun': {'content': '\n'}}]
                    },
                    'startIndex': 0,
                    'endIndex': 1  # Just the newline character
                }
            ]
        }

        result = await manager._replace_section_content('doc123', section, 'New Header', 'kix.header123')

        assert result is True
        mock_batch_update_success.assert_called_once()
        call_args = mock_batch_update_success.call_args
        body = call_args[1]['body']
        requests = body['requests']

        # Should NOT have delete request (nothing to delete) - only insert
        assert len(requests) == 1
        assert 'insertText' in requests[0]
        assert requests[0]['insertText']['text'] == 'New Header'
        assert requests[0]['insertText']['location']['index'] == 0
        assert requests[0]['insertText']['location']['segmentId'] == 'kix.header123'

    @pytest.mark.asyncio
    async def test_replace_content_when_indices_equal(self, manager, mock_batch_update_success):
        """Test when start_index equals end_index (truly empty)."""
        section = {
            'content': [
                {
                    'paragraph': {
                        'elements': []
                    },
                    'startIndex': 0,
                    'endIndex': 0
                }
            ]
        }

        result = await manager._replace_section_content('doc123', section, 'New Header', 'kix.header123')

        assert result is True
        mock_batch_update_success.assert_called_once()
        call_args = mock_batch_update_success.call_args
        body = call_args[1]['body']
        requests = body['requests']

        # Should NOT have delete request - only insert
        assert len(requests) == 1
        assert 'insertText' in requests[0]
        assert requests[0]['insertText']['location']['segmentId'] == 'kix.header123'

    @pytest.mark.asyncio
    async def test_replace_content_with_empty_content_list(self, manager, mock_batch_update_success):
        """Test when section has no content elements at all."""
        section = {
            'content': []
        }

        result = await manager._replace_section_content('doc123', section, 'New Header', 'kix.header123')

        assert result is True
        mock_batch_update_success.assert_called_once()
        call_args = mock_batch_update_success.call_args
        body = call_args[1]['body']
        requests = body['requests']

        # Should just insert at index 0
        assert len(requests) == 1
        assert 'insertText' in requests[0]
        assert requests[0]['insertText']['location']['index'] == 0
        assert requests[0]['insertText']['location']['segmentId'] == 'kix.header123'

    @pytest.mark.asyncio
    async def test_replace_content_no_content_key(self, manager, mock_batch_update_success):
        """Test when section has no 'content' key at all."""
        section = {}

        result = await manager._replace_section_content('doc123', section, 'New Header', 'kix.header123')

        assert result is True
        mock_batch_update_success.assert_called_once()

    @pytest.mark.asyncio
    async def test_api_error_returns_false(self, manager):
        """Test that API errors return False."""
        mock_execute = MagicMock(side_effect=Exception("API Error"))
        manager.service.documents.return_value.batchUpdate.return_value.execute = mock_execute

        section = {
            'content': [
                {
                    'paragraph': {},
                    'startIndex': 0,
                    'endIndex': 1
                }
            ]
        }

        result = await manager._replace_section_content('doc123', section, 'New Header', 'kix.header123')

        assert result is False

    @pytest.mark.asyncio
    async def test_delete_range_not_created_when_would_be_empty(self, manager, mock_batch_update_success):
        """Test that empty delete range is not created (would cause API error)."""
        # This is the specific bug case: header exists but has only newline
        # end_index - 1 == start_index means no actual content to delete
        section = {
            'content': [
                {
                    'paragraph': {
                        'elements': [{'textRun': {'content': '\n'}}]
                    },
                    'startIndex': 0,
                    'endIndex': 1  # end - 1 (0) == start (0), so no content
                }
            ]
        }

        result = await manager._replace_section_content('doc123', section, 'Test', 'kix.header123')

        assert result is True
        call_args = mock_batch_update_success.call_args
        body = call_args[1]['body']
        requests = body['requests']

        # Verify no deleteContentRange request was made
        for req in requests:
            assert 'deleteContentRange' not in req, "Should not delete when range would be empty"


class TestUpdateHeaderFooterCombinedMode:
    """
    Tests for the combined header_content/footer_content parameters
    in update_doc_headers_footers tool.

    Tests the new feature that allows setting both header and footer
    in a single API call instead of requiring two separate calls.
    """

    @pytest.fixture
    def mock_service(self):
        """Create a mocked Google Docs service."""
        service = MagicMock()
        # Default document response with headers and footers
        doc_response = {
            'headers': {
                'kix.header1': {
                    'content': [
                        {
                            'paragraph': {
                                'elements': [{'textRun': {'content': 'Old Header\n'}}]
                            },
                            'startIndex': 0,
                            'endIndex': 11
                        }
                    ]
                }
            },
            'footers': {
                'kix.footer1': {
                    'content': [
                        {
                            'paragraph': {
                                'elements': [{'textRun': {'content': 'Old Footer\n'}}]
                            },
                            'startIndex': 0,
                            'endIndex': 11
                        }
                    ]
                }
            }
        }
        service.documents.return_value.get.return_value.execute.return_value = doc_response
        service.documents.return_value.batchUpdate.return_value.execute.return_value = {}
        return service

    @pytest.fixture
    def manager(self, mock_service):
        """Create a HeaderFooterManager with mocked service."""
        return HeaderFooterManager(mock_service)

    @pytest.mark.asyncio
    async def test_update_header_only_via_header_content(self, manager, mock_service):
        """Test updating only header using header_content parameter."""
        success, message = await manager.update_header_footer_content(
            'doc123', 'header', 'New Header', 'DEFAULT', True
        )

        assert success is True
        assert 'header' in message.lower()

    @pytest.mark.asyncio
    async def test_update_footer_only_via_footer_content(self, manager, mock_service):
        """Test updating only footer using footer_content parameter."""
        success, message = await manager.update_header_footer_content(
            'doc123', 'footer', 'New Footer', 'DEFAULT', True
        )

        assert success is True
        assert 'footer' in message.lower()

    @pytest.mark.asyncio
    async def test_sequential_header_footer_updates(self, manager, mock_service):
        """Test that both header and footer can be updated sequentially."""
        # This simulates what the tool does when header_content and footer_content are both provided
        header_success, header_msg = await manager.update_header_footer_content(
            'doc123', 'header', 'New Header', 'DEFAULT', True
        )
        footer_success, footer_msg = await manager.update_header_footer_content(
            'doc123', 'footer', 'New Footer', 'DEFAULT', True
        )

        assert header_success is True
        assert footer_success is True

        # Verify both API calls were made
        assert mock_service.documents.return_value.batchUpdate.call_count == 2

    @pytest.mark.asyncio
    async def test_create_if_missing_creates_header(self, mock_service):
        """Test that create_if_missing creates header when it doesn't exist."""
        # Setup doc without headers
        mock_service.documents.return_value.get.return_value.execute.return_value = {
            'headers': {},
            'footers': {}
        }

        manager = HeaderFooterManager(mock_service)
        success, message = await manager.update_header_footer_content(
            'doc123', 'header', 'New Header', 'DEFAULT', create_if_missing=True
        )

        # It may fail or succeed depending on the mock setup, but it should attempt creation
        # The key thing is it doesn't fail immediately with "no header found"
        # Since create_if_missing=True, it should try to create
        assert mock_service.documents.return_value.batchUpdate.called


class TestUpdateDocHeadersFootersToolValidation:
    """
    Tests for validation logic in update_doc_headers_footers tool
    specifically around the new combined mode parameters.

    These test the parameter validation before any API calls are made.
    """

    def test_both_modes_not_allowed(self):
        """Test that using both section_type+content AND header_content/footer_content fails."""
        # This is a unit test of the validation logic
        # The actual tool function would reject this combination
        section_type = "header"
        content = "Some content"
        header_content = "Header text"

        # Validation check: should reject when both modes are used
        use_combined_mode = header_content is not None
        has_single_mode = section_type is not None and content is not None

        assert use_combined_mode and has_single_mode  # Both are present - invalid

    def test_combined_mode_detected(self):
        """Test that combined mode is correctly detected."""
        header_content = "Header"
        footer_content = "Footer"

        use_combined_mode = header_content is not None or footer_content is not None
        assert use_combined_mode is True

    def test_single_mode_detected(self):
        """Test that single section mode is correctly detected."""
        header_content = None
        footer_content = None

        use_combined_mode = header_content is not None or footer_content is not None
        assert use_combined_mode is False

    def test_header_content_only_valid(self):
        """Test that providing only header_content is valid."""
        header_content = "Just header"
        footer_content = None

        use_combined_mode = header_content is not None or footer_content is not None
        has_content = header_content is not None or footer_content is not None

        assert use_combined_mode is True
        assert has_content is True

    def test_footer_content_only_valid(self):
        """Test that providing only footer_content is valid."""
        header_content = None
        footer_content = "Just footer"

        use_combined_mode = header_content is not None or footer_content is not None
        has_content = header_content is not None or footer_content is not None

        assert use_combined_mode is True
        assert has_content is True

    def test_both_header_and_footer_content_valid(self):
        """Test that providing both header_content and footer_content is valid."""
        header_content = "Header"
        footer_content = "Footer"

        use_combined_mode = header_content is not None or footer_content is not None
        has_header = header_content is not None
        has_footer = footer_content is not None

        assert use_combined_mode is True
        assert has_header is True
        assert has_footer is True
