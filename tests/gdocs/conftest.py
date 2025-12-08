"""
Pytest configuration and fixtures for Google Docs integration tests.

This module provides fixtures for creating and cleaning up test documents.
All test documents are:
- Named with "[TEST]" prefix for easy identification
- Automatically trashed after test completion (pass or fail)
- Created fresh for each test to ensure isolation
"""
import pytest
import asyncio
import uuid
import os
from typing import Any, Dict
from unittest.mock import Mock

# Import the actual tools we need (not the wrapped versions)
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../..'))

# Import the core modules to get the actual functions
import gdocs.docs_tools as docs_tools_module
import gdrive.drive_tools as drive_tools_module
from auth.google_auth import get_authenticated_google_service


@pytest.fixture(scope="session")
def user_google_email():
    """
    Get the user's Google email for testing.
    
    Set via GOOGLE_TEST_EMAIL environment variable.
    """
    email = os.environ.get("GOOGLE_TEST_EMAIL")
    if not email:
        pytest.skip("GOOGLE_TEST_EMAIL environment variable not set")
    return email


@pytest.fixture(scope="session")
def event_loop():
    """Create an event loop for the session."""
    loop = asyncio.get_event_loop_policy().new_event_loop()
    yield loop
    loop.close()


@pytest.fixture
async def docs_service(user_google_email):
    """Get authenticated Google Docs service."""
    service = await get_authenticated_google_service(
        service_name="docs",
        version="v1",
        tool_name="test_docs_service",
        user_google_email=user_google_email,
        required_scopes=["https://www.googleapis.com/auth/documents"]
    )
    return service


@pytest.fixture
async def drive_service(user_google_email):
    """Get authenticated Google Drive service."""
    service = await get_authenticated_google_service(
        service_name="drive",
        version="v3",
        tool_name="test_drive_service",
        user_google_email=user_google_email,
        required_scopes=["https://www.googleapis.com/auth/drive.file"]
    )
    return service


@pytest.fixture
async def test_document(user_google_email, docs_service, drive_service):
    """
    Create a fresh test document for each test.
    
    The document is:
    - Named with [TEST] prefix and unique ID
    - Automatically trashed after the test completes (pass or fail)
    - Empty by default (tests can populate as needed)
    
    Yields:
        dict: Contains 'document_id' and 'title'
    """
    # Generate unique test name
    test_id = str(uuid.uuid4())[:8]
    title = f"[TEST] GDocs Integration Test {test_id}"
    
    # Create the document - don't pass service, it's injected by decorator
    result = await docs_tools_module.create_doc.fn(
        user_google_email=user_google_email,
        title=title
    )
    
    # Parse the result - it's a string message, extract the ID
    # Format: "Created Google Doc 'Title' (ID: doc_id) for email. Link: url"
    import re
    match = re.search(r'\(ID: ([a-zA-Z0-9_-]+)\)', result)
    if not match:
        raise ValueError(f"Could not extract document ID from result: {result}")
    document_id = match.group(1)
    
    print(f"\n✓ Created test document: {title} (ID: {document_id})")
    
    # Yield to the test
    yield {
        'document_id': document_id,
        'title': title
    }
    
    # Cleanup: Always trash the document, even if test failed
    try:
        await drive_tools_module.update_drive_file.fn(
            user_google_email=user_google_email,
            file_id=document_id,
            trashed=True
        )
        print(f"✓ Trashed test document: {document_id}")
    except Exception as e:
        print(f"⚠ Warning: Failed to trash test document {document_id}: {e}")


@pytest.fixture
async def populated_test_document(user_google_email, docs_service, drive_service):
    """
    Create a test document with common test content structure.
    
    The document includes:
    - Multiple heading levels
    - Paragraphs with various formatting
    - Lists (bullet and numbered)
    - Tables
    - Various text styles
    
    This is useful for tests that need to query/modify existing content.
    
    Yields:
        dict: Contains 'document_id', 'title', and 'structure' info
    """
    # Generate unique test name
    test_id = str(uuid.uuid4())[:8]
    title = f"[TEST] Populated Test Doc {test_id}"
    
    # Create the document - don't pass service, it's injected by decorator
    result = await docs_tools_module.create_doc.fn(
        user_google_email=user_google_email,
        title=title
    )
    
    # Parse the result - extract the document ID
    import re
    match = re.search(r'\(ID: ([a-zA-Z0-9_-]+)\)', result)
    if not match:
        raise ValueError(f"Could not extract document ID from result: {result}")
    document_id = match.group(1)
    
    # TODO: Populate with standard test content
    # For now, just return empty doc - we'll add population in future
    
    print(f"\n✓ Created populated test document: {title} (ID: {document_id})")
    
    structure = {
        'heading_positions': {},
        'list_positions': {},
        'table_positions': {}
    }
    
    yield {
        'document_id': document_id,
        'title': title,
        'structure': structure
    }
    
    # Cleanup
    try:
        await drive_tools_module.update_drive_file.fn(
            user_google_email=user_google_email,
            file_id=document_id,
            trashed=True
        )
        print(f"✓ Trashed populated test document: {document_id}")
    except Exception as e:
        print(f"⚠ Warning: Failed to trash test document {document_id}: {e}")


# Helper function for tests to get document content
async def get_document(docs_service, document_id: str) -> Dict[str, Any]:
    """
    Fetch the full document from Google Docs API.
    
    Args:
        docs_service: Authenticated Docs service
        document_id: ID of document to fetch
        
    Returns:
        dict: Full document structure from API
    """
    result = await asyncio.to_thread(
        docs_service.documents().get(documentId=document_id).execute
    )
    return result


# Helper function to execute batch updates
async def batch_update(docs_service, document_id: str, requests: list) -> Dict[str, Any]:
    """
    Execute a batch of updates on a document.
    
    Args:
        docs_service: Authenticated Docs service
        document_id: ID of document to update
        requests: List of update request objects
        
    Returns:
        dict: API response
    """
    body = {'requests': requests}
    return await asyncio.to_thread(
        docs_service.documents().batchUpdate(
            documentId=document_id,
            body=body
        ).execute
    )

