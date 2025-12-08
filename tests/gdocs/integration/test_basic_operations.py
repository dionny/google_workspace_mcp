"""
Integration tests for basic Google Docs operations.

These tests verify the core functionality works with real API calls:
- Creating and inserting text
- Basic list creation
- Simple formatting
"""
import pytest
import json
import gdocs.docs_tools as docs_tools_module


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_text_at_start(user_google_email, test_document):
    """Test inserting plain text at start of document."""
    doc_id = test_document['document_id']
    
    # Insert some text
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Hello World from Integration Test!"
    )
    
    result_data = json.loads(result)
    assert result_data.get('success') or result_data.get('modified'), f"Insert should succeed: {result}"


@pytest.mark.asyncio
@pytest.mark.integration
async def test_create_simple_bullet_list(user_google_email, docs_service, test_document):
    """Test creating a basic bullet list."""
    doc_id = test_document['document_id']
    
    # Insert text and convert to bullet list
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="First item\nSecond item\nThird item",
        convert_to_list="bullet"
    )
    
    result_data = json.loads(result)
    assert result_data.get('success') or result_data.get('modified'), f"Should create list: {result}"


@pytest.mark.asyncio
@pytest.mark.integration
async def test_create_numbered_list(user_google_email, docs_service, test_document):
    """Test creating a numbered list."""
    doc_id = test_document['document_id']
    
    # Insert text and convert to numbered list
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Step one\nStep two\nStep three",
        convert_to_list="numbered"
    )
    
    result_data = json.loads(result)
    assert result_data.get('success') or result_data.get('modified'), f"Should succeed: {result}"


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_text_at_end(user_google_email, test_document):
    """Test inserting text at end of document."""
    doc_id = test_document['document_id']
    
    # Insert at start first
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Beginning text"
    )
    
    # Now insert at end
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="end",
        text="\nEnding text"
    )
    
    result_data = json.loads(result)
    assert result_data.get('success') or result_data.get('modified'), f"Should succeed: {result}"


@pytest.mark.asyncio
@pytest.mark.integration
async def test_apply_bold_formatting(user_google_email, docs_service, test_document):
    """Test applying bold formatting to text using search."""
    doc_id = test_document['document_id']
    
    # Insert text first
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This text will be bold"
    )
    
    # Apply bold using search and extend
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "This text will be bold", "extend": "paragraph"},
        bold=True
    )
    
    result_data = json.loads(result)
    assert result_data.get('success') or result_data.get('modified'), f"Format should succeed: {result}"

