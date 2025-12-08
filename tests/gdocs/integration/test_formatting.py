"""
Integration tests for text formatting operations in Google Docs.

Tests verify:
- Text styling (bold, italic, underline)
- Font changes
- Text color
- Multiple formatting combinations
"""
import pytest
import json
import gdocs.docs_tools as docs_tools_module


@pytest.mark.asyncio
@pytest.mark.integration
async def test_apply_italic_formatting(user_google_email, test_document):
    """Test applying italic formatting."""
    doc_id = test_document['document_id']
    
    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This text will be italic"
    )
    
    # Apply italic
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "This text will be italic", "extend": "paragraph"},
        italic=True
    )
    
    result_data = json.loads(result)
    assert result_data.get('success') or result_data.get('modified'), f"Should apply italic: {result}"


@pytest.mark.asyncio
@pytest.mark.integration
async def test_apply_underline_formatting(user_google_email, test_document):
    """Test applying underline formatting."""
    doc_id = test_document['document_id']
    
    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This text will be underlined"
    )
    
    # Apply underline
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "This text will be underlined", "extend": "paragraph"},
        underline=True
    )
    
    result_data = json.loads(result)
    assert result_data.get('success') or result_data.get('modified'), f"Should apply underline: {result}"


@pytest.mark.asyncio
@pytest.mark.integration
async def test_apply_multiple_formats(user_google_email, test_document):
    """Test applying multiple formats at once (bold + italic)."""
    doc_id = test_document['document_id']
    
    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This text will be bold and italic"
    )
    
    # Apply bold and italic together
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "bold and italic", "extend": "paragraph"},
        bold=True,
        italic=True
    )
    
    result_data = json.loads(result)
    assert result_data.get('success') or result_data.get('modified'), f"Should apply multiple formats: {result}"


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_text_with_formatting(user_google_email, test_document):
    """Test inserting text that is already formatted."""
    doc_id = test_document['document_id']
    
    # Insert text with bold applied immediately
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This is bold text from the start",
        bold=True
    )
    
    result_data = json.loads(result)
    assert result_data.get('success') or result_data.get('modified'), f"Should insert formatted text: {result}"


@pytest.mark.asyncio
@pytest.mark.integration
async def test_change_font_size(user_google_email, test_document):
    """Test changing font size."""
    doc_id = test_document['document_id']
    
    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This text will be large"
    )
    
    # Change font size
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "This text will be large", "extend": "paragraph"},
        font_size=18
    )
    
    result_data = json.loads(result)
    assert result_data.get('success') or result_data.get('modified'), f"Should change font size: {result}"


@pytest.mark.asyncio
@pytest.mark.integration
async def test_partial_text_formatting(user_google_email, test_document):
    """Test formatting only part of text using search with before/after chars."""
    doc_id = test_document['document_id']
    
    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Make only this word bold in the sentence"
    )
    
    # Format just one word using before_chars/after_chars
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "bold", "before_chars": 0, "after_chars": 0},
        bold=True
    )
    
    result_data = json.loads(result)
    assert result_data.get('success') or result_data.get('modified'), f"Should format partial text: {result}"

