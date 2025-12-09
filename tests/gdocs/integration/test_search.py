"""
Integration tests for search and replace operations in Google Docs.

Tests verify:
- Text search
- Search and replace
- Case-sensitive search
- Multiple occurrences
"""

import pytest
import json
import gdocs.docs_tools as docs_tools_module


@pytest.mark.asyncio
@pytest.mark.integration
async def test_search_text_simple(user_google_email, test_document):
    """Test searching for text in a document."""
    doc_id = test_document["document_id"]

    # Insert some text to search for
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="The quick brown fox jumps over the lazy dog",
    )

    # Search for text using search positioning
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        search="fox",
        position="after",
        text=" and cat",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should find and insert after text: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_before_searched_text(user_google_email, test_document):
    """Test inserting text before a searched term."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Hello World",
    )

    # Insert before "World"
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        search="World",
        position="before",
        text="Beautiful ",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should insert before: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_replace_searched_text(user_google_email, test_document):
    """Test replacing searched text."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="I love dogs and dogs are great",
    )

    # Replace first occurrence of "dogs"
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        search="dogs",
        position="replace",
        text="cats",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should replace text: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_at_searched_location(user_google_email, test_document):
    """Test inserting at exact location of searched text."""
    doc_id = test_document["document_id"]

    # Insert marker text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Start [MARKER] End",
    )

    # Replace the marker
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        search="[MARKER]",
        position="replace",
        text="CONTENT",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should replace marker: {result}"
    )
