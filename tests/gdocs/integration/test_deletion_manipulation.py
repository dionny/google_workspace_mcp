"""
Integration tests for deletion and text manipulation operations.

Tests verify:
- Deleting text sections
- Clearing formatting
- Text replacement with different lengths
- Empty text handling
"""

import pytest
import json
import gdocs.docs_tools as docs_tools_module


@pytest.mark.asyncio
@pytest.mark.integration
async def test_delete_section_by_search(user_google_email, test_document):
    """Test deleting a section of text found by search."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Keep this [DELETE ME] Keep this too",
    )

    # Delete the marked section
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        search="[DELETE ME]",
        position="replace",
        text="",  # Empty string deletes
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should delete section: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_replace_with_longer_text(user_google_email, test_document):
    """Test replacing short text with longer text."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="The cat sat",
    )

    # Replace with longer text
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        search="cat",
        position="replace",
        text="very fluffy and adorable cat",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should replace with longer text: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_replace_with_shorter_text(user_google_email, test_document):
    """Test replacing long text with shorter text."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="The quick brown fox jumps",
    )

    # Replace with shorter text
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        search="quick brown fox",
        position="replace",
        text="cat",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should replace with shorter text: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_multiple_text_insertions(user_google_email, test_document):
    """Test multiple sequential text insertions."""
    doc_id = test_document["document_id"]

    # Insert first text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="First. ",
    )

    # Insert second text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="end",
        text="Second. ",
    )

    # Insert third text
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="end",
        text="Third.",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should insert multiple texts: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_special_characters(user_google_email, test_document):
    """Test inserting text with special characters."""
    doc_id = test_document["document_id"]

    # Insert text with special characters
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Special chars: @#$%^&*() <> [] {} | \\ / ? ! ~ `",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should handle special chars: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_unicode_characters(user_google_email, test_document):
    """Test inserting Unicode characters."""
    doc_id = test_document["document_id"]

    # Insert Unicode text
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Unicode: 你好 🎉 café naïve résumé",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should handle Unicode: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_newline_handling(user_google_email, test_document):
    """Test inserting text with newlines."""
    doc_id = test_document["document_id"]

    # Insert multiline text
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Line 1\nLine 2\nLine 3\n\nLine 5",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should handle newlines: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_long_text(user_google_email, test_document):
    """Test inserting a longer block of text."""
    doc_id = test_document["document_id"]

    # Create a longer text block
    long_text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " * 10

    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text=long_text,
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should handle long text: {result}"
    )
