"""
Integration tests for link and color operations in Google Docs.

Tests verify:
- Inserting hyperlinks
- Text color changes
- Background color
- Link formatting
"""

import pytest
import json
import gdocs.docs_tools as docs_tools_module


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_hyperlink(user_google_email, test_document):
    """Test inserting a hyperlink."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Click here to visit Google",
    )

    # Add link to "Click here"
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "Click here", "before_chars": 0, "after_chars": 0},
        link="https://www.google.com",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should add hyperlink: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_change_text_color(user_google_email, test_document):
    """Test changing text foreground color."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This text will be red",
    )

    # Change color to red
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "red", "before_chars": 0, "after_chars": 0},
        foreground_color="#ff0000",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should change text color: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_change_background_color(user_google_email, test_document):
    """Test changing text background color (highlight)."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This text will be highlighted",
    )

    # Add yellow highlight
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "highlighted", "extend": "paragraph"},
        background_color="#ffff00",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should change background color: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_colored_text(user_google_email, test_document):
    """Test inserting text with color already applied."""
    doc_id = test_document["document_id"]

    # Insert blue text directly
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This is blue text",
        foreground_color="#0000ff",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should insert colored text: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_strikethrough_formatting(user_google_email, test_document):
    """Test applying strikethrough formatting."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This text will be struck through",
    )

    # Apply strikethrough
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "struck through", "extend": "paragraph"},
        strikethrough=True,
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should apply strikethrough: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_font_family_change(user_google_email, test_document):
    """Test changing font family."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This text will use Arial",
    )

    # Change to Arial
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "Arial", "extend": "paragraph"},
        font_family="Arial",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should change font family: {result}"
    )
