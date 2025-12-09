"""
Integration tests for alignment and paragraph formatting in Google Docs.

Tests verify:
- Text alignment (left, center, right)
- Line spacing
- Indentation
- Paragraph spacing
"""

import pytest
import json
import gdocs.docs_tools as docs_tools_module


@pytest.mark.asyncio
@pytest.mark.integration
async def test_center_alignment(user_google_email, test_document):
    """Test center-aligning text."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This text will be centered",
    )

    # Center align
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "centered", "extend": "paragraph"},
        alignment="CENTER",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should center align: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_right_alignment(user_google_email, test_document):
    """Test right-aligning text."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This text will be right-aligned",
    )

    # Right align
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "right-aligned", "extend": "paragraph"},
        alignment="END",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should right align: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_line_spacing(user_google_email, test_document):
    """Test changing line spacing."""
    doc_id = test_document["document_id"]

    # Insert multiline text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Line one\nLine two\nLine three",
    )

    # Change line spacing
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "Line", "extend": "paragraph"},
        line_spacing=1.5,
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should change line spacing: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_indent_first_line(user_google_email, test_document):
    """Test indenting first line of paragraph."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This paragraph will have first line indented",
    )

    # Indent first line
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "paragraph", "extend": "paragraph"},
        indent_first_line=36.0,  # 0.5 inch in points
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should indent first line: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_paragraph_start_indent(user_google_email, test_document):
    """Test indenting entire paragraph from left."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="This paragraph will be indented from the left",
    )

    # Indent from left
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "indented", "extend": "paragraph"},
        indent_start=72.0,  # 1 inch in points
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should indent paragraph: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_paragraph_spacing(user_google_email, test_document):
    """Test adding space above and below paragraphs."""
    doc_id = test_document["document_id"]

    # Insert text
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Paragraph with spacing",
    )

    # Add space above and below
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        range={"search": "spacing", "extend": "paragraph"},
        space_above=12.0,
        space_below=12.0,
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should add paragraph spacing: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_insert_with_formatting_preset(user_google_email, test_document):
    """Test inserting text with multiple formatting options at once."""
    doc_id = test_document["document_id"]

    # Insert formatted text
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Fancy formatted text",
        bold=True,
        italic=True,
        font_size=14,
        foreground_color="#ff6600",
        alignment="CENTER",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should insert with preset formatting: {result}"
    )
