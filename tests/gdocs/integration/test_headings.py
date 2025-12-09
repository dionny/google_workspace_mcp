"""
Integration tests for heading and navigation operations in Google Docs.

Tests verify:
- Creating headings at different levels (H1-H6)
- Inserting text with heading styles
- Navigation and document structure
"""

import pytest
import json
import gdocs.docs_tools as docs_tools_module


@pytest.mark.asyncio
@pytest.mark.integration
async def test_create_heading_h1(user_google_email, test_document):
    """Test creating a Heading 1."""
    doc_id = test_document["document_id"]

    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Main Title",
        heading_style="HEADING_1",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should create H1: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_create_heading_h2(user_google_email, test_document):
    """Test creating a Heading 2."""
    doc_id = test_document["document_id"]

    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Section Title",
        heading_style="HEADING_2",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should create H2: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_create_heading_h3(user_google_email, test_document):
    """Test creating a Heading 3."""
    doc_id = test_document["document_id"]

    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Subsection Title",
        heading_style="HEADING_3",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should create H3: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_multiple_headings_structure(user_google_email, test_document):
    """Test creating a document with multiple heading levels."""
    doc_id = test_document["document_id"]

    # Create H1
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Document Title",
        heading_style="HEADING_1",
    )

    # Create H2
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="end",
        text="\nFirst Section",
        heading_style="HEADING_2",
    )

    # Create H3
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="end",
        text="\nSubsection A",
        heading_style="HEADING_3",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should create structure: {result}"
    )


@pytest.mark.asyncio
@pytest.mark.integration
async def test_heading_with_content(user_google_email, test_document):
    """Test creating heading followed by content."""
    doc_id = test_document["document_id"]

    # Create heading
    await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="start",
        text="Introduction",
        heading_style="HEADING_2",
    )

    # Add normal text after heading
    result = await docs_tools_module.modify_doc_text.fn(
        user_google_email=user_google_email,
        document_id=doc_id,
        location="end",
        text="\nThis is the introduction paragraph with some content.",
    )

    result_data = json.loads(result)
    assert result_data.get("success") or result_data.get("modified"), (
        f"Should add content: {result}"
    )
