#!/usr/bin/env python3
"""Manual test script for gdocs tools."""
import asyncio
import sys
import os

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dotenv import load_dotenv  # noqa: E402
load_dotenv()

# Import after path setup
from tools_cli import init_server, ToolTester  # noqa: E402

DOC_ID = "1A-N-g_mgnDbj7TwcOO4sHUUbtgdAJ_T2VHh9-pChqCc"
USER_EMAIL = "mbradshaw@indeed.com"

async def main():
    """Run manual tests."""
    print("Initializing server...")
    server = init_server()
    tester = ToolTester(server)
    await tester.init_tools()
    print(f"Server initialized ({len(tester.tools)} tools loaded)")

    # Test escape sequences like \n in search text
    print("\n=== Testing escape sequences in search ===")
    # First add text with newlines
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text="\n\n=== NEWLINE TEST ===\nLine 1\nLine 2\nLine 3\n"
    )
    print(f"Insert with newlines result: {result}")

    # Now try to search for text that spans lines (with newline in search)
    # Use modify_doc_text with preview=True instead of removed preview_search_results
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        search="Line 1\nLine 2",
        preview=True
    )
    print(f"Search with newline result: {result}")

    # Test: Can we insert text AFTER a specific section heading?
    print("\n=== Testing insert after heading ===")
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        search="NEWLINE TEST",
        position="after",
        text="\n[Text inserted AFTER the heading]\n"
    )
    print(f"Insert after heading result: {result}")

    # Test: Delete an entire section (heading + content)
    print("\n=== Testing delete section by range ===")
    # First, let's find a section to delete using modify_doc_text with preview=True
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        search="=== NEWLINE TEST ===",
        preview=True
    )
    print(f"Found section: {result}")

    # Test: Inserting text with indentation (tabs/spaces preserved?)
    print("\n=== Testing indentation preservation ===")
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text="\n=== INDENTATION TEST ===\nNormal line\n    4-space indent\n        8-space indent\n\tTab indent\n\t\tDouble tab\n"
    )
    print(f"Indentation test result: {result}")

    # Test: Removing formatting (making bold text NOT bold)
    print("\n=== Testing remove formatting (bold=False) ===")
    # First make some text bold
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text="\n=== REMOVE FORMAT TEST ===\nThis text will be made bold then un-bold.\n"
    )
    print(f"Added text: {result}")

    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        search="made bold then un-bold",
        position="replace",
        text="made bold then un-bold",
        bold=True
    )
    print(f"Made bold: {result}")

    # Now try to remove the bold
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        search="made bold then un-bold",
        position="replace",
        text="made bold then un-bold",
        bold=False
    )
    print(f"Removed bold: {result}")

    # Test inserting a horizontal rule/line
    print("\n=== Testing horizontal rule insertion ===")
    result = await tester.call_tool(
        "insert_doc_elements",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        element_type="horizontal_rule",
        location="end"
    )
    print(f"Horizontal rule result: {result}")

    # Test edge case: empty text insertion
    print("\n=== Testing empty text handling ===")
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text=""
    )
    print(f"Empty text result: {result}")

    # Test edge case: special characters and unicode
    print("\n=== Testing unicode/special characters ===")
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text="\n=== SPECIAL CHARS TEST ===\nEmojis: 🎉 🚀 ✅ 👍\nMath: ∑ ∆ √ π ∞\nCurrency: $ € £ ¥\nLanguages: こんにちは 你好 مرحبا\n"
    )
    print(f"Unicode/special chars result: {result}")

    # Test edge case: very long text
    print("\n=== Testing long text insertion ===")
    long_text = "A" * 5000  # 5000 character string
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text=f"\n=== LONG TEXT TEST (5KB) ===\n{long_text}\n"
    )
    print(f"Long text result: {result}")

    # Test get_doc_section - retrieve a specific section
    print("\n=== Testing get_doc_section ===")
    result = await tester.call_tool(
        "get_doc_section",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        heading="The Problem"
    )
    print(f"Get section result: {result[:500] if len(result) > 500 else result}")

    # Test navigate_heading_siblings - navigate between headings
    print("\n=== Testing navigate_heading_siblings ===")
    result = await tester.call_tool(
        "navigate_heading_siblings",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        current_heading="The Problem",
        direction="next"
    )
    print(f"Navigate siblings result: {result}")

    # Test auto_linkify_doc - auto-detect URLs
    print("\n=== Testing auto_linkify_doc ===")
    # First add some text with URLs
    await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text="\n=== AUTO LINKIFY TEST ===\nVisit https://www.example.com for more info.\nAlso check out https://www.google.com and https://www.github.com\n"
    )
    # Now auto-linkify them
    result = await tester.call_tool(
        "auto_linkify_doc",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
    )
    print(f"Auto linkify result: {result}")

    # Test headers
    print("\n=== Testing update_doc_headers_footers (header) ===")
    result = await tester.call_tool(
        "update_doc_headers_footers",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        section_type="header",
        content="Test Document Header"
    )
    print(f"Header result: {result}")

    # Test footers
    print("\n=== Testing update_doc_headers_footers (footer) ===")
    result = await tester.call_tool(
        "update_doc_headers_footers",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        section_type="footer",
        content="Page Footer - Confidential"
    )
    print(f"Footer result: {result}")

    # Test heading style formatting
    print("\n=== Testing heading style ===")
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text="\n\nNew Section Title\nThis is paragraph text under the new section.\n"
    )
    print(f"Text insertion result: {result}")

    # Now make "New Section Title" a heading
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        search="New Section Title",
        position="replace",
        text="New Section Title",
        heading_style="HEADING_2"
    )
    print(f"Heading style result: {result}")

    # Test alignment
    print("\n=== Testing alignment ===")
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text="\nThis text should be centered.\n"
    )
    print(f"Text insertion result: {result}")

    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        search="This text should be centered",
        position="replace",
        text="This text should be centered",
        alignment="CENTER"
    )
    print(f"Alignment result: {result}")

    # Test creating a hyperlink
    print("\n=== Testing hyperlink creation ===")
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text="\nClick here to visit Google\n",
        link="https://www.google.com"
    )
    print(f"Hyperlink result: {result}")

    # Now convert to bullet list using the convert_to_list parameter
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        search="First list item\nSecond list item\nThird list item",
        position="replace",
        text="First list item\nSecond list item\nThird list item",
        convert_to_list="UNORDERED"
    )
    print(f"Convert to bullet list result: {result}")

    # First add section header
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text="\n\n=== NUMBERED LIST TEST ==="
    )
    print(f"Section header result: {result}")

    # Test insert_doc_elements for numbered list
    print("\n=== Testing insert_doc_elements for numbered list ===")
    result = await tester.call_tool(
        "insert_doc_elements",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        element_type="list",
        location="end",
        list_type="ORDERED",
        items=["Step 1: Do something", "Step 2: Do something else", "Step 3: Finish"]
    )
    print(f"Insert numbered list result: {result}")

    # Test code block formatting
    print("\n=== Testing code block formatting ===")
    result = await tester.call_tool(
        "modify_doc_text",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        location="end",
        text="\n\n=== CODE BLOCK TEST ===\ndef hello():\n    print('Hello World')\n    return True\n",
        code_block=True
    )
    print(f"Code block result: {result}")

    # Test inserting a page break
    print("\n=== Testing page break ===")
    result = await tester.call_tool(
        "insert_doc_elements",
        document_id=DOC_ID,
        user_google_email=USER_EMAIL,
        element_type="page_break",
        location="end"
    )
    print(f"Page break result: {result}")

if __name__ == "__main__":
    asyncio.run(main())
