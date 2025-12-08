#!/usr/bin/env python3
"""
Test runner for systematic testing of Google Docs MCP Tools
"""
import asyncio
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

TEST_DOC_ID = os.environ.get("TEST_DOC_ID", "YOUR_TEST_DOC_ID")
EMAIL = os.environ.get("USER_GOOGLE_EMAIL", "user@example.com")

async def main():
    from tools_cli import init_server, ToolTester

    server = init_server()
    tester = ToolTester(server)
    await tester.init_tools()

    async def call(tool_name, **kwargs):
        """Call a tool with email and doc_id automatically added"""
        if "user_google_email" not in kwargs:
            kwargs["user_google_email"] = EMAIL
        if "document_id" not in kwargs and tool_name not in ["search_docs", "create_doc"]:
            kwargs["document_id"] = TEST_DOC_ID

        tool = tester.tools.get(tool_name)
        if not tool:
            print(f"Tool {tool_name} not found!")
            return None

        try:
            if hasattr(tool, "fn"):
                result = await tool.fn(**kwargs)
            else:
                result = await tool(**kwargs)
            return result
        except Exception as e:
            print(f"Error calling {tool_name}: {e}")
            import traceback
            traceback.print_exc()
            return None

    # ============================================================
    # CLEANUP: Remove test artifacts from document
    # ============================================================
    print("=" * 60)
    print("CLEANUP: Removing test artifacts from document")
    print("=" * 60)

    cleanup_markers = [
        "[REPLACED-START-MARKER]",
        "[BFORMATTED]",
        "[BATCH1]",
        "[BATCH2]",
        "[TEST-142858]",
        "[case-insensitive-match]",
        "[2nd-occurrence]",
        "Test Heading H3 [INSERTED-AFTER]",
        "[HYPERLINK TEST] Click here",
        "[RANGE-DELETE-TEST]",
        "[DELETE-ME-MARKER]",
    ]

    for marker in cleanup_markers:
        result = await call("modify_doc_text", search=marker, position="replace", text="")
        if result and "success" in result and '"success": true' in result:
            print(f"Cleaned: {marker}")
        else:
            print(f"Could not clean '{marker}' (may not exist)")

    print("\nCleanup complete!")

if __name__ == "__main__":
    asyncio.run(main())
