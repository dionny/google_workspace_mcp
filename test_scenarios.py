#!/usr/bin/env python3
"""
Manual Testing Script - Final edge cases
"""
import asyncio
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

TEST_DOC_ID = "1A-N-g_mgnDbj7TwcOO4sHUUbtgdAJ_T2VHh9-pChqCc"
EMAIL = "mbradshaw@indeed.com"


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
            import traceback
            print(f"Error calling {tool_name}: {e}")
            traceback.print_exc()
            return None

    print("=" * 60)
    print("TEST 47: line_spacing with correct values (100, 200)")
    print("=" * 60)
    
    result = await call("modify_doc_text",
                        search="=== Long Text Test ===",
                        position="replace",
                        text="=== Long Text Test ===",
                        line_spacing=200)
    print("Line spacing 200 (double):", result)
    
    print("\n" + "=" * 60)
    print("TEST 48: Test subscript and superscript")
    print("=" * 60)
    
    result = await call("modify_doc_text",
                        location="end",
                        text="\n\nH2O uses subscript",
                        subscript=True)
    print("Subscript:", result)
    
    result = await call("modify_doc_text",
                        location="end",
                        text="\nE=mc2 uses superscript",
                        superscript=True)
    print("Superscript:", result)
    
    print("\n" + "=" * 60)
    print("TEST 49: Invalid document ID")
    print("=" * 60)
    
    result = await call("modify_doc_text",
                        document_id="INVALID_DOC_ID_12345",
                        location="end",
                        text="test")
    print("Invalid doc ID:", result)
    
    print("\n" + "=" * 60)
    print("TEST 50: Extract links from document")
    print("=" * 60)
    
    result = await call("extract_links")
    print("Extract links:", result)
    
    print("\n" + "=" * 60)
    print("TEST 51: Get document section by heading")
    print("=" * 60)
    
    result = await call("get_doc_section",
                        heading="The Problem")
    print("Get section:", result)

if __name__ == "__main__":
    asyncio.run(main())
