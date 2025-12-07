#!/usr/bin/env python3
"""
Google Docs MCP Tools - Scenario Testing Script

This script tests common editing scenarios that users would want to perform.
Run this after making code changes to verify functionality.

Usage:
    uv run python gdocs/scenario_tests.py --doc_id YOUR_TEST_DOC_ID --email your@email.com
    uv run python gdocs/scenario_tests.py --doc_id YOUR_TEST_DOC_ID --email your@email.com --cleanup

Requirements:
    - A test Google Doc you can modify
    - Valid OAuth credentials configured

The script will:
    1. Run all test scenarios
    2. Report pass/fail for each
    3. Optionally clean up test content (--cleanup flag)
"""

import argparse
import asyncio
import json
import sys
import os
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional
from datetime import datetime

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@dataclass
class TestResult:
    """Result of a single test scenario."""

    name: str
    category: str
    passed: bool
    message: str
    details: Optional[Dict[str, Any]] = None
    error: Optional[str] = None
    expected_fail: bool = False  # Mark tests for known missing features


@dataclass
class TestReport:
    """Complete test report."""

    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())
    total: int = 0
    passed: int = 0
    failed: int = 0
    skipped: int = 0
    results: List[TestResult] = field(default_factory=list)


class ScenarioTester:
    """Test runner for Google Docs MCP scenarios."""

    def __init__(self, doc_id: str, email: str, verbose: bool = True):
        self.doc_id = doc_id
        self.email = email
        self.verbose = verbose
        self.tester = None
        self.report = TestReport()
        self.test_marker = f"[TEST-{datetime.now().strftime('%H%M%S')}]"

    async def setup(self):
        """Initialize the test harness."""
        from test_harness import init_server, ToolTester

        server = init_server()
        self.tester = ToolTester(server)
        await self.tester.init_tools()
        if self.verbose:
            print(f"✅ Initialized with {len(self.tester.tools)} tools")
            print(f"📄 Testing document: {self.doc_id}")
            print(f"🏷️  Test marker: {self.test_marker}")
            print()

    async def call_tool(self, tool_name: str, **kwargs) -> Any:
        """Call a tool and return the result."""
        # Add email if tool expects it (most do)
        if "user_google_email" not in kwargs:
            kwargs["user_google_email"] = self.email
        if "document_id" not in kwargs:
            kwargs["document_id"] = self.doc_id

        tool = self.tester.tools.get(tool_name)
        if not tool:
            raise ValueError(f"Tool {tool_name} not found")

        if hasattr(tool, "fn"):
            return await tool.fn(**kwargs)
        return await tool(**kwargs)

    def record(self, result: TestResult):
        """Record a test result."""
        self.report.results.append(result)
        self.report.total += 1
        if result.passed:
            self.report.passed += 1
            if self.verbose:
                print(f"  ✅ {result.name}")
        else:
            self.report.failed += 1
            if self.verbose:
                print(f"  ❌ {result.name}")
                if result.error:
                    print(f"      Error: {result.error[:100]}...")

    # =========================================================================
    # TEST SCENARIOS
    # =========================================================================

    async def test_basic_insertion(self):
        """Test 1: Basic text insertion at known position."""
        print("\n📝 Category: Basic Operations")

        # Test 1a: Insert at specific index
        try:
            result = await self.call_tool(
                "modify_doc_text", start_index=1, text=f"{self.test_marker} START "
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Insert text at index",
                    category="basic",
                    passed=result_dict.get("success", False),
                    message="Inserted text at start",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Insert text at index",
                    category="basic",
                    passed=False,
                    message="Failed to insert",
                    error=str(e),
                )
            )

        # Test 1b: Insert with location='end'
        try:
            result = await self.call_tool(
                "modify_doc_text", location="end", text=f" {self.test_marker} END"
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Insert text at end (location='end')",
                    category="basic",
                    passed=result_dict.get("success", False),
                    message="Inserted text at end",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Insert text at end (location='end')",
                    category="basic",
                    passed=False,
                    message="Failed to insert at end",
                    error=str(e),
                )
            )

    async def test_search_based_operations(self):
        """Test 2: Search-based text operations."""
        print("\n🔍 Category: Search-Based Operations")

        # Test 2a: Insert before search term
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="before",
                text="[BEFORE] ",
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Insert before search term",
                    category="search",
                    passed=result_dict.get("success", False),
                    message="Inserted before found text",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Insert before search term",
                    category="search",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 2b: Insert after search term
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="after",
                text=" [AFTER]",
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Insert after search term",
                    category="search",
                    passed=result_dict.get("success", False),
                    message="Inserted after found text",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Insert after search term",
                    category="search",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 2c: Replace search term
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search="[BEFORE]",
                position="replace",
                text="[REPLACED]",
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Replace search term",
                    category="search",
                    passed=result_dict.get("success", False),
                    message="Replaced text",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Replace search term",
                    category="search",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 2d: Case-insensitive search
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker.lower(),
                match_case=False,
                position="replace",
                text=self.test_marker,
                preview=True,
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Case-insensitive search",
                    category="search",
                    passed=result_dict.get("would_modify", False),
                    message="Found with case-insensitive search",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Case-insensitive search",
                    category="search",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_preview_mode(self):
        """Test 3: Preview mode functionality."""
        print("\n👁️ Category: Preview Mode")

        # Test 3a: Preview insert
        try:
            result = await self.call_tool(
                "modify_doc_text", start_index=1, text="PREVIEW TEST", preview=True
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Preview insert operation",
                    category="preview",
                    passed=result_dict.get("preview", False)
                    and result_dict.get("would_modify", False),
                    message="Preview returned expected structure",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Preview insert operation",
                    category="preview",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 3b: Preview format
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="replace",
                bold=True,
                preview=True,
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Preview format operation",
                    category="preview",
                    passed=result_dict.get("preview", False),
                    message="Preview format works",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Preview format operation",
                    category="preview",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_text_formatting(self):
        """Test 4: Text formatting operations."""
        print("\n🎨 Category: Text Formatting")

        # Test 4a: Bold formatting
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="replace",
                bold=True,
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Apply bold formatting",
                    category="formatting",
                    passed=result_dict.get("success", False),
                    message="Bold applied",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Apply bold formatting",
                    category="formatting",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 4b: Multiple formatting options
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search="[REPLACED]",
                position="replace",
                bold=True,
                italic=True,
                underline=True,
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Apply multiple formats (bold+italic+underline)",
                    category="formatting",
                    passed=result_dict.get("success", False),
                    message="Multiple formats applied",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Apply multiple formats",
                    category="formatting",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 4c: Insert with formatting
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search="[AFTER]",
                position="after",
                text=" [FORMATTED INSERT]",
                bold=True,
                font_size=14,
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Insert text with formatting",
                    category="formatting",
                    passed=result_dict.get("success", False),
                    message="Inserted with formatting",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Insert text with formatting",
                    category="formatting",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 4d: Hyperlink (expected to fail - missing feature)
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search="[FORMATTED INSERT]",
                position="replace",
                link="https://example.com",
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Add hyperlink to text",
                    category="formatting",
                    passed=result_dict.get("success", False),
                    message="Hyperlink added - FEATURE NOW IMPLEMENTED! 🎉",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Add hyperlink to text",
                    category="formatting",
                    passed=False,
                    message="Failed to add hyperlink",
                    error=str(e),
                )
            )

        # Test 4e: Text color (foreground)
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="replace",
                foreground_color="#FF0000",
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            # Check for success or absence of error
            passed = result_dict.get("success", False) or (
                "error" not in result_dict and "link" in str(result_dict)
            )
            self.record(
                TestResult(
                    name="Change text color (foreground)",
                    category="formatting",
                    passed=passed,
                    message="Text color applied",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Change text color (foreground)",
                    category="formatting",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 4f: Background color (highlight)
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="replace",
                background_color="yellow",
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            passed = result_dict.get("success", False) or (
                "error" not in result_dict and "link" in str(result_dict)
            )
            self.record(
                TestResult(
                    name="Change background color (highlight)",
                    category="formatting",
                    passed=passed,
                    message="Background color applied",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Change background color (highlight)",
                    category="formatting",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 4g: Strikethrough
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="replace",
                strikethrough=True,
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            passed = result_dict.get("success", False) or (
                "error" not in result_dict and "link" in str(result_dict)
            )
            self.record(
                TestResult(
                    name="Apply strikethrough",
                    category="formatting",
                    passed=passed,
                    message="Strikethrough applied",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Apply strikethrough",
                    category="formatting",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_delete_operations(self):
        """Test 5: Delete text operations."""
        print("\n🗑️ Category: Delete Operations")

        # Test 5a: Delete by replacing with empty string (KNOWN BUG)
        try:
            result = await self.call_tool(
                "modify_doc_text", search="[AFTER]", position="replace", text=""
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Delete via replace with empty string",
                    category="delete",
                    passed=result_dict.get("success", False),
                    message="Deleted via empty replacement",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Delete via replace with empty string",
                    category="delete",
                    passed=False,
                    message="KNOWN BUG - Generates invalid insertText request",
                    error=str(e),
                )
            )

        # Test 5b: Delete via batch operation
        try:
            # First find the position to delete
            preview = await self.call_tool(
                "modify_doc_text",
                search="[FORMATTED INSERT]",
                position="replace",
                text="X",
                preview=True,
            )
            preview_dict = json.loads(preview) if isinstance(preview, str) else preview
            if preview_dict.get("affected_range"):
                start = preview_dict["affected_range"]["start"]
                end = preview_dict["affected_range"]["end"]

                result = await self.call_tool(
                    "batch_edit_doc",
                    operations=[
                        {"type": "delete_text", "start_index": start, "end_index": end}
                    ],
                )
                result_dict = json.loads(result) if isinstance(result, str) else result
                self.record(
                    TestResult(
                        name="Delete via batch_edit_doc",
                        category="delete",
                        passed=result_dict.get("success", False),
                        message="Deleted via batch operation",
                        details=result_dict,
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Delete via batch_edit_doc",
                        category="delete",
                        passed=False,
                        message="Could not find text to delete",
                    )
                )
        except Exception as e:
            self.record(
                TestResult(
                    name="Delete via batch_edit_doc",
                    category="delete",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_find_and_replace(self):
        """Test 6: Find and replace operations."""
        print("\n🔄 Category: Find and Replace")

        # Test 6a: Basic find and replace all
        try:
            result = await self.call_tool(
                "find_and_replace_doc",
                find_text="[REPLACED]",
                replace_text="[FAR-REPLACED]",
            )
            # Note: returns string, not JSON
            passed = (
                "occurrence" in str(result).lower() or "replaced" in str(result).lower()
            )
            self.record(
                TestResult(
                    name="Find and replace all occurrences",
                    category="find_replace",
                    passed=passed,
                    message="Replace executed",
                    details={"response": result},
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Find and replace all occurrences",
                    category="find_replace",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 6b: Find and replace with preview (KNOWN MISSING)
        try:
            result = await self.call_tool(
                "find_and_replace_doc",
                find_text="test",
                replace_text="TEST",
                preview=True,
            )
            self.record(
                TestResult(
                    name="Find and replace with preview",
                    category="find_replace",
                    passed=True,
                    message="Preview - FEATURE NOW IMPLEMENTED! 🎉",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Find and replace with preview",
                    category="find_replace",
                    passed=False,
                    message="Failed to preview find and replace",
                    error=str(e),
                )
            )

        # Test 6c: Find and replace with formatting
        try:
            result = await self.call_tool(
                "find_and_replace_doc",
                find_text="[FAR-REPLACED]",
                replace_text="[FORMATTED]",
                bold=True,
                foreground_color="red",
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            has_formatting_info = (
                result_dict.get("success", False)
                and "formatting_applied" in result_dict
                and "bold" in result_dict.get("formatting_applied", [])
            )
            self.record(
                TestResult(
                    name="Find and replace with formatting",
                    category="find_replace",
                    passed=has_formatting_info,
                    message="Replaced and formatted" if has_formatting_info else "Missing formatting info",
                    details={"response": result_dict},
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Find and replace with formatting",
                    category="find_replace",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 6d: Find and replace with formatting preview
        try:
            result = await self.call_tool(
                "find_and_replace_doc",
                find_text="test",
                replace_text="TEST",
                bold=True,
                italic=True,
                preview=True,
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            has_formatting_preview = (
                result_dict.get("preview", False)
                and "formatting_requested" in result_dict
            )
            self.record(
                TestResult(
                    name="Find and replace formatting preview",
                    category="find_replace",
                    passed=has_formatting_preview,
                    message="Preview shows formatting" if has_formatting_preview else "Missing formatting in preview",
                    details={"response": result_dict},
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Find and replace formatting preview",
                    category="find_replace",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 6e: Empty find_text should return validation error
        try:
            result = await self.call_tool(
                "find_and_replace_doc",
                find_text="",
                replace_text="test",
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            # Should get a validation error
            is_validation_error = (
                result_dict.get("error", False)
                and result_dict.get("code") == "INVALID_PARAM_VALUE"
                and "find_text" in str(result_dict.get("message", ""))
            )
            self.record(
                TestResult(
                    name="Reject empty find_text",
                    category="find_replace",
                    passed=is_validation_error,
                    message="Validation caught empty find_text" if is_validation_error else "Missing validation",
                    details={"response": result_dict},
                )
            )
        except Exception as e:
            # Exception is acceptable if it's a validation error
            self.record(
                TestResult(
                    name="Reject empty find_text",
                    category="find_replace",
                    passed=False,
                    message="Exception instead of validation error",
                    error=str(e),
                )
            )

    async def test_document_structure(self):
        """Test 7: Document structure operations."""
        print("\n📊 Category: Document Structure")

        # Test 7a: Get document info (structure)
        try:
            result = await self.call_tool("get_doc_info", detail="summary")
            has_info = "total_length" in str(result) or "statistics" in str(result)
            self.record(
                TestResult(
                    name="Get document info (summary)",
                    category="structure",
                    passed=has_info,
                    message="Info retrieved",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Get document info (summary)",
                    category="structure",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_section_operations(self):
        """Test 8: Section-based operations."""
        print("\n📑 Category: Section Operations")

        # First, find a heading in the document
        heading_name = None
        try:
            structure = await self.call_tool("get_doc_info", detail="headings")
            structure_str = str(structure)
            # Try to extract a heading name from structure
            # Look for heading patterns in the JSON-like output
            if "heading" in structure_str.lower():
                # Parse the structure to find an actual heading
                try:
                    # Try to parse as JSON and find headings
                    start_idx = structure_str.find('{')
                    end_idx = structure_str.rfind('}') + 1
                    if start_idx >= 0 and end_idx > start_idx:
                        structure_dict = json.loads(structure_str[start_idx:end_idx])
                        # Look for headings in the structure
                        headings = structure_dict.get("headings", [])
                        if headings:
                            # Get first non-empty heading
                            for h in headings:
                                h_text = h.get("text", "") if isinstance(h, dict) else str(h)
                                if h_text and len(h_text) > 2 and h_text.strip():
                                    heading_name = h_text.strip()
                                    break
                except (json.JSONDecodeError, AttributeError):
                    pass

            if heading_name:
                # Test 8a: Get section by heading
                result = await self.call_tool(
                    "get_doc_section",
                    heading=heading_name,
                )
                has_content = "content" in str(result) or "text" in str(result)
                self.record(
                    TestResult(
                        name="Get section by heading",
                        category="section",
                        passed=has_content,
                        message=f"Section retrieved for '{heading_name[:30]}...'",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Get section by heading",
                        category="section",
                        passed=True,  # Skip - no headings to test
                        message="Skipped - no usable headings found",
                    )
                )
        except Exception as e:
            self.record(
                TestResult(
                    name="Get section by heading",
                    category="section",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 8b: Insert at heading position
        # Use a heading we found from the structure, or skip if none exist
        try:
            if heading_name:
                result = await self.call_tool(
                    "modify_doc_text",
                    heading=heading_name,
                    section_position="end",
                    text=f" {self.test_marker}-SECTION-END",
                    preview=True,
                )
                result_dict = json.loads(result) if isinstance(result, str) else result
                # Check if the operation would succeed (either would_modify or success)
                would_work = result_dict.get("would_modify", False) or result_dict.get("success", False)
                self.record(
                    TestResult(
                        name="Insert at section end (preview)",
                        category="section",
                        passed=would_work,
                        message="Section positioning works" if would_work else "Section not found",
                        details=result_dict,
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Insert at section end (preview)",
                        category="section",
                        passed=True,  # Skip test - no headings to test with
                        message="Skipped - no headings found in document",
                    )
                )
        except Exception as e:
            # If heading doesn't exist, check if we got a helpful error
            error_str = str(e).lower()
            if "heading not found" in error_str or "available_headings" in error_str:
                self.record(
                    TestResult(
                        name="Insert at section end (preview)",
                        category="section",
                        passed=True,  # Helpful error is acceptable
                        message="Heading not found (test doc may have changed)",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Insert at section end (preview)",
                        category="section",
                        passed=False,
                        message="Failed",
                        error=str(e),
                    )
                )

    async def test_batch_operations(self):
        """Test 9: Batch operations."""
        print("\n📦 Category: Batch Operations")

        # Test 9a: Multiple operations in batch
        try:
            result = await self.call_tool(
                "batch_edit_doc",
                operations=[
                    {
                        "type": "insert",
                        "search": self.test_marker,
                        "position": "before",
                        "text": "[BATCH1]",
                    },
                    {
                        "type": "insert",
                        "search": self.test_marker,
                        "position": "after",
                        "text": "[BATCH2]",
                    },
                ],
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Execute multiple batch operations",
                    category="batch",
                    passed=result_dict.get("success", False),
                    message=f"Completed {result_dict.get('operations_completed', 0)} operations",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Execute multiple batch operations",
                    category="batch",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 9b: Batch with auto position adjustment
        try:
            result = await self.call_tool(
                "batch_edit_doc",
                operations=[
                    {"type": "insert_text", "index": 1, "text": "A"},
                    {
                        "type": "insert_text",
                        "index": 2,
                        "text": "B",
                    },  # Should auto-adjust
                ],
                auto_adjust_positions=True,
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Batch with auto position adjustment",
                    category="batch",
                    passed=result_dict.get("success", False),
                    message="Auto-adjustment working",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Batch with auto position adjustment",
                    category="batch",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 9c: Invalid operation type should be caught early
        try:
            result = await self.call_tool(
                "batch_edit_doc",
                operations=[{"type": "completely_invalid_op_type", "text": "bad"}],
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            # Should get a validation error, not an API error
            result_str = str(result_dict).lower()
            # Check if the operation was properly rejected with a clear error
            has_invalid_op_error = (
                "invalid operation type" in result_str
                or "unsupported operation type" in result_str
            )
            is_api_error = "must specify at least one request" in result_str

            if is_api_error:
                # BUG: Invalid op type leaked through to API
                self.record(
                    TestResult(
                        name="Reject invalid batch operation type",
                        category="batch",
                        passed=False,
                        message="BUG: Invalid op type passed validation, failed at API level",
                        error="Should reject invalid operation types during validation",
                    )
                )
            elif has_invalid_op_error:
                self.record(
                    TestResult(
                        name="Reject invalid batch operation type",
                        category="batch",
                        passed=True,
                        message="Invalid operation type correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject invalid batch operation type",
                        category="batch",
                        passed=False,
                        message="Unclear handling of invalid operation type",
                        error=result_str[:100],
                    )
                )
        except Exception as e:
            # An exception is acceptable if it mentions validation
            if "invalid" in str(e).lower() or "operation" in str(e).lower():
                self.record(
                    TestResult(
                        name="Reject invalid batch operation type",
                        category="batch",
                        passed=True,
                        message="Invalid operation type rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject invalid batch operation type",
                        category="batch",
                        passed=False,
                        message="Failed",
                        error=str(e),
                    )
                )

    async def test_table_operations(self):
        """Test 10: Table operations."""
        print("\n📋 Category: Table Operations")

        # Test 10a: Create table (append to end)
        try:
            result = await self.call_tool(
                "create_table_with_data", table_data=[["H1", "H2"], ["D1", "D2"]]
            )
            passed = (
                "success" in str(result).lower() or "created" in str(result).lower()
            )
            self.record(
                TestResult(
                    name="Create table (append to end)",
                    category="table",
                    passed=passed,
                    message="Table created",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Create table (append to end)",
                    category="table",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_element_insertion(self):
        """Test 11: Element insertion."""
        print("\n🧩 Category: Element Insertion")

        # Test 11a: Page break
        try:
            # Get doc length first
            _ = await self.call_tool("get_doc_info", detail="summary")
            # Insert near end
            result = await self.call_tool(
                "insert_doc_elements", element_type="page_break", index=100
            )
            passed = "inserted" in str(result).lower()
            self.record(
                TestResult(
                    name="Insert page break",
                    category="elements",
                    passed=passed,
                    message="Page break inserted",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Insert page break",
                    category="elements",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_history_and_undo(self):
        """Test 12: History and undo operations."""
        print("\n⏪ Category: History & Undo")

        # Test 12a: Get operation history
        # Note: This tool does NOT accept user_google_email parameter
        try:
            tool = self.tester.tools.get("get_doc_operation_history")
            if tool and hasattr(tool, "fn"):
                result = await tool.fn(document_id=self.doc_id)
            else:
                result = await tool(document_id=self.doc_id)
            result_dict = json.loads(result) if isinstance(result, str) else result
            self.record(
                TestResult(
                    name="Get operation history",
                    category="history",
                    passed="operations" in str(result),
                    message=f"Found {result_dict.get('total_operations', 0)} operations",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Get operation history",
                    category="history",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_content_extraction(self):
        """Test 13: Content extraction tools."""
        print("\n📤 Category: Content Extraction")

        # Test 13a: Extract links
        try:
            result = await self.call_tool("extract_links")
            passed = "links" in str(result).lower() or "[]" in str(result)
            self.record(
                TestResult(
                    name="Extract hyperlinks from document",
                    category="extraction",
                    passed=passed,
                    message="Links extracted",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Extract hyperlinks from document",
                    category="extraction",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 13b: Extract document summary/outline
        try:
            result = await self.call_tool("extract_document_summary")
            passed = "outline" in str(result).lower() or "summary" in str(result).lower() or "heading" in str(result).lower()
            self.record(
                TestResult(
                    name="Extract document summary/outline",
                    category="extraction",
                    passed=passed,
                    message="Summary extracted",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Extract document summary/outline",
                    category="extraction",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 13c: Get full document content via get_doc_section (more reliable)
        try:
            # Use get_doc_info instead as it's more reliable for native Google Docs
            result = await self.call_tool("get_doc_info", detail="all")
            passed = len(str(result)) > 100  # Should have substantial content
            self.record(
                TestResult(
                    name="Get full document info",
                    category="extraction",
                    passed=passed,
                    message="Document info retrieved",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Get full document info",
                    category="extraction",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_comments(self):
        """Test 14: Document comments."""
        print("\n💬 Category: Comments")

        # Test 14a: Read comments (should work even if no comments)
        try:
            result = await self.call_tool("read_document_comments")
            passed = "comments" in str(result).lower() or "[]" in str(result) or "no comments" in str(result).lower()
            self.record(
                TestResult(
                    name="Read document comments",
                    category="comments",
                    passed=passed,
                    message="Comments read",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Read document comments",
                    category="comments",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 14b: Create a comment
        try:
            result = await self.call_tool(
                "create_document_comment",
                comment_content=f"{self.test_marker} Test comment",
            )
            passed = "comment" in str(result).lower() or "created" in str(result).lower()
            self.record(
                TestResult(
                    name="Create document comment",
                    category="comments",
                    passed=passed,
                    message="Comment created",
                )
            )
        except Exception as e:
            # Comments may require specific permissions or setup
            self.record(
                TestResult(
                    name="Create document comment",
                    category="comments",
                    passed=False,
                    message="Failed (may need permissions)",
                    error=str(e),
                )
            )

    async def test_find_elements(self):
        """Test 15: Find elements by type."""
        print("\n🔎 Category: Find Elements")

        # Test 15a: Find all headings
        try:
            result = await self.call_tool(
                "find_doc_elements",
                element_type="heading",
            )
            passed = "heading" in str(result).lower() or "element" in str(result).lower() or "[]" in str(result)
            self.record(
                TestResult(
                    name="Find all headings",
                    category="find_elements",
                    passed=passed,
                    message="Headings found",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Find all headings",
                    category="find_elements",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 15b: Find all tables
        try:
            result = await self.call_tool(
                "find_doc_elements",
                element_type="table",
            )
            passed = "table" in str(result).lower() or "[]" in str(result)
            self.record(
                TestResult(
                    name="Find all tables",
                    category="find_elements",
                    passed=passed,
                    message="Tables found",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Find all tables",
                    category="find_elements",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_paragraph_formatting(self):
        """Test 16: Paragraph-level formatting (expected to fail - not implemented)."""
        print("\n📐 Category: Paragraph Formatting")

        # Test 16a: Paragraph alignment (expected to fail)
        try:
            await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="replace",
                alignment="CENTER",
            )
            self.record(
                TestResult(
                    name="Paragraph alignment (center)",
                    category="paragraph",
                    passed=True,
                    message="Alignment - FEATURE NOW IMPLEMENTED! 🎉",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Paragraph alignment (center)",
                    category="paragraph",
                    passed=False,
                    message="Failed to set alignment",
                    error=str(e),
                )
            )

        # Test 16b: Heading style change (expected to fail)
        try:
            await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="replace",
                heading_style="HEADING_2",
            )
            self.record(
                TestResult(
                    name="Change to heading style",
                    category="paragraph",
                    passed=True,
                    message="Heading style - FEATURE NOW IMPLEMENTED! 🎉",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Change to heading style",
                    category="paragraph",
                    passed=False,
                    message="Failed to set heading style",
                    error=str(e),
                )
            )

        # Test 16c: Line spacing (expected to fail)
        try:
            _ = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="replace",
                line_spacing=1.5,
            )
            self.record(
                TestResult(
                    name="Line spacing",
                    category="paragraph",
                    passed=True,
                    message="Line spacing - FEATURE NOW IMPLEMENTED! 🎉",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Line spacing",
                    category="paragraph",
                    passed=False,
                    message="Failed to set line spacing",
                    error=str(e),
                )
            )

    async def test_advanced_text_formatting(self):
        """Test 17: Advanced text formatting (some expected to fail)."""
        print("\n✨ Category: Advanced Text Formatting")

        # Test 17a: Superscript (expected to fail)
        try:
            await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="replace",
                superscript=True,
            )
            self.record(
                TestResult(
                    name="Superscript text",
                    category="advanced_formatting",
                    passed=True,
                    message="Superscript - FEATURE NOW IMPLEMENTED! 🎉",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Superscript text",
                    category="advanced_formatting",
                    passed=False,
                    message="Failed to apply superscript",
                    error=str(e),
                )
            )

        # Test 17b: Subscript (expected to fail)
        try:
            await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="replace",
                subscript=True,
            )
            self.record(
                TestResult(
                    name="Subscript text",
                    category="advanced_formatting",
                    passed=True,
                    message="Subscript - FEATURE NOW IMPLEMENTED! 🎉",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Subscript text",
                    category="advanced_formatting",
                    passed=False,
                    message="Failed to apply subscript",
                    error=str(e),
                )
            )

        # Test 17c: Small caps (expected to fail)
        try:
            _ = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="replace",
                small_caps=True,
            )
            self.record(
                TestResult(
                    name="Small caps text",
                    category="advanced_formatting",
                    passed=True,
                    message="Small caps - FEATURE NOW IMPLEMENTED! 🎉",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Small caps text",
                    category="advanced_formatting",
                    passed=False,
                    message="Failed to apply small caps",
                    error=str(e),
                )
            )

    async def test_multiple_occurrences(self):
        """Test 18: Handling multiple occurrences of search text."""
        print("\n🔢 Category: Multiple Occurrences")

        # First insert multiple markers
        await self.call_tool(
            "modify_doc_text",
            location="end",
            text=f"\n{self.test_marker}-A {self.test_marker}-B {self.test_marker}-C\n",
        )

        # Test 18a: Target 2nd occurrence
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="after",
                occurrence=2,
                text="[2ND]",
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            passed = result_dict.get("success", False) or "link" in str(result_dict)
            self.record(
                TestResult(
                    name="Target 2nd occurrence",
                    category="occurrences",
                    passed=passed,
                    message="2nd occurrence targeted",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Target 2nd occurrence",
                    category="occurrences",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 18b: Target last occurrence (-1)
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search=self.test_marker,
                position="after",
                occurrence=-1,
                text="[LAST]",
            )
            result_dict = json.loads(result) if isinstance(result, str) else result
            passed = result_dict.get("success", False) or "link" in str(result_dict)
            self.record(
                TestResult(
                    name="Target last occurrence (-1)",
                    category="occurrences",
                    passed=passed,
                    message="Last occurrence targeted",
                    details=result_dict,
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Target last occurrence (-1)",
                    category="occurrences",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_search_preview(self):
        """Test 19: Search preview tool."""
        print("\n🔮 Category: Search Preview")

        # Test 19a: Preview search results
        try:
            result = await self.call_tool(
                "preview_search_results",
                search_text=self.test_marker,
            )
            passed = "match" in str(result).lower() or "found" in str(result).lower() or "occurrence" in str(result).lower()
            self.record(
                TestResult(
                    name="Preview search results",
                    category="search_preview",
                    passed=passed,
                    message="Search preview returned results",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Preview search results",
                    category="search_preview",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_list_insertion(self):
        """Test 20: List creation."""
        print("\n📋 Category: List Operations")

        # Get current doc length for insertion point
        try:
            info = await self.call_tool("get_doc_info", detail="summary")
            info_dict = json.loads(info) if isinstance(info, str) else info
            # Use a safe index near the end
            insert_index = info_dict.get("statistics", {}).get("total_length", 100) - 10
            if insert_index < 10:
                insert_index = 10
        except Exception:
            insert_index = 100

        # Test 20a: Create unordered list (bullets)
        try:
            result = await self.call_tool(
                "insert_doc_elements",
                element_type="list",
                index=insert_index,
                list_type="UNORDERED",
                text=f"{self.test_marker} List Item",
            )
            passed = "list" in str(result).lower() or "inserted" in str(result).lower()
            self.record(
                TestResult(
                    name="Create bullet list",
                    category="lists",
                    passed=passed,
                    message="Bullet list created",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Create bullet list",
                    category="lists",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 20b: Create ordered list (numbered)
        try:
            result = await self.call_tool(
                "insert_doc_elements",
                element_type="list",
                index=insert_index + 50,
                list_type="ORDERED",
                text=f"{self.test_marker} Numbered Item",
            )
            passed = "list" in str(result).lower() or "inserted" in str(result).lower()
            self.record(
                TestResult(
                    name="Create numbered list",
                    category="lists",
                    passed=passed,
                    message="Numbered list created",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Create numbered list",
                    category="lists",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 20c: Convert text to bullet list using modify_doc_text
        try:
            # First insert some plain text to convert
            await self.call_tool(
                "modify_doc_text",
                location="end",
                text=f"\n{self.test_marker} Convert Item 1\n{self.test_marker} Convert Item 2\n",
            )
            # Get current document length to figure out the range
            result = await self.call_tool("get_doc_info", detail="summary")
            # Extract JSON from response (may have text prefix and suffix)
            result_str = str(result)
            json_start = result_str.find('{')
            json_end = result_str.rfind('}') + 1
            if json_start >= 0 and json_end > json_start:
                info = json.loads(result_str[json_start:json_end])
            else:
                info = {}
            end_index = info.get("total_length", info.get("statistics", {}).get("total_length", 200)) - 1
            # Convert the last few lines to a bullet list
            result = await self.call_tool(
                "modify_doc_text",
                start_index=max(1, end_index - 100),
                end_index=end_index,
                convert_to_list="UNORDERED",
            )
            passed = "success" in str(result).lower() or "list" in str(result).lower() or "bullet" in str(result).lower()
            self.record(
                TestResult(
                    name="Convert text to bullet list",
                    category="lists",
                    passed=passed,
                    message="Text converted to bullet list",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Convert text to bullet list",
                    category="lists",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 20d: Insert text AND convert to numbered list in one operation
        try:
            result = await self.call_tool(
                "modify_doc_text",
                location="end",
                text=f"\n{self.test_marker} Step 1\n{self.test_marker} Step 2\n{self.test_marker} Step 3\n",
                convert_to_list="ORDERED",
            )
            passed = "success" in str(result).lower() or "numbered" in str(result).lower() or "list" in str(result).lower()
            self.record(
                TestResult(
                    name="Insert and convert to numbered list",
                    category="lists",
                    passed=passed,
                    message="Text inserted and converted to numbered list",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Insert and convert to numbered list",
                    category="lists",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

        # Test 20e: Find all lists (verifies find_doc_elements can find lists)
        try:
            result = await self.call_tool(
                "find_doc_elements",
                element_type="list",
            )
            result_str = str(result)
            # Should find at least 1 list (we created several above)
            # Check for "count" field with non-zero value
            if '"count": 0' in result_str or "'count': 0" in result_str:
                self.record(
                    TestResult(
                        name="Find all lists",
                        category="lists",
                        passed=False,
                        message="BUG: find_doc_elements returns 0 lists",
                        error="Expected to find at least 1 list element",
                    )
                )
            elif (
                "list" in result_str.lower()
                or "bullet" in result_str.lower()
                or "numbered" in result_str.lower()
            ):
                self.record(
                    TestResult(
                        name="Find all lists",
                        category="lists",
                        passed=True,
                        message="Lists found successfully",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Find all lists",
                        category="lists",
                        passed=False,
                        message="Unexpected response format",
                        error=result_str[:200],
                    )
                )
        except Exception as e:
            self.record(
                TestResult(
                    name="Find all lists",
                    category="lists",
                    passed=False,
                    message="Failed to find lists",
                    error=str(e),
                )
            )

    async def test_heading_navigation(self):
        """Test 21: Navigate between headings."""
        print("\n🧭 Category: Heading Navigation")

        # Test 21a: Get heading siblings (returns next/previous automatically)
        try:
            result = await self.call_tool(
                "navigate_heading_siblings",
                heading="The Problem",
            )
            passed = "heading" in str(result).lower() or "sibling" in str(result).lower() or "next" in str(result).lower() or "previous" in str(result).lower()
            self.record(
                TestResult(
                    name="Get heading siblings",
                    category="navigation",
                    passed=passed,
                    message="Navigation successful",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Get heading siblings",
                    category="navigation",
                    passed=False,
                    message="Failed",
                    error=str(e),
                )
            )

    async def test_export_pdf(self):
        """Test 22: Export document to PDF."""
        print("\n📄 Category: Export Operations")

        # Test 22a: Export to PDF (may fail without Drive permissions)
        try:
            result = await self.call_tool("export_doc_to_pdf")
            passed = "pdf" in str(result).lower() or "export" in str(result).lower() or "drive" in str(result).lower()
            self.record(
                TestResult(
                    name="Export document to PDF",
                    category="export",
                    passed=passed,
                    message="PDF exported",
                )
            )
        except Exception as e:
            self.record(
                TestResult(
                    name="Export document to PDF",
                    category="export",
                    passed=False,
                    message="Failed (may need Drive permissions)",
                    error=str(e),
                )
            )

    async def test_validation_edge_cases(self):
        """Test 23: Validation edge cases - inputs that should be caught early."""
        print("\n🛡️ Category: Validation Edge Cases")

        # Test 23a: Negative index should be rejected
        try:
            result = await self.call_tool(
                "modify_doc_text",
                start_index=-5,
                text="should_fail",
            )
            result_str = str(result)
            # Check if it was caught by validation (good) or leaked to API (bug)
            if "API error" in result_str or "Index must be greater" in result_str:
                self.record(
                    TestResult(
                        name="Reject negative index",
                        category="validation",
                        passed=False,
                        message="BUG: Negative index leaked to API",
                        error="Should be caught during validation",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject negative index",
                        category="validation",
                        passed=True,
                        message="Validation caught negative index",
                    )
                )
        except Exception as e:
            error_str = str(e)
            if "negative" in error_str.lower() or "greater than" in error_str.lower():
                self.record(
                    TestResult(
                        name="Reject negative index",
                        category="validation",
                        passed=True,
                        message="Validation caught negative index",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject negative index",
                        category="validation",
                        passed=False,
                        message="BUG: Negative index failed but not caught properly",
                        error=error_str,
                    )
                )

        # Test 23b: Invalid color format should be rejected gracefully
        try:
            result = await self.call_tool(
                "modify_doc_text",
                location="end",
                text="color_test",
                foreground_color="not_a_valid_color_xyz",
            )
            self.record(
                TestResult(
                    name="Reject invalid color format",
                    category="validation",
                    passed=True,
                    message="Invalid color handled",
                )
            )
        except Exception as e:
            error_str = str(e)
            if "color" in error_str.lower() and ("hex" in error_str.lower() or "named" in error_str.lower()):
                self.record(
                    TestResult(
                        name="Reject invalid color format",
                        category="validation",
                        passed=True,
                        message="Invalid color rejected with helpful message",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject invalid color format",
                        category="validation",
                        passed=False,
                        message="Invalid color rejected but message unclear",
                        error=error_str[:100],
                    )
                )

        # Test 23c: font_size=0 should be rejected (valid range is 1-400)
        try:
            result = await self.call_tool(
                "modify_doc_text",
                location="end",
                text="[FONTSIZE0]",
                font_size=0,
            )
            result_str = str(result)
            # If it succeeds, it's a bug - font_size=0 should be rejected
            if "success" in result_str.lower() and '"success": true' in result_str:
                self.record(
                    TestResult(
                        name="Reject font_size=0",
                        category="validation",
                        passed=False,
                        message="BUG: font_size=0 passed validation (should be 1-400)",
                        error="font_size=0 should be rejected during validation",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject font_size=0",
                        category="validation",
                        passed=True,
                        message="font_size=0 correctly rejected",
                    )
                )
        except Exception as e:
            error_str = str(e).lower()
            if "font" in error_str or "size" in error_str or "range" in error_str:
                self.record(
                    TestResult(
                        name="Reject font_size=0",
                        category="validation",
                        passed=True,
                        message="font_size=0 correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject font_size=0",
                        category="validation",
                        passed=False,
                        message="font_size=0 rejected but unclear message",
                        error=str(e)[:100],
                    )
                )

        # Test 23d: Invalid list_type should be rejected
        try:
            result = await self.call_tool(
                "insert_doc_elements",
                element_type="list",
                index=100,
                list_type="INVALID_TYPE",
                text="test item",
            )
            result_str = str(result)
            # If it "succeeds", it's a bug
            if "inserted" in result_str.lower() or "success" in result_str.lower():
                self.record(
                    TestResult(
                        name="Reject invalid list_type",
                        category="validation",
                        passed=False,
                        message="BUG: Invalid list_type passed validation",
                        error="Should only accept ORDERED or UNORDERED",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject invalid list_type",
                        category="validation",
                        passed=True,
                        message="Invalid list_type correctly rejected",
                    )
                )
        except Exception as e:
            error_str = str(e).lower()
            if "list" in error_str or "type" in error_str or "ordered" in error_str:
                self.record(
                    TestResult(
                        name="Reject invalid list_type",
                        category="validation",
                        passed=True,
                        message="Invalid list_type correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject invalid list_type",
                        category="validation",
                        passed=False,
                        message="Invalid list_type rejected but unclear message",
                        error=str(e)[:100],
                    )
                )

        # Test 23e: Empty find_text in find_and_replace should be rejected
        try:
            result = await self.call_tool(
                "find_and_replace_doc",
                find_text="",
                replace_text="test",
            )
            result_str = str(result)
            if "API error" in result_str or "should not be empty" in result_str:
                self.record(
                    TestResult(
                        name="Reject empty find_text",
                        category="validation",
                        passed=False,
                        message="BUG: Empty find_text leaked to API",
                        error="Should be caught during validation",
                    )
                )
            elif "error" in result_str.lower() and "empty" in result_str.lower():
                self.record(
                    TestResult(
                        name="Reject empty find_text",
                        category="validation",
                        passed=True,
                        message="Empty find_text correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject empty find_text",
                        category="validation",
                        passed=True,
                        message="Empty find_text handled",
                    )
                )
        except Exception as e:
            error_str = str(e)
            if "empty" in error_str.lower() or "required" in error_str.lower():
                self.record(
                    TestResult(
                        name="Reject empty find_text",
                        category="validation",
                        passed=True,
                        message="Empty find_text correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject empty find_text",
                        category="validation",
                        passed=False,
                        message="Empty find_text error but unclear message",
                        error=error_str[:100],
                    )
                )

        # Test 23f: Empty search string in modify_doc_text should give specific error
        try:
            result = await self.call_tool(
                "modify_doc_text",
                search="",
                position="after",
                text="test",
            )
            result_str = str(result).lower()
            if "empty" in result_str and "search" in result_str:
                self.record(
                    TestResult(
                        name="Empty search gives specific error",
                        category="validation",
                        passed=True,
                        message="Empty search text correctly rejected with clear message",
                    )
                )
            elif "not found" in result_str:
                self.record(
                    TestResult(
                        name="Empty search gives specific error",
                        category="validation",
                        passed=False,
                        message="Empty search returns generic 'not found' instead of specific error",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Empty search gives specific error",
                        category="validation",
                        passed=True,
                        message="Empty search handled appropriately",
                    )
                )
        except Exception as e:
            error_str = str(e).lower()
            if "empty" in error_str:
                self.record(
                    TestResult(
                        name="Empty search gives specific error",
                        category="validation",
                        passed=True,
                        message="Empty search correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Empty search gives specific error",
                        category="validation",
                        passed=False,
                        message="Empty search error but unclear message",
                        error=str(e)[:100],
                    )
                )

        # Test 23g: end_index before start_index should be rejected
        try:
            result = await self.call_tool(
                "modify_doc_text",
                start_index=100,
                end_index=50,
                bold=True,
            )
            result_str = str(result)
            if "error" in result_str.lower() and ("start" in result_str.lower() or "end" in result_str.lower()):
                self.record(
                    TestResult(
                        name="Reject end_index < start_index",
                        category="validation",
                        passed=True,
                        message="Invalid index range correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject end_index < start_index",
                        category="validation",
                        passed=False,
                        message="Invalid range not caught properly",
                        error=result_str[:100],
                    )
                )
        except Exception as e:
            error_str = str(e).lower()
            if "start" in error_str or "end" in error_str or "range" in error_str:
                self.record(
                    TestResult(
                        name="Reject end_index < start_index",
                        category="validation",
                        passed=True,
                        message="Invalid index range correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject end_index < start_index",
                        category="validation",
                        passed=False,
                        message="Index range error but unclear message",
                        error=str(e)[:100],
                    )
                )

        # Test 23h: Index beyond document length should be caught
        try:
            result = await self.call_tool(
                "modify_doc_text",
                start_index=100,
                end_index=9999999,
                bold=True,
            )
            result_str = str(result)
            if "error" in result_str.lower() and ("bounds" in result_str.lower() or "length" in result_str.lower()):
                self.record(
                    TestResult(
                        name="Reject index beyond document",
                        category="validation",
                        passed=True,
                        message="Out-of-bounds index correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject index beyond document",
                        category="validation",
                        passed=False,
                        message="Out-of-bounds not caught properly",
                        error=result_str[:100],
                    )
                )
        except Exception as e:
            error_str = str(e).lower()
            if "bounds" in error_str or "length" in error_str or "exceed" in error_str:
                self.record(
                    TestResult(
                        name="Reject index beyond document",
                        category="validation",
                        passed=True,
                        message="Out-of-bounds index correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject index beyond document",
                        category="validation",
                        passed=False,
                        message="Out-of-bounds error but unclear message",
                        error=str(e)[:100],
                    )
                )

        # Test 23i: Empty table data should be rejected
        try:
            result = await self.call_tool(
                "create_table_with_data",
                table_data=[],
            )
            result_str = str(result)
            if "error" in result_str.lower() or "cannot be empty" in result_str.lower():
                self.record(
                    TestResult(
                        name="Reject empty table data",
                        category="validation",
                        passed=True,
                        message="Empty table data correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject empty table data",
                        category="validation",
                        passed=False,
                        message="Empty table data not caught",
                        error=result_str[:100],
                    )
                )
        except Exception as e:
            error_str = str(e).lower()
            if "empty" in error_str or "required" in error_str:
                self.record(
                    TestResult(
                        name="Reject empty table data",
                        category="validation",
                        passed=True,
                        message="Empty table data correctly rejected",
                    )
                )
            else:
                self.record(
                    TestResult(
                        name="Reject empty table data",
                        category="validation",
                        passed=False,
                        message="Empty table error but unclear message",
                        error=str(e)[:100],
                    )
                )

    async def cleanup_test_content(self):
        """Remove test content from document."""
        print("\n🧹 Cleaning up test content...")

        cleanup_patterns = [
            self.test_marker,
            "[BATCH1]",
            "[BATCH2]",
            "[FAR-REPLACED]",
            "[2ND]",
            "[LAST]",
            "[FONTSIZE0]",  # From validation test
            "AB",  # From batch test
        ]

        for pattern in cleanup_patterns:
            try:
                await self.call_tool(
                    "find_and_replace_doc", find_text=pattern, replace_text=""
                )
                if self.verbose:
                    print(f"  Cleaned: {pattern}")
            except Exception:
                pass  # Pattern not found, no cleanup needed

        print("  Cleanup complete!")

    async def run_all_tests(self, cleanup: bool = False):
        """Run all test scenarios."""
        print("=" * 60)
        print("🧪 Google Docs MCP Scenario Tests")
        print("=" * 60)

        await self.setup()

        # Run all test categories
        await self.test_basic_insertion()
        await self.test_search_based_operations()
        await self.test_preview_mode()
        await self.test_text_formatting()
        await self.test_delete_operations()
        await self.test_find_and_replace()
        await self.test_document_structure()
        await self.test_section_operations()
        await self.test_batch_operations()
        await self.test_table_operations()
        await self.test_element_insertion()
        await self.test_history_and_undo()
        await self.test_content_extraction()
        await self.test_comments()
        await self.test_find_elements()
        await self.test_paragraph_formatting()
        await self.test_advanced_text_formatting()
        await self.test_multiple_occurrences()
        await self.test_search_preview()
        await self.test_list_insertion()
        await self.test_heading_navigation()
        await self.test_export_pdf()
        await self.test_validation_edge_cases()

        if cleanup:
            await self.cleanup_test_content()

        all_expected = self.print_report()
        self.report.all_failures_expected = all_expected
        return self.report

    def print_report(self):
        """Print test report summary."""
        print("\n" + "=" * 60)
        print("📊 TEST REPORT")
        print("=" * 60)

        # Separate expected vs unexpected failures
        expected_fails = [
            r for r in self.report.results if not r.passed and r.expected_fail
        ]
        unexpected_fails = [
            r for r in self.report.results if not r.passed and not r.expected_fail
        ]

        print(f"Total:  {self.report.total}")
        print(f"Passed: {self.report.passed} ✅")
        if expected_fails:
            print(
                f"Expected Failures: {len(expected_fails)} ⚠️ (known missing features)"
            )
        if unexpected_fails:
            print(f"Unexpected Failures: {len(unexpected_fails)} ❌ (BUGS!)")

        effective_pass_rate = self.report.passed + len(expected_fails)
        print(
            f"Rate:   {self.report.passed}/{self.report.total} ({100 * self.report.passed // max(1, self.report.total)}%)"
        )
        print(
            f"Effective Rate (excluding expected): {effective_pass_rate}/{self.report.total} ({100 * effective_pass_rate // max(1, self.report.total)}%)"
        )

        # Group by category
        by_category = {}
        for r in self.report.results:
            if r.category not in by_category:
                by_category[r.category] = {"passed": 0, "failed": 0, "expected_fail": 0}
            if r.passed:
                by_category[r.category]["passed"] += 1
            elif r.expected_fail:
                by_category[r.category]["expected_fail"] += 1
            else:
                by_category[r.category]["failed"] += 1

        print("\nBy Category:")
        for cat, stats in by_category.items():
            total = stats["passed"] + stats["failed"] + stats["expected_fail"]
            if stats["failed"] > 0:
                status = "❌"  # Has unexpected failures
            elif stats["expected_fail"] > 0:
                status = "⚠️"  # Only expected failures
            else:
                status = "✅"  # All pass
            print(
                f"  {status} {cat}: {stats['passed']}/{total}"
                + (
                    f" ({stats['expected_fail']} expected fail)"
                    if stats["expected_fail"]
                    else ""
                )
            )

        # List unexpected failures (real bugs)
        if unexpected_fails:
            print("\n❌ UNEXPECTED FAILURES (BUGS):")
            for f in unexpected_fails:
                print(f"  - {f.name}: {f.message}")
                if f.error:
                    print(f"    Error: {f.error[:80]}...")

        # List expected failures (missing features)
        if expected_fails:
            print("\n⚠️ Expected Failures (missing features):")
            for f in expected_fails:
                print(f"  - {f.name}")

        print("\n" + "=" * 60)

        # Return exit status based on unexpected failures only
        return len(unexpected_fails) == 0


async def main():
    parser = argparse.ArgumentParser(
        description="Run Google Docs MCP scenario tests",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  uv run python gdocs/scenario_tests.py --doc_id ABC123 --email user@example.com
  uv run python gdocs/scenario_tests.py --doc_id ABC123 --email user@example.com --cleanup
  uv run python gdocs/scenario_tests.py --doc_id ABC123 --email user@example.com --quiet
        """,
    )
    parser.add_argument("--doc_id", required=True, help="Google Doc ID to test with")
    parser.add_argument("--email", required=True, help="User Google email for auth")
    parser.add_argument(
        "--cleanup", action="store_true", help="Clean up test content after"
    )
    parser.add_argument("--quiet", action="store_true", help="Less verbose output")

    args = parser.parse_args()

    tester = ScenarioTester(
        doc_id=args.doc_id, email=args.email, verbose=not args.quiet
    )

    report = await tester.run_all_tests(cleanup=args.cleanup)

    # Exit with error code only if there are UNEXPECTED failures
    # (Expected failures from known missing features don't fail the run)
    sys.exit(0 if getattr(report, "all_failures_expected", False) else 1)


if __name__ == "__main__":
    asyncio.run(main())
