#!/usr/bin/env python3
"""
Google Calendar Feature Tests

Tests two main features:
1. Recurring Event Support (recurrence parameter in create_event)
2. get_events_times_only Tool (weekday grouping & de-duplication)

Usage:
    python tests/test_calendar_features.py --email user@example.com
    USER_GOOGLE_EMAIL=user@example.com python tests/test_calendar_features.py

Requirements:
    - Valid Google OAuth credentials for the specified email
    - Run from project root directory
"""

import argparse
import asyncio
import os
import sys
from datetime import datetime, timedelta
from typing import List, Optional

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class TestResult:
    """Represents a single test result."""
    
    def __init__(self, name: str, passed: bool, message: str = "", warning: bool = False):
        self.name = name
        self.passed = passed
        self.message = message
        self.warning = warning
    
    def __str__(self):
        if self.warning:
            status = "⚠️  WARNING"
        elif self.passed:
            status = "✅ PASSED"
        else:
            status = "❌ FAILED"
        return f"{status}: {self.name}" + (f" - {self.message}" if self.message else "")


class CalendarFeatureTests:
    """Test suite for Google Calendar features."""
    
    def __init__(self, email: str):
        self.email = email
        self.results: List[TestResult] = []
        self.created_event_ids: List[str] = []
        self.tester = None
        
        # Calculate dates
        today = datetime.now()
        days_until_monday = (7 - today.weekday()) % 7
        if days_until_monday == 0:
            days_until_monday = 7
        self.next_monday = (today + timedelta(days=days_until_monday)).strftime("%Y-%m-%d")
        self.next_week = (today + timedelta(days=days_until_monday + 7)).strftime("%Y-%m-%d")
    
    async def setup(self):
        """Initialize the test environment."""
        from tools_cli import init_server, ToolTester
        
        print("🔧 Initializing test environment...")
        server = init_server()
        self.tester = ToolTester(server)
        await self.tester.init_tools()
        print(f"✅ Loaded {len(self.tester.tools)} tools")
    
    async def call_tool(self, tool_name: str, **kwargs) -> str:
        """Call a tool and return the result."""
        kwargs['user_google_email'] = self.email
        tool = self.tester.tools[tool_name]
        try:
            if hasattr(tool, 'fn'):
                return await tool.fn(**kwargs)
            return await tool(**kwargs)
        except Exception as e:
            return f"ERROR: {e}"
    
    def add_result(self, name: str, passed: bool, message: str = "", warning: bool = False):
        """Add a test result."""
        result = TestResult(name, passed, message, warning)
        self.results.append(result)
        print(result)
    
    async def extract_event_id(self, result: str) -> Optional[str]:
        """Extract event ID from create_event result."""
        if "Successfully created" in result:
            # The event ID is in the URL
            import re
            match = re.search(r"event\?eid=([a-zA-Z0-9_-]+)", result)
            if match:
                return match.group(1)
        return None

    # ============================================
    # PART 1: Recurring Event Support
    # ============================================
    
    async def test_1_1_weekly_recurring(self):
        """Test 1.1 - Weekly Recurring Meeting"""
        result = await self.call_tool(
            "create_event",
            summary="TEST: Weekly Test Meeting",
            start_time=f"{self.next_monday}T10:00:00-06:00",
            end_time=f"{self.next_monday}T11:00:00-06:00",
            recurrence=["RRULE:FREQ=WEEKLY"]
        )
        passed = "Successfully created" in result
        self.add_result("1.1 Weekly Recurring Meeting", passed)
        if event_id := await self.extract_event_id(result):
            self.created_event_ids.append(event_id)
    
    async def test_1_2_biweekly(self):
        """Test 1.2 - Bi-Weekly Meeting"""
        result = await self.call_tool(
            "create_event",
            summary="TEST: Bi-Weekly Sync",
            start_time=f"{self.next_monday}T14:00:00-06:00",
            end_time=f"{self.next_monday}T15:00:00-06:00",
            recurrence=["RRULE:FREQ=WEEKLY;INTERVAL=2"]
        )
        passed = "Successfully created" in result
        self.add_result("1.2 Bi-Weekly Meeting", passed)
    
    async def test_1_3_weekday_standup(self):
        """Test 1.3 - Weekday-Only Standup"""
        result = await self.call_tool(
            "create_event",
            summary="TEST: Daily Standup",
            start_time=f"{self.next_monday}T09:00:00-06:00",
            end_time=f"{self.next_monday}T09:15:00-06:00",
            recurrence=["RRULE:FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR"]
        )
        passed = "Successfully created" in result
        self.add_result("1.3 Weekday-Only Standup", passed)
    
    async def test_1_4_monthly_15th(self):
        """Test 1.4 - Monthly Meeting on 15th"""
        result = await self.call_tool(
            "create_event",
            summary="TEST: Monthly Review",
            start_time="2025-12-15T13:00:00-06:00",
            end_time="2025-12-15T14:00:00-06:00",
            recurrence=["RRULE:FREQ=MONTHLY;BYMONTHDAY=15"]
        )
        passed = "Successfully created" in result
        self.add_result("1.4 Monthly Meeting on 15th", passed)
    
    async def test_1_5_limited_count(self):
        """Test 1.5 - Limited Series (10 occurrences)"""
        result = await self.call_tool(
            "create_event",
            summary="TEST: Training Session",
            start_time=f"{self.next_monday}T15:00:00-06:00",
            end_time=f"{self.next_monday}T16:00:00-06:00",
            recurrence=["RRULE:FREQ=WEEKLY;COUNT=10"]
        )
        passed = "Successfully created" in result
        self.add_result("1.5 Limited Series (COUNT=10)", passed)
    
    async def test_1_6_until_date(self):
        """Test 1.6 - Event with End Date"""
        result = await self.call_tool(
            "create_event",
            summary="TEST: Project Meetings",
            start_time=f"{self.next_monday}T16:00:00-06:00",
            end_time=f"{self.next_monday}T17:00:00-06:00",
            recurrence=["RRULE:FREQ=WEEKLY;UNTIL=20250630T235959Z"]
        )
        passed = "Successfully created" in result
        self.add_result("1.6 Event with UNTIL Date", passed)
    
    async def test_1_7_non_recurring(self):
        """Test 1.7 - Non-Recurring Event (Regression)"""
        result = await self.call_tool(
            "create_event",
            summary="TEST: Single Meeting",
            start_time=f"{self.next_monday}T11:00:00-06:00",
            end_time=f"{self.next_monday}T12:00:00-06:00"
        )
        passed = "Successfully created" in result
        self.add_result("1.7 Non-Recurring Event (Regression)", passed)
    
    async def test_1_8_recurring_with_meet(self):
        """Test 1.8 - Recurring Event with Google Meet"""
        result = await self.call_tool(
            "create_event",
            summary="TEST: Virtual Team Meeting",
            start_time=f"{self.next_monday}T17:00:00-06:00",
            end_time=f"{self.next_monday}T18:00:00-06:00",
            recurrence=["RRULE:FREQ=WEEKLY"],
            add_google_meet=True
        )
        if "Successfully created" in result:
            self.add_result("1.8 Recurring + Google Meet", True)
        elif "Invalid conference type" in result:
            self.add_result("1.8 Recurring + Google Meet", False, 
                          "Google API limitation", warning=True)
        else:
            self.add_result("1.8 Recurring + Google Meet", False, result)

    # ============================================
    # PART 2: Get Events Times Only Tool
    # ============================================
    
    async def test_2_1_weekday_grouping(self):
        """Test 2.1 - Weekday Grouping"""
        result = await self.call_tool(
            "get_events_times_only",
            time_min=self.next_monday,
            time_max=self.next_week
        )
        passed = any(day in result for day in ["Monday:", "Tuesday:", "Wednesday:", "Thursday:", "Friday:"])
        self.add_result("2.1 Weekday Grouping", passed)
    
    async def test_2_2_recurring_dedup(self):
        """Test 2.2 - Recurring Event De-duplication"""
        # Need a 2-week range to see recurring event de-dup
        two_weeks = (datetime.strptime(self.next_monday, "%Y-%m-%d") + timedelta(days=14)).strftime("%Y-%m-%d")
        result = await self.call_tool(
            "get_events_times_only",
            time_min=self.next_monday,
            time_max=two_weeks,
            max_results=50
        )
        passed = "(recurring," in result
        self.add_result("2.2 Recurring De-duplication", passed,
                       "" if passed else "No recurring markers found")
    
    async def test_2_3_allday_format(self):
        """Test 2.3 - All-Day Events Format"""
        # Create an all-day event
        next_day = (datetime.strptime(self.next_monday, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")
        await self.call_tool(
            "create_event",
            summary="TEST: All-Day Event",
            start_time=self.next_monday,
            end_time=next_day
        )
        result = await self.call_tool(
            "get_events_times_only",
            time_min=self.next_monday,
            time_max=self.next_week
        )
        passed = "(All day)" in result
        self.add_result("2.3 All-Day Events Format", passed)
    
    async def test_2_4_default_time_min(self):
        """Test 2.4 - Default Time Min"""
        result = await self.call_tool("get_events_times_only")
        passed = "Events for" in result
        self.add_result("2.4 Default Time Min", passed)
    
    async def test_2_5_max_results(self):
        """Test 2.5 - Max Results Limit"""
        result = await self.call_tool(
            "get_events_times_only",
            time_min=self.next_monday,
            time_max=self.next_week,
            max_results=3
        )
        event_count = result.count("  - ")
        passed = event_count <= 3
        self.add_result("2.5 Max Results Limit", passed, f"Got {event_count} events")

    # ============================================
    # PART 3: Regression Tests
    # ============================================
    
    async def test_3_1_list_calendars(self):
        """Test 3.1 - List Calendars"""
        result = await self.call_tool("list_calendars")
        passed = "Successfully listed" in result
        self.add_result("3.1 List Calendars", passed)
    
    async def test_3_2_get_events(self):
        """Test 3.2 - Get Events (Original Tool)"""
        result = await self.call_tool(
            "get_events",
            time_min=self.next_monday,
            time_max=self.next_week,
            max_results=5
        )
        passed = "Successfully retrieved" in result or "No events found" in result
        self.add_result("3.2 Get Events (Original)", passed)
    
    async def test_3_3_delete_event(self):
        """Test 3.3 - Delete Event"""
        # Get a test event to delete
        events = await self.call_tool(
            "get_events",
            time_min="2025-01-01",
            time_max="2026-12-31",
            query="TEST:",
            max_results=1
        )
        import re
        match = re.search(r"ID: ([a-zA-Z0-9_]+)", events)
        if match:
            event_id = match.group(1)
            result = await self.call_tool("delete_event", event_id=event_id)
            passed = "Successfully deleted" in result
            self.add_result("3.3 Delete Event", passed)
        else:
            self.add_result("3.3 Delete Event", False, "No test event found", warning=True)

    # ============================================
    # Cleanup
    # ============================================
    
    async def cleanup(self):
        """Delete all TEST: events."""
        print("\n🧹 Cleaning up test events...")
        deleted = 0
        while True:
            events = await self.call_tool(
                "get_events",
                time_min="2025-01-01",
                time_max="2026-12-31",
                query="TEST:",
                max_results=50
            )
            import re
            event_ids = re.findall(r"ID: ([a-zA-Z0-9_]+)", events)
            if not event_ids:
                break
            for event_id in event_ids:
                try:
                    await self.call_tool("delete_event", event_id=event_id)
                    deleted += 1
                except Exception:
                    pass
        print(f"✅ Deleted {deleted} test events")
    
    async def run_all(self):
        """Run all tests."""
        print("\n" + "=" * 50)
        print("PART 1: Recurring Event Support")
        print("=" * 50)
        await self.test_1_1_weekly_recurring()
        await self.test_1_2_biweekly()
        await self.test_1_3_weekday_standup()
        await self.test_1_4_monthly_15th()
        await self.test_1_5_limited_count()
        await self.test_1_6_until_date()
        await self.test_1_7_non_recurring()
        await self.test_1_8_recurring_with_meet()
        
        print("\n" + "=" * 50)
        print("PART 2: Get Events Times Only Tool")
        print("=" * 50)
        await self.test_2_1_weekday_grouping()
        await self.test_2_2_recurring_dedup()
        await self.test_2_3_allday_format()
        await self.test_2_4_default_time_min()
        await self.test_2_5_max_results()
        
        print("\n" + "=" * 50)
        print("PART 3: Regression Tests")
        print("=" * 50)
        await self.test_3_1_list_calendars()
        await self.test_3_2_get_events()
        await self.test_3_3_delete_event()
        
        await self.cleanup()
        
        # Summary
        print("\n" + "=" * 50)
        print("TEST SUMMARY")
        print("=" * 50)
        passed = sum(1 for r in self.results if r.passed and not r.warning)
        warnings = sum(1 for r in self.results if r.warning)
        failed = sum(1 for r in self.results if not r.passed and not r.warning)
        print(f"✅ Passed: {passed}")
        print(f"⚠️  Warnings: {warnings}")
        print(f"❌ Failed: {failed}")
        print(f"📊 Total: {len(self.results)}")
        
        return failed == 0


async def main():
    parser = argparse.ArgumentParser(description="Google Calendar Feature Tests")
    parser.add_argument("--email", "-e", 
                       default=os.getenv("USER_GOOGLE_EMAIL"),
                       help="Google email address")
    args = parser.parse_args()
    
    if not args.email:
        print("❌ Error: Please provide --email or set USER_GOOGLE_EMAIL")
        sys.exit(1)
    
    tests = CalendarFeatureTests(args.email)
    await tests.setup()
    success = await tests.run_all()
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    asyncio.run(main())

