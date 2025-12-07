#!/bin/bash
#
# Google Calendar Feature Testing Script
# 
# Tests two main features:
# 1. Recurring Event Support (recurrence parameter in create_event)
# 2. get_events_times_only Tool (weekday grouping & de-duplication)
#
# Prerequisites:
# - Set USER_GOOGLE_EMAIL environment variable or pass as first argument
# - Run from project root: ./tests/test_calendar_features.sh [email]
#
# Usage:
#   ./tests/test_calendar_features.sh mbradshaw@indeed.com
#   USER_GOOGLE_EMAIL=mbradshaw@indeed.com ./tests/test_calendar_features.sh
#

set -e

# Configuration
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
EMAIL="${1:-$USER_GOOGLE_EMAIL}"

if [ -z "$EMAIL" ]; then
    echo "❌ Error: Please provide email address as argument or set USER_GOOGLE_EMAIL"
    echo "Usage: $0 <email@example.com>"
    exit 1
fi

cd "$PROJECT_ROOT"

# Helper function to run tools_cli
run_tool() {
    local tool_name="$1"
    shift
    timeout 60 uv run python tools_cli.py --tool "$tool_name" --user_google_email "$EMAIL" "$@" 2>&1
}

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

pass() { echo -e "${GREEN}✅ PASSED${NC}: $1"; }
fail() { echo -e "${RED}❌ FAILED${NC}: $1"; }
warn() { echo -e "${YELLOW}⚠️  WARNING${NC}: $1"; }
info() { echo -e "ℹ️  $1"; }

# Track created event IDs for cleanup
declare -a EVENT_IDS

# Calculate next Monday
NEXT_MONDAY=$(date -v+1d -v+mon +%Y-%m-%d 2>/dev/null || date -d "next monday" +%Y-%m-%d)
NEXT_WEEK=$(date -v+7d +%Y-%m-%d 2>/dev/null || date -d "+7 days" +%Y-%m-%d)

echo "=============================================="
echo "Google Calendar Feature Tests"
echo "=============================================="
echo "Email: $EMAIL"
echo "Date Range: $NEXT_MONDAY to $NEXT_WEEK"
echo ""

# ============================================
# PART 1: Recurring Event Support
# ============================================
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "PART 1: Recurring Event Support"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# Test 1.1 - Weekly Recurring Meeting
echo ""
echo "Test 1.1 - Weekly Recurring Meeting"
result=$(run_tool create_event \
    --summary "TEST: Weekly Test Meeting" \
    --start_time "${NEXT_MONDAY}T10:00:00-06:00" \
    --end_time "${NEXT_MONDAY}T11:00:00-06:00" \
    --recurrence '["RRULE:FREQ=WEEKLY"]')
if echo "$result" | grep -q "Successfully created"; then
    pass "Weekly recurring event created"
    EVENT_IDS+=("$(echo "$result" | grep -oE 'ID: [a-zA-Z0-9]+' | head -1 | cut -d' ' -f2)")
else
    fail "Weekly recurring event"
    echo "$result"
fi

# Test 1.2 - Bi-Weekly Meeting
echo ""
echo "Test 1.2 - Bi-Weekly Meeting"
result=$(run_tool create_event \
    --summary "TEST: Bi-Weekly Sync" \
    --start_time "${NEXT_MONDAY}T14:00:00-06:00" \
    --end_time "${NEXT_MONDAY}T15:00:00-06:00" \
    --recurrence '["RRULE:FREQ=WEEKLY;INTERVAL=2"]')
if echo "$result" | grep -q "Successfully created"; then
    pass "Bi-weekly recurring event created"
else
    fail "Bi-weekly recurring event"
fi

# Test 1.3 - Weekday-Only Standup
echo ""
echo "Test 1.3 - Weekday-Only Standup"
result=$(run_tool create_event \
    --summary "TEST: Daily Standup" \
    --start_time "${NEXT_MONDAY}T09:00:00-06:00" \
    --end_time "${NEXT_MONDAY}T09:15:00-06:00" \
    --recurrence '["RRULE:FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR"]')
if echo "$result" | grep -q "Successfully created"; then
    pass "Weekday-only recurring event created"
else
    fail "Weekday-only recurring event"
fi

# Test 1.4 - Monthly Meeting on 15th
echo ""
echo "Test 1.4 - Monthly Meeting on 15th"
result=$(run_tool create_event \
    --summary "TEST: Monthly Review" \
    --start_time "2025-12-15T13:00:00-06:00" \
    --end_time "2025-12-15T14:00:00-06:00" \
    --recurrence '["RRULE:FREQ=MONTHLY;BYMONTHDAY=15"]')
if echo "$result" | grep -q "Successfully created"; then
    pass "Monthly recurring event created"
else
    fail "Monthly recurring event"
fi

# Test 1.5 - Limited Series (10 occurrences)
echo ""
echo "Test 1.5 - Limited Series (10 occurrences)"
result=$(run_tool create_event \
    --summary "TEST: Training Session" \
    --start_time "${NEXT_MONDAY}T15:00:00-06:00" \
    --end_time "${NEXT_MONDAY}T16:00:00-06:00" \
    --recurrence '["RRULE:FREQ=WEEKLY;COUNT=10"]')
if echo "$result" | grep -q "Successfully created"; then
    pass "Limited series event created"
else
    fail "Limited series event"
fi

# Test 1.6 - Event with End Date
echo ""
echo "Test 1.6 - Event with End Date"
result=$(run_tool create_event \
    --summary "TEST: Project Meetings" \
    --start_time "${NEXT_MONDAY}T16:00:00-06:00" \
    --end_time "${NEXT_MONDAY}T17:00:00-06:00" \
    --recurrence '["RRULE:FREQ=WEEKLY;UNTIL=20250630T235959Z"]')
if echo "$result" | grep -q "Successfully created"; then
    pass "Event with end date created"
else
    fail "Event with end date"
fi

# Test 1.7 - Non-Recurring Event (Regression)
echo ""
echo "Test 1.7 - Non-Recurring Event (Regression)"
result=$(run_tool create_event \
    --summary "TEST: Single Meeting" \
    --start_time "${NEXT_MONDAY}T11:00:00-06:00" \
    --end_time "${NEXT_MONDAY}T12:00:00-06:00")
if echo "$result" | grep -q "Successfully created"; then
    pass "Non-recurring event created"
else
    fail "Non-recurring event"
fi

# Test 1.8 - Recurring Event with Google Meet
echo ""
echo "Test 1.8 - Recurring Event with Google Meet"
result=$(run_tool create_event \
    --summary "TEST: Virtual Team Meeting" \
    --start_time "${NEXT_MONDAY}T17:00:00-06:00" \
    --end_time "${NEXT_MONDAY}T18:00:00-06:00" \
    --recurrence '["RRULE:FREQ=WEEKLY"]' \
    --add_google_meet true 2>&1) || true
if echo "$result" | grep -q "Successfully created"; then
    pass "Recurring event with Google Meet created"
elif echo "$result" | grep -q "Invalid conference type"; then
    warn "Google API limitation - recurring + Google Meet not supported"
else
    fail "Recurring event with Google Meet"
fi

# ============================================
# PART 2: Get Events Times Only Tool
# ============================================
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "PART 2: Get Events Times Only Tool"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# Test 2.1 - Get Events for a Week (Weekday Grouping)
echo ""
echo "Test 2.1 - Get Events for a Week (Weekday Grouping)"
result=$(run_tool get_events_times_only \
    --time_min "$NEXT_MONDAY" \
    --time_max "$NEXT_WEEK")
if echo "$result" | grep -qE "(Monday|Tuesday|Wednesday|Thursday|Friday):"; then
    pass "Events grouped by weekday"
else
    fail "Weekday grouping"
    echo "$result"
fi

# Test 2.2 - Verify Test Events Appear
echo ""
echo "Test 2.2 - Verify Test Events Appear"
if echo "$result" | grep -q "TEST:"; then
    pass "Test events appear in output"
else
    warn "Test events may not be in the time range"
fi

# Test 2.3 - All-Day Events Format
echo ""
echo "Test 2.3 - All-Day Events Format"
# Create an all-day test event
all_day_result=$(run_tool create_event \
    --summary "TEST: All-Day Event" \
    --start_time "$NEXT_MONDAY" \
    --end_time "$(date -v+1d -j -f %Y-%m-%d "$NEXT_MONDAY" +%Y-%m-%d 2>/dev/null || date -d "$NEXT_MONDAY +1 day" +%Y-%m-%d)")
if echo "$all_day_result" | grep -q "Successfully created"; then
    # Check if it shows as "(All day)"
    result=$(run_tool get_events_times_only --time_min "$NEXT_MONDAY" --time_max "$NEXT_WEEK")
    if echo "$result" | grep -q "(All day)"; then
        pass "All-day events show correctly"
    else
        warn "All-day event may not show '(All day)' format"
    fi
else
    fail "All-day event creation"
fi

# Test 2.4 - Default Time Min
echo ""
echo "Test 2.4 - Default Time Min"
result=$(run_tool get_events_times_only)
if echo "$result" | grep -q "Events for"; then
    pass "Default time_min works"
else
    fail "Default time_min"
fi

# Test 2.5 - Max Results Limit
echo ""
echo "Test 2.5 - Max Results Limit"
result=$(run_tool get_events_times_only --time_min "$NEXT_MONDAY" --time_max "$NEXT_WEEK" --max_results 3)
event_count=$(echo "$result" | grep -c "  - " || true)
if [ "$event_count" -le 3 ]; then
    pass "Max results limit works (got $event_count events)"
else
    fail "Max results limit (got $event_count events, expected ≤3)"
fi

# ============================================
# PART 3: Regression Tests
# ============================================
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "PART 3: Regression Tests"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# Test 3.1 - List Calendars
echo ""
echo "Test 3.1 - List Calendars"
result=$(run_tool list_calendars)
if echo "$result" | grep -q "Successfully listed"; then
    pass "list_calendars works"
else
    fail "list_calendars"
fi

# Test 3.2 - Get Events (Original Tool)
echo ""
echo "Test 3.2 - Get Events (Original Tool)"
result=$(run_tool get_events --time_min "$NEXT_MONDAY" --time_max "$NEXT_WEEK" --max_results 5)
if echo "$result" | grep -q "Successfully retrieved"; then
    pass "get_events works"
else
    fail "get_events"
fi

# Test 3.3 - Delete Event
echo ""
echo "Test 3.3 - Delete Event"
# Get a test event ID to delete
test_event=$(run_tool get_events --time_min "$NEXT_MONDAY" --time_max "$NEXT_WEEK" --query "TEST:" --max_results 1)
event_id=$(echo "$test_event" | grep -oE 'ID: [a-zA-Z0-9_]+' | head -1 | cut -d' ' -f2)
if [ -n "$event_id" ]; then
    result=$(run_tool delete_event --event_id "$event_id")
    if echo "$result" | grep -q "Successfully deleted"; then
        pass "delete_event works"
    else
        fail "delete_event"
    fi
else
    warn "No test event found to delete"
fi

# ============================================
# CLEANUP
# ============================================
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "CLEANUP"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

echo ""
echo "Deleting all TEST: events..."
deleted=0
while true; do
    test_events=$(run_tool get_events --time_min "2025-01-01" --time_max "2026-12-31" --query "TEST:" --max_results 50)
    event_ids=$(echo "$test_events" | grep -oE 'ID: [a-zA-Z0-9_]+' | cut -d' ' -f2)
    
    if [ -z "$event_ids" ]; then
        break
    fi
    
    for id in $event_ids; do
        run_tool delete_event --event_id "$id" > /dev/null 2>&1 && ((deleted++)) || true
    done
done

echo "Deleted $deleted test events"
pass "Cleanup complete"

echo ""
echo "=============================================="
echo "Test Summary"
echo "=============================================="
echo "All tests completed. Review output above for results."
echo ""

