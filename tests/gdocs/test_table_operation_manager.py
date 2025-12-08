"""
Tests for TableOperationManager.

Focus on _find_table_at_index method which is critical for finding
newly created tables when using after_heading parameter.
"""
from gdocs.managers.table_operation_manager import TableOperationManager


class TestFindTableAtIndex:
    """Tests for _find_table_at_index method."""

    def setup_method(self):
        """Create a TableOperationManager for testing."""
        # Service is not needed for _find_table_at_index
        self.manager = TableOperationManager(service=None)

    def test_find_table_at_exact_index(self):
        """Table at exact insertion index should be found."""
        tables = [
            {'start_index': 10, 'end_index': 50},
            {'start_index': 100, 'end_index': 150},
            {'start_index': 200, 'end_index': 250},
        ]

        result = self.manager._find_table_at_index(tables, 100)
        assert result == 1

    def test_find_table_closest_to_index(self):
        """Table closest to insertion index should be found."""
        tables = [
            {'start_index': 10, 'end_index': 50},
            {'start_index': 105, 'end_index': 150},  # 5 away from 100
            {'start_index': 200, 'end_index': 250},
        ]

        result = self.manager._find_table_at_index(tables, 100)
        assert result == 1

    def test_find_table_with_small_offset(self):
        """Table slightly offset from insertion index should be found."""
        tables = [
            {'start_index': 10, 'end_index': 50},
            {'start_index': 100, 'end_index': 150},  # First table
            {'start_index': 103, 'end_index': 160},  # Second table, 3 away from 100
        ]

        # Should find the table at index 100 exactly
        result = self.manager._find_table_at_index(tables, 100)
        assert result == 1

    def test_empty_tables_returns_none(self):
        """Empty table list returns None."""
        result = self.manager._find_table_at_index([], 100)
        assert result is None

    def test_single_table_found(self):
        """Single table in document is found."""
        tables = [
            {'start_index': 42, 'end_index': 100},
        ]

        result = self.manager._find_table_at_index(tables, 42)
        assert result == 0

    def test_fallback_to_last_table_when_far_from_insertion(self):
        """When no table is close, falls back to last table."""
        tables = [
            {'start_index': 10, 'end_index': 50},
            {'start_index': 200, 'end_index': 250},  # 100 away from insertion point
        ]

        # With insertion at 100, both tables are > 50 away
        # Should fall back to the last table
        result = self.manager._find_table_at_index(tables, 100)
        assert result == 1  # Last table as fallback

    def test_find_first_table_when_insertion_at_start(self):
        """Table at start of document is found when insertion index is 1."""
        tables = [
            {'start_index': 2, 'end_index': 50},  # Table near start
            {'start_index': 100, 'end_index': 150},
        ]

        result = self.manager._find_table_at_index(tables, 1)
        assert result == 0

    def test_multiple_tables_finds_closest(self):
        """With multiple tables, finds the one closest to insertion point."""
        tables = [
            {'start_index': 10, 'end_index': 50},
            {'start_index': 80, 'end_index': 120},   # 20 away from 100
            {'start_index': 95, 'end_index': 140},   # 5 away from 100
            {'start_index': 150, 'end_index': 200},  # 50 away from 100
        ]

        result = self.manager._find_table_at_index(tables, 100)
        assert result == 2  # Table at 95 is closest to 100

    def test_after_heading_scenario(self):
        """Simulate after_heading scenario where table is inserted mid-document."""
        # Document with multiple sections and tables
        tables = [
            {'start_index': 50, 'end_index': 100},    # Table in first section
            {'start_index': 250, 'end_index': 300},   # Table in third section
            {'start_index': 150, 'end_index': 200},   # Newly inserted table after heading at index 148
        ]

        # Sort tables by start_index as find_tables does
        tables_sorted = sorted(tables, key=lambda t: t['start_index'])

        # Insertion was at index 148, so table at 150 should be found
        result = self.manager._find_table_at_index(tables_sorted, 148)
        assert tables_sorted[result]['start_index'] == 150

    def test_table_inserted_between_existing_tables(self):
        """Table inserted between two existing tables is found correctly."""
        tables = [
            {'start_index': 10, 'end_index': 50},
            {'start_index': 100, 'end_index': 150},   # Newly inserted table
            {'start_index': 200, 'end_index': 250},
        ]

        result = self.manager._find_table_at_index(tables, 100)
        assert result == 1  # Middle table

    def test_within_threshold_distance(self):
        """Tables within 50 index threshold are valid matches."""
        tables = [
            {'start_index': 145, 'end_index': 200},  # 45 away from 100
        ]

        result = self.manager._find_table_at_index(tables, 100)
        assert result == 0  # Should match since within 50

    def test_exactly_at_threshold(self):
        """Table exactly at threshold distance (50) is still matched."""
        tables = [
            {'start_index': 150, 'end_index': 200},  # Exactly 50 away from 100
        ]

        result = self.manager._find_table_at_index(tables, 100)
        assert result == 0
