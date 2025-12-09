"""
Unit tests for escape sequence interpretation in Google Docs tools.

Tests verify that literal escape sequences (backslash followed by character)
are properly converted to actual characters when text is inserted or replaced.
"""

from gdocs.docs_helpers import interpret_escape_sequences
from gdocs.managers.batch_operation_manager import normalize_operation


class TestInterpretEscapeSequences:
    """Tests for the interpret_escape_sequences helper function."""

    def test_newline_escape_sequence(self):
        """Test that \\n is converted to actual newline."""
        result = interpret_escape_sequences("Hello\\nWorld")
        assert result == "Hello\nWorld"
        assert len(result) == 11  # H-e-l-l-o-\n-W-o-r-l-d

    def test_tab_escape_sequence(self):
        """Test that \\t is converted to actual tab."""
        result = interpret_escape_sequences("Column1\\tColumn2")
        assert result == "Column1\tColumn2"
        assert "\t" in result

    def test_carriage_return_escape_sequence(self):
        """Test that \\r is converted to actual carriage return."""
        result = interpret_escape_sequences("Line1\\rLine2")
        assert result == "Line1\rLine2"
        assert "\r" in result

    def test_backslash_escape_sequence(self):
        """Test that \\\\ is converted to single backslash."""
        result = interpret_escape_sequences("path\\\\to\\\\file")
        assert result == "path\\to\\file"
        # Original has 4 backslashes (escaped as \\\\), result has 2

    def test_multiple_newlines(self):
        """Test multiple consecutive newlines."""
        result = interpret_escape_sequences("Para1\\n\\nPara2")
        assert result == "Para1\n\nPara2"
        assert result.count("\n") == 2

    def test_mixed_escape_sequences(self):
        """Test combination of different escape sequences."""
        result = interpret_escape_sequences("Line1\\nTab:\\tEnd\\\\Done")
        assert result == "Line1\nTab:\tEnd\\Done"

    def test_no_escape_sequences(self):
        """Test that text without backslash is unchanged."""
        text = "Hello World no escapes here"
        result = interpret_escape_sequences(text)
        assert result == text

    def test_none_input(self):
        """Test that None input returns None."""
        result = interpret_escape_sequences(None)
        assert result is None

    def test_empty_string(self):
        """Test that empty string returns empty string."""
        result = interpret_escape_sequences("")
        assert result == ""

    def test_unknown_escape_sequence_preserved(self):
        """Test that unknown escape sequences are preserved as-is."""
        result = interpret_escape_sequences("Hello\\xWorld")
        assert result == "Hello\\xWorld"

    def test_trailing_backslash(self):
        """Test that trailing backslash is preserved."""
        result = interpret_escape_sequences("ends with\\")
        assert result == "ends with\\"

    def test_real_newlines_unchanged(self):
        """Test that actual newlines in input are preserved."""
        text = "Already\nhas\nnewlines"
        result = interpret_escape_sequences(text)
        assert result == text

    def test_crlf_conversion(self):
        """Test Windows-style line endings."""
        result = interpret_escape_sequences("Line1\\r\\nLine2")
        assert result == "Line1\r\nLine2"


class TestNormalizeOperationEscapeSequences:
    """Tests for escape sequence handling in batch operations."""

    def test_normalize_operation_interprets_text_escape(self):
        """Test that normalize_operation interprets escape sequences in text field."""
        op = {"type": "insert", "text": "Hello\\nWorld"}
        normalized = normalize_operation(op)
        assert normalized["text"] == "Hello\nWorld"

    def test_normalize_operation_interprets_replace_text_escape(self):
        """Test that normalize_operation interprets escape sequences in replace_text field."""
        op = {"type": "find_replace", "find_text": "old", "replace_text": "new\\nline"}
        normalized = normalize_operation(op)
        assert normalized["replace_text"] == "new\nline"

    def test_normalize_operation_preserves_none_text(self):
        """Test that None text fields are preserved."""
        op = {"type": "format", "text": None}
        normalized = normalize_operation(op)
        assert normalized["text"] is None

    def test_normalize_operation_without_text(self):
        """Test operations without text field work correctly."""
        op = {"type": "delete", "index": 100, "end_index": 110}
        normalized = normalize_operation(op)
        assert "text" not in normalized

    def test_normalize_operation_normalizes_type_and_escapes(self):
        """Test that both type normalization and escape interpretation work together."""
        op = {"type": "insert", "text": "Line1\\nLine2"}
        normalized = normalize_operation(op)
        assert normalized["type"] == "insert_text"  # Type normalized
        assert normalized["text"] == "Line1\nLine2"  # Escapes interpreted


class TestEscapeSequenceEdgeCases:
    """Edge case tests for escape sequence handling."""

    def test_only_escape_sequences(self):
        """Test string containing only escape sequences."""
        result = interpret_escape_sequences("\\n\\t\\r")
        assert result == "\n\t\r"

    def test_consecutive_backslashes(self):
        """Test multiple consecutive backslashes."""
        # Four backslashes (escaped as \\\\\\\\) should become two
        result = interpret_escape_sequences("a\\\\\\\\b")
        assert result == "a\\\\b"

    def test_escape_at_start(self):
        """Test escape sequence at start of string."""
        result = interpret_escape_sequences("\\nStarts with newline")
        assert result == "\nStarts with newline"

    def test_escape_at_end(self):
        """Test escape sequence at end of string."""
        result = interpret_escape_sequences("Ends with newline\\n")
        assert result == "Ends with newline\n"

    def test_very_long_string(self):
        """Test escape sequences in long string."""
        # Create a long string with escape sequences throughout
        text = "\\n".join(["Paragraph " + str(i) for i in range(100)])
        result = interpret_escape_sequences(text)
        assert result.count("\n") == 99

    def test_unicode_preserved(self):
        """Test that Unicode characters are preserved."""
        result = interpret_escape_sequences("Hello\\n🌍 World\\n日本語")
        assert "🌍" in result
        assert "日本語" in result
        assert result.count("\n") == 2
