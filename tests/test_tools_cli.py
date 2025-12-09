"""Tests for tools_cli.py argument parsing and value conversion."""


# Extract the functions from tools_cli.py for testing
# These are defined inside main(), so we need to replicate them here for unit testing
def unescape_shell_chars(value: str) -> str:
    """Unescape common shell-escaped characters in string values."""
    if not isinstance(value, str):
        return value
    value = value.replace(r"\\", "\x00")  # Temporarily protect \\
    value = value.replace(r"\!", "!")
    value = value.replace(r"\$", "$")
    value = value.replace(r"\`", "`")
    value = value.replace(r"\#", "#")
    value = value.replace(r"\"", '"')
    value = value.replace(r"\'", "'")
    value = value.replace("\x00", "\\")  # Restore single backslash from \\
    return value


class TestUnescapeShellChars:
    """Test cases for unescape_shell_chars function."""

    def test_unescape_exclamation_mark(self):
        """Test that escaped exclamation marks are unescaped."""
        assert unescape_shell_chars(r"Sheet1\!A1:B2") == "Sheet1!A1:B2"

    def test_unescape_exclamation_mark_in_middle(self):
        """Test unescaping ! in the middle of a string."""
        assert unescape_shell_chars(r"prefix\!suffix") == "prefix!suffix"

    def test_unescape_multiple_exclamation_marks(self):
        """Test unescaping multiple exclamation marks."""
        assert unescape_shell_chars(r"a\!b\!c") == "a!b!c"

    def test_unescape_dollar_sign(self):
        """Test that escaped dollar signs are unescaped."""
        assert unescape_shell_chars(r"\$HOME") == "$HOME"

    def test_unescape_backtick(self):
        """Test that escaped backticks are unescaped."""
        assert unescape_shell_chars(r"\`command\`") == "`command`"

    def test_unescape_hash(self):
        """Test that escaped hash signs are unescaped."""
        assert unescape_shell_chars(r"\#comment") == "#comment"

    def test_unescape_double_quote(self):
        """Test that escaped double quotes are unescaped."""
        assert unescape_shell_chars(r"\"quoted\"") == '"quoted"'

    def test_unescape_single_quote(self):
        """Test that escaped single quotes are unescaped."""
        assert unescape_shell_chars(r"\'quoted\'") == "'quoted'"

    def test_preserve_double_backslash_as_single(self):
        """Test that double backslash becomes single backslash."""
        assert unescape_shell_chars(r"path\\to\\file") == "path\\to\\file"

    def test_no_change_for_normal_text(self):
        """Test that normal text without escapes is unchanged."""
        assert unescape_shell_chars("normal text") == "normal text"

    def test_no_change_for_unescaped_exclamation(self):
        """Test that unescaped exclamation marks are preserved."""
        assert unescape_shell_chars("Sheet1!A1:B2") == "Sheet1!A1:B2"

    def test_non_string_returns_unchanged(self):
        """Test that non-string values are returned unchanged."""
        assert unescape_shell_chars(123) == 123
        assert unescape_shell_chars(None) is None
        assert unescape_shell_chars(["list"]) == ["list"]

    def test_empty_string(self):
        """Test that empty string returns empty string."""
        assert unescape_shell_chars("") == ""

    def test_complex_range_name(self):
        """Test a complex Google Sheets range name."""
        # This is the exact use case from the bug report
        assert unescape_shell_chars(r"'Sheet Name'\!A1:Z100") == "'Sheet Name'!A1:Z100"

    def test_mixed_escapes(self):
        """Test multiple different escape sequences in one string."""
        assert unescape_shell_chars(r"\!\$\#") == "!$#"
