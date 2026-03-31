"""Tests for models (write gate, constants)."""

from __future__ import annotations

from unittest.mock import patch


class TestWriteGate:
    def test_write_disabled_by_default(self):
        from office365_blade_mcp.models import is_write_enabled, require_write

        with patch.dict("os.environ", {}, clear=True):
            assert is_write_enabled() is False
            assert require_write() is not None
            assert "disabled" in require_write().lower()

    def test_write_enabled(self):
        from office365_blade_mcp.models import is_write_enabled, require_write

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "true"}):
            assert is_write_enabled() is True
            assert require_write() is None

    def test_write_enabled_case_insensitive(self):
        from office365_blade_mcp.models import is_write_enabled

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "True"}):
            assert is_write_enabled() is True

    def test_write_disabled_with_other_value(self):
        from office365_blade_mcp.models import is_write_enabled

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "yes"}):
            assert is_write_enabled() is False


class TestScopes:
    def test_read_scopes(self):
        from office365_blade_mcp.models import get_scopes

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "false"}):
            scopes = get_scopes()
            assert "Mail.Read" in scopes
            assert "Mail.ReadWrite" not in scopes

    def test_write_scopes(self):
        from office365_blade_mcp.models import get_scopes

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "true"}):
            scopes = get_scopes()
            assert "Mail.ReadWrite" in scopes
            assert "Mail.Send" in scopes


class TestConstants:
    def test_defaults(self):
        from office365_blade_mcp.models import DEFAULT_LIMIT, MAX_BATCH_SIZE, MAX_BODY_CHARS

        assert DEFAULT_LIMIT == 20
        assert MAX_BATCH_SIZE == 50
        assert MAX_BODY_CHARS == 50_000
