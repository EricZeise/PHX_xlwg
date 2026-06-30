"""Shared test fixtures — one Excel instance for the entire session."""

from __future__ import annotations

import pytest


@pytest.fixture(scope="session", autouse=True)
def _shared_excel():
    """Launch Excel once, share it across all tests, quit when done."""
    try:
        from phpp_tool.excel_app import excel_app, set_shared_app

        app = excel_app()
        set_shared_app(app)
        yield app
        set_shared_app(None)
        app.quit()
    except Exception:
        yield None


@pytest.fixture()
def require_excel(_shared_excel):
    """Skip the test if Excel is not available."""
    if _shared_excel is None:
        pytest.skip("Excel not available")
