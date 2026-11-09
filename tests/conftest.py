"""Shared test fixtures — one Excel instance for the entire session, started only when a test asks for it."""

from __future__ import annotations

import pytest


@pytest.fixture(scope="session")
def _shared_excel():
    """Launch Excel once (on the first test that needs it), share it, quit when done.

    Not autouse: the unit tests never touch Excel, so a plain `pytest -m "not excel"` run
    (the review gate) does not launch it or trigger macOS file-access prompts.
    """
    try:
        from phpp_tool.excel_app import excel_app, set_shared_app

        app = excel_app()
        set_shared_app(app)
        yield app
        set_shared_app(None)
        app.quit()
    except Exception:  # noqa: BLE001 -- Excel unavailable for any reason: the Excel tests skip
        yield None


@pytest.fixture()
def require_excel(_shared_excel):
    """Skip the test if Excel is not available."""
    if _shared_excel is None:
        pytest.skip("Excel not available")
