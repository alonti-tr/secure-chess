"""Shared pytest fixtures for the secure-chess test suite."""

from __future__ import annotations

from pathlib import Path

import pytest


@pytest.fixture
def chess_fixtures_dir() -> Path:
    return Path(__file__).parent / "fixtures"
