"""Tests for the core conversion pipeline."""

from __future__ import annotations

from pathlib import Path

import pytest

from html_to_pptx import extract_measurements



@pytest.mark.asyncio
async def test_file_not_found_raises():
    with pytest.raises(FileNotFoundError):
        await extract_measurements("/nonexistent/path.html")


@pytest.mark.asyncio
async def test_empty_html_raises(tmp_path: Path):
    empty = tmp_path / "empty.html"
    empty.write_text("<html><body></body></html>")
    with pytest.raises(ValueError, match="No slides found"):
        await extract_measurements(str(empty))


