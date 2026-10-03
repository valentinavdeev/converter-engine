"""Comparison exports must not erase an existing user directory."""
from pathlib import Path

import pytest

from html_to_pptx.compare import process_file


@pytest.mark.asyncio
async def test_existing_comparison_directory_is_not_deleted(tmp_path: Path):
    source = tmp_path / 'deck.html'
    source.write_text('<section class="slide"></section>', encoding='utf-8')
    target = tmp_path / 'results' / 'deck'
    target.mkdir(parents=True)
    original = target / 'user-notes.txt'
    original.write_text('Keep this unrelated material.', encoding='utf-8')
    with pytest.raises(FileExistsError):
        await process_file(source, tmp_path / 'results', only_convert=True)
    assert original.read_text(encoding='utf-8') == 'Keep this unrelated material.'
