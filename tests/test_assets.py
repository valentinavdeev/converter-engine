"""Local image embedding and hidden-slide SVG capture contracts."""

from __future__ import annotations

import base64
import io
from pathlib import Path

import pytest
from PIL import Image
from pptx import Presentation

from html_to_pptx import convert
from html_to_pptx.converter import (
    EXTRACTION_JS, _rasterize_inline_svgs, _walk_elements,
)


@pytest.mark.asyncio
async def test_local_image_survives_full_conversion(tmp_path: Path):
    image_path = tmp_path / "synthetic image.png"
    Image.new("RGB", (32, 24), (25, 100, 200)).save(image_path)
    source = tmp_path / "deck # local.html"
    source.write_text(
        '<style>.slide {width:1280px;height:720px}</style>'
        '<section class="slide"><img src="synthetic image.png" '
        'style="width:64px;height:48px"></section>',
        encoding="utf-8",
    )
    output = tmp_path / "local.pptx"
    await convert(str(source), str(output))
    presentation = Presentation(str(output))
    pictures = [shape for shape in presentation.slides[0].shapes if hasattr(shape, "image")]
    assert len(pictures) == 1
    assert pictures[0].image.blob == image_path.read_bytes()
    assert pictures[0].image.size == (32, 24)


@pytest.mark.asyncio
async def test_remote_image_is_rejected_without_output(tmp_path: Path):
    source = tmp_path / "remote.html"
    source.write_text(
        '<section class="slide"><img src="https://example.invalid/image.png"></section>',
        encoding="utf-8",
    )
    output = tmp_path / "remote.pptx"
    with pytest.raises(ValueError, match="Unsupported image source"):
        await convert(str(source), str(output))
    assert not output.exists()


@pytest.mark.asyncio
async def test_svg_capture_preserves_authored_layout_and_slide_state(tmp_path: Path):
    from playwright.async_api import async_playwright

    html = """
        <style>
            .slide { width:1280px; height:720px; gap:20px; }
            .grid { display:grid; grid-template-columns:200px 80px; align-items:start; }
            .flex { display:flex; align-items:flex-start; }
            .slide > div { width:100px; height:40px; flex-shrink:0; }
            .slide svg { width:100%; height:40px; }
            .flex svg { width:80px; flex-shrink:0; }
        </style>
        <section class="slide grid preserved" style="display:none;color:red">
            <div></div><svg viewBox="0 0 80 40"><rect width="80" height="40" fill="#19c864"/></svg>
        </section>
        <section class="slide flex preserved" style="display:none;color:blue">
            <div></div><svg viewBox="0 0 80 40"><rect width="80" height="40" fill="#19c864"/></svg>
        </section>
    """
    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(headless=True)
        try:
            page = await browser.new_page(viewport={"width": 1280, "height": 720})
            await page.set_content(html)
            snapshot_js = "() => Array.from(document.querySelectorAll('.slide'), s => [s.getAttribute('style'), s.getAttribute('class')])"
            original = await page.evaluate(snapshot_js)
            measurements = await page.evaluate(EXTRACTION_JS)
            assert await page.evaluate(snapshot_js) == original
            svgs = [
                next(element for element in _walk_elements(slide["elements"]) if element.get("isSvg"))
                for slide in measurements
            ]
            assert [svg["x"] for svg in svgs] == [220, 120]
            assert [svg["width"] for svg in svgs] == [80, 80]
            await _rasterize_inline_svgs(page, measurements)
            assert await page.evaluate(snapshot_js) == original
            for svg in svgs:
                assert svg["isImage"] is True
                assert svg["isSvg"] is False
                pixels = base64.b64decode(svg["src"].split(",", 1)[1])
                with Image.open(io.BytesIO(pixels)) as image:
                    assert image.size == (80, 40)
                    assert image.convert("RGB").getpixel((40, 20)) == (25, 200, 100)
        finally:
            await browser.close()
    source = tmp_path / "inline-svg.html"
    source.write_text(html, encoding="utf-8")
    output = tmp_path / "inline-svg.pptx"
    await convert(str(source), str(output))
    presentation = Presentation(str(output))
    assert len(presentation.slides) == 2
    for slide in presentation.slides:
        pictures = [shape for shape in slide.shapes if hasattr(shape, "image")]
        assert len(pictures) == 1
        assert pictures[0].image.size == (80, 40)


@pytest.mark.asyncio
@pytest.mark.parametrize("alias", ["same", "symlink", "hardlink"])
async def test_convert_rejects_source_as_output(tmp_path: Path, alias: str):
    source = tmp_path / "source.html"
    original = b'<section class="slide">Preserved</section>'
    source.write_bytes(original)
    output = source
    if alias != "same":
        output = tmp_path / "alias.pptx"
        if alias == "symlink":
            output.symlink_to(source)
        else:
            output.hardlink_to(source)
    with pytest.raises(ValueError, match="Input and output must be different files"):
        await convert(str(source), str(output))
    assert source.read_bytes() == original
    assert output.read_bytes() == original
