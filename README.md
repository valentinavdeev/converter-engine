# converter-engine

A maintained fork of [Design-Arena/html-to-pptx](https://github.com/Design-Arena/html-to-pptx), based on upstream commit `3b54c7ffb02b453ac85222d7c7ac9024d6677614`.

Converts a constrained HTML slide format into editable PowerPoint text, shapes and images. Chromium measures layout; `python-pptx` writes OOXML. This is **not** a general-purpose browser-to-PowerPoint renderer and does not promise pixel-identical output across Office applications.

Distribution name: `converter-engine`. Python import: `html_to_pptx`. Commands: `html-to-pptx` and `html-to-pptx-compare`.

## Install from source

The fidelity changes are proposed on `fix/html-pptx-fidelity`; until its PR is merged, use that branch rather than the fork's unchanged `main`.

```sh
git clone --branch fix/html-pptx-fidelity https://github.com/valentinavdeev/converter-engine.git
cd converter-engine
python -m venv .venv
. .venv/bin/activate
python -m pip install '.[compare]'
python -m playwright install chromium
```

Python 3.10+ is required. Dependencies and Chromium must be installed explicitly; conversion does not install them. No PyPI release is implied by the distribution name. The upstream MIT license and Design Arena attribution are preserved.

## Convert

```sh
html-to-pptx deck.html deck.pptx
```

```python
import asyncio
from html_to_pptx import convert

asyncio.run(convert('deck.html', 'deck.pptx'))
```

Input/output aliases are rejected to protect the HTML source. An existing, distinct output PPTX may be replaced. Local `<img>` sources are embedded in memory without rewriting the HTML. Fonts and images are awaited before measurement. HTTP(S) resources are blocked: provide local fonts and local or embedded images. Convert trusted HTML; network restrictions are not a security sandbox for arbitrary documents.

### Two-stage API and existing browsers

```python
import asyncio
from html_to_pptx import extract_measurements, render_pptx

async def main():
    measurements = await extract_measurements('deck.html')
    render_pptx(measurements).save('deck.pptx')

asyncio.run(main())
```

For a host that already provides Chromium:

```sh
html-to-pptx --extract-js
html-to-pptx --measurements measurements.json deck.pptx
```

Execute the printed function in the browser after fonts/images load and serialize its return value as JSON. The measurement function alone does not read local image bytes or rasterize SVG: supply `data:image/...` image sources in the DOM and rasterized SVG data when using this route. The full Python extraction path performs these preparation steps. Measurement JSON can contain embedded document images; treat it as document data, not public diagnostic output.

## HTML input profile

Use a fixed 1920×1080 CSS-pixel canvas per `.slide`. The resulting PPTX is 13⅓×7.5 inches (16:9). The profile constrains exportable features, not the artistic arrangement of slides.

```html
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  .slide {
    width: 1920px; height: 1080px; position: relative;
    display: none; flex-direction: column; justify-content: center;
    padding: 100px; background: white; font-family: Arial, sans-serif;
  }
  .slide.active { display: flex; }
  .card { background: #edf4fa; border-radius: 24px; padding: 40px; }
  h1 { font-size: 64px; line-height: 1.1; }
  p { font-size: 32px; line-height: 1.4; }
</style>
</head>
<body>
<section class="slide active">
  <div class="card" role="group" aria-label="Introduction">
    <h1>An editable presentation</h1>
    <p>Keep the source HTML and review the actual PPTX render.</p>
  </div>
</section>
</body>
</html>
```

- Normal block, flex and grid layout are measured in the browser. Slide activation preserves authored display rules; it does not force every slide into flex. Prefer a shallow tree and explicit dimensions where they matter.
- Absolute positioning is supported, but do not assume a tightly measured browser text box has identical Office metrics.
- Text remains native. Authored sizes are retained instead of automatically shrinking them. Fix overflow in HTML; changing text inside PowerPoint does not rerun HTML layout.
- CSS tracking, preformatted spaces/line breaks, mixed inline text and basic alignment are retained. Use fonts installed in both the browser environment and the target viewer; family substitutions remain, fonts are not embedded automatically.
- `role="group"` or `data-pptx-group` marks a native component group. `aria-label` or `data-pptx-group` supplies its name. Ordinary layout wrappers do not automatically become groups.
- Positive fractional-size rules are retained. Uniform rounded rectangles and equal rounded top corners with square bottoms have native geometry.
- Solid fills, borders and supported linear gradients retain native alpha. Overlays are not automatically baked into underlying images. Default Office theme shadows are disabled; a simple authored outer shadow is supported.
- 2D rotations account for CSS `transform-origin`; skew/reflection/3D transforms are rejected. Arbitrary transformations are not supported.
- Local/base64 `<img>` sources become pictures. Inline SVG and conic gradients become raster images, not editable vector diagrams. `background-image: url(...)` is unsupported.

### Linear gradients

Supported: a single linear gradient with computed RGB/RGBA colors (comma or space/slash syntax), `transparent`, percentage positions, omitted positions, two positions per color, repeated stops/hard transitions, degree angles and directional keywords. Missing positions interpolate between anchors; decreasing positions follow CSS fixup. Stops outside 0–100% are clipped with alpha-aware boundary colors.

Not supported: pixel/other length positions, `calc()`, interpolation hints, non-RGB color spaces, non-degree angle units, radial or multi-layer gradients. Unsupported linear syntax falls back to an existing solid background rather than inventing evenly spaced stops. This module is not a comprehensive CSS validator: do not interpret successful export as proof that every authored effect was reproduced.

### Remaining limitations

- Complex `z-index`/stacking contexts, filters and pseudo-element transforms are not fully reproduced.
- Group opacity multiplies member alpha, rather than isolated CSS subtree compositing; overlapping translucent descendants may differ.
- Arbitrary four-corner/elliptical radii, advanced inline baseline shifts, tab-stop metrics and vertical text layout are not universally reproduced.
- Animations and transitions are not exported.
- Font substitution and Office rendering can change text geometry. A real PPTX render is required; browser screenshots and XML assertions alone are insufficient.

## Visual comparison

Requires LibreOffice on `PATH` and the `compare` extra:

```sh
html-to-pptx-compare tests/fixtures/fidelity.html -o results
```

This produces source HTML screenshots, a PPTX, LibreOffice PDF/PNG renders and side-by-side comparisons under `results/fidelity/`. The output subdirectory must not already exist: the tool will not recursively delete an existing directory. LibreOffice uses a separate temporary profile, not a user's active profile.

Both `examples/demo.html` and `tests/fixtures/fidelity.html` are synthetic examples. Do not commit private slide decks, source images, measurement JSON or rendered customer documents.

## Changes relative to upstream

Bug fixes address:

- [#4](https://github.com/Design-Arena/html-to-pptx/issues/4): significant preformatted boundary spaces;
- [#5](https://github.com/Design-Arena/html-to-pptx/issues/5): discarded thin geometry;
- [#6](https://github.com/Design-Arena/html-to-pptx/issues/6): lost text tracking;
- [#7](https://github.com/Design-Arena/html-to-pptx/issues/7): discarded gradient stop positions;
- [#2](https://github.com/Design-Arena/html-to-pptx/issues/2): unauthored theme shadows (upstream also has [PR #3](https://github.com/Design-Arena/html-to-pptx/pull/3)).

Native alpha, explicit groups, authored font-size preservation and noncentral rotation origins are deliberate fork choices/extensions, not claims that the upstream's documented behavior was accidental. The changes are reviewed through a PR in this fork; no automatic merge or upstream PR submission is implied.

## Development

```sh
python -m pip install -e '.[dev]'
python -m playwright install chromium
pytest -v
html-to-pptx tests/fixtures/fidelity.html /tmp/fidelity.pptx
```

For a browser installation contained inside the virtual environment, set `PLAYWRIGHT_BROWSERS_PATH=0` both during `playwright install` and while running tests/commands. CI tests Python 3.10–3.13, exercises the installed CLI and builds the package. It does not certify Microsoft PowerPoint rendering. PyPI publication automation is intentionally absent.

### Local verification — 2026-10-03

- Python 3.12: 65 tests passed, including real Chromium extraction and saved-PPTX regressions.
- Editable installation and sdist/wheel builds succeeded.
- The installed comparison CLI converted the one-slide fidelity fixture and the five-slide upstream demo; all six HTML/LibreOffice render pairs were visually inspected.
- The fidelity fixture retained preformatted text, tracking, thin lines, connected rotated marks, native groups and the translucent overlay. The demo still shows renderer/font-metric and gradient differences; this is not pixel-identical conversion.
- Microsoft PowerPoint and interactive editing were not tested. The next integration step is review/merge of the fork PR before connecting any presentation workflow or deployed environment.

## License

[MIT](LICENSE). Original work: Copyright (c) 2026 Design Arena. This fork retains the original attribution and license.
