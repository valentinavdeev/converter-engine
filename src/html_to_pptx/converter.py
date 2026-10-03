"""HTML slide deck → editable PPTX converter.

Converts a structured HTML slide deck into an editable PowerPoint file
by measuring DOM elements in a headless browser (Playwright) and mapping
them to python-pptx shapes.

This module has no Agon framework dependencies and can be used standalone.

Pipeline:
  1. Load the HTML in headless Chromium via Playwright
  2. For each <section class="slide">, measure every visible element's
     bounding box, computed styles, text content, and image data
  3. Rasterize inline <svg> elements to PNG via Playwright screenshots
  4. Map each measurement to a python-pptx shape: text boxes, picture
     shapes, gradient fills, rounded rectangles, and translucent overlays
  5. Save the Presentation as a .pptx file

Supported HTML features:
  - CSS linear-gradient backgrounds (solid, transparent stops, any angle)
  - Gradient text fill (-webkit-background-clip: text) → PPTX gradient run
  - Inline <svg> icons → rasterized PNG pictures
  - Rounded corners on images and shapes (border-radius, including %)
  - Translucent overlays (rgba backgrounds with alpha)
  - Mixed inline content (<p>text <strong>bold</strong> more</p>)
  - RTL text direction
  - List markers (bullets and ordered numbers)
  - CSS pseudo-elements (::before/::after) for decorative accents
  - Border-left accent bars

The HTML canvas is fixed at 1920×1080 CSS pixels. Measurements are
converted to PowerPoint's coordinate system (13.333 × 7.5 inches).

Usage as CLI:
    html-to-pptx input.html [output.pptx]

Usage as library:
    from html_to_pptx import convert, extract_measurements, render_pptx

    # Full pipeline
    await convert("slides.html", "output.pptx")

    # Or step by step
    measurements = await extract_measurements("slides.html")
    prs = render_pptx(measurements)
    prs.save("output.pptx")

Requirements:
    pip install python-pptx playwright lxml
    python -m playwright install chromium
"""

from __future__ import annotations

import base64
import logging
import math
import mimetypes
import os
import re
import tempfile
from pathlib import Path
from urllib.parse import unquote, urlsplit

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Standard widescreen slide dimensions (16:9)
SLIDE_WIDTH_INCHES = 40 / 3
SLIDE_HEIGHT_INCHES = 7.5

# HTML canvas dimensions — the deck is authored at this fixed resolution
SLIDE_CANVAS_WIDTH_PX = 1920
SLIDE_CANVAS_HEIGHT_PX = 1080

# px-to-inch conversion: Playwright reports CSS pixels, PPTX uses inches
PIXELS_TO_INCHES_X = SLIDE_WIDTH_INCHES / SLIDE_CANVAS_WIDTH_PX
PIXELS_TO_INCHES_Y = SLIDE_HEIGHT_INCHES / SLIDE_CANVAS_HEIGHT_PX

# Font size scaling: CSS px -> PowerPoint pt, adjusted for the same
# layout scale as position coordinates so text proportions match.
# 96 CSS px = 1 inch at standard DPI; the ratio ensures font sizes
# and element positions use the same scale factor.
_LAYOUT_SCALE = SLIDE_WIDTH_INCHES / (SLIDE_CANVAS_WIDTH_PX / 96.0)
CSS_PX_TO_PT = 0.75 * _LAYOUT_SCALE


MAX_HTML_SIZE_MB = 10

PLAYWRIGHT_TIMEOUT_MS = 30_000


# ---------------------------------------------------------------------------
# DOM measurement script (injected into the page via Playwright)
#
# Shows each slide in isolation (hiding siblings), clears ancestor CSS
# transforms so getBoundingClientRect returns layout-space coordinates,
# then recursively measures every visible element.
#
# Pure containers (no text, no image, no visible background, single child)
# are collapsed to reduce nesting without losing visual information.
# ---------------------------------------------------------------------------

_SLIDE_STATE_JS = """
    function captureSlideState(slides) {
        const states = Array.from(slides, slide => ({
            slide, style: slide.getAttribute('style'), className: slide.getAttribute('class'),
        }));
        function restore() {
            for (const state of states) {
                if (state.style === null) state.slide.removeAttribute('style');
                else state.slide.setAttribute('style', state.style);
                if (state.className === null) state.slide.removeAttribute('class');
                else state.slide.setAttribute('class', state.className);
            }
        }
        function show(index) {
            restore();
            states.forEach(({slide}, i) => {
                if (i === index) {
                    if (slide.style.display === 'none') slide.style.removeProperty('display');
                    slide.classList.add('active');
                    if (getComputedStyle(slide).display === 'none') slide.style.display = 'block';
                } else {
                    slide.style.display = 'none';
                    slide.classList.remove('active');
                }
            });
        }
        return {show, restore};
    }
"""


EXTRACTION_JS = """
() => {
""" + _SLIDE_STATE_JS + """
    const slides = document.querySelectorAll('.slide');
    const results = [];
    let _svgCounter = 0;
    const INLINE_TAGS = new Set([
        'span','strong','em','b','i','a','code','mark','sub','sup',
        'small','u','s','del','abbr','cite','q','time','var','kbd',
    ]);

    function normalizeText(text, whiteSpace) {
        text = text.replace(/\\r\\n?/g, '\\n');
        if (['pre', 'pre-wrap', 'break-spaces'].includes(whiteSpace)) return text;
        if (whiteSpace === 'pre-line') {
            return text.replace(/[\\t\\f ]+/g, ' ').replace(/ *\\n */g, '\\n');
        }
        return text.replace(/[\\t\\n\\f\\r ]+/g, ' ');
    }

    function getDirectText(el) {
        const whiteSpace = getComputedStyle(el).whiteSpace;
        let text = '';
        for (const node of el.childNodes) {
            if (node.nodeType === Node.TEXT_NODE) text += node.textContent;
        }
        text = normalizeText(text, whiteSpace);
        return ['pre', 'pre-wrap', 'break-spaces'].includes(whiteSpace)
            ? text : text.replace(/^ +| +$/g, '');
    }

    function collectInlineRuns(el) {
        const runs = [];
        function visit(parent, href = null, opacity = 1) {
            const cs = getComputedStyle(parent);
            const childBgImage = cs.backgroundImage !== 'none' ? cs.backgroundImage : null;
            const fillColor = cs.webkitTextFillColor || '';
            const gradient = (fillColor === 'transparent' || fillColor === 'rgba(0, 0, 0, 0)')
                && childBgImage && childBgImage.includes('gradient');
            const properties = {
                color: cs.color,
                fontSize: parseFloat(cs.fontSize),
                fontFamily: cs.fontFamily,
                fontWeight: cs.fontWeight,
                fontStyle: cs.fontStyle,
                letterSpacing: cs.letterSpacing === 'normal' ? 0 : (parseFloat(cs.letterSpacing) || 0),
                whiteSpace: cs.whiteSpace,
                opacity: opacity,
                textTransform: cs.textTransform,
                href: href,
                isGradientText: !!gradient,
                backgroundImage: gradient ? childBgImage : null,
            };
            for (const node of parent.childNodes) {
                if (node.nodeType === Node.TEXT_NODE) {
                    const segments = normalizeText(node.textContent, cs.whiteSpace).split('\\n');
                    segments.forEach((text, index) => {
                        if (index) runs.push({...properties, text: '\\n'});
                        if (text) runs.push({...properties, text: text});
                    });
                } else if (node.nodeType === Node.ELEMENT_NODE) {
                    const tag = node.tagName.toLowerCase();
                    if (['script', 'style', 'link', 'meta'].includes(tag)) continue;
                    const style = getComputedStyle(node);
                    if (style.display === 'none' || style.visibility === 'hidden' || ['absolute', 'fixed'].includes(style.position)) continue;
                    if (tag === 'br') {
                        runs.push({...properties, text: '\\n'});
                    } else if (style.display.startsWith('inline') || INLINE_TAGS.has(tag)) {
                        visit(node, tag === 'a' ? node.getAttribute('href') : href, opacity * (parseFloat(style.opacity) || 0));
                    }
                }
            }
        }
        visit(el);
        // Collapsible whitespace is shared across inline boundaries, but NBSP
        // and preformatted boundary spaces remain literal text.
        let atLineStart = true;
        let previousSpace = false;
        for (let index = 0; index < runs.length; index++) {
            const run = runs[index];
            if (run.text === '\\n') {
                for (let previous = index - 1; previous >= 0; previous--) {
                    const preceding = runs[previous];
                    if (preceding.text === '\\n' || ['pre', 'pre-wrap', 'break-spaces'].includes(preceding.whiteSpace)) break;
                    preceding.text = preceding.text.replace(/ +$/, '');
                    if (preceding.text) break;
                }
                atLineStart = true;
                previousSpace = false;
                continue;
            }
            const preserve = ['pre', 'pre-wrap', 'break-spaces'].includes(run.whiteSpace);
            if (!preserve) {
                if (atLineStart || previousSpace) run.text = run.text.replace(/^ +/, '');
            }
            if (run.text) {
                atLineStart = false;
                previousSpace = run.text.endsWith(' ');
            }
        }
        for (let index = runs.length - 1; index >= 0; index--) {
            const run = runs[index];
            if (run.text === '\\n') continue;
            if (!['pre', 'pre-wrap', 'break-spaces'].includes(run.whiteSpace)) {
                run.text = run.text.replace(/ +$/, '');
            }
            if (run.text) break;
        }
        return runs.filter(run => run.text.length);
    }

    const textMeasureCanvas = document.createElement('canvas');
    const textMeasureContext = textMeasureCanvas.getContext('2d');
    function measureTextGeometry(el) {
        const style = getComputedStyle(el);
        const result = {whiteSpace: style.whiteSpace};
        if (style.writingMode.startsWith('vertical')) return result;
        const elementRect = el.getBoundingClientRect();
        const rects = [];
        let unspacedWidth = 0;
        function visit(parent) {
            const cs = getComputedStyle(parent);
            textMeasureContext.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
            const metrics = textMeasureContext.measureText('Mg');
            const ascent = metrics.fontBoundingBoxAscent;
            const descent = metrics.fontBoundingBoxDescent;
            if (!Number.isFinite(ascent) || !Number.isFinite(descent)) return;
            for (const node of parent.childNodes) {
                if (node.nodeType === Node.TEXT_NODE && node.textContent) {
                    const range = document.createRange();
                    range.selectNodeContents(node);
                    const nodeRects = Array.from(range.getClientRects());
                    for (const rect of nodeRects) {
                        if (!rect.height || (!rect.width && !['pre', 'pre-wrap', 'break-spaces', 'pre-line'].includes(cs.whiteSpace))) continue;
                        rects.push({
                            x: rect.left - elementRect.left,
                            y: rect.top - elementRect.top,
                            width: rect.width, height: rect.height,
                            baseline: rect.top - elementRect.top + ascent,
                            ascent: ascent, descent: descent,
                        });
                    }
                    textMeasureContext.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
                    let text = normalizeText(node.textContent, cs.whiteSpace);
                    if (cs.textTransform === 'uppercase') text = text.toUpperCase();
                    else if (cs.textTransform === 'lowercase') text = text.toLowerCase();
                    unspacedWidth += textMeasureContext.measureText(text).width;
                } else if (node.nodeType === Node.ELEMENT_NODE) {
                    const childStyle = getComputedStyle(node);
                    if (childStyle.display === 'none' || childStyle.visibility === 'hidden' || ['absolute', 'fixed'].includes(childStyle.position)) continue;
                    if (childStyle.display.startsWith('inline') || INLINE_TAGS.has(node.tagName.toLowerCase())) visit(node);
                }
            }
        }
        visit(el);
        if (!rects.length) return result;
        const baselines = [];
        for (const rect of rects) {
            if (!baselines.some(baseline => Math.abs(baseline - rect.baseline) < 1)) baselines.push(rect.baseline);
        }
        baselines.sort((a, b) => a - b);
        const first = rects.reduce((a, b) => a.baseline <= b.baseline ? a : b);
        const left = Math.min(...rects.map(rect => rect.x));
        const top = Math.min(...rects.map(rect => rect.y));
        const right = Math.max(...rects.map(rect => rect.x + rect.width));
        const bottom = Math.max(...rects.map(rect => rect.y + rect.height));
        result.textGeometry = {
            x: left, y: top, width: right - left, height: bottom - top,
            baseline: first.baseline, ascent: first.ascent, descent: first.descent,
            lineHeight: parseFloat(style.lineHeight) || (baselines.length > 1 ? baselines[1] - baselines[0] : first.height),
            lineCount: baselines.length,
            unspacedWidth: baselines.length === 1 ? unspacedWidth : 0,
        };
        return result;
    }

    function hasChildElementText(el) {
        for (const child of el.children) {
            if (child.textContent && child.textContent.trim()) return true;
        }
        return false;
    }

    function clearAncestorTransforms(el) {
        const saved = [];
        let current = el;
        while (current && current !== document.documentElement) {
            const computed = getComputedStyle(current);
            if (computed.transform && computed.transform !== 'none') {
                saved.push({el: current, original: current.style.transform});
                current.style.transform = 'none';
            }
            current = current.parentElement;
        }
        return saved;
    }

    function restoreTransforms(saved) {
        for (const item of saved) {
            item.el.style.transform = item.original;
        }
    }

    function measureElement(el, slideRect, depth) {
        if (depth > 15) return null;

        const style = getComputedStyle(el);
        if (style.display === 'none' || style.visibility === 'hidden') return null;
        if (parseFloat(style.opacity) === 0) return null;

        // OOXML rotates around a shape's center. The transformed DOM rectangle's
        // center is also the transformed local center, even for a CSS origin on
        // an edge. Keep that center but use unrotated local dimensions.
        const visualRect = el.getBoundingClientRect();
        let matrix = new DOMMatrix();
        let ancestor = el;
        while (ancestor && !ancestor.classList.contains('slide')) {
            const transform = getComputedStyle(ancestor).transform;
            if (transform && transform !== 'none') {
                matrix = new DOMMatrix(transform).multiply(matrix);
            }
            ancestor = ancestor.parentElement;
        }
        if (!matrix.is2D || Math.abs(matrix.a * matrix.c + matrix.b * matrix.d) > 0.0001 ||
            matrix.a * matrix.d - matrix.b * matrix.c < 0) {
            throw new Error('Unsupported CSS skew, reflection, or 3D transform: ' + el.tagName);
        }
        const savedTransforms = clearAncestorTransforms(el);
        const localRect = el.getBoundingClientRect();
        const textMeasurement = measureTextGeometry(el);
        restoreTransforms(savedTransforms);
        const ownRotation = Math.atan2(matrix.b, matrix.a) * 180 / Math.PI;
        const width = localRect.width * Math.hypot(matrix.a, matrix.b);
        const height = localRect.height * Math.hypot(matrix.c, matrix.d);
        const rect = {width, height};
        const relX = (visualRect.left + visualRect.right - width) / 2 - slideRect.left;
        const relY = (visualRect.top + visualRect.bottom - height) / 2 - slideRect.top;

        if (width <= 0 || height <= 0) return null;
        if (visualRect.right <= slideRect.left || visualRect.bottom <= slideRect.top) return null;
        if (visualRect.left >= slideRect.right || visualRect.top >= slideRect.bottom) return null;

        let directText = getDirectText(el);
        const tag = el.tagName.toLowerCase();

        let markerColor = null;
        let markerPrefix = '';
        if (tag === 'li') {
            const parentTag = el.parentElement ? el.parentElement.tagName.toLowerCase() : '';
            const listStyle = getComputedStyle(el).listStyleType;
            if (parentTag === 'ol') {
                const index = Array.from(el.parentElement.children).indexOf(el) + 1;
                markerPrefix = index + '. ';
                directText = markerPrefix + directText;
            } else if (listStyle !== 'none') {
                markerPrefix = '\\u2022 ';
                directText = markerPrefix + directText;
            }
            try {
                const ms = getComputedStyle(el, '::marker');
                if (ms && ms.color) markerColor = ms.color;
            } catch(e) {}
        }

        const isImg = tag === 'img';
        const isSvg = tag === 'svg';
        const hasVisibleBg = style.backgroundColor !== 'rgba(0, 0, 0, 0)' &&
                             style.backgroundColor !== 'transparent';
        const hasBorder = style.borderWidth && style.borderWidth !== '0px' &&
                         style.borderStyle !== 'none';

        const bgImage = style.backgroundImage !== 'none' ? style.backgroundImage : null;
        const textFillColor = style.webkitTextFillColor || style.WebkitTextFillColor || '';
        const isFillTransparent = textFillColor === 'transparent' || textFillColor === 'rgba(0, 0, 0, 0)';
        const isGradientText = isFillTransparent && bgImage && bgImage.includes('gradient');

        const data = {
            tag: tag,
            x: relX,
            y: relY,
            width: rect.width,
            height: rect.height,
            text: directText,
            color: style.color,
            backgroundColor: hasVisibleBg ? style.backgroundColor : null,
            fontSize: parseFloat(style.fontSize),
            fontFamily: style.fontFamily,
            fontWeight: style.fontWeight,
            fontStyle: style.fontStyle,
            textAlign: style.textAlign,
            lineHeight: style.lineHeight,
            letterSpacing: (style.letterSpacing === 'normal' ? 0 : (parseFloat(style.letterSpacing) || 0)),
            direction: style.direction,
            position: style.position,
            display: style.display,
            justifyContent: style.justifyContent,
            alignItems: style.alignItems,
            writingMode: style.writingMode,
            rotation: ownRotation,
            opacity: parseFloat(style.opacity),
            isGroup: (el.getAttribute('role') || '').split(/\\s+/).includes('group') ||
                     el.hasAttribute('data-pptx-group'),
            groupName: el.getAttribute('aria-label') || el.getAttribute('data-pptx-group') || '',
            boxShadow: style.boxShadow,
            filter: style.filter,
            paddingLeft: parseFloat(style.paddingLeft) || 0,
            paddingRight: parseFloat(style.paddingRight) || 0,
            paddingTop: parseFloat(style.paddingTop) || 0,
            paddingBottom: parseFloat(style.paddingBottom) || 0,
            objectFit: isImg ? style.objectFit : null,
            naturalWidth: isImg ? (el.naturalWidth || 0) : 0,
            naturalHeight: isImg ? (el.naturalHeight || 0) : 0,
            borderRadius: style.borderRadius,
            borderColor: hasBorder ? style.borderColor : null,
            borderWidth: hasBorder ? parseFloat(style.borderWidth) : 0,
            borderStyle: hasBorder ? style.borderStyle : null,
            borderLeftColor: style.borderLeftColor !== style.borderColor ? style.borderLeftColor : null,
            borderLeftWidth: parseFloat(style.borderLeftWidth) || 0,
            borderLeftStyle: style.borderLeftColor !== style.borderColor ? style.borderLeftStyle : null,
            textTransform: style.textTransform,
            backgroundImage: bgImage,
            isGradientText: isGradientText,
            isImage: isImg,
            isSvg: isSvg,
            src: isImg ? el.getAttribute('src') : null,
            href: tag === 'a' ? el.getAttribute('href') : null,
            markerColor: markerColor,
            children: []
        };
        Object.assign(data, textMeasurement);

        // Mark inline SVGs for rasterization in the Python post-pass
        if (isSvg) {
            const svgId = 'pptx-svg-' + (_svgCounter++);
            el.setAttribute('data-pptx-id', svgId);
            data.svgId = svgId;
        }

        for (const child of el.children) {
            if (['script', 'style', 'link', 'meta'].includes(child.tagName.toLowerCase())) continue;
            const childData = measureElement(child, slideRect, depth + 1);
            if (childData) data.children.push(childData);
        }

        // Measure ::before and ::after pseudo-elements as synthetic children
        for (const pseudo of ['::before', '::after']) {
            try {
                const ps = getComputedStyle(el, pseudo);
                if (!ps.content || ps.content === 'none' || ps.content === 'normal') continue;
                if (ps.display === 'none') continue;

                const psBg = ps.backgroundColor !== 'rgba(0, 0, 0, 0)' && ps.backgroundColor !== 'transparent';
                const pw = parseFloat(ps.width) || 0;
                const ph = parseFloat(ps.height) || 0;
                if (pw <= 0 && ph <= 0 && !psBg) continue;

                let px = relX, py = relY;
                const pt = parseFloat(ps.top); const pl = parseFloat(ps.left);
                const pr = parseFloat(ps.right); const pb = parseFloat(ps.bottom);
                if (ps.position === 'absolute') {
                    if (!isNaN(pl)) px = relX + pl;
                    else if (!isNaN(pr)) px = relX + rect.width - pw - pr;
                    if (!isNaN(pt)) py = relY + pt;
                    else if (!isNaN(pb)) py = relY + rect.height - ph - pb;
                }

                if (pw > 0 && ph > 0) {
                    data.children.push({
                        tag: '_pseudo',
                        x: px, y: py, width: pw, height: ph,
                        text: '',
                        color: ps.color,
                        backgroundColor: psBg ? ps.backgroundColor : null,
                        backgroundImage: ps.backgroundImage !== 'none' ? ps.backgroundImage : null,
                        fontSize: 0, fontFamily: '', fontWeight: '400', fontStyle: 'normal',
                        textAlign: 'left', lineHeight: 'normal', direction: 'ltr',
                        opacity: parseFloat(ps.opacity), borderRadius: ps.borderRadius,
                        boxShadow: ps.boxShadow, filter: ps.filter,
                        borderColor: null, borderWidth: 0, borderStyle: null,
                        borderLeftColor: null, borderLeftWidth: 0, borderLeftStyle: null,
                        textTransform: 'none', isImage: false, src: null, href: null,
                        markerColor: null, children: [], isGradientText: false,
                    });
                }
            } catch(e) {}
        }

        const inlineOnly = el.children.length > 0 && Array.from(el.children).every(child => {
            const childStyle = getComputedStyle(child);
            return !['absolute', 'fixed'].includes(childStyle.position) &&
                   (child.tagName.toLowerCase() === 'br' || childStyle.display.startsWith('inline'));
        });
        if ((directText && hasChildElementText(el)) ||
            (inlineOnly && (el.textContent || '').length > 0)) {
            const runs = collectInlineRuns(el);
            if (runs.length > 0) {
                if (markerPrefix && runs.length > 0) {
                    runs[0].text = markerPrefix + runs[0].text;
                }
                data.inlineRuns = runs;
                data.text = '';
            }
        }

        const isContainer = !data.text && !isImg && !isSvg && !hasVisibleBg && !hasBorder &&
                           data.backgroundImage === null && !data.inlineRuns &&
                           (!data.boxShadow || data.boxShadow === 'none');
        if (isContainer && !data.isGroup && data.children.length === 1 && depth > 0) {
            const child = data.children[0];
            // A pure wrapper is collapsed away, but its visual effects must be
            // folded into the surviving child, or they are silently lost:
            //  - opacity (e.g. a faint hero-image wrapper at opacity:0.18)
            //  - border-radius + overflow:hidden clipping (rounded image frames)
            if (data.opacity < 1) {
                child.opacity = (child.opacity == null ? 1 : child.opacity) * data.opacity;
            }
            const childHasRadius = child.borderRadius && child.borderRadius !== '0px'
                && child.borderRadius !== '0';
            if (!childHasRadius && data.borderRadius && data.borderRadius !== '0px'
                && data.borderRadius !== '0') {
                child.borderRadius = data.borderRadius;
            }
            return child;
        }

        return data;
    }

    const slideState = captureSlideState(slides);
    try {
    slides.forEach((slide, slideIndex) => {
        slideState.show(slideIndex);

        const savedTransforms = clearAncestorTransforms(slide);
        slide.offsetHeight;

        const slideRect = slide.getBoundingClientRect();
        const slideStyle = getComputedStyle(slide);
        const slideData = {
            index: slideIndex,
            width: slideRect.width,
            height: slideRect.height,
            backgroundColor: slideStyle.backgroundColor,
            backgroundImage: slideStyle.backgroundImage !== 'none' ? slideStyle.backgroundImage : null,
            elements: []
        };

        for (const child of slide.children) {
            if (['script', 'style', 'link', 'meta'].includes(child.tagName.toLowerCase())) continue;
            const measured = measureElement(child, slideRect, 0);
            if (measured) slideData.elements.push(measured);
        }

        restoreTransforms(savedTransforms);
        results.push(slideData);
    });
    } finally {
        slideState.restore();
    }

    return results;
}
"""


# ---------------------------------------------------------------------------
# Color and font parsing
# ---------------------------------------------------------------------------

_RGB_RE = re.compile(
    r"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)"
)


def _css_color_to_rgb(css_color: str) -> tuple[RGBColor, float] | None:
    """Parse a CSS color string to (RGBColor, opacity). Returns None for transparent."""
    if not css_color or css_color == "transparent":
        return None
    match = _RGB_RE.match(css_color)
    if match:
        r, g, b = int(match.group(1)), int(match.group(2)), int(match.group(3))
        a = float(match.group(4)) if match.group(4) else 1.0
        if a <= 0:
            return None
        return RGBColor(r, g, b), a
    if css_color.startswith("#"):
        h = css_color.lstrip("#")
        if len(h) == 3:
            h = "".join(c * 2 for c in h)
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)), 1.0
    return None


# Default backdrop used when a slide has no resolvable background color.
_DEFAULT_BACKDROP = RGBColor(0xFF, 0xFF, 0xFF)




def _resolve_backdrop(slide_data: dict) -> RGBColor:
    """Best-effort opaque backdrop color for a slide (for flattening alpha)."""
    grad = slide_data.get("backgroundImage", "")
    if grad:
        first = _first_gradient_color(grad)
        if first is not None:
            return first
    bg = _css_color_to_rgb(slide_data.get("backgroundColor", ""))
    if bg is not None:
        return bg[0]
    return _DEFAULT_BACKDROP


_CSS_NUMBER = r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?"
_GRADIENT_ANGLE_RE = re.compile(rf"linear-gradient\(\s*({_CSS_NUMBER})deg", re.I)
_GRADIENT_DIR_RE = re.compile(r"linear-gradient\(\s*to\s+([\w\s]+?)\s*,", re.I)

_CSS_DIR_TO_DEG = {
    "top": 0, "right": 90, "bottom": 180, "left": 270,
    "top right": 45, "right top": 45,
    "bottom right": 135, "right bottom": 135,
    "bottom left": 225, "left bottom": 225,
    "top left": 315, "left top": 315,
}


def _parse_css_gradient(
    css_val: str,
) -> list[tuple[RGBColor, float, float]] | None:
    """Resolve a single CSS linear-gradient's color stops.

    Supports computed rgb/rgba colors, transparent, percentage positions (one
    or two per color), and omitted positions. CSS stop fixup precedes clipping
    to OOXML's 0..100% range. Lengths such as px, calc(), color hints, other
    color spaces, and layered backgrounds are not representable here and are
    rejected, rather than silently replacing their positions with equal gaps.
    """
    if not css_val:
        return None
    match = re.fullmatch(r"\s*linear-gradient\((.*)\)\s*", css_val, re.I | re.S)
    if not match:
        return None

    # Split only commas outside color functions. Full-token parsing below also
    # rejects extra background layers and nested unsupported gradient functions.
    parts: list[str] = []
    depth, start = 0, 0
    body = match.group(1)
    for i, char in enumerate(body):
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth < 0:
                return None
        elif char == "," and depth == 0:
            parts.append(body[start:i].strip())
            start = i + 1
    if depth:
        return None
    parts.append(body[start:].strip())

    direction = parts[0].lower()
    if re.fullmatch(rf"{_CSS_NUMBER}deg", direction):
        if not math.isfinite(float(direction[:-3])):
            return None
        parts.pop(0)
    elif direction.startswith("to "):
        if direction[3:].strip() not in _CSS_DIR_TO_DEG:
            return None
        parts.pop(0)
    if len(parts) < 2:
        return None

    colors: list[tuple[RGBColor, float]] = []
    positions: list[float | None] = []
    for part in parts:
        color_match = re.match(r"(rgba?\([^()]*\)|transparent)(.*)\Z", part, re.I)
        if not color_match:
            return None
        color_text, position_text = color_match.groups()
        if color_text.lower() == "transparent":
            color, alpha = RGBColor(0, 0, 0), 0.0
        else:
            channels = re.split(r"[\s,/]+", color_text[color_text.index("(") + 1:-1].strip())
            if len(channels) not in (3, 4):
                return None
            if any(not re.fullmatch(rf"{_CSS_NUMBER}%?", channel) for channel in channels):
                return None
            if any(not math.isfinite(float(channel.rstrip("%"))) for channel in channels):
                return None
            rgb = [
                round(max(0.0, min(255.0, float(channel.rstrip("%"))
                    * (255 / 100 if channel.endswith("%") else 1))))
                for channel in channels[:3]
            ]
            color = RGBColor(*rgb)
            alpha = (
                max(0.0, min(1.0, float(channels[3].rstrip("%"))
                    / (100 if channels[3].endswith("%") else 1)))
                if len(channels) == 4 else 1.0
            )
        tokens = position_text.split()
        if len(tokens) > 2:
            return None
        if not tokens:
            colors.append((color, alpha))
            positions.append(None)
        for token in tokens:
            # Unitless zero is a CSS length of zero; nonzero lengths require
            # element geometry and must not be mistaken for percentages.
            if re.fullmatch(rf"{_CSS_NUMBER}%", token):
                position = float(token[:-1]) / 100
            elif re.fullmatch(_CSS_NUMBER, token) and float(token) == 0:
                position = 0.0
            else:
                return None
            if not math.isfinite(position):
                return None
            colors.append((color, alpha))
            positions.append(position)

    # CSS Images stop fixup: default endpoints, clamp decreasing explicit
    # positions to the greatest preceding one, then interpolate omitted runs.
    if positions[0] is None:
        positions[0] = 0.0
    if positions[-1] is None:
        positions[-1] = 1.0
    previous = positions[0]
    for i, position in enumerate(positions):
        if position is not None:
            previous = max(previous, position)
            positions[i] = previous
    anchor = 0
    for i in range(1, len(positions)):
        if positions[i] is None:
            continue
        left, right = positions[anchor], positions[i]
        for j in range(anchor + 1, i):
            positions[j] = left + (right - left) * (j - anchor) / (i - anchor)
        anchor = i
    stops = [
        (color, position, alpha)
        for (color, alpha), position in zip(colors, positions)
    ]
    return _clip_gradient_stops(stops)


def _clip_gradient_stops(
    stops: list[tuple[RGBColor, float, float]],
) -> list[tuple[RGBColor, float, float]]:
    """Clip out-of-range CSS stops without shifting the visible gradient."""
    if stops[0][1] >= 0 and stops[-1][1] <= 1:
        return stops

    def at_boundary(position: float) -> tuple[RGBColor, float, float]:
        if position <= stops[0][1]:
            return stops[0][0], position, stops[0][2]
        if position >= stops[-1][1]:
            return stops[-1][0], position, stops[-1][2]
        for left, right in zip(stops, stops[1:]):
            if left[1] < position < right[1]:
                fraction = (position - left[1]) / (right[1] - left[1])
                alpha = left[2] * (1 - fraction) + right[2] * fraction
                # CSS interpolation uses premultiplied alpha; transparent
                # black must not darken a neighboring opaque color.
                color = RGBColor(*(
                    round((left[0][channel] * left[2] * (1 - fraction)
                           + right[0][channel] * right[2] * fraction) / alpha)
                    if alpha else 0
                    for channel in range(3)
                ))
                return color, position, alpha
        raise ValueError("Gradient boundary is not between resolved stops")

    clipped = [stop for stop in stops if 0 <= stop[1] <= 1]
    if not clipped or clipped[0][1] > 0:
        clipped.insert(0, at_boundary(0.0))
    if clipped[-1][1] < 1:
        clipped.append(at_boundary(1.0))
    return clipped


def _gradient_angle_emu(css_val: str) -> int:
    """Extract CSS gradient angle and convert to OOXML EMU angle (60000ths of degree).

    CSS: 0deg=to-top, 90deg=to-right, 180deg=to-bottom (clockwise from top).
    OOXML: 0=right-to-left, 5400000=top-to-bottom (clockwise from right).
    Supports both degree angles and direction keywords (to right, to bottom left, etc.).
    """
    m = _GRADIENT_ANGLE_RE.search(css_val or "")
    if m:
        css_deg = float(m.group(1))
    else:
        dm = _GRADIENT_DIR_RE.search(css_val or "")
        css_deg = _CSS_DIR_TO_DEG.get(dm.group(1).strip().lower(), 180) if dm else 180.0
    ooxml_deg = (css_deg + 270) % 360
    return int(ooxml_deg * 60000)


def _apply_gradient_fill(
    xml_ancestor, css_val: str, backdrop: RGBColor | None = None, opacity: float = 1.0,
) -> bool:
    """Apply a native linear-gradient, retaining stop alpha over actual artwork.

    The optional backdrop does not determine native fill colors or alpha.
    """
    stops = _parse_css_gradient(css_val)
    if not stops:
        return False
    angle = _gradient_angle_emu(css_val)

    # Find or create the <a:gradFill> element
    grad_fill = xml_ancestor.find(".//" + qn("a:gradFill"))
    if grad_fill is None:
        # Find the properties container (spPr for shapes, bgPr for backgrounds).
        # Use explicit "is not None" checks: lxml elements raise a FutureWarning
        # on truth-testing, and an empty element is falsy, so `or` is unsafe.
        props = xml_ancestor.find(qn("p:spPr"))
        if props is None:
            props = xml_ancestor.find(qn("p:bgPr"))
        if props is None:
            props = xml_ancestor.find(".//" + qn("a:spPr"))
        if props is None:
            props = xml_ancestor
        # Remove any existing fill
        for tag in ("a:solidFill", "a:noFill", "a:pattFill", "a:blipFill"):
            for old in props.findall(qn(tag)):
                props.remove(old)
        grad_fill = etree.SubElement(props, qn("a:gradFill"))

    grad_fill.set("rotWithShape", "0")

    # Replace gradient stops
    gs_lst = grad_fill.find(qn("a:gsLst"))
    if gs_lst is None:
        gs_lst = etree.SubElement(grad_fill, qn("a:gsLst"))
    for old in gs_lst.findall(qn("a:gs")):
        gs_lst.remove(old)

    for color, pos, alpha in stops:
        gs = etree.SubElement(gs_lst, qn("a:gs"))
        gs.set("pos", str(round(pos * 100000)))
        srgb = etree.SubElement(gs, qn("a:srgbClr"))
        srgb.set("val", f"{color[0]:02X}{color[1]:02X}{color[2]:02X}")
        alpha = max(0.0, min(1.0, alpha * opacity))
        if alpha < 1.0:
            alpha_el = etree.SubElement(srgb, qn("a:alpha"))
            alpha_el.set("val", str(round(alpha * 100000)))

    # Set angle
    lin = grad_fill.find(qn("a:lin"))
    if lin is None:
        lin = etree.SubElement(grad_fill, qn("a:lin"))
    lin.set("ang", str(angle))
    lin.set("scaled", "1")

    return True


def _apply_gradient_text(run, css_val: str, opacity: float = 1.0) -> bool:
    """Apply a CSS gradient as text fill color via OOXML <a:gradFill> on the run.

    PPTX supports gradient text — the gradient goes inside <a:rPr> on the run,
    replacing the solid <a:solidFill>.
    """
    stops = _parse_css_gradient(css_val)
    if not stops:
        return False
    angle = _gradient_angle_emu(css_val)

    rPr = run._r.get_or_add_rPr()

    # Remove existing solid fill
    for old in rPr.findall(qn("a:solidFill")):
        rPr.remove(old)

    grad = etree.SubElement(rPr, qn("a:gradFill"))
    gs_lst = etree.SubElement(grad, qn("a:gsLst"))
    for color, pos, alpha in stops:
        gs = etree.SubElement(gs_lst, qn("a:gs"))
        gs.set("pos", str(round(pos * 100000)))
        srgb = etree.SubElement(gs, qn("a:srgbClr"))
        srgb.set("val", f"{color[0]:02X}{color[1]:02X}{color[2]:02X}")
        effective_alpha = max(0.0, min(1.0, alpha * opacity))
        if effective_alpha < 1.0:
            alpha_el = etree.SubElement(srgb, qn("a:alpha"))
            alpha_el.set("val", str(round(effective_alpha * 100000)))

    lin = etree.SubElement(grad, qn("a:lin"))
    lin.set("ang", str(angle))
    lin.set("scaled", "1")

    return True


def _first_gradient_color(css_val: str) -> RGBColor | None:
    """Extract the first opaque color from a CSS gradient for use as a solid fallback."""
    stops = _parse_css_gradient(css_val)
    if not stops:
        return None
    for color, _, alpha in stops:
        if alpha > 0.3:
            return color
    return stops[0][0]


def _resolve_font_color(el: dict) -> RGBColor:
    """Resolve the RGB component; run-level OOXML preserves its CSS alpha."""
    if el.get("isGradientText"):
        grad_color = _first_gradient_color(el.get("backgroundImage", ""))
        if grad_color:
            return grad_color

    result = _css_color_to_rgb(el.get("color", "rgb(255,255,255)"))
    if not result:
        return RGBColor(0xFF, 0xFF, 0xFF)
    return result[0]


def _parse_font_family(css_font: str) -> str:
    """Extract the first font family name from a CSS font-family string."""
    first = css_font.split(",")[0].strip()
    return first.strip("'\"")


# Fonts that ship with Windows/Office (and are Microsoft-metric-compatible on
# macOS/LibreOffice), so we can rely on them rendering without substitution.
_SAFE_FONTS = {
    "arial", "arial black", "calibri", "cambria", "candara", "consolas",
    "constantia", "corbel", "courier new", "georgia", "times new roman",
    "trebuchet ms", "verdana", "segoe ui", "tahoma", "garamond",
    "book antiqua", "century gothic", "palatino linotype", "gill sans",
    "franklin gothic medium", "lucida sans", "impact",
}

# Explicit web-font -> metric/style-compatible safe equivalent. Grouped so the
# substitute stays in the SAME visual family (display-serif, humanist-sans,
# geometric-sans, monospace); keeping the family keeps glyph advance widths
# close, which stops headings from reflowing onto an extra line.
_FONT_EQUIVALENTS = {
    # ---- display / body serifs ----
    "playfair display": "Georgia",
    "playfair": "Georgia",
    "merriweather": "Georgia",
    "lora": "Georgia",
    "pt serif": "Georgia",
    "noto serif": "Georgia",
    "source serif pro": "Cambria",
    "source serif 4": "Cambria",
    "roboto slab": "Cambria",
    "dm serif display": "Georgia",
    "dm serif text": "Georgia",
    "cormorant": "Cambria",
    "cormorant garamond": "Cambria",
    "eb garamond": "Garamond",
    "crimson text": "Garamond",
    "crimson pro": "Garamond",
    "libre baskerville": "Georgia",
    "bitter": "Georgia",
    "spectral": "Cambria",
    "frank ruhl libre": "Georgia",
    # ---- humanist / grotesque sans ----
    "inter": "Segoe UI",
    "roboto": "Arial",
    "open sans": "Segoe UI",
    "lato": "Calibri",
    "noto sans": "Segoe UI",
    "source sans pro": "Segoe UI",
    "source sans 3": "Segoe UI",
    "work sans": "Segoe UI",
    "dm sans": "Segoe UI",
    "manrope": "Segoe UI",
    "ibm plex sans": "Segoe UI",
    "pt sans": "Segoe UI",
    "rubik": "Segoe UI",
    "karla": "Segoe UI",
    "mulish": "Segoe UI",
    "barlow": "Segoe UI",
    "titillium web": "Segoe UI",
    "figtree": "Segoe UI",
    "plus jakarta sans": "Segoe UI",
    "ubuntu": "Segoe UI",
    "helvetica": "Arial",
    "helvetica neue": "Arial",
    "nunito": "Calibri",
    "nunito sans": "Calibri",
    # ---- geometric sans ----
    "montserrat": "Century Gothic",
    "poppins": "Century Gothic",
    "raleway": "Century Gothic",
    "quicksand": "Century Gothic",
    "josefin sans": "Century Gothic",
    "comfortaa": "Century Gothic",
    # ---- monospace ----
    "jetbrains mono": "Consolas",
    "fira code": "Consolas",
    "fira mono": "Consolas",
    "source code pro": "Consolas",
    "roboto mono": "Consolas",
    "ibm plex mono": "Consolas",
    "space mono": "Consolas",
    "ubuntu mono": "Consolas",
    "inconsolata": "Consolas",
    "menlo": "Consolas",
    "monaco": "Consolas",
    "courier": "Courier New",
}

# CSS generic keyword -> concrete safe default (same family class).
_GENERIC_FALLBACK = {
    "serif": "Georgia",
    "sans-serif": "Calibri",
    "monospace": "Consolas",
    "cursive": "Segoe Script",
    "system-ui": "Segoe UI",
    "-apple-system": "Segoe UI",
    "blinkmacsystemfont": "Segoe UI",
    "ui-sans-serif": "Segoe UI",
    "ui-serif": "Georgia",
    "ui-monospace": "Consolas",
}


def _resolve_pptx_font(css_font: str) -> str:
    """Pick a rendering-safe font that stays in the source's visual family.

    Walks the CSS font-family stack in declared order and returns the first of:
      1. a family already known to be installed everywhere, else
      2. a known web font mapped to a metric-compatible safe equivalent, else
      3. the CSS generic keyword (serif/sans-serif/monospace) default.
    Falls back to Calibri. Staying in the same family keeps advance widths close
    so a substituted heading does not wrap onto an unwanted extra line.
    """
    if not css_font:
        return "Calibri"
    generic_seen: str | None = None
    for raw in css_font.split(","):
        name = raw.strip().strip("'\"")
        if not name:
            continue
        low = name.lower()
        if low in _SAFE_FONTS:
            return name
        if low in _FONT_EQUIVALENTS:
            return _FONT_EQUIVALENTS[low]
        if low in _GENERIC_FALLBACK and generic_seen is None:
            generic_seen = _GENERIC_FALLBACK[low]
    return generic_seen or "Calibri"


# ---------------------------------------------------------------------------
# Image handling
# ---------------------------------------------------------------------------


def _apply_image_opacity(pic, opacity: float) -> None:
    """Make a picture translucent via ``<a:alphaModFix>`` in its blipFill.

    Used for faint background/hero images (e.g. an image wrapper at
    ``opacity: 0.18``) so overlaid text stays readable, matching the source.
    """
    if opacity >= 1.0:
        return
    blip = pic._element.find(".//" + qn("a:blip"))
    if blip is None:
        return
    for old in blip.findall(qn("a:alphaModFix")):
        blip.remove(old)
    amod = etree.SubElement(blip, qn("a:alphaModFix"))
    amod.set("amt", str(int(max(0.0, min(1.0, opacity)) * 100000)))


def _apply_object_fit_cover(pic, box_w_px: float, box_h_px: float,
                            nat_w: float, nat_h: float) -> None:
    """Emulate CSS ``object-fit: cover`` by cropping (never stretching).

    ``add_picture`` with explicit width+height stretches the image to the box,
    which distorts any image whose aspect ratio differs from the box (e.g. a
    square hero/product image placed in a wide frame). CSS ``cover`` instead
    scales to fill and crops the overflow, so we replicate that with a centered
    crop on the longer axis — the picture keeps the box geometry but is no longer
    distorted.
    """
    if nat_w <= 0 or nat_h <= 0 or box_w_px <= 0 or box_h_px <= 0:
        return
    box_ar = box_w_px / box_h_px
    img_ar = nat_w / nat_h
    if abs(img_ar - box_ar) < 1e-3:
        return
    if img_ar > box_ar:  # image too wide -> crop left/right
        crop = (1 - box_ar / img_ar) / 2
        pic.crop_left = crop
        pic.crop_right = crop
    else:  # image too tall -> crop top/bottom
        crop = (1 - img_ar / box_ar) / 2
        pic.crop_top = crop
        pic.crop_bottom = crop


def _add_image_from_data_uri(slide, data_uri: str, left, top, width, height,
                             border_radius_px: float = 0,
                             width_px: float = 0, height_px: float = 0,
                             opacity: float = 1.0,
                             object_fit: str | None = None,
                             natural_w: float = 0, natural_h: float = 0):
    """Decode a base64 data URI and add it as a picture shape.

    When border_radius_px > 0, clips the image to a rounded rectangle
    by swapping the shape geometry from 'rect' to 'roundRect'.
    When opacity < 1, the picture is made translucent to match the source.
    When object_fit == 'cover', the picture is cropped (not stretched) to fill.
    """
    match = re.match(r"data:image/(\w+);base64,(.*)", data_uri, re.DOTALL)
    if not match:
        return None
    ext = match.group(1)
    try:
        img_bytes = base64.b64decode(match.group(2))
    except Exception:
        logger.warning("Failed to decode base64 image data")
        return None

    tmp = tempfile.NamedTemporaryFile(suffix=f".{ext}", delete=False)
    try:
        tmp.write(img_bytes)
        tmp.close()
        pic = slide.shapes.add_picture(tmp.name, left, top, width, height)

        if object_fit == "cover":
            _apply_object_fit_cover(pic, width_px, height_px, natural_w, natural_h)

        if border_radius_px > 0 and width_px > 0 and height_px > 0:
            sp_pr = pic._element.find(qn("p:spPr"))
            if sp_pr is not None:
                prst_geom = sp_pr.find(qn("a:prstGeom"))
                if prst_geom is not None:
                    prst_geom.set("prst", "roundRect")
                    _set_corner_radius(pic, border_radius_px, width_px, height_px)

        _apply_image_opacity(pic, opacity)

        return pic
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Element tree helpers
# ---------------------------------------------------------------------------


def _any_descendant_has_text(el: dict) -> bool:
    """True if any child/grandchild has non-empty direct text."""
    for child in el.get("children", []):
        if child.get("text", "").strip():
            return True
        if _any_descendant_has_text(child):
            return True
    return False


def _count_elements(elements: list[dict]) -> int:
    """Count total elements in a measurement tree (including nested children)."""
    total = len(elements)
    for el in elements:
        total += _count_elements(el.get("children", []))
    return total


# ---------------------------------------------------------------------------
# PPTX rendering helpers
# ---------------------------------------------------------------------------


def _apply_font(
    run,
    *,
    size_pt: float,
    color: RGBColor,
    bold: bool,
    italic: bool,
    family: str,
    letter_spacing_px: float = 0.0,
    alpha: float = 1.0,
) -> None:
    """Apply font properties to a text run."""
    run.font.size = Pt(size_pt)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = family
    r_pr = run._r.get_or_add_rPr()
    # DrawingML tracking is hundredths of a point, not CSS pixels.
    r_pr.set("spc", str(round(letter_spacing_px * CSS_PX_TO_PT * 100)))
    r_pr.set("kern", "0")
    if alpha < 1.0:
        color_node = r_pr.find(qn("a:solidFill") + "/" + qn("a:srgbClr"))
        if color_node is not None:
            alpha_el = etree.SubElement(color_node, qn("a:alpha"))
            alpha_el.set("val", str(round(max(0.0, alpha) * 100000)))


def _parse_border_radius_px(el: dict) -> float:
    """Extract border-radius in pixels from an element measurement.

    Handles both pixel values ('16px') and percentages ('50%').
    Percentages are resolved against the element's smaller dimension.
    """
    br = str(el.get("borderRadius", "0")).strip().split()[0]
    if not br or br == "0":
        return 0
    try:
        if "%" in br:
            pct = float(br.replace("%", "")) / 100
            min_dim = min(el.get("width", 0), el.get("height", 0))
            return pct * min_dim
        return float(br.replace("px", ""))
    except (ValueError, IndexError):
        return 0


def _set_corner_radius(
    shape, radius_px: float, width_px: float, height_px: float,
) -> None:
    """Set rounded-rectangle corner radius via the OOXML 'adj' guide.

    The guide value is a ratio of the minimum dimension:
    50000 = fully rounded (pill shape), 0 = sharp corners.
    """
    min_dim = min(width_px, height_px)
    if min_dim <= 0:
        return
    ratio = min(radius_px / min_dim, 0.5)
    adj_val = int(ratio * 100000)

    prstGeom = shape._element.find(".//" + qn("a:prstGeom"))
    if prstGeom is None:
        return
    avLst = prstGeom.find(qn("a:avLst"))
    if avLst is None:
        avLst = etree.SubElement(prstGeom, qn("a:avLst"))
    for old in avLst.findall(qn("a:gd")):
        avLst.remove(old)
    gd = etree.SubElement(avLst, qn("a:gd"))
    gd.set("name", "adj")
    gd.set("fmla", f"val {adj_val}")


def _apply_corner_geometry(shape, el: dict, radius_px: float) -> None:
    """Keep equal rounded top corners and square bottom corners natively."""
    tokens = str(el.get("borderRadius") or "0").split()
    if 1 <= len(tokens) <= 4 and all(re.fullmatch(r"[\d.]+(?:px)?", token) for token in tokens):
        values = [float(token.removesuffix("px")) for token in tokens]
        if len(values) == 1:
            values *= 4
        elif len(values) == 2:
            values *= 2
        elif len(values) == 3:
            values.append(values[1])
        if values[0] == values[1] and values[0] > 0 and values[2] == values[3] == 0:
            geom = shape._element.find(".//" + qn("a:prstGeom"))
            if geom is not None:
                geom.set("prst", "round2SameRect")
                guides = geom.find(qn("a:avLst"))
                if guides is None:
                    guides = etree.SubElement(geom, qn("a:avLst"))
                for guide in list(guides):
                    guides.remove(guide)
                minimum = min(el.get("width", 0), el.get("height", 0))
                adjustment = round(min(values[0] / minimum, 0.5) * 100000) if minimum > 0 else 0
                for name, value in (("adj1", adjustment), ("adj2", 0)):
                    guide = etree.SubElement(guides, qn("a:gd"))
                    guide.set("name", name)
                    guide.set("fmla", f"val {value}")
                return
    _set_corner_radius(shape, radius_px, el.get("width", 0), el.get("height", 0))


def _resolve_alignment(
    el: dict, is_single_line: bool, has_visual_bg: bool,
):
    """Use explicit CSS alignment; a short box is not evidence of centering."""
    is_rtl = el.get("direction") == "rtl"
    text_align = el.get("textAlign", "right" if is_rtl else "left")
    h = {
        "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT, "left": PP_ALIGN.LEFT,
        "justify": PP_ALIGN.JUSTIFY,
        "start": PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT,
        "end": PP_ALIGN.LEFT if is_rtl else PP_ALIGN.RIGHT,
    }.get(text_align, PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT)
    if el.get("textGeometry"):
        return h, MSO_ANCHOR.TOP

    display = el.get("display", "") or ""
    is_flex = "flex" in display or "grid" in display
    v = MSO_ANCHOR.TOP

    if is_flex:
        jc = el.get("justifyContent", "") or ""
        if "center" in jc or "space" in jc:
            h = PP_ALIGN.CENTER
        elif jc in ("flex-end", "end", "right"):
            h = PP_ALIGN.RIGHT
        elif jc in ("flex-start", "start", "left"):
            h = PP_ALIGN.LEFT
        ai = el.get("alignItems", "") or ""
        if "center" in ai:
            v = MSO_ANCHOR.MIDDLE
        elif ai in ("flex-end", "end"):
            v = MSO_ANCHOR.BOTTOM
    return h, v


def _line_height_ratio(el: dict) -> float | None:
    """CSS line-height as a unitless multiple of the font size, or None.

    PowerPoint's default line spacing (~1.2) is looser than tight display
    line-heights (e.g. ``line-height: 0.92`` on a big headline), which makes a
    multi-line heading grow taller than its box and overlap the element below.
    Reproducing the CSS ratio keeps line count and vertical extent faithful.
    """
    lh = el.get("lineHeight", "normal")
    fs = el.get("fontSize", 0)
    if not fs or not lh or lh == "normal":
        return None
    try:
        px = float(str(lh).replace("px", "").strip())
    except ValueError:
        return None
    ratio = px / fs
    if 0.5 <= ratio <= 3.0:
        return ratio
    return None


def _apply_line_spacing(paragraph, el: dict) -> None:
    """Set exact line spacing (in points) to match the CSS line-height.

    A float ``line_spacing`` in PPTX multiplies the font's *natural* line height
    (~1.2×), not the font size, so passing the CSS ratio (e.g. 0.92) still comes
    out ~1.1× too tall and a multi-line heading creeps into the element below.
    Converting the measured px line-height to an absolute point value reproduces
    the CSS box exactly.
    """
    paragraph.space_before = Pt(0)
    paragraph.space_after = Pt(0)
    geometry = el.get("textGeometry") or {}
    lh = geometry.get("lineHeight") or el.get("lineHeight", "normal")
    if not lh or lh == "normal":
        return
    try:
        px = float(str(lh).replace("px", "").strip())
    except ValueError:
        return
    if px > 0:
        paragraph.line_spacing = Pt(px * PIXELS_TO_INCHES_Y * 72.0)



def _apply_text_padding(tf, el: dict, extra_left_px: float = 0.0) -> None:
    """Match CSS padding by insetting the text inside its box.

    python-pptx text frames default to ~0.1in/0.05in internal margins; the
    element box we place is the CSS border box, so without this the text hugs the
    box edge (e.g. bulleted text overprints its ``::before`` dot, card labels
    touch the card edge). Setting the frame margins to the measured padding keeps
    the box geometry identical while placing the glyphs where the browser did.
    """
    left_px = max(el.get("paddingLeft", 0), extra_left_px)
    tf.margin_left = Inches(left_px * PIXELS_TO_INCHES_X)
    tf.margin_right = Inches(el.get("paddingRight", 0) * PIXELS_TO_INCHES_X)
    tf.margin_top = Inches(el.get("paddingTop", 0) * PIXELS_TO_INCHES_Y)
    tf.margin_bottom = Inches(el.get("paddingBottom", 0) * PIXELS_TO_INCHES_Y)


def _text_box_layout(
    el: dict, x_in: float, y_in: float, w_in: float, h_in: float,
) -> tuple[float, float, float, float, dict]:
    """Place text by measured font bounds, independently of its CSS border box."""
    geometry = el.get("textGeometry")
    if not geometry:
        return x_in, y_in, w_in, h_in, el
    original_center = (x_in + w_in / 2, y_in + h_in / 2)
    text_el = dict(el)
    # CSS line-height can be smaller than a font's ascent+descent. A border-box
    # top is therefore not a text top: retain the measured baseline instead.
    text_top = geometry["baseline"] - geometry["ascent"]
    y_in += text_top * PIXELS_TO_INCHES_Y
    natural_height = geometry["ascent"] + geometry["descent"]
    line_count = max(1, geometry.get("lineCount", 1))
    height_px = max(geometry["height"], natural_height + (line_count - 1) * geometry["lineHeight"])
    h_in = height_px * PIXELS_TO_INCHES_Y
    text_el["paddingTop"] = 0
    text_el["paddingBottom"] = 0
    if line_count == 1:
        # A no-wrap frame must also accommodate the font's untracked advance:
        # some importers lay out before applying negative DrawingML tracking.
        width_px = max(geometry["width"], geometry.get("unspacedWidth", 0))
        x_in += geometry["x"] * PIXELS_TO_INCHES_X
        w_in = max(width_px * PIXELS_TO_INCHES_X, 1 / 914400)
        text_el["paddingLeft"] = 0
        text_el["paddingRight"] = 0
        text_el["textAlign"] = "left"
        text_el["display"] = "block"
    rotation = el.get("rotation", 0) or 0
    if rotation:
        angle = math.radians(rotation)
        dx = x_in + w_in / 2 - original_center[0]
        dy = y_in + h_in / 2 - original_center[1]
        x_in = original_center[0] + dx * math.cos(angle) - dy * math.sin(angle) - w_in / 2
        y_in = original_center[1] + dx * math.sin(angle) + dy * math.cos(angle) - h_in / 2
    return x_in, y_in, w_in, h_in, text_el


def _text_is_single_line(el: dict, text: str) -> bool:
    geometry = el.get("textGeometry")
    if geometry:
        return geometry.get("lineCount", 1) == 1 and "\n" not in text
    if "\n" in text:
        return False
    line_px = (_line_height_ratio(el) or 1.3) * el.get("fontSize", 20)
    content_px = el.get("height", 0) - el.get("paddingTop", 0) - el.get("paddingBottom", 0)
    return content_px <= line_px * 1.6


def _text_alpha(el: dict, opacity: float) -> float:
    color = _css_color_to_rgb(el.get("color", "rgb(255,255,255)"))
    return opacity * (color[1] if color else 1.0)


def _apply_fill_alpha(shape_or_txbox, alpha: float) -> None:
    """Set fill opacity via OOXML alpha child element.

    OOXML represents alpha as a child of srgbClr, not an attribute:
      <a:srgbClr val="0F172A"><a:alpha val="70000"/></a:srgbClr>
    Value is in 1/100000ths (70000 = 70% opacity).
    """
    props = shape_or_txbox._element.find(qn("p:spPr"))
    if props is None:
        return
    srgb = props.find(qn("a:solidFill") + "/" + qn("a:srgbClr"))
    if srgb is not None:
        _set_color_alpha(srgb, alpha)


def _set_color_alpha(color, alpha: float) -> None:
    for old in color.findall(qn("a:alpha")):
        color.remove(old)
    if alpha < 1.0:
        etree.SubElement(color, qn("a:alpha")).set(
            "val", str(round(max(0.0, min(1.0, alpha)) * 100000)),
        )


def _apply_shape_effects(shape, el: dict) -> None:
    """Disable theme effects; translate a single authored outer CSS shadow."""
    style = shape._element.find(qn("p:style"))
    if style is not None:
        ref = style.find(qn("a:effectRef"))
        if ref is not None:
            ref.set("idx", "0")
    props = shape._element.find(qn("p:spPr"))
    if props is None:
        props = shape._element.find(qn("p:grpSpPr"))
    if props is None:
        return
    for tag in ("a:effectLst", "a:effectDag"):
        for old in props.findall(qn(tag)):
            props.remove(old)
    effects = etree.SubElement(props, qn("a:effectLst"))
    shadow = el.get("boxShadow") or "none"
    if shadow != "none":
        color_match = _RGB_RE.search(shadow)
        lengths = re.findall(r"(-?[\d.]+)px", _RGB_RE.sub("", shadow))
        if color_match and len(lengths) >= 2 and "inset" not in shadow and (
            shadow.count("rgb") == 1 and (len(lengths) < 4 or float(lengths[3]) == 0)
        ):
            color, alpha = _css_color_to_rgb(color_match.group(0)) or (RGBColor(0, 0, 0), 0)
            dx, dy = map(float, lengths[:2])
            blur = float(lengths[2]) if len(lengths) > 2 else 0
            outer = etree.SubElement(effects, qn("a:outerShdw"))
            outer.set("blurRad", str(round(max(0, blur) * PIXELS_TO_INCHES_X * 914400)))
            outer.set("dist", str(round(math.hypot(dx, dy) * PIXELS_TO_INCHES_X * 914400)))
            outer.set("dir", str(round((math.degrees(math.atan2(dy, dx)) % 360) * 60000)))
            outer.set("algn", "ctr")
            outer.set("rotWithShape", "0")
            rgb = etree.SubElement(outer, qn("a:srgbClr"))
            rgb.set("val", str(color))
            _set_color_alpha(rgb, alpha * el.get("opacity", 1.0))
        else:
            logger.warning("Unsupported CSS box-shadow: %s", shadow)
    css_filter = el.get("filter")
    if css_filter and css_filter != "none":
        logger.warning("Unsupported CSS filter: %s", css_filter)


def _apply_rotation(shape, el: dict) -> None:
    """Rotate a shape/picture to match a CSS transform:rotate (degrees, CW)."""
    rot = el.get("rotation", 0) or 0
    if rot:
        shape.rotation = float(rot)
    _apply_shape_effects(shape, el)


def _apply_vertical_text(tf, el: dict) -> None:
    """Map CSS writing-mode:vertical-* to a vertical PPTX text body."""
    wm = el.get("writingMode", "") or ""
    if not wm.startswith("vertical"):
        return
    try:
        bodyPr = tf._txBody.find(qn("a:bodyPr"))
        if bodyPr is not None:
            # vert270 = bottom-to-top (matches vertical-rl side labels / rotate180)
            bodyPr.set("vert", "vert270")
    except Exception:
        pass


def _conic_angle(tok: str) -> float | None:
    tok = tok.strip()
    try:
        if tok.endswith("%"):
            return float(tok[:-1]) * 3.6
        if tok.endswith("deg"):
            return float(tok[:-3])
        if tok.endswith("turn"):
            return float(tok[:-4]) * 360.0
        return float(tok)
    except ValueError:
        return None


def _parse_conic_gradient(css: str):
    """Parse a conic-gradient into [(RGBColor, start_deg, end_deg), ...]."""
    if not css or "conic-gradient" not in css:
        return None
    i = css.find("conic-gradient(") + len("conic-gradient(")
    depth, end = 1, len(css)
    for j in range(i, len(css)):
        if css[j] == "(":
            depth += 1
        elif css[j] == ")":
            depth -= 1
            if depth == 0:
                end = j
                break
    body = css[i:end]
    parts, buf, d = [], "", 0
    for ch in body:
        if ch == "(":
            d += 1
        elif ch == ")":
            d -= 1
        if ch == "," and d == 0:
            parts.append(buf); buf = ""
        else:
            buf += ch
    if buf.strip():
        parts.append(buf)

    segs: list[list] = []
    cursor = 0.0
    for part in parts:
        m = re.match(r"\s*(rgba?\([^)]*\)|#[0-9a-fA-F]{3,6})", part)
        if not m:
            continue
        cres = _css_color_to_rgb(m.group(1))
        if not cres:
            continue
        nums = [t for t in part[m.end():].split() if t]
        start = _conic_angle(nums[0]) if len(nums) >= 1 else cursor
        end_a = _conic_angle(nums[1]) if len(nums) >= 2 else None
        if start is None:
            start = cursor
        segs.append([cres[0], start, end_a, cres[1]])
        cursor = end_a if end_a is not None else start
    for k in range(len(segs)):
        if segs[k][2] is None or segs[k][2] <= segs[k][1]:
            segs[k][2] = segs[k + 1][1] if k + 1 < len(segs) else 360.0
    return segs or None


def _render_conic_gradient(
    slide, el: dict, x_in, y_in, w_in, h_in, opacity: float = 1.0,
) -> bool:
    """Rasterize a conic-gradient; retain its stop and inherited opacity."""
    segs = _parse_conic_gradient(el.get("backgroundImage", ""))
    if not segs:
        return False
    try:
        from io import BytesIO

        from PIL import Image, ImageChops, ImageDraw
    except Exception:
        return False
    S = 700
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    for color, a0, a1, alpha in segs:
        # PIL angles start at 3 o'clock; CSS conic starts at 12 o'clock → -90.
        d.pieslice([0, 0, S - 1, S - 1], a0 - 90, a1 - 90,
                   fill=(int(color[0]), int(color[1]), int(color[2]), round(alpha * 255)))
    radius_px = _parse_border_radius_px(el)
    if radius_px > 0 and radius_px >= 0.5 * min(el.get("width", 0), el.get("height", 0)):
        mask = Image.new("L", (S, S), 0)
        ImageDraw.Draw(mask).ellipse([0, 0, S - 1, S - 1], fill=255)
        img.putalpha(ImageChops.multiply(img.getchannel("A"), mask))
    buf = BytesIO()
    img.save(buf, "PNG")
    uri = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
    pic = _add_image_from_data_uri(
        slide, uri, Inches(x_in), Inches(y_in), Inches(w_in), Inches(h_in),
        opacity=opacity,
    )
    if pic is not None:
        _apply_rotation(pic, el)
    return True


def _find_top_accent(el: dict, radius_px: float) -> dict | None:
    """A thin, full-width, text-free colored child flush with the card's top edge
    (e.g. a ::before accent strip) — rendered specially so it inherits the card's
    rounded top corners instead of poking square corners past them."""
    ew, eh = el.get("width", 0), el.get("height", 0)
    if ew <= 0 or eh <= 0:
        return None
    ex, ey = el.get("x", 0), el.get("y", 0)
    for ch in el.get("children", []):
        if ch.get("_skip") or ch.get("children"):
            continue
        if ch.get("text", "").strip() or _any_descendant_has_text(ch):
            continue
        has_fill = (
            _css_color_to_rgb(ch.get("backgroundColor", "")) is not None
            or bool(_parse_css_gradient(ch.get("backgroundImage", "")))
        )
        if not has_fill:
            continue
        chh, cw = ch.get("height", 0), ch.get("width", 0)
        rel_x, rel_y = ch.get("x", 0) - ex, ch.get("y", 0) - ey
        if abs(rel_y) > 3 or chh > eh * 0.25 or chh > radius_px * 2 + 6:
            continue
        if cw < ew * 0.85 or rel_x > ew * 0.1:
            continue
        return ch
    return None


def _fill_accent(
    shape, el: dict, backdrop: RGBColor | None, inherited_opacity: float = 1.0,
) -> None:
    """Fill an accent with native color/gradient alpha."""
    opacity = el.get("opacity", 1.0) * inherited_opacity
    if _parse_css_gradient(el.get("backgroundImage", "")):
        _apply_gradient_fill(shape._element, el["backgroundImage"], backdrop, opacity)
        return
    res = _css_color_to_rgb(el.get("backgroundColor", ""))
    if res:
        col, alpha = res
        shape.fill.solid()
        shape.fill.fore_color.rgb = col
        _apply_fill_alpha(shape, alpha * opacity)
    else:
        shape.fill.background()


def _render_bg_shape(
    slide,
    el: dict,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    opacity: float,
    backdrop: RGBColor | None = None,
) -> None:
    """Render a background/border rectangle (optionally rounded)."""
    bg_result = _css_color_to_rgb(el.get("backgroundColor", ""))
    has_fill = bg_result is not None
    has_gradient = bool(_parse_css_gradient(el.get("backgroundImage", "")))

    border_color_str = el.get("borderColor")
    border_width = el.get("borderWidth", 0)
    has_border = bool(border_color_str) and border_width > 0

    # A left-border accent (border-left: Npx solid <color>) is a common
    # decorative bar. On a rounded card a plain rectangle bar would poke square
    # corners past the rounded edge, so we reproduce it as a same-radius rounded
    # shape behind the card, with the card inset to the right so only the left
    # rounded sliver shows (the only way to get a rounded-left accent in PPTX).
    left_border_color = el.get("borderLeftColor")
    left_border_width = el.get("borderLeftWidth", 0)
    left_border_style = el.get("borderLeftStyle")
    left_result = _css_color_to_rgb(left_border_color) if left_border_color else None
    has_left_accent = bool(
        left_result
        and left_border_width > 0
        and left_border_style not in (None, "none")
    )

    if not has_fill and not has_border and not has_gradient and not has_left_accent and (
        not el.get("boxShadow") or el["boxShadow"] == "none"
    ):
        return

    radius_px = _parse_border_radius_px(el)
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius_px > 0 else MSO_SHAPE.RECTANGLE

    card_x, card_w = x_in, w_in
    card_y, card_h = y_in, h_in
    if has_left_accent:
        accent_color = left_result[0]
        bar_w = max(left_border_width * PIXELS_TO_INCHES_X, 0.03)
        if radius_px > 0:
            # Rounded accent: full-size rounded rect underneath, card inset right.
            accent = slide.shapes.add_shape(
                shape_type, Inches(x_in), Inches(y_in), Inches(w_in), Inches(h_in),
            )
            _set_corner_radius(
                accent, radius_px, el.get("width", 100), el.get("height", 100),
            )
            accent.fill.solid()
            accent.fill.fore_color.rgb = accent_color
            _apply_fill_alpha(accent, left_result[1] * opacity)
            accent.line.fill.background()
            _apply_rotation(accent, dict(el, boxShadow="none"))
            card_x = x_in + bar_w
            card_w = max(w_in - bar_w, 0.05)

    # A top accent bar (e.g. a ::before strip or a thin full-width child at the
    # card's top edge) drawn as a plain rectangle would poke square corners past
    # the card's rounded top. Reproduce it the same way as the left accent: a
    # same-radius rounded rect the full card size UNDER the card, with the card
    # pushed down by the bar height so only the rounded top sliver shows.
    if radius_px > 0:
        top_accent = _find_top_accent(el, radius_px)
        if top_accent is not None:
            bar_h = max(top_accent.get("height", 0) * PIXELS_TO_INCHES_Y, 0.03)
            acc = slide.shapes.add_shape(
                shape_type, Inches(card_x), Inches(y_in), Inches(card_w), Inches(h_in),
            )
            _set_corner_radius(acc, radius_px, el.get("width", 100), el.get("height", 100))
            _fill_accent(acc, top_accent, backdrop, opacity)
            acc.line.fill.background()
            _apply_rotation(acc, dict(el, boxShadow="none"))
            card_y = y_in + bar_h
            card_h = max(h_in - bar_h, 0.05)
            top_accent["_skip"] = True

    shape = slide.shapes.add_shape(
        shape_type,
        Inches(card_x), Inches(card_y),
        Inches(card_w), Inches(card_h),
    )

    if radius_px > 0:
        _apply_corner_geometry(shape, el, radius_px)

    if has_gradient:
        _apply_gradient_fill(shape._element, el["backgroundImage"], backdrop, opacity)
    elif has_fill:
        bg_color, bg_alpha = bg_result
        effective_alpha = bg_alpha * opacity
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
        _apply_fill_alpha(shape, effective_alpha)
    else:
        shape.fill.background()

    if has_border:
        border_result = _css_color_to_rgb(border_color_str)
        if border_result:
            shape.line.color.rgb = border_result[0]
            shape.line.width = Pt(border_width * CSS_PX_TO_PT)
            rgb = shape._element.find(
                qn("p:spPr") + "/" + qn("a:ln") + "/" + qn("a:solidFill") + "/" + qn("a:srgbClr")
            )
            if rgb is not None:
                _set_color_alpha(rgb, border_result[1] * opacity)
        else:
            shape.line.fill.background()
    else:
        shape.line.fill.background()

    _apply_rotation(shape, el)

    # Square left accent bar (non-rounded cards): a thin filled rectangle on top.
    if has_left_accent and radius_px <= 0:
        bar_w = max(left_border_width * PIXELS_TO_INCHES_X, 0.03)
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(x_in), Inches(y_in),
            Inches(bar_w), Inches(h_in),
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent_color
        _apply_fill_alpha(bar, left_result[1] * opacity)
        bar.line.fill.background()
        _apply_rotation(bar, dict(el, boxShadow="none"))


def _apply_shape_bg(
    shape_or_txbox,
    el: dict,
    opacity: float,
    backdrop: RGBColor | None = None,
) -> None:
    """Apply background fill (solid or gradient) to a shape or textbox."""
    gradient_css = el.get("backgroundImage", "")
    is_gradient_text = el.get("isGradientText", False)

    # Don't apply gradient as bg fill when the gradient is for text coloring
    if not is_gradient_text and _parse_css_gradient(gradient_css):
        _apply_gradient_fill(shape_or_txbox._element, gradient_css, backdrop, opacity)
        return

    bg_result = _css_color_to_rgb(el.get("backgroundColor", ""))
    if bg_result:
        bg_color, bg_alpha = bg_result
        effective_alpha = bg_alpha * opacity
        shape_or_txbox.fill.solid()
        shape_or_txbox.fill.fore_color.rgb = bg_color
        _apply_fill_alpha(shape_or_txbox, effective_alpha)


def _render_text_element(
    slide,
    el: dict,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    has_visual_bg: bool,
    opacity: float,
    backdrop: RGBColor | None = None,
    extra_left_px: float = 0.0,
) -> None:
    """Render a text box with optional background (solid, gradient, or translucent)."""
    # Gradient-text elements use backgroundImage for text coloring, not as a fill
    if el.get("isGradientText"):
        has_visual_bg = False

    text = el["text"]
    if el.get("whiteSpace", "normal") not in ("pre", "pre-wrap", "break-spaces"):
        text = text.strip(" ")

    transform = el.get("textTransform", "none")
    if transform == "uppercase":
        text = text.upper()
    elif transform == "lowercase":
        text = text.lower()
    elif transform == "capitalize":
        text = text.title()

    font_size_pt = el.get("fontSize", 20) * CSS_PX_TO_PT

    font_color = _resolve_font_color(el)
    font_family = _resolve_pptx_font(el.get("fontFamily", "Calibri"))
    is_bold = el.get("fontWeight", "400") in ("bold", "600", "700", "800", "900")
    is_italic = el.get("fontStyle", "normal") == "italic"

    is_single_line = _text_is_single_line(el, text)
    has_border = bool(el.get("borderColor")) and el.get("borderWidth", 0) > 0
    if has_visual_bg or has_border or el.get("boxShadow", "none") != "none":
        _render_bg_shape(slide, el, x_in, y_in, w_in, h_in, opacity, backdrop)
    x_in, y_in, w_in, h_in, text_el = _text_box_layout(el, x_in, y_in, w_in, h_in)
    alignment, vertical_anchor = _resolve_alignment(text_el, is_single_line, False)
    box_holder = slide.shapes.add_textbox(
        Inches(x_in), Inches(y_in), Inches(w_in), Inches(h_in),
    )
    tf = box_holder.text_frame
    white_space = el.get("whiteSpace", "normal")
    tf.word_wrap = not is_single_line and white_space not in ("pre", "nowrap")
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = vertical_anchor
    _apply_text_padding(tf, text_el, 0 if el.get("textGeometry") else extra_left_px)
    _apply_vertical_text(tf, el)
    _apply_rotation(box_holder, {**el, "boxShadow": "none"})

    p = tf.paragraphs[0]
    _apply_line_spacing(p, el)

    marker_color_str = el.get("markerColor")
    has_bullet = text.startswith("\u2022 ") or (
        len(text) >= 3 and text[0].isdigit() and text[:3].rstrip().endswith(".")
    )

    if marker_color_str and has_bullet:
        marker_result = _css_color_to_rgb(marker_color_str)
        if marker_result and text.startswith("\u2022 "):
            bullet_run = p.add_run()
            bullet_run.text = "\u2022 "
            _apply_font(
                bullet_run,
                size_pt=font_size_pt,
                color=marker_result[0],
                bold=is_bold,
                italic=False,
                family=font_family,
                letter_spacing_px=el.get("letterSpacing", 0),
                alpha=marker_result[1] * opacity,
            )
            text = text[2:]

    p.alignment = alignment
    href = el.get("href")
    for index, line in enumerate(text.split("\n")):
        if index:
            p.add_line_break()
        if not line:
            continue
        run = p.add_run()
        run.text = line
        _apply_font(
            run, size_pt=font_size_pt, color=font_color,
            bold=is_bold, italic=is_italic, family=font_family,
            letter_spacing_px=el.get("letterSpacing", 0),
            alpha=_text_alpha(el, opacity),
        )
        if el.get("isGradientText") and el.get("backgroundImage"):
            _apply_gradient_text(run, el["backgroundImage"], opacity)
        if href and not href.startswith("#"):
            try:
                run.hyperlink.address = href
            except Exception:
                pass


def _render_inline_runs(
    slide,
    el: dict,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    has_bg: bool,
    opacity: float,
    backdrop: RGBColor | None = None,
) -> None:
    """Render a text box with multiple styled runs from inline mixed content.

    Handles elements like <p>Hello <strong>bold</strong> more text</p>
    where text nodes and inline children are interleaved.
    """
    runs_data = el.get("inlineRuns", [])
    if not runs_data:
        return

    text = "".join(run.get("text", "") for run in runs_data)
    is_single_line = _text_is_single_line(el, text)
    if has_bg or el.get("boxShadow", "none") != "none":
        _render_bg_shape(slide, el, x_in, y_in, w_in, h_in, opacity, backdrop)
    x_in, y_in, w_in, h_in, text_el = _text_box_layout(el, x_in, y_in, w_in, h_in)
    alignment, vertical_anchor = _resolve_alignment(text_el, is_single_line, False)
    box_holder = slide.shapes.add_textbox(
        Inches(x_in), Inches(y_in), Inches(w_in), Inches(h_in),
    )
    tf = box_holder.text_frame
    tf.vertical_anchor = vertical_anchor
    white_space = el.get("whiteSpace", "normal")
    tf.word_wrap = not is_single_line and white_space not in ("pre", "nowrap")
    tf.auto_size = MSO_AUTO_SIZE.NONE
    _apply_text_padding(tf, text_el)
    _apply_vertical_text(tf, el)
    _apply_rotation(box_holder, {**el, "boxShadow": "none"})

    p = tf.paragraphs[0]
    p.alignment = alignment
    _apply_line_spacing(p, el)

    for run_data in runs_data:
        text = run_data.get("text", "")
        if text == "\n":
            p.add_line_break()
            continue
        if not text:
            continue

        transform = run_data.get("textTransform", "none")
        if transform == "uppercase":
            text = text.upper()
        elif transform == "lowercase":
            text = text.lower()
        elif transform == "capitalize":
            text = text.title()

        run_size = run_data.get("fontSize", el.get("fontSize", 20)) * CSS_PX_TO_PT

        run_is_gradient = run_data.get("isGradientText", False)
        run_bg_image = run_data.get("backgroundImage", "")

        color_result = _css_color_to_rgb(run_data.get("color", el.get("color", "rgb(255,255,255)")))
        run_alpha = color_result[1] if color_result else 1.0
        if run_is_gradient and run_bg_image:
            run_color = _first_gradient_color(run_bg_image) or RGBColor(0xFF, 0xFF, 0xFF)
        elif color_result:
            run_color = color_result[0]
        else:
            run_color = _resolve_font_color(el)

        run_bold = run_data.get("fontWeight", "400") in ("bold", "600", "700", "800", "900")
        run_italic = run_data.get("fontStyle", "normal") == "italic"
        run_family = _resolve_pptx_font(run_data.get("fontFamily", el.get("fontFamily", "Calibri")))

        r = p.add_run()
        r.text = text
        _apply_font(
            r, size_pt=run_size, color=run_color, bold=run_bold,
            italic=run_italic, family=run_family,
            letter_spacing_px=run_data.get("letterSpacing", el.get("letterSpacing", 0)),
            alpha=run_alpha * opacity * run_data.get("opacity", 1),
        )
        if run_is_gradient and run_bg_image:
            _apply_gradient_text(r, run_bg_image, opacity * run_data.get("opacity", 1))

        href = run_data.get("href")
        if href and not href.startswith("#"):
            try:
                r.hyperlink.address = href
            except Exception:
                pass


# ---------------------------------------------------------------------------
# PPTX rendering (element tree -> slide shapes)
# ---------------------------------------------------------------------------


def _render_measured_element(
    slide, el: dict, backdrop: RGBColor | None = None, inherited_opacity: float = 1.0,
) -> None:
    """Render an explicit DOM group without moving or reordering its members."""
    if not el.get("isGroup"):
        _render_element_content(slide, el, backdrop, inherited_opacity)
        return
    first = len(slide.shapes)
    _render_element_content(slide, el, backdrop, inherited_opacity)
    members = [shape for index, shape in enumerate(slide.shapes) if index >= first]
    if members:
        group = slide.shapes.add_group_shape(members)
        if el.get("groupName"):
            group.name = el["groupName"]
        # off/chOff are equal: members retain their global DOM coordinates.
        _apply_shape_effects(group, {})


def _render_element_content(
    slide, el: dict, backdrop: RGBColor | None = None, inherited_opacity: float = 1.0,
) -> None:
    """Render a single measured element onto a PPTX slide.

    Traversal strategy:
    - Image → render picture shape, stop
    - Inline runs (mixed content like <p>text <strong>bold</strong> more</p>)
        → render background if any, render multi-run text box, then
          recurse only non-inline (block) children
    - Text leaf (has text, no descendant text) → render text box, stop
    - Non-leaf with background/gradient → render shape, recurse children
    - Pure container → recurse children only
    """
    if el.get("_skip"):
        return

    x_in = el["x"] * PIXELS_TO_INCHES_X
    y_in = el["y"] * PIXELS_TO_INCHES_Y
    w_in = el["width"] * PIXELS_TO_INCHES_X
    h_in = el["height"] * PIXELS_TO_INCHES_Y

    # Keep the unrotated local box, including off-slide coordinates; resizing
    # would move the rotation center. The slide viewport clips during export.
    if w_in <= 0 or h_in <= 0:
        return
    w_in = max(w_in, 1 / 914400)
    h_in = max(h_in, 1 / 914400)

    is_image = el.get("isImage", False)
    has_bg = el.get("backgroundColor") is not None
    is_gradient_text = el.get("isGradientText", False)
    has_gradient = (
        bool(_parse_css_gradient(el.get("backgroundImage", "")))
        and not is_gradient_text
    )
    has_border = bool(el.get("borderColor")) and el.get("borderWidth", 0) > 0
    has_visual_bg = has_bg or has_gradient or (
        bool(el.get("boxShadow")) and el["boxShadow"] != "none"
    )
    has_text = bool(el.get("text", "").strip())
    has_inline_runs = bool(el.get("inlineRuns"))
    children = el.get("children", [])
    # A CSS parent opacity visually multiplies its descendants. Containers are
    # not always collapsed (e.g. a top-level faint image wrapper), so fold the
    # inherited opacity in here and pass it down, or it is silently dropped.
    opacity = el.get("opacity", 1.0) * inherited_opacity
    el = dict(el, opacity=opacity)

    # Detect standalone left-border accents (common decorative pattern)
    has_left_accent = (
        el.get("borderLeftWidth", 0) > 0
        and el.get("borderLeftStyle") not in (None, "none")
        and bool(_css_color_to_rgb(el.get("borderLeftColor", "")))
    )

    child_backdrop = backdrop

    if (is_image or el.get("isSvg")) and (el.get("src") or "").startswith("data:image/"):
        pic = _add_image_from_data_uri(
            slide, el["src"],
            Inches(x_in), Inches(y_in),
            Inches(w_in), Inches(h_in),
            border_radius_px=_parse_border_radius_px(el),
            width_px=el.get("width", 0),
            height_px=el.get("height", 0),
            opacity=opacity,
            object_fit=el.get("objectFit"),
            natural_w=el.get("naturalWidth", 0),
            natural_h=el.get("naturalHeight", 0),
        )
        if pic is not None:
            _apply_rotation(pic, el)
        return

    # Unreasterized SVGs: skip (no python-pptx SVG support)
    if el.get("isSvg"):
        return

    # conic-gradient (pie/donut) — rasterize to a picture, then draw children
    # (e.g. the center hole + label) on top.
    if "conic-gradient" in (el.get("backgroundImage") or ""):
        _render_conic_gradient(slide, el, x_in, y_in, w_in, h_in, opacity)
        for child in children:
            _render_measured_element(slide, child, child_backdrop, opacity)
        return

    if has_inline_runs:
        _render_inline_runs(slide, el, x_in, y_in, w_in, h_in, has_visual_bg, opacity, backdrop)
        for child in children:
            if child.get("position") in ("absolute", "fixed") or child.get("tag") not in (
                "span", "strong", "em", "b", "i", "a", "code", "mark",
                "sub", "sup", "small", "u", "s", "del",
            ):
                _render_measured_element(slide, child, child_backdrop, opacity)
        return

    is_text_leaf = has_text and not _any_descendant_has_text(el)
    if is_text_leaf:
        # A leading, in-flow decorative child (e.g. an inline flag dot before a
        # city name: <h4><span class="flag-dot"></span>Tokyo</h4>) occupies
        # horizontal space in the browser, so the text must start after it or the
        # dot overprints the first letter. Absolutely-positioned accents (bullet
        # ::before) sit in the padding and are handled by padding instead.
        extra_left_px = 0.0
        el_left = el.get("x", 0)
        el_mid_y = el.get("y", 0) + el.get("height", 0) / 2
        for child in children:
            if child.get("text", "").strip() or _any_descendant_has_text(child):
                continue
            if child.get("position") in ("absolute", "fixed"):
                continue
            cx = child.get("x", 0)
            cright = cx + child.get("width", 0)
            c_top, c_bot = child.get("y", 0), child.get("y", 0) + child.get("height", 0)
            near_left = cx <= el_left + el.get("width", 0) * 0.4
            vertically_on_line = c_top <= el_mid_y <= c_bot
            if near_left and cright > el_left and vertically_on_line:
                extra_left_px = max(extra_left_px, cright - el_left + 8)
        _render_text_element(
            slide, el, x_in, y_in, w_in, h_in, has_visual_bg, opacity, backdrop,
            extra_left_px,
        )
        # Render the decorative, text-free children themselves (e.g. bullet dots).
        for child in children:
            _render_measured_element(slide, child, child_backdrop, opacity)
        return

    if has_visual_bg or has_border or has_left_accent:
        _render_bg_shape(slide, el, x_in, y_in, w_in, h_in, opacity, backdrop)

    for child in children:
        _render_measured_element(slide, child, child_backdrop, opacity)


# ---------------------------------------------------------------------------
# Pipeline stages (public API)
# ---------------------------------------------------------------------------


def _walk_elements(elements: list[dict]):
    """Yield every element in a measurement tree (depth-first)."""
    for el in elements:
        yield el
        yield from _walk_elements(el.get("children", []))


async def _rasterize_inline_svgs(page, measurements: list[dict]) -> None:
    """Screenshot inline <svg> elements and patch them as raster images.

    SVGs can't be added directly to python-pptx. This second pass shows
    each slide, finds marked SVGs, screenshots them as PNGs, and converts
    the measurement to an image element.
    """
    slide_state = None
    try:
        for i, slide_data in enumerate(measurements):
            svg_els = [e for e in _walk_elements(slide_data.get("elements", []))
                       if e.get("isSvg") and e.get("svgId")]
            if not svg_els:
                continue

            if slide_state is None:
                slide_state = await page.evaluate_handle(
                    "() => {" + _SLIDE_STATE_JS
                    + "return captureSlideState(document.querySelectorAll('.slide'));}"
                )
            await slide_state.evaluate("(state, index) => state.show(index)", i)

            for el in svg_els:
                try:
                    handle = await page.query_selector(
                        f'[data-pptx-id="{el["svgId"]}"]'
                    )
                    if not handle:
                        continue
                    png_bytes = await handle.screenshot(type="png")
                    encoded = base64.b64encode(png_bytes).decode()
                    el["isImage"] = True
                    el["isSvg"] = False
                    el["src"] = f"data:image/png;base64,{encoded}"
                    el["children"] = []
                except Exception:
                    logger.debug("Failed to rasterize SVG %s", el.get("svgId", "?"))
    finally:
        if slide_state is not None:
            try:
                await slide_state.evaluate("state => state.restore()")
            finally:
                await slide_state.dispose()


async def extract_measurements(html_path: str) -> list[dict]:
    """Open an HTML slide deck in headless Chromium and measure every element.

    Args:
        html_path: Path to the HTML file.

    Returns:
        List of slide measurement dicts, each containing:
          - index, width, height, backgroundColor
          - elements: recursive tree of measured DOM nodes

    Raises:
        FileNotFoundError: If html_path does not exist.
        ValueError: If the file exceeds the size limit or contains no slides.
        ImportError: If Playwright is not installed.
    """
    abs_path = os.path.abspath(html_path)
    if not os.path.isfile(abs_path):
        raise FileNotFoundError(f"HTML file not found: {abs_path}")

    file_size_mb = os.path.getsize(abs_path) / (1024 * 1024)
    if file_size_mb > MAX_HTML_SIZE_MB:
        raise ValueError(
            f"HTML file is {file_size_mb:.1f} MB, exceeds {MAX_HTML_SIZE_MB} MB limit"
        )

    from playwright.async_api import async_playwright

    logger.info("Loading %s (%.1f MB)", abs_path, file_size_mb)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(
            viewport={
                "width": SLIDE_CANVAS_WIDTH_PX,
                "height": SLIDE_CANVAS_HEIGHT_PX,
            },
            service_workers="block",
        )

        async def block_network(route):
            if urlsplit(route.request.url).scheme in ("http", "https"):
                await route.abort()
            else:
                await route.continue_()

        await page.route("**/*", block_network)

        await page.goto(
            Path(abs_path).as_uri(),
            wait_until="networkidle",
            timeout=PLAYWRIGHT_TIMEOUT_MS,
        )
        await page.evaluate("() => document.fonts.ready")
        image_sources = await page.evaluate(
            "() => Array.from(document.images, img => img.currentSrc || img.src)"
        )
        image_data = {}
        for source in image_sources:
            if source.startswith("data:image/"):
                continue
            url = urlsplit(source)
            if url.scheme != "file" or url.netloc not in ("", "localhost"):
                raise ValueError(f"Unsupported image source (embed a data URI or use a local file): {source}")
            image_path = Path(unquote(url.path))
            mime = mimetypes.guess_type(image_path.name)[0]
            if not mime or not mime.startswith("image/"):
                raise ValueError(f"Unsupported image type: {image_path}")
            image_data[source] = f"data:{mime};base64," + base64.b64encode(image_path.read_bytes()).decode("ascii")
        await page.evaluate(
            """async imageData => {
                await Promise.all(Array.from(document.images, async img => {
                    const source = img.currentSrc || img.src;
                    if (imageData[source]) {
                        img.removeAttribute('srcset');
                        const picture = img.closest('picture');
                        if (picture) picture.querySelectorAll('source').forEach(source => source.remove());
                        img.src = imageData[source];
                    }
                    await img.decode();
                }));
            }""",
            image_data,
        )
        await page.evaluate("() => document.fonts.ready")

        measurements = await page.evaluate(EXTRACTION_JS)

        await _rasterize_inline_svgs(page, measurements)

        await browser.close()

    if not measurements:
        raise ValueError(
            'No slides found. Ensure the HTML contains <section class="slide"> elements.'
        )

    total_elements = sum(
        _count_elements(s.get("elements", [])) for s in measurements
    )
    logger.info(
        "Extracted %d slides, %d total elements", len(measurements), total_elements,
    )
    return measurements




def render_pptx(measurements: list[dict]) -> Presentation:
    """Convert DOM measurements into a python-pptx Presentation.

    Args:
        measurements: Output of extract_measurements() — one dict per slide,
            each with a backgroundColor and a recursive elements tree.

    Returns:
        A python-pptx Presentation ready to be saved.
    """
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_INCHES)
    prs.slide_height = Inches(SLIDE_HEIGHT_INCHES)
    blank_layout = prs.slide_layouts[6]

    for slide_data in measurements:
        slide = prs.slides.add_slide(blank_layout)

        gradient_css = slide_data.get("backgroundImage", "")
        if gradient_css and _parse_css_gradient(gradient_css):
            slide.background.fill.gradient()
            _apply_gradient_fill(slide.background._element, gradient_css)
        else:
            bg_result = _css_color_to_rgb(slide_data.get("backgroundColor", ""))
            if bg_result:
                bg_color, _ = bg_result
                fill = slide.background.fill
                fill.solid()
                fill.fore_color.rgb = bg_color

        backdrop = _resolve_backdrop(slide_data)
        for el in slide_data.get("elements", []):
            _render_measured_element(slide, el, backdrop)

    return prs


def _check_distinct_paths(input_path: str | Path, output_path: str | Path) -> None:
    """Reject direct paths, symlinks, and hardlinks that would overwrite input."""
    source, output = Path(input_path), Path(output_path)
    if source.resolve() == output.resolve() or (
        source.exists() and output.exists() and source.samefile(output)
    ):
        raise ValueError("Input and output must be different files; output was not written")


async def convert(input_html: str, output_pptx: str = "output.pptx") -> str:
    """Convert an HTML slide deck to an editable PPTX file.

    Args:
        input_html: Path to the HTML file containing <section class="slide"> elements.
        output_pptx: Path where the .pptx file will be written.

    Returns:
        The output_pptx path.

    Raises:
        FileNotFoundError: If input_html does not exist.
        ValueError: If input and output refer to the same file, the file exceeds
            the size limit, or it contains no slides.
        ImportError: If Playwright is not installed.
    """
    _check_distinct_paths(input_html, output_pptx)
    measurements = await extract_measurements(input_html)
    prs = render_pptx(measurements)
    prs.save(output_pptx)

    file_size_mb = os.path.getsize(output_pptx) / (1024 * 1024)
    logger.info(
        "Saved %s (%.1f MB, %d slides)",
        output_pptx, file_size_mb, len(measurements),
    )
    return output_pptx


