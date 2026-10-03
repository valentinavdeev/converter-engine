"""Behavioral checks for font fallback, supported gradients and text alignment."""

from __future__ import annotations

from pptx.enum.text import MSO_ANCHOR
from pptx.enum.text import PP_ALIGN

from html_to_pptx import converter as C


# ---------------------------------------------------------------------------
# Font substitution stays in-family and rendering-safe
# ---------------------------------------------------------------------------


def test_font_resolution_in_family():
    assert C._resolve_pptx_font("'Playfair Display', Georgia, serif") == "Georgia"
    assert C._resolve_pptx_font("'Inter', 'Segoe UI', system-ui, sans-serif") == "Segoe UI"
    assert C._resolve_pptx_font("Roboto, Arial, sans-serif") == "Arial"
    assert C._resolve_pptx_font("'JetBrains Mono', monospace") == "Consolas"
    assert C._resolve_pptx_font("Montserrat, sans-serif") == "Century Gothic"
    assert C._resolve_pptx_font("'Totally Unknown', serif") == "Georgia"




def test_radial_and_conic_and_multi_gradients_rejected():
    assert C._parse_css_gradient("radial-gradient(circle, rgb(0,0,0), rgb(1,1,1))") is None
    assert C._parse_css_gradient("conic-gradient(rgb(0,0,0) 0deg, rgb(1,1,1) 180deg)") is None
    assert C._parse_css_gradient(
        "radial-gradient(rgba(0,0,0,0.2), transparent), linear-gradient(rgb(1,1,1), rgb(2,2,2))"
    ) is None


def test_translucent_gradient_stops_preserve_alpha():
    stops = C._parse_css_gradient(
        "linear-gradient(90deg, rgba(232,119,46,0.06), rgba(232,119,46,0.01))"
    )
    assert stops
    assert abs(stops[0][2] - 0.06) < 1e-6
    assert abs(stops[1][2] - 0.01) < 1e-6


# ---------------------------------------------------------------------------
# Conic-gradient parsing (for rasterized pies/donuts)
# ---------------------------------------------------------------------------


def test_conic_gradient_parse_segments():
    segs = C._parse_conic_gradient(
        "conic-gradient(rgb(79,70,229) 0deg 223.2deg, "
        "rgb(6,182,212) 223.2deg 302.4deg, rgb(228,231,236) 302.4deg 360deg)"
    )
    assert segs and len(segs) == 3
    assert segs[0][1] == 0 and abs(segs[0][2] - 223.2) < 0.1
    assert abs(segs[-1][2] - 360.0) < 0.1


def test_conic_gradient_percent_positions():
    segs = C._parse_conic_gradient(
        "conic-gradient(rgb(0,0,0) 0 62%, rgb(1,1,1) 62% 84%, rgb(2,2,2) 84% 100%)"
    )
    assert segs and len(segs) == 3
    assert abs(segs[0][2] - 62 * 3.6) < 0.1  # 62% -> 223.2deg


# ---------------------------------------------------------------------------
# Line-height ratio
# ---------------------------------------------------------------------------


def test_line_height_ratio():
    assert abs(C._line_height_ratio({"lineHeight": "101.2px", "fontSize": 110}) - 0.92) < 1e-3
    assert C._line_height_ratio({"lineHeight": "normal", "fontSize": 20}) is None
    assert C._line_height_ratio({"lineHeight": "12px", "fontSize": 0}) is None


# ---------------------------------------------------------------------------
# CSS-driven alignment (flex, chips, tables)
# ---------------------------------------------------------------------------


def test_flex_center_maps_to_center_middle():
    el = {"display": "flex", "justifyContent": "center", "alignItems": "center",
          "textAlign": "start"}
    h, v = C._resolve_alignment(el, is_single_line=False, has_visual_bg=True)
    assert h == PP_ALIGN.CENTER and v == MSO_ANCHOR.MIDDLE


def test_table_cell_keeps_column_alignment():
    el = {"tag": "th", "display": "table-cell", "textAlign": "left", "backgroundColor": "rgb(0,0,0)"}
    h, _ = C._resolve_alignment(el, is_single_line=True, has_visual_bg=True)
    assert h == PP_ALIGN.LEFT  # not hijacked by chip-centering




