"""CLI errors must not destroy source data or an existing destination."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

from html_to_pptx.cli import main


def assert_cli_error(monkeypatch, capsys, arguments: list[str], message: str) -> None:
    monkeypatch.setattr(sys, "argv", ["html-to-pptx", *arguments])
    with pytest.raises(SystemExit) as exc:
        main()
    assert exc.value.code == 2
    assert message in capsys.readouterr().err


def test_missing_input_reports_usage(monkeypatch, capsys):
    assert_cli_error(monkeypatch, capsys, [], "input is required")


def test_missing_file_preserves_destination(tmp_path: Path, monkeypatch, capsys):
    output = tmp_path / "existing.pptx"
    output.write_bytes(b"destination must survive")
    assert_cli_error(
        monkeypatch, capsys,
        [str(tmp_path / "missing.html"), str(output)],
        "Input file not found",
    )
    assert output.read_bytes() == b"destination must survive"


@pytest.mark.parametrize("contents", ["[]", "null", "{}"])
def test_empty_measurements_preserve_files(tmp_path: Path, monkeypatch, capsys, contents: str):
    source = tmp_path / "measurements.json"
    source.write_text(contents, encoding="utf-8")
    output = tmp_path / "existing.pptx"
    output.write_bytes(b"destination must survive")
    assert_cli_error(
        monkeypatch, capsys,
        ["--measurements", str(source), str(output)],
        "No slides found",
    )
    assert source.read_text(encoding="utf-8") == contents
    assert output.read_bytes() == b"destination must survive"


@pytest.mark.parametrize("contents", ["not JSON", '["not a slide"]', '{"slides": [{}]}'])
def test_invalid_measurements_preserve_destination(tmp_path: Path, monkeypatch, capsys, contents: str):
    source = tmp_path / "measurements.json"
    source.write_text(contents, encoding="utf-8")
    output = tmp_path / "existing.pptx"
    output.write_bytes(b"destination must survive")
    assert_cli_error(
        monkeypatch, capsys,
        ["--measurements", str(source), str(output)],
        "error:",
    )
    assert source.read_text(encoding="utf-8") == contents
    assert output.read_bytes() == b"destination must survive"


@pytest.mark.parametrize("alias", ["same", "symlink", "hardlink"])
@pytest.mark.parametrize("mode", [[], ["--measurements"]])
def test_source_cannot_be_output(tmp_path: Path, monkeypatch, capsys, alias: str, mode: list[str]):
    source = tmp_path / "source.json"
    original = b'[{"elements": []}]'
    source.write_bytes(original)
    output = source
    if alias != "same":
        output = tmp_path / "alias.pptx"
        if alias == "symlink":
            output.symlink_to(source)
        else:
            output.hardlink_to(source)
    assert_cli_error(
        monkeypatch, capsys,
        [*mode, str(source), str(output)],
        "Input and output must be different files",
    )
    assert source.read_bytes() == original
    assert output.read_bytes() == original


def test_extract_js_rejects_file_arguments(tmp_path: Path, monkeypatch, capsys):
    source = tmp_path / "source.html"
    source.write_text("<section class='slide'>Keep me</section>", encoding="utf-8")
    assert_cli_error(
        monkeypatch, capsys,
        ["--extract-js", str(source), str(source)],
        "--extract-js does not take file arguments",
    )
    assert source.read_text(encoding="utf-8") == "<section class='slide'>Keep me</section>"
