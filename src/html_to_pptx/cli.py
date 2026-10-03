"""Command-line interface for html-to-pptx."""

from __future__ import annotations

import argparse
import asyncio
import json
import logging
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="html-to-pptx",
        description="Convert an HTML slide deck to an editable PPTX file.",
    )
    parser.add_argument(
        "input", nargs="?", type=Path,
        help="HTML slide deck, or measurement JSON with --measurements",
    )
    parser.add_argument(
        "output", nargs="?", type=Path,
        help="Output .pptx path (default: output.pptx)",
    )
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--extract-js", action="store_true",
        help="Print the browser measurement function; run it after fonts and images load",
    )
    mode.add_argument(
        "--measurements", action="store_true",
        help="Render measurement JSON from an existing Chromium browser",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Enable debug logging",
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(name)s: %(message)s",
    )

    from html_to_pptx.converter import (
        EXTRACTION_JS, _check_distinct_paths, extract_measurements, render_pptx,
    )

    if args.extract_js:
        if args.input is not None or args.output is not None:
            parser.error("--extract-js does not take file arguments")
        print(EXTRACTION_JS)
        return
    if args.input is None:
        parser.error("input is required")
    if not args.input.is_file():
        parser.error(f"Input file not found: {args.input}")
    output = args.output or Path("output.pptx")

    try:
        _check_distinct_paths(args.input, output)
        if args.measurements:
            measurements = json.loads(args.input.read_text(encoding="utf-8"))
            if not measurements:
                parser.error("No slides found; output was not written")
            if not isinstance(measurements, list) or not all(
                isinstance(slide, dict) for slide in measurements
            ):
                parser.error("Measurement JSON must be an array of slide objects")
        else:
            measurements = asyncio.run(extract_measurements(str(args.input)))
    except ModuleNotFoundError as exc:
        if exc.name and exc.name.startswith("playwright"):
            parser.exit(
                2,
                "HTML extraction requires Playwright and an installed Chromium. "
                "Nothing was installed. Alternatively run --extract-js in an "
                "existing browser and render its JSON with --measurements.\n",
            )
        raise
    except (OSError, ValueError) as exc:
        parser.error(str(exc))

    if not measurements:
        parser.error("No slides found; output was not written")
    presentation = render_pptx(measurements)
    output.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(str(output))
    print(f"Saved {len(presentation.slides)} slides: {output}")


if __name__ == "__main__":
    main()
