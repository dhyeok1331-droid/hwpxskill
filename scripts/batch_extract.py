#!/usr/bin/env python3
"""Extract text from every HWPX file in a folder.

Wraps text_extract.py's conversion logic to process a whole directory at once.

Usage:
    python batch_extract.py ./hwpx_files
    python batch_extract.py ./hwpx_files --format markdown
    python batch_extract.py ./hwpx_files --format markdown --output-dir ./md_out
    python batch_extract.py ./hwpx_files --recursive
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from text_extract import extract_markdown, extract_plain  # noqa: E402


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract text from all .hwpx files in a folder"
    )
    parser.add_argument("input_dir", help="Folder containing .hwpx files")
    parser.add_argument(
        "--format", "-f",
        choices=["plain", "markdown"],
        default="markdown",
        help="Output format (default: markdown)",
    )
    parser.add_argument(
        "--include-tables",
        action="store_true",
        help="Include text from tables and nested objects (plain mode)",
    )
    parser.add_argument(
        "--output-dir", "-o",
        help="Folder to write converted files to (default: same folder as each .hwpx)",
    )
    parser.add_argument(
        "--recursive", "-r",
        action="store_true",
        help="Search subfolders as well",
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    if not input_dir.is_dir():
        print(f"Error: Folder not found: {input_dir}", file=sys.stderr)
        sys.exit(1)

    pattern = "**/*.hwpx" if args.recursive else "*.hwpx"
    files = sorted(input_dir.glob(pattern))
    if not files:
        print(f"No .hwpx files found in: {input_dir}", file=sys.stderr)
        sys.exit(1)

    ext = ".md" if args.format == "markdown" else ".txt"
    output_dir = Path(args.output_dir) if args.output_dir else None
    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)

    ok, failed = 0, 0
    for src in files:
        if output_dir:
            dest_dir = output_dir / src.parent.relative_to(input_dir)
            dest_dir.mkdir(parents=True, exist_ok=True)
        else:
            dest_dir = src.parent
        dest = dest_dir / src.with_suffix(ext).name
        try:
            if args.format == "markdown":
                result = extract_markdown(str(src))
            else:
                result = extract_plain(str(src), include_tables=args.include_tables)
            dest.write_text(result, encoding="utf-8")
            print(f"OK   {src} -> {dest}", flush=True)
            ok += 1
        except Exception as exc:
            print(f"FAIL {src}: {exc}", file=sys.stderr, flush=True)
            failed += 1

    print(f"\nDone: {ok} converted, {failed} failed.", file=sys.stderr)
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
