#!/usr/bin/env python3
"""Batch-convert legacy .hwp files to .hwpx using the installed Hancom Office.

.hwp is a proprietary binary format. There is no reliable pure-Python
writer for it, so this script drives the real 한글(Hancom Office) app
via COM automation and asks it to "Save As" each file in HWPX format.
Requires Windows with Hancom Office (한글) installed.

Usage:
    python convert_hwp_to_hwpx.py ./hwp_files
    python convert_hwp_to_hwpx.py ./hwp_files --output-dir ./hwpx_out
    python convert_hwp_to_hwpx.py ./hwp_files --recursive
"""

import argparse
import sys
from pathlib import Path


def convert_folder(input_dir: Path, *, output_dir: Path | None, recursive: bool, skip_existing: bool) -> tuple[int, int, int]:
    import win32com.client

    pattern = "**/*.hwp" if recursive else "*.hwp"
    files = sorted(p for p in input_dir.glob(pattern) if p.suffix.lower() == ".hwp")
    if not files:
        print(f"No .hwp files found in: {input_dir}", file=sys.stderr)
        sys.exit(1)

    hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
    hwp.XHwpWindows.Item(0).Visible = False

    ok, failed, skipped = 0, 0, 0
    try:
        for src in files:
            if output_dir:
                dest_dir = output_dir / src.parent.relative_to(input_dir)
                dest_dir.mkdir(parents=True, exist_ok=True)
            else:
                dest_dir = src.parent
            dest = dest_dir / src.with_suffix(".hwpx").name

            if skip_existing and dest.exists():
                print(f"SKIP {src} (already converted)", flush=True)
                skipped += 1
                continue

            try:
                hwp.Open(str(src))
                hwp.SaveAs(str(dest), "HWPX")
                hwp.Clear(1)
                print(f"OK   {src} -> {dest}", flush=True)
                ok += 1
            except Exception as exc:
                print(f"FAIL {src}: {exc}", file=sys.stderr, flush=True)
                failed += 1
    finally:
        hwp.Quit()

    return ok, failed, skipped


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Batch-convert .hwp files to .hwpx via Hancom Office automation"
    )
    parser.add_argument("input_dir", help="Folder containing .hwp files")
    parser.add_argument(
        "--output-dir", "-o",
        help="Folder to write .hwpx files to (default: same folder as each .hwp)",
    )
    parser.add_argument(
        "--recursive", "-r",
        action="store_true",
        help="Search subfolders as well",
    )
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        help="Skip files that already have a matching .hwpx output (for resuming an interrupted run)",
    )
    args = parser.parse_args()

    if sys.platform != "win32":
        print("Error: this script requires Windows + Hancom Office.", file=sys.stderr)
        sys.exit(1)

    input_dir = Path(args.input_dir)
    if not input_dir.is_dir():
        print(f"Error: Folder not found: {input_dir}", file=sys.stderr)
        sys.exit(1)

    output_dir = Path(args.output_dir) if args.output_dir else None
    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)

    ok, failed, skipped = convert_folder(
        input_dir, output_dir=output_dir, recursive=args.recursive, skip_existing=args.skip_existing
    )

    print(f"\nDone: {ok} converted, {skipped} skipped, {failed} failed.", file=sys.stderr)
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
