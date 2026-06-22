#!/usr/bin/env python3
"""Convert PDF files to Markdown text, with OCR fallback for scanned pages.

Each page becomes a "## Page N" section. Pages with extractable text use
PyMuPDF directly; pages with little/no extractable text (scanned images)
are rendered to an image and run through Tesseract OCR (Korean+English).

Usage:
    python pdf_extract.py document.pdf
    python pdf_extract.py ./pdf_files --recursive --output-dir ./md_out
    python pdf_extract.py ./pdf_files --recursive --exclude-list dupes.txt
"""

import argparse
import sys
from pathlib import Path

TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
MIN_TEXT_CHARS = 20
OCR_DPI = 200
MAX_PAGES = 200
TIMEOUT_SECONDS = 120


def _ocr_page(page) -> str:
    import io

    import pytesseract
    from PIL import Image

    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    pix = page.get_pixmap(dpi=OCR_DPI)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    return pytesseract.image_to_string(img, lang="kor+eng")


def extract_markdown(path: str) -> str:
    """Convert a PDF to Markdown, one '## Page N' section per page."""

    import fitz

    sections = []
    with fitz.open(path) as doc:
        for i, page in enumerate(doc):
            if i >= MAX_PAGES:
                sections.append(f"## Page {i + 1}+\n\n*(문서가 {MAX_PAGES}페이지를 넘어 이후는 생략됨)*")
                break
            text = page.get_text()
            used_ocr = False
            if len(text.strip()) < MIN_TEXT_CHARS:
                ocr_text = _ocr_page(page)
                if len(ocr_text.strip()) > len(text.strip()):
                    text = ocr_text
                    used_ocr = True
            note = " *(OCR)*" if used_ocr else ""
            body = text.strip() or "(빈 페이지)"
            sections.append(f"## Page {i + 1}{note}\n\n{body}")

    return "\n\n---\n\n".join(sections)


def _convert_one(src: str, dest: str) -> None:
    """Run in a separate process so the parent can enforce a timeout."""
    result = extract_markdown(src)
    Path(dest).write_text(result, encoding="utf-8")


def _load_exclude_set(exclude_list_path: str | None) -> set[str]:
    if not exclude_list_path:
        return set()
    lines = Path(exclude_list_path).read_text(encoding="utf-8").splitlines()
    return {str(Path(line.strip())) for line in lines if line.strip()}


def convert_folder(
    input_dir: Path, *, output_dir: Path | None, recursive: bool, skip_existing: bool, exclude: set[str]
) -> tuple[int, int, int, int]:
    import multiprocessing

    pattern = "**/*.pdf" if recursive else "*.pdf"
    files = sorted(p for p in input_dir.glob(pattern) if p.is_file())
    if not files:
        print(f"No .pdf files found in: {input_dir}", file=sys.stderr)
        sys.exit(1)

    ok, failed, skipped, excluded = 0, 0, 0, 0
    for src in files:
        if str(src) in exclude:
            print(f"EXCLUDE {src} (중복 후보, hwpx에서 이미 변환됨)", flush=True)
            excluded += 1
            continue

        if output_dir:
            dest_dir = output_dir / src.parent.relative_to(input_dir)
            dest_dir.mkdir(parents=True, exist_ok=True)
        else:
            dest_dir = src.parent
        dest = dest_dir / (src.stem + ".md")

        if skip_existing and dest.exists():
            print(f"SKIP {src} (already converted)", flush=True)
            skipped += 1
            continue

        proc = multiprocessing.Process(target=_convert_one, args=(str(src), str(dest)))
        proc.start()
        proc.join(TIMEOUT_SECONDS)
        if proc.is_alive():
            proc.terminate()
            proc.join(5)
            print(f"FAIL {src}: timed out after {TIMEOUT_SECONDS}s, skipped", file=sys.stderr, flush=True)
            failed += 1
        elif proc.exitcode == 0 and dest.exists():
            print(f"OK   {src} -> {dest}", flush=True)
            ok += 1
        else:
            print(f"FAIL {src}: worker process exited with code {proc.exitcode}", file=sys.stderr, flush=True)
            failed += 1

    return ok, failed, skipped, excluded


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert PDF files to Markdown (with OCR fallback)")
    parser.add_argument("input", help="Path to .pdf file, or a folder")
    parser.add_argument("--output-dir", "-o", help="Folder to write .md files to")
    parser.add_argument("--recursive", "-r", action="store_true", help="Search subfolders as well")
    parser.add_argument("--skip-existing", action="store_true", help="Skip files with an existing .md output")
    parser.add_argument(
        "--exclude-list",
        help="Text file with one PDF path per line to skip (e.g. confirmed duplicates of a hwp/hwpx)",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir) if args.output_dir else None

    if input_path.is_file():
        result = extract_markdown(str(input_path))
        if output_dir:
            output_dir.mkdir(parents=True, exist_ok=True)
            dest = output_dir / (input_path.stem + ".md")
        else:
            dest = input_path.with_suffix(".md")
        dest.write_text(result, encoding="utf-8")
        print(f"Converted: {dest}")
        return

    if not input_path.is_dir():
        print(f"Error: Path not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)

    exclude = _load_exclude_set(args.exclude_list)
    ok, failed, skipped, excluded = convert_folder(
        input_path, output_dir=output_dir, recursive=args.recursive, skip_existing=args.skip_existing, exclude=exclude
    )
    print(f"\nDone: {ok} converted, {skipped} skipped, {excluded} excluded (dupes), {failed} failed.", file=sys.stderr)
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
