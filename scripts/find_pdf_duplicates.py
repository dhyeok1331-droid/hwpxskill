#!/usr/bin/env python3
"""Find PDF files that are likely duplicates of a .hwp/.hwpx file in the same folder.

This does NOT delete or move anything. It only writes a Markdown report of
candidate duplicate pairs (filename similarity + text-content similarity) so
a human can review and confirm before any file is archived/removed.

Usage:
    python find_pdf_duplicates.py ./folder --recursive --output report.md
"""

import argparse
import difflib
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

NAME_THRESHOLD = 0.6
CONTENT_SAMPLE_CHARS = 3000


def normalize_stem(stem: str) -> str:
    s = stem.lower()
    s = re.sub(r"[\s\-_()\[\].,·~!@#$%^&*+=]", "", s)
    return s


def name_similarity(a: str, b: str) -> float:
    na, nb = normalize_stem(a), normalize_stem(b)
    if not na or not nb:
        return 0.0
    return difflib.SequenceMatcher(None, na, nb).ratio()


def extract_pdf_text(path: Path, max_pages: int = 5) -> str:
    import fitz

    text_parts = []
    with fitz.open(path) as doc:
        for i, page in enumerate(doc):
            if i >= max_pages:
                break
            text_parts.append(page.get_text())
    return "".join(text_parts)


def extract_hwpx_text(path: Path) -> str:
    from text_extract import extract_plain

    return extract_plain(str(path), include_tables=False)


def content_similarity(text_a: str, text_b: str) -> float:
    a = "".join(text_a.split())[:CONTENT_SAMPLE_CHARS]
    b = "".join(text_b.split())[:CONTENT_SAMPLE_CHARS]
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).quick_ratio()


def find_candidates(root: Path, recursive: bool):
    pattern = "**/*" if recursive else "*"
    all_files = list(root.glob(pattern))
    by_dir: dict[Path, list[Path]] = {}
    for p in all_files:
        if p.is_file():
            by_dir.setdefault(p.parent, []).append(p)

    candidates = []
    for directory, files in by_dir.items():
        pdfs = [f for f in files if f.suffix.lower() == ".pdf"]
        docs = [f for f in files if f.suffix.lower() in (".hwp", ".hwpx")]
        if not pdfs or not docs:
            continue
        for pdf in pdfs:
            best = None
            for doc in docs:
                score = name_similarity(pdf.stem, doc.stem)
                if score >= NAME_THRESHOLD and (best is None or score > best[1]):
                    best = (doc, score)
            if best:
                candidates.append((pdf, best[0], best[1]))
    return candidates


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Find PDF files that look like duplicates of a HWP/HWPX file in the same folder"
    )
    parser.add_argument("input_dir", help="Folder to scan")
    parser.add_argument("--recursive", "-r", action="store_true", help="Scan subfolders as well")
    parser.add_argument("--output", "-o", default="pdf_duplicate_candidates.md", help="Report output path")
    parser.add_argument(
        "--skip-content-check",
        action="store_true",
        help="Skip text-extraction confirmation (filename match only, faster but less reliable)",
    )
    args = parser.parse_args()

    root = Path(args.input_dir)
    if not root.is_dir():
        print(f"Error: Folder not found: {root}", file=sys.stderr)
        sys.exit(1)

    print("Scanning for filename-based candidates...", file=sys.stderr)
    candidates = find_candidates(root, args.recursive)
    print(f"Found {len(candidates)} filename-based candidate pairs.", file=sys.stderr)

    rows = []
    for i, (pdf, doc, name_score) in enumerate(candidates, 1):
        content_score = None
        error = None
        if not args.skip_content_check:
            try:
                pdf_text = extract_pdf_text(pdf)
                doc_text = extract_hwpx_text(doc) if doc.suffix.lower() == ".hwpx" else None
                if doc_text is not None:
                    content_score = content_similarity(pdf_text, doc_text)
                if not pdf_text.strip():
                    error = "PDF에서 텍스트 추출 안 됨 (스캔본/이미지 PDF로 추정)"
            except Exception as exc:
                error = f"비교 실패: {exc}"
        rows.append((pdf, doc, name_score, content_score, error))
        print(f"[{i}/{len(candidates)}] {pdf.name}", flush=True, file=sys.stderr)

    def sort_key(row):
        _, _, name_score, content_score, _ = row
        return content_score if content_score is not None else name_score

    rows.sort(key=sort_key, reverse=True)

    lines = ["# PDF 중복 후보 리포트", "", f"스캔 대상: {root}", f"후보 쌍: {len(rows)}개", "",
             "자동 삭제는 하지 않음. 검토 후 확정된 항목만 수동으로 archive/삭제 처리할 것.", ""]
    for pdf, doc, name_score, content_score, error in rows:
        lines.append(f"## {pdf.name}")
        lines.append(f"- PDF: `{pdf}`")
        lines.append(f"- 후보 원본: `{doc}`")
        lines.append(f"- 파일명 유사도: {name_score:.0%}")
        if content_score is not None:
            lines.append(f"- 본문 유사도(앞부분 샘플 기준): {content_score:.0%}")
        if error:
            lines.append(f"- 비고: {error}")
        lines.append("")

    Path(args.output).write_text("\n".join(lines), encoding="utf-8")
    print(f"\n리포트 작성 완료: {args.output}", file=sys.stderr)


if __name__ == "__main__":
    main()
