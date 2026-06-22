#!/usr/bin/env python3
"""Run the full document pipeline on a folder, on demand.

Chains every step built for the predecessor-handover backup conversion:
  1. hwp -> hwpx (Hancom Office automation)
  2. validate every hwpx
  3. archive hwp originals whose hwpx passed validation (never deletes)
  4. hwpx -> md
  5. xlsx/xls -> md
  6. detect pdf duplicates of a hwp/hwpx in the same folder; auto-exclude
     high-confidence (>=80% content similarity) matches from conversion
  7. pdf -> md (with OCR fallback for scanned pages)

Designed to be re-run safely on the same folder as new files show up --
every step skips work it already did (--skip-existing / archive-aware).

Usage:
    python run_pipeline.py "C:\\Users\\User\\Desktop\\업무자료" --vault-dir "C:\\Users\\User\\Documents\\dh llm wiki\\raw\\업무-인수인계"
"""

import argparse
import subprocess
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent


def run_step(name: str, args: list[str]) -> None:
    print(f"\n{'=' * 10} {name} {'=' * 10}", flush=True)
    result = subprocess.run([sys.executable, *args])
    if result.returncode not in (0, 1):
        # 0 = clean, 1 = "Done, but some files failed" (already logged) -- both OK to continue.
        print(f"WARNING: {name} exited with code {result.returncode}", file=sys.stderr, flush=True)


def archive_converted_hwp(target_dir: Path) -> None:
    sys.path.insert(0, str(SCRIPTS_DIR))
    from validate import validate
    import shutil

    archive_dir = target_dir.parent / (target_dir.name + "_hwp원본_archive")
    hwp_files = sorted(p for p in target_dir.rglob("*.hwp") if p.suffix.lower() == ".hwp")
    moved = 0
    for hwp in hwp_files:
        hwpx = hwp.with_suffix(".hwpx")
        if not hwpx.exists():
            continue
        if validate(str(hwpx)):
            continue  # invalid hwpx, leave the hwp alone for inspection
        rel = hwp.relative_to(target_dir)
        dest = archive_dir / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(hwp), str(dest))
        moved += 1
    print(f"Archived {moved} validated hwp originals -> {archive_dir}", flush=True)


def build_pdf_exclude_list(target_dir: Path, report_path: Path) -> Path:
    sys.path.insert(0, str(SCRIPTS_DIR))
    from find_pdf_duplicates import find_candidates, extract_pdf_text, extract_hwpx_text, content_similarity

    candidates = find_candidates(target_dir, recursive=True)
    excludes = []
    rows = []
    for pdf, doc, name_score in candidates:
        score = None
        if doc.suffix.lower() == ".hwpx":
            try:
                score = content_similarity(extract_pdf_text(pdf), extract_hwpx_text(doc))
            except Exception:
                score = None
        if score is not None and score >= 0.8:
            excludes.append(str(pdf))
        rows.append((pdf, doc, name_score, score))

    report_path.write_text(
        "\n".join(
            f"{p} | candidate={d} | name={n:.0%} | content={'' if s is None else f'{s:.0%}'}"
            for p, d, n, s in rows
        ),
        encoding="utf-8",
    )
    exclude_list_path = report_path.with_suffix(".exclude.txt")
    exclude_list_path.write_text("\n".join(excludes), encoding="utf-8")
    print(f"PDF dup candidates: {len(rows)}, auto-excluded (>=80% content match): {len(excludes)}", flush=True)
    return exclude_list_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the full hwp/excel/pdf -> Markdown pipeline on a folder")
    parser.add_argument("target_dir", help="Folder to process (scanned recursively)")
    parser.add_argument("--vault-dir", required=True, help="Output folder for all generated .md files")
    parser.add_argument(
        "--skip-hwp",
        action="store_true",
        help="Skip the hwp->hwpx conversion + archive step (e.g. while deferring a large one-time backlog)",
    )
    args = parser.parse_args()

    target_dir = Path(args.target_dir)
    vault_dir = Path(args.vault_dir)
    if not target_dir.is_dir():
        print(f"Error: folder not found: {target_dir}", file=sys.stderr)
        sys.exit(1)
    vault_dir.mkdir(parents=True, exist_ok=True)

    if args.skip_hwp:
        print("Skipping hwp -> hwpx (--skip-hwp)", flush=True)
    else:
        run_step("1. hwp -> hwpx", [
            str(SCRIPTS_DIR / "convert_hwp_to_hwpx.py"), str(target_dir), "--recursive", "--skip-existing",
        ])

        print(f"\n{'=' * 10} 2-3. validate + archive hwp originals {'=' * 10}", flush=True)
        archive_converted_hwp(target_dir)

    run_step("4. hwpx -> md", [
        str(SCRIPTS_DIR / "batch_extract.py"), str(target_dir), "--recursive", "--skip-existing",
        "--format", "markdown", "--output-dir", str(vault_dir),
    ])

    run_step("5. excel -> md", [
        str(SCRIPTS_DIR / "excel_extract.py"), str(target_dir), "--recursive", "--skip-existing",
        "--output-dir", str(vault_dir),
    ])

    print(f"\n{'=' * 10} 6. pdf duplicate detection {'=' * 10}", flush=True)
    report_path = target_dir.parent / (target_dir.name + "_pdf중복후보검토.txt")
    exclude_list_path = build_pdf_exclude_list(target_dir, report_path)

    run_step("7. pdf -> md", [
        str(SCRIPTS_DIR / "pdf_extract.py"), str(target_dir), "--recursive", "--skip-existing",
        "--exclude-list", str(exclude_list_path), "--output-dir", str(vault_dir),
    ])

    print("\nPipeline run complete.", flush=True)


if __name__ == "__main__":
    main()
