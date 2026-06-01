import argparse
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path


def safe_stem(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", value).strip("._") or "page"


def parse_ranges(raw: str, page_count: int) -> list[tuple[int, int]]:
    if not raw.strip():
        return [(page, page) for page in range(1, page_count + 1)]

    ranges = []
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_raw, end_raw = part.split("-", 1)
            if not start_raw.strip().isdigit() or not end_raw.strip().isdigit():
                raise ValueError(f"Invalid page range: {part}")
            start, end = int(start_raw), int(end_raw)
        else:
            if not part.isdigit():
                raise ValueError(f"Invalid page number: {part}")
            start = end = int(part)

        if start < 1 or end < 1 or start > page_count or end > page_count or start > end:
            raise ValueError(f"Page range out of bounds: {part}")
        ranges.append((start, end))

    if not ranges:
        raise ValueError("No valid page ranges provided.")
    return ranges


@dataclass
class SplitOptions:
    input_pdf: Path
    output_folder: Path
    ranges: str
    yes: bool
    open_folder: bool


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Split one PDF into separate PDF files.")
    parser.add_argument("--input", required=True, help="PDF file to split.")
    parser.add_argument(
        "--output-folder",
        default="",
        help="Folder where split PDFs will be saved. Defaults to the input PDF folder.",
    )
    parser.add_argument(
        "--ranges",
        default="",
        help="Optional page ranges like 1-3,4,5-7. Blank means one PDF per page.",
    )
    parser.add_argument("--yes", action="store_true", help="Skip confirmation.")
    parser.add_argument("--open-folder", action="store_true", help="Open output folder when finished.")
    return parser


def resolve_options(args: argparse.Namespace) -> SplitOptions:
    input_pdf = Path(args.input.strip('"')).expanduser().resolve()
    if not input_pdf.exists() or not input_pdf.is_file():
        raise FileNotFoundError(f"PDF not found: {input_pdf}")
    if input_pdf.suffix.lower() != ".pdf":
        raise ValueError("Input file must be a PDF.")

    output_folder_raw = args.output_folder.strip('"')
    output_folder = (
        Path(output_folder_raw).expanduser().resolve()
        if output_folder_raw
        else input_pdf.parent
    )
    output_folder.mkdir(parents=True, exist_ok=True)

    return SplitOptions(
        input_pdf=input_pdf,
        output_folder=output_folder,
        ranges=args.ranges,
        yes=args.yes,
        open_folder=args.open_folder,
    )


def output_path_for(input_pdf: Path, output_folder: Path, start: int, end: int) -> Path:
    stem = safe_stem(input_pdf.stem)
    label = f"p{start:03d}" if start == end else f"p{start:03d}-{end:03d}"
    candidate = output_folder / f"{stem}_{label}.pdf"
    if not candidate.exists():
        return candidate

    counter = 2
    while True:
        candidate = output_folder / f"{stem}_{label}_{counter}.pdf"
        if not candidate.exists():
            return candidate
        counter += 1


def split_pdf(options: SplitOptions) -> int:
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        print("ERROR: PDF splitting needs pypdf. Run: venv\\Scripts\\python.exe -m pip install pypdf")
        return 1

    reader = PdfReader(str(options.input_pdf))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:
            pass
        if reader.is_encrypted:
            print("ERROR: PDF is encrypted and could not be opened.")
            return 1

    page_count = len(reader.pages)
    ranges = parse_ranges(options.ranges, page_count)

    print("\nSplit settings:")
    print(f"Input : {options.input_pdf}")
    print(f"Output: {options.output_folder}")
    print(f"Pages : {page_count}")
    print(f"Ranges: {options.ranges or 'one file per page'}")

    if not options.yes:
        proceed = input("\nCreate split PDF files? [Y/n]: ").strip().lower()
        if proceed in {"n", "no"}:
            print("Cancelled.")
            return 0

    created = []
    print("\nCreating split PDFs...")
    for index, (start, end) in enumerate(ranges, start=1):
        writer = PdfWriter()
        for page_number in range(start, end + 1):
            writer.add_page(reader.pages[page_number - 1])

        output_pdf = output_path_for(options.input_pdf, options.output_folder, start, end)
        with output_pdf.open("wb") as handle:
            writer.write(handle)
        writer.close()
        created.append(output_pdf)
        label = f"page {start}" if start == end else f"pages {start}-{end}"
        print(f"[{index}/{len(ranges)}] OK: {label} -> {output_pdf.name}")

    print("\nSummary:")
    print(f"Created: {len(created)}")
    print(f"Output : {options.output_folder}")

    if options.open_folder:
        try:
            os.startfile(str(options.output_folder))
        except Exception as exc:
            print(f"WARNING: Could not open output folder: {exc}")

    return 0


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        options = resolve_options(args)
        return split_pdf(options)
    except Exception as exc:
        print(f"ERROR: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
