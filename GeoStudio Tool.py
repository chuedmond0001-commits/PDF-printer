import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from pypdf import PdfReader, PdfWriter, Transformation
from pypdf._page import PageObject


SCRIPT_DIR = Path(__file__).resolve().parent
GEOCMD_PATH = Path(r"C:\Program Files\Seequent\GeoStudio 2025.1\Bin\GeoCmd.exe")
PRINT_SCRIPT = SCRIPT_DIR / "Print code.py"
BROWSER_CANDIDATES = [
    Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
    Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
    Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate GeoStudio reports and optional PDFs.")
    parser.add_argument("--project", required=True, help="GeoStudio .gsz project file.")
    parser.add_argument(
        "--analyses",
        default="",
        help="Optional comma-separated analysis names. Blank means all analyses.",
    )
    parser.add_argument("--output-folder", required=True, help="Folder for reports/PDFs.")
    parser.add_argument("--solve", action="store_true", help="Solve selected analyses before reporting.")
    parser.add_argument("--results", action="store_true", help="Include results reports.")
    parser.add_argument("--pdf", action="store_true", help="Convert generated HTML reports to PDF.")
    parser.add_argument(
        "--pages-per-sheet",
        type=int,
        choices=[1, 2, 4],
        default=2,
        help="Pages per sheet for generated report PDFs.",
    )
    parser.add_argument("--print-pdf", action="store_true", help="Print generated report PDFs after conversion.")
    parser.add_argument("--printer", default="", help="Printer for --print-pdf.")
    parser.add_argument("--print-pages", default="", help="Optional page range for printing generated PDFs.")
    parser.add_argument(
        "--print-backend",
        choices=["pdfxchange", "shell"],
        default="pdfxchange",
        help="Print generated PDFs through PDF-XChange or Windows shell printto.",
    )
    parser.add_argument("--print-script", default=str(PRINT_SCRIPT), help="Path to Print code.py.")
    parser.add_argument("--open-folder", action="store_true", help="Open output folder when finished.")
    parser.add_argument("--geocmd", default=str(GEOCMD_PATH), help="Path to GeoCmd.exe.")
    return parser


def parse_analyses(raw: str) -> list[str]:
    return [part.strip() for part in raw.split(",") if part.strip()]


def find_browser() -> Path | None:
    for candidate in BROWSER_CANDIDATES:
        if candidate.exists():
            return candidate
    return None


def find_new_html_files(folder: Path, before: set[Path]) -> list[Path]:
    after = {path.resolve() for path in folder.rglob("*.htm*") if path.is_file()}
    return sorted(after - before, key=lambda path: str(path).lower())


def snapshot_html_files(folders: list[Path]) -> dict[Path, tuple[int, int]]:
    snapshot = {}
    for folder in folders:
        if not folder.exists():
            continue
        for path in folder.glob("*.htm*"):
            if path.is_file():
                stat = path.stat()
                snapshot[path.resolve()] = (stat.st_mtime_ns, stat.st_size)
    return snapshot


def changed_html_files(folders: list[Path], before: dict[Path, tuple[int, int]]) -> list[Path]:
    changed = []
    for folder in folders:
        if not folder.exists():
            continue
        for path in folder.glob("*.htm*"):
            if not path.is_file():
                continue
            resolved = path.resolve()
            stat = path.stat()
            current = (stat.st_mtime_ns, stat.st_size)
            if before.get(resolved) != current:
                changed.append(resolved)
    return sorted(set(changed), key=lambda path: str(path).lower())


def collect_reports(html_files: list[Path], output_folder: Path) -> list[Path]:
    collected = []
    for html in html_files:
        html = html.resolve()
        destination = output_folder / html.name
        if html.parent.resolve() != output_folder.resolve():
            shutil.copy2(html, destination)
            collected.append(destination.resolve())
        else:
            collected.append(html)
    return sorted(set(collected), key=lambda path: str(path).lower())


def html_to_pdf(browser: Path, html_path: Path, output_pdf: Path) -> None:
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    command = [
        str(browser),
        "--headless",
        "--disable-gpu",
        f"--print-to-pdf={output_pdf}",
        str(html_path.resolve().as_uri()),
    ]
    result = subprocess.run(command, capture_output=True, text=True, timeout=120)
    if result.returncode != 0:
        details = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(f"HTML to PDF failed for {html_path.name}: {details}")


def nup_pdf(input_pdf: Path, output_pdf: Path, pages_per_sheet: int) -> None:
    if pages_per_sheet == 1:
        if input_pdf != output_pdf:
            output_pdf.write_bytes(input_pdf.read_bytes())
        return

    reader = PdfReader(str(input_pdf))
    writer = PdfWriter()
    if not reader.pages:
        return

    first_page = reader.pages[0]
    source_width = float(first_page.mediabox.width)
    source_height = float(first_page.mediabox.height)
    if pages_per_sheet == 2:
        sheet_width = max(source_width, source_height)
        sheet_height = min(source_width, source_height)
        columns = 2
        rows = 1
    else:
        sheet_width = source_width
        sheet_height = source_height
        columns = 2
        rows = 2

    cell_width = sheet_width / columns
    cell_height = sheet_height / rows

    for start in range(0, len(reader.pages), pages_per_sheet):
        sheet = PageObject.create_blank_page(width=sheet_width, height=sheet_height)
        for offset, page in enumerate(reader.pages[start : start + pages_per_sheet]):
            page_width = float(page.mediabox.width)
            page_height = float(page.mediabox.height)
            scale = min(cell_width / page_width, cell_height / page_height)
            scaled_width = page_width * scale
            scaled_height = page_height * scale

            column = offset % columns
            row = offset // columns
            x = column * cell_width + (cell_width - scaled_width) / 2
            y = sheet_height - ((row + 1) * cell_height) + (cell_height - scaled_height) / 2

            transform = Transformation().scale(scale).translate(tx=x, ty=y)
            sheet.merge_transformed_page(page, transform)

        writer.add_page(sheet)

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with output_pdf.open("wb") as handle:
        writer.write(handle)


def print_generated_pdfs(args: argparse.Namespace, pdf_files: list[Path]) -> int:
    printer = args.printer.strip()
    if not printer:
        print("ERROR: Choose a printer before printing GeoStudio PDFs.")
        return 1
    if not pdf_files:
        print("ERROR: No generated PDFs were found to print.")
        return 1

    print_script = Path(args.print_script.strip('"')).expanduser().resolve()
    if not print_script.exists():
        print(f"ERROR: Print script not found: {print_script}")
        return 1

    temp_list = None
    try:
        with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json", delete=False) as handle:
            json.dump([str(path) for path in pdf_files], handle)
            temp_list = Path(handle.name)

        command = [
            sys.executable,
            str(print_script),
            "--file-list",
            str(temp_list),
            "--mode",
            "pdf",
            "--printer",
            printer,
            "--backend",
            args.print_backend,
            "--yes",
        ]
        if args.print_pages.strip():
            command.extend(["--pages", args.print_pages.strip()])

        print("\nPrinting generated PDFs...")
        result = subprocess.run(command, capture_output=True, text=True, timeout=None)
        if result.stdout:
            print(result.stdout.strip())
        if result.stderr:
            print(result.stderr.strip())
        if result.returncode != 0:
            print(f"ERROR: Print task failed with exit code {result.returncode}")
        return result.returncode
    finally:
        if temp_list is not None:
            try:
                temp_list.unlink(missing_ok=True)
            except OSError:
                pass


def run_geocmd(args: argparse.Namespace) -> int:
    project = Path(args.project.strip('"')).expanduser().resolve()
    if not project.exists() or not project.is_file():
        print(f"ERROR: GeoStudio project not found: {project}")
        return 1
    if project.suffix.lower() != ".gsz":
        print("ERROR: Project must be a .gsz file.")
        return 1

    output_folder = Path(args.output_folder.strip('"')).expanduser().resolve()
    output_folder.mkdir(parents=True, exist_ok=True)

    geocmd = Path(args.geocmd.strip('"')).expanduser().resolve()
    if not geocmd.exists():
        print(f"ERROR: GeoCmd.exe not found: {geocmd}")
        return 1
    if args.print_pdf and not args.pdf:
        print("Printing GeoStudio reports needs PDF conversion, so PDF output has been enabled.")
        args.pdf = True

    command = [str(geocmd), str(project)]
    command.extend(parse_analyses(args.analyses))
    if args.solve:
        command.append("/solve")
    command.append("/report")
    if args.results:
        command.append("/results")

    html_search_folders = [output_folder, project.parent]
    before_html = snapshot_html_files(html_search_folders)

    print("\nGeoStudio report settings:")
    print(f"Project : {project}")
    print(f"Output  : {output_folder}")
    print(f"Analyses: {args.analyses or 'all'}")
    print(f"Solve   : {'yes' if args.solve else 'no'}")
    print(f"Results : {'yes' if args.results else 'no'}")
    print(f"PDF     : {'yes' if args.pdf else 'no'}")
    if args.pdf:
        print(f"N-up PDF: {args.pages_per_sheet} page(s) per sheet")
    print("\nRunning GeoCmd...")
    print(" ".join(f'"{part}"' if " " in part else part for part in command))

    result = subprocess.run(command, cwd=output_folder, capture_output=True, text=True, timeout=None)
    if result.stdout:
        print(result.stdout.strip())
    if result.stderr:
        print(result.stderr.strip())
    if result.returncode != 0:
        print(f"ERROR: GeoCmd failed with exit code {result.returncode}")
        return result.returncode

    time.sleep(1)
    html_files = collect_reports(changed_html_files(html_search_folders, before_html), output_folder)
    if not html_files:
        html_files = sorted(output_folder.rglob("*.htm*"), key=lambda path: str(path).lower())

    print("\nHTML reports:")
    if html_files:
        for html in html_files:
            print(f"- {html}")
    else:
        print("- No HTML reports found. GeoCmd may have written reports beside the project file.")

    if args.pdf:
        browser = find_browser()
        if browser is None:
            print("ERROR: Could not find Edge or Chrome for HTML-to-PDF conversion.")
            return 1

        pdf_folder = output_folder / "pdf"
        pdf_files = []
        print("\nConverting HTML reports to PDF...")
        for html in html_files:
            output_pdf = pdf_folder / f"{html.stem}.pdf"
            if args.pages_per_sheet == 1:
                html_to_pdf(browser, html, output_pdf)
            else:
                raw_pdf = pdf_folder / f"{html.stem}.__raw.pdf"
                html_to_pdf(browser, html, raw_pdf)
                nup_pdf(raw_pdf, output_pdf, args.pages_per_sheet)
                raw_pdf.unlink(missing_ok=True)
            pdf_files.append(output_pdf)
            print(f"- {output_pdf}")
    else:
        pdf_files = []

    if args.print_pdf:
        print_result = print_generated_pdfs(args, pdf_files)
        if print_result != 0:
            return print_result

    if args.open_folder:
        try:
            os.startfile(str(output_folder))
        except Exception as exc:
            print(f"WARNING: Could not open output folder: {exc}")

    print("\nSummary:")
    print("GeoStudio report task complete.")
    return 0


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        return run_geocmd(args)
    except Exception as exc:
        print(f"ERROR: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
