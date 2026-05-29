import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import win32api
import win32com.client
import win32print


SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "print_config.json"

DEFAULT_CONFIG = {
    "pdfxedit_path": r"C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe",
    "default_folder": "",
    "default_printer": "",
    "duplex": "none",
    "color": "yes",
    "mode": "all",
    "delay_seconds": 2.0,
    "recursive": False,
    "temp_dir": "",
    "show_print_dialog": False,
    "pages": "",
    "backend": "pdfxchange",
    "keep_temp_pdfs": True,
}

MODE_EXTENSIONS = {
    "all": (".docx", ".doc", ".rtf", ".xlsx", ".xls", ".pdf"),
    "pdf": (".pdf",),
    "word": (".docx", ".doc", ".rtf"),
    "excel": (".xlsx", ".xls"),
}

DUPLEX_CHOICES = {
    "1": "none",
    "2": "vertical",
    "3": "horizontal",
    "none": "none",
    "single": "none",
    "simplex": "none",
    "vertical": "vertical",
    "long": "vertical",
    "long-edge": "vertical",
    "horizontal": "horizontal",
    "short": "horizontal",
    "short-edge": "horizontal",
}

COLOR_CHOICES = {
    "1": "yes",
    "yes": "yes",
    "y": "yes",
    "color": "yes",
    "colour": "yes",
    "2": "no",
    "no": "no",
    "n": "no",
    "bw": "no",
    "black": "no",
    "grayscale": "no",
    "greyscale": "no",
}


@dataclass
class RunOptions:
    folder: Path
    selected_files: list[Path] | None
    printer: str
    duplex: str
    color: str
    mode: str
    delay_seconds: float
    recursive: bool
    dry_run: bool
    yes: bool
    temp_dir: Path
    pdfxedit_path: Path
    exclude_indices: set[int]
    show_print_dialog: bool
    pages: str
    backend: str
    keep_temp_pdfs: bool
    save_pdf_folder: Path | None


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        return DEFAULT_CONFIG.copy()

    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError) as exc:
        print(f"WARNING: Could not read {CONFIG_PATH.name}: {exc}")
        return DEFAULT_CONFIG.copy()

    config = DEFAULT_CONFIG.copy()
    config.update({key: value for key, value in data.items() if key in config})
    return config


def save_config(config: dict) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as handle:
        json.dump(config, handle, indent=2)
        handle.write("\n")
    print(f"Saved defaults to {CONFIG_PATH}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert Word/Excel/RTF files to PDF and send PDFs to a printer."
    )
    parser.add_argument("--folder", help="Folder that contains files to print.")
    parser.add_argument(
        "--file-list",
        help="JSON file containing exact files to print. Used by the local web app.",
    )
    parser.add_argument(
        "--pick-files",
        action="store_true",
        help="Open a Windows file picker and print only the selected files.",
    )
    parser.add_argument("--printer", help="Printer name. Use --list-printers to see choices.")
    parser.add_argument(
        "--duplex",
        choices=["none", "vertical", "horizontal"],
        help="none, vertical/long-edge, or horizontal/short-edge.",
    )
    parser.add_argument(
        "--color",
        choices=["yes", "no"],
        help="yes for color, no for grayscale/black-and-white.",
    )
    parser.add_argument(
        "--mode",
        choices=sorted(MODE_EXTENSIONS),
        help="File type group to print.",
    )
    parser.add_argument("--delay", type=float, help="Seconds to wait after each print job.")
    parser.add_argument("--recursive", action="store_true", help="Include files in subfolders.")
    parser.add_argument("--dry-run", action="store_true", help="Preview the run without converting or printing.")
    parser.add_argument("--yes", action="store_true", help="Skip final confirmation.")
    parser.add_argument("--exclude", default="", help="Comma-separated file numbers to exclude, such as 1,3,5.")
    parser.add_argument("--list-printers", action="store_true", help="Show available printers and exit.")
    parser.add_argument("--save-defaults", action="store_true", help="Save provided options to print_config.json.")
    parser.add_argument("--temp-dir", help="Folder for temporary converted PDFs.")
    parser.add_argument("--pdfxedit", help="Path to PDF-XChange Editor executable.")
    parser.add_argument(
        "--print-dialog",
        action="store_true",
        help="Show PDF-XChange print dialog for each PDF before printing.",
    )
    parser.add_argument("--pages", help="Page range to print, such as 1-3,5.")
    parser.add_argument(
        "--backend",
        choices=["pdfxchange", "shell"],
        help="Print through PDF-XChange or Windows shell printto.",
    )
    parser.add_argument(
        "--keep-temp-pdfs",
        action="store_true",
        help="Keep converted temporary PDFs for troubleshooting.",
    )
    parser.add_argument(
        "--save-pdf-folder",
        help="Save converted/copied PDFs to this folder instead of sending them to a printer.",
    )
    return parser


def get_printers() -> list[str]:
    printers = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )
    return sorted(p[2] for p in printers)


def normalize_duplex(value: str | None, default: str) -> str:
    raw = (value or default or "none").strip().lower()
    return DUPLEX_CHOICES.get(raw, "none")


def normalize_color(value: str | None, default: str) -> str:
    raw = (value or default or "yes").strip().lower()
    return COLOR_CHOICES.get(raw, "yes")


def prompt_with_default(message: str, default: str = "") -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{message}{suffix}: ").strip().strip('"')
    return value or default


def choose_printer(printers: list[str], default_printer: str = "") -> str:
    print("\nAvailable printers:")
    for index, name in enumerate(printers, start=1):
        marker = " (default)" if default_printer and name == default_printer else ""
        print(f"[{index}] {name}{marker}")

    while True:
        choice = prompt_with_default("Choose printer number or exact name", default_printer)
        if not choice:
            print("Printer is required.")
            continue

        if choice.isdigit():
            selected_index = int(choice) - 1
            if 0 <= selected_index < len(printers):
                return printers[selected_index]

        matches = [name for name in printers if choice.lower() in name.lower()]
        if len(matches) == 1:
            return matches[0]
        if choice in printers:
            return choice

        print("Could not match that printer. Try the number or paste the full name.")


def choose_mode(default_mode: str) -> str:
    options = [("all", "All supported files"), ("pdf", "PDF only"), ("word", "Word/RTF only"), ("excel", "Excel only")]
    print("\nFile type mode:")
    for index, (mode, label) in enumerate(options, start=1):
        marker = " (default)" if mode == default_mode else ""
        print(f"[{index}] {label}{marker}")

    choice = prompt_with_default("Choose mode number or name", default_mode).lower()
    if choice.isdigit() and 1 <= int(choice) <= len(options):
        return options[int(choice) - 1][0]
    return choice if choice in MODE_EXTENSIONS else default_mode


def choose_yes_no(message: str, default: bool) -> bool:
    default_text = "yes" if default else "no"
    value = prompt_with_default(message, default_text).strip().lower()
    return value in {"1", "y", "yes", "true"}


def parse_exclude_indices(raw: str) -> set[int]:
    indices = set()
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        if not part.isdigit():
            raise ValueError(f"Invalid exclude index: {part}")
        indices.add(int(part) - 1)
    return indices


def pick_files_with_dialog(initial_dir: str = "") -> list[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:
        raise RuntimeError(f"Could not open Windows file picker: {exc}") from exc

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_types = [
        ("Supported files", "*.pdf *.docx *.doc *.rtf *.xlsx *.xls"),
        ("PDF files", "*.pdf"),
        ("Word/RTF files", "*.docx *.doc *.rtf"),
        ("Excel files", "*.xlsx *.xls"),
        ("All files", "*.*"),
    ]
    selected = filedialog.askopenfilenames(
        title="Choose files to print",
        initialdir=initial_dir or str(Path.home()),
        filetypes=file_types,
    )
    root.destroy()
    return [Path(path).resolve() for path in selected]


def common_parent(paths: list[Path]) -> Path:
    if not paths:
        return Path.cwd()
    return Path(os.path.commonpath([str(path.parent) for path in paths]))


def resolve_options(args: argparse.Namespace, config: dict, printers: list[str]) -> RunOptions:
    selected_files = None
    initial_dir = args.folder or config.get("default_folder") or ""

    if args.file_list:
        list_path = Path(args.file_list.strip('"')).expanduser().resolve()
        with list_path.open("r", encoding="utf-8") as handle:
            selected_files = [Path(path).expanduser().resolve() for path in json.load(handle)]
        if not selected_files:
            raise RuntimeError("No files selected.")
        folder = common_parent(selected_files)
    elif args.pick_files or not args.folder:
        selected_files = pick_files_with_dialog(initial_dir)
        if not selected_files:
            raise RuntimeError("No files selected.")
        folder = common_parent(selected_files)
    else:
        folder = Path(args.folder.strip('"')).expanduser().resolve()
        if not folder.exists() or not folder.is_dir():
            raise FileNotFoundError(f"Folder not found: {folder}")

    printer = args.printer or config.get("default_printer") or ""
    if not printer:
        printer = choose_printer(printers, "")
    elif printer not in printers:
        matches = [name for name in printers if printer.lower() in name.lower()]
        if len(matches) == 1:
            printer = matches[0]
        else:
            print(f"Configured printer was not found: {printer}")
            printer = choose_printer(printers, config.get("default_printer", ""))

    duplex = normalize_duplex(args.duplex, config.get("duplex", "none"))
    color = normalize_color(args.color, config.get("color", "yes"))
    mode = args.mode or config.get("mode") or "all"
    if mode not in MODE_EXTENSIONS:
        mode = "all"
    show_print_dialog = bool(args.print_dialog or config.get("show_print_dialog", False))
    pages = (args.pages or config.get("pages") or "").strip()
    backend = args.backend or config.get("backend") or "pdfxchange"
    if backend not in {"pdfxchange", "shell"}:
        backend = "pdfxchange"
    keep_temp_pdfs = bool(args.keep_temp_pdfs or config.get("keep_temp_pdfs", False))

    if not (args.duplex or args.color or args.mode or args.yes):
        print("\nPrint preferences:")
        backend = prompt_with_default("Print engine: pdfxchange or shell", backend).lower()
        if backend not in {"pdfxchange", "shell"}:
            backend = "pdfxchange"
        show_print_dialog = choose_yes_no(
            "Open print dialog for each file? Use this for duplex/color/copies/page setup",
            show_print_dialog,
        )
        duplex = normalize_duplex(
            prompt_with_default("Duplex: 1=single, 2=long-edge, 3=short-edge", duplex),
            duplex,
        )
        color = normalize_color(prompt_with_default("Color: 1=color, 2=black/white", color), color)
        mode = choose_mode(mode)
        pages = prompt_with_default("Pages to print, blank means all pages", pages)

    delay_seconds = args.delay if args.delay is not None else float(config.get("delay_seconds", 2.0))
    recursive = bool(args.recursive or config.get("recursive", False))

    temp_dir_raw = args.temp_dir or config.get("temp_dir") or tempfile.gettempdir()
    temp_dir = Path(temp_dir_raw).expanduser().resolve() / "office_pdf_print"
    temp_dir.mkdir(parents=True, exist_ok=True)

    pdfxedit_path = Path(args.pdfxedit or config.get("pdfxedit_path") or DEFAULT_CONFIG["pdfxedit_path"])
    exclude_indices = parse_exclude_indices(args.exclude)
    save_pdf_folder = None
    if args.save_pdf_folder:
        save_pdf_folder = Path(args.save_pdf_folder.strip('"')).expanduser().resolve()
        save_pdf_folder.mkdir(parents=True, exist_ok=True)

    return RunOptions(
        folder=folder,
        selected_files=selected_files,
        printer=printer,
        duplex=duplex,
        color=color,
        mode=mode,
        delay_seconds=max(delay_seconds, 0),
        recursive=recursive,
        dry_run=args.dry_run,
        yes=args.yes,
        temp_dir=temp_dir,
        pdfxedit_path=pdfxedit_path,
        exclude_indices=exclude_indices,
        show_print_dialog=show_print_dialog,
        pages=pages,
        backend=backend,
        keep_temp_pdfs=keep_temp_pdfs,
        save_pdf_folder=save_pdf_folder,
    )


def find_files(folder: Path, mode: str, recursive: bool) -> list[Path]:
    valid_exts = MODE_EXTENSIONS[mode]
    entries = folder.rglob("*") if recursive else folder.iterdir()
    files = [
        path
        for path in entries
        if path.is_file()
        and path.suffix.lower() in valid_exts
        and not path.name.startswith("~$")
        and not path.name.endswith("_temp_print.pdf")
    ]
    return sorted(files, key=lambda path: str(path).lower())


def filter_selected_files(files: list[Path], mode: str) -> list[Path]:
    valid_exts = MODE_EXTENSIONS[mode]
    return sorted(
        [
            path
            for path in files
            if path.is_file()
            and path.suffix.lower() in valid_exts
            and not path.name.startswith("~$")
            and not path.name.endswith("_temp_print.pdf")
        ],
        key=lambda path: str(path).lower(),
    )


def safe_temp_pdf_path(source: Path, temp_dir: Path) -> Path:
    safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", source.stem).strip("._") or "converted"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return temp_dir / f"{safe_stem}_{stamp}.pdf"


def unique_pdf_output_path(source: Path, output_folder: Path) -> Path:
    stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", source.stem).strip("._") or "output"
    candidate = output_folder / f"{stem}.pdf"
    if not candidate.exists():
        return candidate

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return output_folder / f"{stem}_{stamp}.pdf"


def print_file_list(files: list[Path], folder: Path) -> None:
    print("\nFiles to process:")
    for index, path in enumerate(files, start=1):
        try:
            label = path.relative_to(folder)
        except ValueError:
            label = path
        print(f"[{index}] {label}")


def confirm_run(options: RunOptions, files: list[Path]) -> bool:
    print_file_list(files, options.folder)
    print("\nSettings:")
    print(f"Folder   : {options.folder}")
    print(f"Printer  : {options.printer}")
    print(f"Mode     : {options.mode}")
    print(f"Duplex   : {options.duplex}")
    print(f"Color    : {options.color}")
    print(f"Backend  : {options.backend}")
    if options.save_pdf_folder:
        print(f"Save PDF : {options.save_pdf_folder}")
    print(f"Dialog   : {'yes' if options.show_print_dialog else 'no'}")
    print(f"Pages    : {options.pages or 'all'}")
    print(f"Delay    : {options.delay_seconds:g}s")
    print(f"Recursive: {'yes' if options.recursive else 'no'}")
    if options.dry_run:
        print("Dry run  : yes")

    if options.yes or options.dry_run:
        return True
    return input("\nStart printing? [Y/n]: ").strip().lower() not in {"n", "no"}


def convert_office_to_pdf(path: Path, temp_pdf: Path, word_app, excel_app):
    ext = path.suffix.lower()
    if ext in (".docx", ".doc", ".rtf"):
        if word_app is None:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
        doc = word_app.Documents.Open(str(path))
        try:
            doc.ExportAsFixedFormat(str(temp_pdf), 17)
        finally:
            doc.Close(False)
        return temp_pdf, word_app, excel_app

    if ext in (".xlsx", ".xls"):
        if excel_app is None:
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
        workbook = excel_app.Workbooks.Open(str(path))
        try:
            workbook.ExportAsFixedFormat(0, str(temp_pdf))
        finally:
            workbook.Close(False)
        return temp_pdf, word_app, excel_app

    return path, word_app, excel_app


def submit_pdf_to_printer(pdf_path: Path, options: RunOptions) -> None:
    if options.backend == "shell" or not options.pdfxedit_path.exists():
        win32api.ShellExecute(0, "printto", str(pdf_path), f'"{options.printer}"', ".", 0)
        return

    if options.backend == "pdfxchange":
        print_options = [
            "default=yes",
            f"showui={'yes' if options.show_print_dialog else 'no'}",
            f'printer="{options.printer}"',
        ]
        if options.pages:
            print_options.append(f"pages={options.pages}")
        command = [
            str(options.pdfxedit_path),
            f"/print:{';'.join(print_options)}",
            str(pdf_path),
        ]
        result = subprocess.run(
            command,
            shell=False,
            capture_output=True,
            text=True,
            timeout=None if options.show_print_dialog else 120,
        )
        if result.returncode != 0:
            details = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(
                f"PDF-XChange print command failed with exit code {result.returncode}. {details}"
            )
        return

    raise ValueError(f"Unsupported backend: {options.backend}")


def snapshot_print_jobs(printer: str) -> dict[int, tuple[str, int]]:
    handle = None
    try:
        handle = win32print.OpenPrinter(printer)
        jobs = win32print.EnumJobs(handle, 0, -1, 1)
    except Exception:
        return {}
    finally:
        if handle is not None:
            win32print.ClosePrinter(handle)

    snapshot = {}
    for job in jobs:
        job_id = job.get("JobId")
        status = str(job.get("Status", ""))
        size = int(job.get("Size", 0) or 0)
        if job_id is not None:
            snapshot[int(job_id)] = (status, size)
    return snapshot


def wait_for_submission(printer: str, before: dict[int, tuple[str, int]], timeout: float = 10.0) -> str:
    deadline = time.time() + timeout
    latest_message = "queue did not show a new job"

    while time.time() < deadline:
        after = snapshot_print_jobs(printer)
        new_ids = sorted(set(after) - set(before))
        if new_ids:
            details = []
            for job_id in new_ids:
                status, size = after[job_id]
                details.append(f"job {job_id}, status={status or 'unknown'}, size={size}")
            latest_message = "; ".join(details)
            if any(size > 0 for _, size in (after[job_id] for job_id in new_ids)):
                return latest_message
        time.sleep(0.5)

    return latest_message


def cleanup_file(path: Path) -> None:
    for _ in range(6):
        try:
            if path.exists():
                path.unlink()
            return
        except PermissionError:
            time.sleep(0.5)


def run_print_job(options: RunOptions, files: list[Path]) -> int:
    if options.dry_run:
        print("\nDry run complete. No files were converted or printed.")
        return 0

    original_default_printer = None
    word_app = None
    excel_app = None
    printed = []
    saved = []
    failed = []

    try:
        try:
            original_default_printer = win32print.GetDefaultPrinter()
            win32print.SetDefaultPrinter(options.printer)
        except Exception as exc:
            print(f"WARNING: Could not set default printer: {exc}")

        print("\nStarting print job...")
        for number, path in enumerate(files, start=1):
            temp_pdf = safe_temp_pdf_path(path, options.temp_dir)
            temp_created = False
            print(f"[{number}/{len(files)}] {path.name}")

            try:
                target_pdf, word_app, excel_app = convert_office_to_pdf(path, temp_pdf, word_app, excel_app)
                temp_created = target_pdf == temp_pdf

                if not target_pdf.exists():
                    raise FileNotFoundError(f"Converted PDF was not created: {target_pdf}")

                if options.save_pdf_folder:
                    output_pdf = unique_pdf_output_path(path, options.save_pdf_folder)
                    if target_pdf.resolve() != output_pdf.resolve():
                        shutil.copy2(target_pdf, output_pdf)
                    saved.append(output_pdf)
                    print(f"  OK: saved PDF to {output_pdf}")
                else:
                    queue_before = snapshot_print_jobs(options.printer)
                    submit_pdf_to_printer(target_pdf, options)
                    queue_message = wait_for_submission(options.printer, queue_before)
                    printed.append(path)
                    print(f"  OK: sent to printer ({queue_message})")
                if options.delay_seconds:
                    time.sleep(options.delay_seconds)
            except Exception as exc:
                failed.append((path, exc))
                print(f"  FAILED: {exc}")
            finally:
                if temp_created and options.keep_temp_pdfs:
                    print(f"  Kept temp PDF: {temp_pdf}")
                elif temp_created:
                    cleanup_file(temp_pdf)
    finally:
        if word_app is not None:
            word_app.Quit()
        if excel_app is not None:
            excel_app.Quit()
        if original_default_printer:
            try:
                win32print.SetDefaultPrinter(original_default_printer)
            except Exception:
                pass

    print("\nSummary:")
    print(f"Printed: {len(printed)}")
    if options.save_pdf_folder:
        print(f"Saved  : {len(saved)}")
    print(f"Failed : {len(failed)}")
    if failed:
        print("\nFailed files:")
        for path, exc in failed:
            print(f"- {path.name}: {exc}")
    return 1 if failed else 0


def update_config_from_args(args: argparse.Namespace, config: dict) -> dict:
    updated = config.copy()
    mappings = {
        "folder": "default_folder",
        "printer": "default_printer",
        "duplex": "duplex",
        "color": "color",
        "mode": "mode",
        "temp_dir": "temp_dir",
        "pdfxedit": "pdfxedit_path",
        "pages": "pages",
        "backend": "backend",
    }
    for arg_name, config_name in mappings.items():
        value = getattr(args, arg_name)
        if value:
            updated[config_name] = value
    if args.delay is not None:
        updated["delay_seconds"] = args.delay
    if args.recursive:
        updated["recursive"] = True
    if args.print_dialog:
        updated["show_print_dialog"] = True
    if args.keep_temp_pdfs:
        updated["keep_temp_pdfs"] = True
    return updated


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    config = load_config()

    if args.list_printers:
        try:
            printers = get_printers()
        except Exception as exc:
            print(f"ERROR: Could not read printer list: {exc}")
            return 1
        for index, printer in enumerate(printers, start=1):
            print(f"[{index}] {printer}")
        return 0

    if args.save_defaults:
        save_config(update_config_from_args(args, config))
        if not args.folder:
            return 0

    try:
        printers = get_printers()
    except Exception as exc:
        print(f"ERROR: Could not read printer list: {exc}")
        return 1

    try:
        options = resolve_options(args, config, printers)
        if options.selected_files is not None:
            files = filter_selected_files(options.selected_files, options.mode)
        else:
            files = find_files(options.folder, options.mode, options.recursive)
        if options.exclude_indices:
            files = [path for index, path in enumerate(files) if index not in options.exclude_indices]
    except Exception as exc:
        print(f"ERROR: {exc}")
        return 1

    if not files:
        print("No matching files found.")
        return 0

    if not confirm_run(options, files):
        print("Cancelled.")
        return 0

    return run_print_job(options, files)


if __name__ == "__main__":
    sys.exit(main())
