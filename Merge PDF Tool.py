import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "merge_config.json"

DEFAULT_CONFIG = {
    "pdfxedit_path": r"C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe",
    "default_folder": "",
    "output_folder": "",
    "recursive": False,
    "temp_dir": "",
    "open_merged_pdf": True,
}

IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp")
MERGE_EXTENSIONS = (".pdf",) + IMAGE_EXTENSIONS


@dataclass
class MergeOptions:
    folder: Path
    selected_files: list[Path] | None
    output_pdf: Path
    recursive: bool
    dry_run: bool
    yes: bool
    temp_dir: Path
    pdfxedit_path: Path
    exclude_indices: set[int]
    open_merged_pdf: bool


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
    parser = argparse.ArgumentParser(description="Merge PDFs and photos into one PDF.")
    parser.add_argument("--folder", help="Folder that contains PDFs/photos to merge.")
    parser.add_argument(
        "--file-list",
        help="JSON file containing exact PDFs/photos to merge. Used by the local web app.",
    )
    parser.add_argument(
        "--pick-files",
        action="store_true",
        help="Open a Windows file picker and merge only the selected files.",
    )
    parser.add_argument("--output", help="Output PDF path. If omitted, a timestamped PDF is created.")
    parser.add_argument("--recursive", action="store_true", help="Include files in subfolders.")
    parser.add_argument("--dry-run", action="store_true", help="Preview the run without creating a PDF.")
    parser.add_argument("--yes", action="store_true", help="Skip final confirmation.")
    parser.add_argument("--exclude", default="", help="Comma-separated file numbers to exclude, such as 1,3,5.")
    parser.add_argument("--save-defaults", action="store_true", help="Save provided options to merge_config.json.")
    parser.add_argument("--temp-dir", help="Folder for temporary image PDFs.")
    parser.add_argument("--pdfxedit", help="Path to PDF-XChange Editor executable.")
    parser.add_argument(
        "--open-merged",
        action="store_true",
        help="Open the merged PDF in PDF-XChange Editor after creating it.",
    )
    parser.add_argument(
        "--no-open-merged",
        action="store_true",
        help="Do not open the merged PDF after creating it.",
    )
    return parser


def prompt_with_default(message: str, default: str = "") -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{message}{suffix}: ").strip().strip('"')
    return value or default


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


def pick_merge_files_with_dialog(initial_dir: str = "") -> list[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:
        raise RuntimeError(f"Could not open Windows file picker: {exc}") from exc

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_types = [
        ("PDF and photo files", "*.pdf *.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp"),
        ("PDF files", "*.pdf"),
        ("Photo files", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp"),
        ("All files", "*.*"),
    ]
    selected = filedialog.askopenfilenames(
        title="Choose PDFs/photos to merge",
        initialdir=initial_dir or str(Path.home()),
        filetypes=file_types,
    )
    root.destroy()
    return [Path(path).resolve() for path in selected]


def common_parent(paths: list[Path]) -> Path:
    if not paths:
        return Path.cwd()
    return Path(os.path.commonpath([str(path.parent) for path in paths]))


def default_merge_output_path(folder: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return folder / f"merged_{stamp}.pdf"


def resolve_output_pdf(raw_output: str | None, folder: Path, config: dict) -> Path:
    if raw_output:
        output_pdf = Path(raw_output.strip('"')).expanduser()
        if not output_pdf.is_absolute():
            if output_pdf.parent == Path("."):
                output_pdf = folder / output_pdf
            else:
                output_pdf = Path.cwd() / output_pdf
    else:
        output_folder_raw = config.get("output_folder") or ""
        output_folder = Path(output_folder_raw).expanduser() if output_folder_raw else folder
        output_pdf = default_merge_output_path(output_folder)

    if output_pdf.suffix.lower() != ".pdf":
        output_pdf = output_pdf.with_suffix(".pdf")
    return output_pdf.resolve()


def resolve_options(args: argparse.Namespace, config: dict) -> MergeOptions:
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
        selected_files = pick_merge_files_with_dialog(initial_dir)
        if not selected_files:
            raise RuntimeError("No files selected.")
        folder = common_parent(selected_files)
    else:
        folder = Path(args.folder.strip('"')).expanduser().resolve()
        if not folder.exists() or not folder.is_dir():
            raise FileNotFoundError(f"Folder not found: {folder}")

    if not (args.folder or args.pick_files or args.yes):
        recursive = choose_yes_no("Include files in subfolders?", bool(config.get("recursive", False)))
    else:
        recursive = bool(args.recursive or config.get("recursive", False))

    temp_dir_raw = args.temp_dir or config.get("temp_dir") or tempfile.gettempdir()
    temp_dir = Path(temp_dir_raw).expanduser().resolve() / "pdf_merge_tool"
    temp_dir.mkdir(parents=True, exist_ok=True)

    output_pdf = resolve_output_pdf(args.output, folder, config)
    pdfxedit_path = Path(args.pdfxedit or config.get("pdfxedit_path") or DEFAULT_CONFIG["pdfxedit_path"])
    open_merged_pdf = bool(config.get("open_merged_pdf", True))
    if args.open_merged:
        open_merged_pdf = True
    if args.no_open_merged:
        open_merged_pdf = False

    return MergeOptions(
        folder=folder,
        selected_files=selected_files,
        output_pdf=output_pdf,
        recursive=recursive,
        dry_run=args.dry_run,
        yes=args.yes,
        temp_dir=temp_dir,
        pdfxedit_path=pdfxedit_path,
        exclude_indices=parse_exclude_indices(args.exclude),
        open_merged_pdf=open_merged_pdf,
    )


def find_merge_files(folder: Path, recursive: bool) -> list[Path]:
    entries = folder.rglob("*") if recursive else folder.iterdir()
    files = [
        path
        for path in entries
        if path.is_file()
        and path.suffix.lower() in MERGE_EXTENSIONS
        and not path.name.startswith("~$")
        and not path.name.endswith("_temp_print.pdf")
        and not path.name.lower().startswith("merged_")
        and "merged" not in path.stem.lower()
    ]
    return sorted(files, key=lambda path: str(path).lower())


def filter_selected_merge_files(files: list[Path]) -> list[Path]:
    return sorted(
        [
            path
            for path in files
            if path.is_file()
            and path.suffix.lower() in MERGE_EXTENSIONS
            and not path.name.startswith("~$")
            and not path.name.endswith("_temp_print.pdf")
        ],
        key=lambda path: str(path).lower(),
    )


def safe_temp_pdf_path(source: Path, temp_dir: Path) -> Path:
    safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", source.stem).strip("._") or "converted"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return temp_dir / f"{safe_stem}_{stamp}.pdf"


def print_file_list(files: list[Path], folder: Path) -> None:
    print("\nFiles to merge:")
    for index, path in enumerate(files, start=1):
        try:
            label = path.relative_to(folder)
        except ValueError:
            label = path
        print(f"[{index}] {label}")


def confirm_run(options: MergeOptions, files: list[Path]) -> bool:
    print_file_list(files, options.folder)
    print("\nMerge settings:")
    print(f"Folder   : {options.folder}")
    print(f"Output   : {options.output_pdf}")
    print(f"Recursive: {'yes' if options.recursive else 'no'}")
    print(f"Open PDF : {'yes' if options.open_merged_pdf else 'no'}")
    if options.dry_run:
        print("Dry run  : yes")

    if options.yes or options.dry_run:
        return True
    return input("\nCreate merged PDF? [Y/n]: ").strip().lower() not in {"n", "no"}


def open_pdf_in_pdfxchange(pdf_path: Path, pdfxedit_path: Path) -> None:
    if pdfxedit_path.exists():
        subprocess.Popen([str(pdfxedit_path), str(pdf_path)], shell=False)
        return
    os.startfile(str(pdf_path))


def image_to_pdf(image_path: Path, temp_pdf: Path) -> Path:
    try:
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("Photo merging needs Pillow. Run: venv\\Scripts\\python.exe -m pip install Pillow") from exc

    images = []
    try:
        with Image.open(image_path) as source:
            frame_count = getattr(source, "n_frames", 1)
            for frame_index in range(frame_count):
                if frame_count > 1:
                    source.seek(frame_index)
                frame = source.copy()
                if frame.mode in ("RGBA", "LA", "P"):
                    background = Image.new("RGB", frame.size, "white")
                    if frame.mode == "P":
                        frame = frame.convert("RGBA")
                    alpha = frame.getchannel("A") if frame.mode in ("RGBA", "LA") else None
                    background.paste(frame.convert("RGB"), mask=alpha)
                    frame = background
                else:
                    frame = frame.convert("RGB")
                images.append(frame)

        if not images:
            raise RuntimeError("No image frames found.")

        first, rest = images[0], images[1:]
        first.save(str(temp_pdf), "PDF", save_all=bool(rest), append_images=rest, resolution=300.0)
        return temp_pdf
    finally:
        for image in images:
            image.close()


def add_pdf_to_writer(pdf_path: Path, writer) -> int:
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise RuntimeError("PDF merging needs pypdf. Run: venv\\Scripts\\python.exe -m pip install pypdf") from exc

    reader = PdfReader(str(pdf_path))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:
            pass
        if reader.is_encrypted:
            raise RuntimeError("PDF is encrypted and could not be opened.")

    page_count = len(reader.pages)
    for page in reader.pages:
        writer.add_page(page)
    return page_count


def cleanup_file(path: Path) -> None:
    for _ in range(6):
        try:
            if path.exists():
                path.unlink()
            return
        except PermissionError:
            time.sleep(0.5)


def run_merge_job(options: MergeOptions, files: list[Path]) -> int:
    if options.dry_run:
        print("\nDry run complete. No PDF was created.")
        return 0

    try:
        from pypdf import PdfWriter
    except ImportError:
        print("ERROR: PDF merging needs pypdf. Run: venv\\Scripts\\python.exe -m pip install pypdf")
        return 1

    options.output_pdf.parent.mkdir(parents=True, exist_ok=True)
    if options.output_pdf.exists():
        if options.yes:
            cleanup_file(options.output_pdf)
        else:
            overwrite = input(f"\nOutput already exists. Replace {options.output_pdf.name}? [y/N]: ").strip().lower()
            if overwrite not in {"y", "yes"}:
                print("Cancelled.")
                return 0

    writer = PdfWriter()
    temp_pdfs = []
    added_pages = 0
    merged = []
    failed = []

    print("\nCreating merged PDF...")
    for number, path in enumerate(files, start=1):
        print(f"[{number}/{len(files)}] {path.name}")
        try:
            source_pdf = path
            if path.suffix.lower() in IMAGE_EXTENSIONS:
                source_pdf = safe_temp_pdf_path(path, options.temp_dir)
                image_to_pdf(path, source_pdf)
                temp_pdfs.append(source_pdf)

            pages = add_pdf_to_writer(source_pdf, writer)
            added_pages += pages
            merged.append(path)
            print(f"  OK: added {pages} page(s)")
        except Exception as exc:
            failed.append((path, exc))
            print(f"  FAILED: {exc}")

    if not merged:
        print("\nNo files were merged.")
        for temp_pdf in temp_pdfs:
            cleanup_file(temp_pdf)
        return 1

    try:
        with options.output_pdf.open("wb") as handle:
            writer.write(handle)
    finally:
        writer.close()
        for temp_pdf in temp_pdfs:
            cleanup_file(temp_pdf)

    print("\nSummary:")
    print(f"Merged : {len(merged)} file(s)")
    print(f"Pages  : {added_pages}")
    print(f"Failed : {len(failed)}")
    print(f"Output : {options.output_pdf}")
    if failed:
        print("\nFailed files:")
        for path, exc in failed:
            print(f"- {path.name}: {exc}")

    if options.open_merged_pdf:
        try:
            open_pdf_in_pdfxchange(options.output_pdf, options.pdfxedit_path)
            print("Opened merged PDF.")
        except Exception as exc:
            print(f"WARNING: Could not open merged PDF: {exc}")

    return 1 if failed else 0


def update_config_from_args(args: argparse.Namespace, config: dict) -> dict:
    updated = config.copy()
    mappings = {
        "folder": "default_folder",
        "temp_dir": "temp_dir",
        "pdfxedit": "pdfxedit_path",
    }
    for arg_name, config_name in mappings.items():
        value = getattr(args, arg_name)
        if value:
            updated[config_name] = value
    if args.output:
        updated["output_folder"] = str(Path(args.output).expanduser().parent)
    if args.recursive:
        updated["recursive"] = True
    if args.open_merged:
        updated["open_merged_pdf"] = True
    if args.no_open_merged:
        updated["open_merged_pdf"] = False
    return updated


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    config = load_config()

    if args.save_defaults:
        save_config(update_config_from_args(args, config))
        if not args.folder and not args.pick_files:
            return 0

    try:
        options = resolve_options(args, config)
        if options.selected_files is not None:
            files = filter_selected_merge_files(options.selected_files)
        else:
            files = find_merge_files(options.folder, options.recursive)
        files = [path for path in files if path.resolve() != options.output_pdf]
        if options.exclude_indices:
            files = [path for index, path in enumerate(files) if index not in options.exclude_indices]
    except Exception as exc:
        print(f"ERROR: {exc}")
        return 1

    if not files:
        print("No matching PDF/photo files found.")
        return 0

    if not confirm_run(options, files):
        print("Cancelled.")
        return 0

    return run_merge_job(options, files)


if __name__ == "__main__":
    sys.exit(main())
