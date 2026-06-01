import argparse
import csv
import io
import json
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader, PdfWriter, Transformation
from pypdf._page import PageObject


SCRIPT_DIR = Path(__file__).resolve().parent
GEOCMD_PATH = Path(r"C:\Program Files\Seequent\GeoStudio 2025.1\Bin\GeoCmd.exe")
GEOSTUDIO_PATH = Path(r"C:\Program Files\Seequent\GeoStudio 2025.1\Bin\GeoStudio.exe")
PRINT_SCRIPT = SCRIPT_DIR / "Print code.py"
MICROSOFT_PDF_PRINTER = "Microsoft Print to PDF"
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
    parser.add_argument(
        "--output-folder",
        default="",
        help="Folder for reports/PDFs. Defaults to the project file folder.",
    )
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
    parser.add_argument(
        "--layout-pdf",
        action="store_true",
        help="Print GeoStudio Page Layout PDFs using GeoStudio's native print dialog.",
    )
    parser.add_argument(
        "--layout-pdf-fallback",
        action="store_true",
        help="Generate approximate geometry PDFs from project data, or use this as fallback if native printing fails.",
    )
    parser.add_argument(
        "--layout-printer",
        default=MICROSOFT_PDF_PRINTER,
        help="Printer to use for native GeoStudio Page Layout PDFs.",
    )
    parser.add_argument(
        "--layout-print-timeout",
        type=int,
        default=300,
        help="Maximum seconds to wait for native GeoStudio Page Layout PDF printing.",
    )
    parser.add_argument(
        "--print-pdf",
        action="store_true",
        help="Deprecated alias for --layout-pdf. Does not send files to a printer.",
    )
    parser.add_argument("--printer", default="", help=argparse.SUPPRESS)
    parser.add_argument("--print-pages", default="", help=argparse.SUPPRESS)
    parser.add_argument(
        "--print-backend",
        choices=["pdfxchange", "shell"],
        default="pdfxchange",
        help="Print generated PDFs through PDF-XChange or Windows shell printto.",
    )
    parser.add_argument("--print-script", default=str(PRINT_SCRIPT), help="Path to Print code.py.")
    parser.add_argument("--open-folder", action="store_true", help="Open output folder when finished.")
    parser.add_argument("--geocmd", default=str(GEOCMD_PATH), help="Path to GeoCmd.exe.")
    parser.add_argument(
        "--geocmd-timeout",
        type=int,
        default=900,
        help="Maximum seconds to wait for GeoCmd report generation.",
    )
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


def create_run_output_folder(base_folder: Path) -> Path:
    stamp = time.strftime("%Y%m%d-%H%M%S")
    candidate = base_folder / f"{stamp}-Output pdf"
    counter = 2
    while candidate.exists():
        candidate = base_folder / f"{stamp}-Output pdf {counter}"
        counter += 1
    candidate.mkdir(parents=True, exist_ok=False)
    return candidate


def pdf_output_folder(output_folder: Path) -> Path:
    return output_folder


def html_output_folder(output_folder: Path) -> Path:
    return output_folder / "html"


def layout_pdf_output_folder(output_folder: Path) -> Path:
    return output_folder


def prepare_pdf_output_path(folder: Path, stem: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    candidate = folder / f"{safe_filename(stem)}.pdf"
    if not candidate.exists():
        return candidate
    if candidate.is_file():
        try:
            candidate.unlink()
            return candidate
        except OSError:
            pass

    stamp = time.strftime("%Y%m%d_%H%M%S")
    return folder / f"{safe_filename(stem)}_{stamp}.pdf"


def remove_existing_output_path(path: Path) -> None:
    if not path.exists():
        return
    if path.is_dir():
        shutil.rmtree(path, onerror=remove_readonly_path)
    else:
        path.unlink()


def overwrite_pdf_output_path(folder: Path, stem: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    candidate = folder / f"{safe_filename(stem)}.pdf"
    remove_existing_output_path(candidate)
    return candidate


def expected_layout_pdf_paths(output_folder: Path, analyses: list[str]) -> list[Path]:
    layout_folder = layout_pdf_output_folder(output_folder)
    return [layout_folder / f"{safe_filename(analysis_name)}.pdf" for analysis_name in analyses]


def missing_or_empty(paths: list[Path]) -> list[Path]:
    return [path for path in paths if not path.exists() or not path.is_file() or path.stat().st_size <= 0]


def new_pdf_output_path(folder: Path, stem: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    safe_stem = safe_filename(stem)
    candidate = folder / f"{safe_stem}.pdf"
    if not candidate.exists():
        return candidate

    stamp = time.strftime("%Y%m%d_%H%M%S")
    candidate = folder / f"{safe_stem}_{stamp}.pdf"
    counter = 2
    while candidate.exists():
        candidate = folder / f"{safe_stem}_{stamp}_{counter}.pdf"
        counter += 1
    return candidate


def remove_readonly_path(func, path, exc_info) -> None:
    try:
        os.chmod(path, 0o700)
        func(path)
    except OSError:
        raise exc_info[1]


def collect_reports(html_files: list[Path], output_folder: Path) -> list[Path]:
    html_folder = html_output_folder(output_folder)
    html_folder.mkdir(parents=True, exist_ok=True)
    collected = []
    for html in html_files:
        html = html.resolve()
        destination = html_folder / html.name
        if html.parent.resolve() != html_folder.resolve():
            shutil.copy2(html, destination)
            collected.append(destination.resolve())
        else:
            collected.append(html)
    return sorted(set(collected), key=lambda path: str(path).lower())


def unique_work_project_path(project: Path, output_folder: Path) -> Path:
    work_root = output_folder.parent / "_geostudio_work"
    work_root.mkdir(parents=True, exist_ok=True)
    work_folder = Path(tempfile.mkdtemp(prefix="run_", dir=work_root))
    return work_folder / project.name


def html_print_copy(html_path: Path, temp_folder: Path, page_size: tuple[int, int] | None = None) -> Path:
    source = html_path.read_text(encoding="utf-8", errors="replace")
    page_rule = "size: A4 portrait;"
    if page_size is not None:
        page_width, page_height = page_size
        page_width_mm = max(page_width / 10, 1)
        page_height_mm = max(page_height / 10, 1)
        page_rule = f"size: {page_width_mm:.1f}mm {page_height_mm:.1f}mm;"
    css = f"""
<style>
@page {{
    {page_rule}
    margin: 12mm;
}}
html, body {{
    margin: 0;
    padding: 0;
}}
body {{
    box-sizing: border-box;
    color-adjust: exact;
    -webkit-print-color-adjust: exact;
}}
table {{
    page-break-inside: avoid;
}}
h1, h2, h3 {{
    page-break-after: avoid;
}}
</style>
"""
    if "</head>" in source.lower():
        lower_source = source.lower()
        insert_at = lower_source.rfind("</head>")
        source = source[:insert_at] + css + source[insert_at:]
    else:
        source = css + source

    copy_path = temp_folder / html_path.name
    copy_path.write_text(source, encoding="utf-8")
    return copy_path


def html_to_pdf(browser: Path, html_path: Path, output_pdf: Path, page_size: tuple[int, int] | None = None) -> None:
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="geostudio_html_pdf_") as temp_dir:
        temp_folder = Path(temp_dir)
        user_data_dir = temp_folder / "profile"
        print_html = html_print_copy(html_path, temp_folder, page_size)
        command = [
            str(browser),
            "--headless=new",
            "--disable-gpu",
            "--disable-background-networking",
            "--disable-extensions",
            "--disable-sync",
            "--no-first-run",
            "--no-default-browser-check",
            f"--user-data-dir={user_data_dir}",
            "--allow-file-access-from-files",
            "--no-pdf-header-footer",
            "--print-to-pdf-no-header",
            f"--print-to-pdf={output_pdf}",
            str(print_html.resolve().as_uri()),
        ]
        result = subprocess.run(command, capture_output=True, text=True, timeout=180)
    if result.returncode != 0:
        details = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(f"HTML to PDF failed for {html_path.name}: {details}")
    if not output_pdf.exists() or output_pdf.stat().st_size <= 0:
        details = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(f"HTML to PDF did not create {output_pdf.name}: {details}")


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


def safe_filename(name: str) -> str:
    cleaned = "".join(char if char.isalnum() or char in "._- " else "_" for char in name).strip()
    return cleaned or "geometry"


def xml_tag_name(element: ET.Element) -> str:
    return element.tag.rsplit("}", 1)[-1] if "}" in element.tag else element.tag


def xml_child(element: ET.Element, tag_name: str) -> ET.Element | None:
    for child in element:
        if xml_tag_name(child) == tag_name:
            return child
    return None


def xml_child_text(element: ET.Element, tag_name: str) -> str:
    child = xml_child(element, tag_name)
    return (child.text or "").strip() if child is not None else ""


def analysis_elements(root: ET.Element) -> list[ET.Element]:
    analyses = xml_child(root, "Analyses")
    if analyses is None:
        return []
    return [analysis for analysis in analyses if xml_tag_name(analysis) == "Analysis"]


def analysis_name(analysis: ET.Element) -> str:
    return xml_child_text(analysis, "Name")


def selected_analysis_names(project: Path, requested: list[str]) -> list[str]:
    root = extract_project_root(project)
    if root is None:
        return requested

    names = [analysis_name(analysis) for analysis in analysis_elements(root)]
    names = [name for name in names if name]
    if not requested:
        return names

    by_key = {name.lower(): name for name in names}
    return [by_key.get(name.lower(), name) for name in requested]


def wait_for_geostudio_window(desktop, project_name: str, timeout: int = 120):
    deadline = time.time() + timeout
    project_stem = Path(project_name).stem.lower()
    while time.time() < deadline:
        matches = [
            window
            for window in desktop.windows()
            if "GeoStudio" in window.window_text()
            and (project_name.lower() in window.window_text().lower() or project_stem in window.window_text().lower())
        ]
        if matches:
            return matches[0]
        time.sleep(1)
    raise RuntimeError(f"GeoStudio did not open the project in time. Open windows: {window_diagnostics(desktop)}")


def find_descendant(parent, predicate, timeout: int = 20, description: str = "GeoStudio control"):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for child in parent.descendants():
            if predicate(child):
                return child
        time.sleep(0.5)
    raise RuntimeError(f"Expected {description} was not found.")


def find_geostudio_print_button(main):
    return find_descendant(
        main,
        lambda child: child.friendly_class_name() == "Button"
        and getattr(child.element_info, "automation_id", "") == "2182",
        timeout=20,
        description="GeoStudio Page Layout Print button",
    )


def activate_geostudio_page_layout(main) -> None:
    if "Page Layout" in main.window_text():
        return

    page_layout_button = find_descendant(
        main,
        lambda child: child.friendly_class_name() == "Button" and child.window_text().strip() == "Page Layout",
        timeout=20,
        description="GeoStudio Page Layout mode button",
    )
    try:
        page_layout_button.invoke()
    except Exception:
        page_layout_button.click_input()

    find_descendant(
        main,
        lambda child: child.friendly_class_name() == "Button"
        and getattr(child.element_info, "automation_id", "") == "2182",
        timeout=20,
        description="GeoStudio Page Layout Print button after switching modes",
    )


def selected_state(control) -> bool | None:
    try:
        return bool(control.iface_selection_item.CurrentIsSelected)
    except Exception:
        pass
    try:
        return bool(control.is_selected())
    except Exception:
        return None


def find_analysis_tree_item(main, analysis_name: str):
    expected = analysis_name.strip()
    for item in main.descendants(control_type="TreeItem"):
        if item.window_text().strip() == expected:
            return item
    return None


def select_geostudio_analysis(main, analysis_name: str, keyboard, timeout: int = 25) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        main.set_focus()
        activate_geostudio_page_layout(main)
        matching_item = find_analysis_tree_item(main, analysis_name)
        if matching_item is None:
            keyboard.send_keys("^f")
            keyboard.send_keys(analysis_name, with_spaces=True)
            keyboard.send_keys("{ESC}")
            time.sleep(0.5)
            matching_item = find_analysis_tree_item(main, analysis_name)
        if matching_item is None:
            time.sleep(0.5)
            continue

        try:
            matching_item.scroll_into_view()
        except Exception:
            pass

        try:
            matching_item.select()
        except Exception:
            matching_item.click_input()

        time.sleep(0.3)
        try:
            matching_item.click_input()
        except Exception:
            pass

        # GeoStudio sometimes highlights the tree item before the layout view changes.
        # Opening the item and re-entering Page Layout below prevents printing stale output.
        try:
            matching_item.double_click_input()
        except Exception:
            matching_item.set_focus()
            keyboard.send_keys("{ENTER}")

        time.sleep(1.5)
        activate_geostudio_page_layout(main)
        try:
            find_geostudio_print_button(main)
            return
        except RuntimeError:
            time.sleep(0.5)

    raise RuntimeError(f"Could not select GeoStudio analysis: {analysis_name}")


def window_diagnostics(desktop) -> str:
    titles = []
    for window in desktop.windows():
        title = window.window_text()
        if title:
            titles.append(title)
    diagnostics = "; ".join(titles[:20]) or "no titled windows"
    return diagnostics.encode("ascii", errors="backslashreplace").decode("ascii")


def wait_for_save_pdf_dialog(desktop, printer: str, timeout: int = 30):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for window in desktop.windows():
            title = window.window_text()
            if "Save Print Output As" in title or title == "Save As":
                return window
        time.sleep(0.5)
    raise RuntimeError(f"{printer} did not open a PDF save dialog. Open windows: {window_diagnostics(desktop)}")


def find_open_print_dialog(desktop, main=None):
    for window in desktop.windows():
        if window.window_text().strip() == "Print":
            return window
    if main is not None:
        for child in main.descendants():
            if child.friendly_class_name() == "Dialog" and child.window_text().strip() == "Print":
                return child
    return None


def wait_for_geostudio_print_dialog(desktop, main, timeout: int = 10):
    deadline = time.time() + timeout
    while time.time() < deadline:
        print_dialog = find_open_print_dialog(desktop, main)
        if print_dialog is not None:
            return print_dialog
        time.sleep(0.5)
    raise RuntimeError(f"GeoStudio Print dialog did not open. Open windows: {window_diagnostics(desktop)}")


def open_geostudio_page_layout_print_dialog(main, desktop):
    print_button = find_geostudio_print_button(main)
    for attempt in range(3):
        main.set_focus()
        time.sleep(0.3)
        print(f"Clicking GeoStudio Page Layout Print (attempt {attempt + 1}/3)...")

        try:
            print_button.invoke()
        except Exception:
            print_button.click_input()

        try:
            return wait_for_geostudio_print_dialog(desktop, main, timeout=10)
        except RuntimeError:
            print_button = find_geostudio_print_button(main)

    return wait_for_geostudio_print_dialog(desktop, main, timeout=10)


def printer_type_text(print_dialog) -> str:
    for child in print_dialog.descendants():
        if getattr(child.element_info, "automation_id", "") == "1098":
            return child.window_text()
    return ""


def printer_matches_requested(printer: str, type_text: str) -> bool:
    expected = printer.lower()
    actual = type_text.lower()
    if "microsoft print to pdf" in expected:
        return "microsoft print to pdf" in actual
    if "pdf-xchange" in expected:
        return "pdf-xchange" in actual
    return expected in actual


def choose_geostudio_print_dialog_printer(print_dialog, printer: str, keyboard) -> str:
    combo = find_descendant(
        print_dialog,
        lambda child: child.friendly_class_name() == "ComboBox"
        and getattr(child.element_info, "automation_id", "") == "1139",
        description="GeoStudio printer name dropdown",
    )

    try:
        combo.select(printer)
    except Exception:
        combo.click_input()
        keyboard.send_keys("^a")
        keyboard.send_keys(printer, with_spaces=True)
        keyboard.send_keys("{ENTER}")

    deadline = time.time() + 10
    type_text = printer_type_text(print_dialog)
    while time.time() < deadline:
        type_text = printer_type_text(print_dialog)
        if printer_matches_requested(printer, type_text):
            return type_text
        time.sleep(0.5)

    raise RuntimeError(
        f'GeoStudio print dialog is not using requested printer "{printer}". '
        f'Current printer type shown: "{type_text or "unknown"}".'
    )


def confirm_replace_if_needed(desktop) -> None:
    for window in desktop.windows():
        title = window.window_text()
        if "Confirm Save As" in title or "Replace" in title:
            for button in window.descendants(control_type="Button"):
                if button.window_text().replace("&", "") in {"Yes", "OK"}:
                    button.click_input()
                    return


def dismiss_replace_prompt_with_no(desktop) -> None:
    for window in desktop.windows():
        title = window.window_text()
        if "Confirm Save As" in title or "Replace" in title:
            for button in window.descendants(control_type="Button"):
                if button.window_text().replace("&", "").strip() == "No":
                    button.click_input()
                    return


def dismiss_save_error_prompt(desktop) -> str | None:
    for window in desktop.windows():
        title = window.window_text().strip()
        if title not in {"Microsoft Print to PDF", "Save As", "Folder In Use"}:
            continue
        message_parts = []
        for text in window.descendants(control_type="Text"):
            value = text.window_text().strip()
            if value:
                message_parts.append(value)
        for button in window.descendants(control_type="Button"):
            label = button.window_text().replace("&", "").strip().lower()
            if label in {"ok", "cancel"}:
                try:
                    button.click_input()
                except Exception:
                    button.invoke()
                return " ".join(message_parts) or title
    return None


def set_clipboard_text(text: str) -> None:
    import win32clipboard

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


def visible_enabled_edits(window):
    return [
        control
        for control in window.descendants(control_type="Edit")
        if control.is_enabled() and control.is_visible()
    ]


def automation_id(control) -> str:
    try:
        return getattr(control.element_info, "automation_id", "") or ""
    except Exception:
        return ""


def control_name(control) -> str:
    try:
        return control.window_text().replace("&", "").strip().lower()
    except Exception:
        return ""


def find_save_filename_field(save_dialog):
    edits = visible_enabled_edits(save_dialog)
    if not edits:
        raise RuntimeError("Save dialog filename field was not found.")

    for edit in edits:
        if automation_id(edit) == "1001":
            return edit

    for edit in edits:
        name = control_name(edit)
        if name in {"file name:", "file name"} or "file name" in name:
            return edit

    # In the common Windows Save dialog, the filename box is normally the
    # lowest enabled edit field; address/search fields appear above it.
    try:
        return sorted(edits, key=lambda edit: edit.rectangle().top)[-1]
    except Exception:
        return edits[-1]


def set_save_filename_text(save_dialog, text: str, keyboard) -> None:
    edit = find_save_filename_field(save_dialog)

    edit.set_focus()
    try:
        edit.set_edit_text(text)
    except Exception:
        set_clipboard_text(text)
        keyboard.send_keys("^a")
        keyboard.send_keys("^v")


def click_save_dialog_save(save_dialog, keyboard) -> None:
    for button in save_dialog.descendants(control_type="Button"):
        label = button.window_text().replace("&", "").strip().lower()
        if label == "save":
            try:
                button.click_input()
            except Exception:
                button.invoke()
            return
    keyboard.send_keys("{ENTER}")


def wait_for_dialog_to_close(save_dialog, timeout: int = 30) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if not save_dialog.exists(timeout=0.2):
                return True
        except Exception:
            return True
        time.sleep(0.5)
    return False


def enter_save_pdf_path(save_dialog, output_pdf: Path, keyboard) -> None:
    save_dialog.set_focus()
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    set_clipboard_text(str(output_pdf.parent))
    keyboard.send_keys("%d")
    keyboard.send_keys("^v")
    keyboard.send_keys("{ENTER}")
    time.sleep(1)
    save_dialog.set_focus()
    set_save_filename_text(save_dialog, output_pdf.name, keyboard)
    click_save_dialog_save(save_dialog, keyboard)


def dismiss_geostudio_save_prompt(desktop) -> None:
    for window in desktop.windows():
        title = window.window_text()
        if "GeoStudio" not in title and "Save" not in title:
            continue
        buttons = window.descendants(control_type="Button")
        for button in buttons:
            label = button.window_text().replace("&", "").strip().lower()
            if label in {"no", "don't save", "dont save"}:
                button.click_input()
                return


def close_geostudio_process(app, project_name: str) -> None:
    if app is None:
        return

    try:
        from pywinauto import Desktop

        desktop = Desktop(backend="uia")
        for window in desktop.windows():
            if "GeoStudio" in window.window_text() and project_name in window.window_text():
                window.close()
                time.sleep(1)
                dismiss_geostudio_save_prompt(desktop)
        time.sleep(1)
        app.kill()
    except Exception:
        pass

    try:
        subprocess.run(["taskkill", "/PID", str(app.process), "/T", "/F"], capture_output=True, text=True)
    except Exception:
        pass


def geostudio_process_ids() -> set[int]:
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq GeoStudio.exe", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=10,
        )
    except Exception:
        return set()

    process_ids = set()
    for row in csv.reader(io.StringIO(result.stdout or "")):
        if len(row) >= 2 and row[0].lower() == "geostudio.exe":
            try:
                process_ids.add(int(row[1]))
            except ValueError:
                pass
    return process_ids


def stop_geostudio_processes(process_ids: set[int]) -> None:
    for process_id in process_ids:
        subprocess.run(["taskkill", "/PID", str(process_id), "/T", "/F"], capture_output=True, text=True)


def wait_for_pdf(path: Path, timeout: int = 90) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if path.exists() and path.stat().st_size > 0:
            return
        time.sleep(1)
    raise RuntimeError(f"Native GeoStudio print did not create PDF: {path}")


def wait_for_pdf_or_replace_prompt(path: Path, desktop, timeout: int = 90) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if path.exists() and path.stat().st_size > 0:
            return True
        for window in desktop.windows():
            title = window.window_text()
            if "Confirm Save As" in title or "Replace" in title:
                return False
        time.sleep(1)
    raise RuntimeError(f"Native GeoStudio print did not create PDF: {path}")


def print_geostudio_layout_pdfs(project: Path, output_folder: Path, analyses: list[str], printer: str) -> list[Path]:
    try:
        import win32print
        from pywinauto import Application, Desktop, keyboard
    except ImportError as exc:
        raise RuntimeError("Native GeoStudio printing needs pywinauto and pywin32 in the virtual environment.") from exc

    if not GEOSTUDIO_PATH.exists():
        raise RuntimeError(f"GeoStudio.exe not found: {GEOSTUDIO_PATH}")

    layout_folder = layout_pdf_output_folder(output_folder)
    layout_folder.mkdir(parents=True, exist_ok=True)

    original_printer = win32print.GetDefaultPrinter()
    app = None
    printed = []
    try:
        win32print.SetDefaultPrinter(printer)
        app = Application(backend="uia").start(f'"{GEOSTUDIO_PATH}" "{project}"')
        desktop = Desktop(backend="uia")
        main = wait_for_geostudio_window(desktop, project.name)
        time.sleep(3)
        activate_geostudio_page_layout(main)

        for analysis_name in analyses:
            print(f"Preparing GeoStudio page layout: {analysis_name}")
            select_geostudio_analysis(main, analysis_name, keyboard)
            print("Waiting for GeoStudio Page Layout Print button...")
            find_geostudio_print_button(main)
            time.sleep(1.0)

            print_dialog = open_geostudio_page_layout_print_dialog(main, desktop)
            type_text = choose_geostudio_print_dialog_printer(print_dialog, printer, keyboard)
            print(f'GeoStudio print dialog printer verified: {printer} ({type_text})')
            ok_button = find_descendant(
                print_dialog,
                lambda child: child.friendly_class_name() == "Button" and child.window_text() == "OK",
                description="GeoStudio Print OK button",
            )
            try:
                ok_button.click_input()
            except Exception:
                print_dialog.set_focus()
                keyboard.send_keys("{ENTER}")

            save_dialog = wait_for_save_pdf_dialog(desktop, printer)
            output_pdf = overwrite_pdf_output_path(layout_folder, analysis_name)
            print(f"Saving native GeoStudio layout PDF to: {output_pdf}")
            for save_attempt in range(3):
                enter_save_pdf_path(save_dialog, output_pdf, keyboard)
                time.sleep(1)
                if wait_for_dialog_to_close(save_dialog, timeout=30) and output_pdf.exists() and output_pdf.stat().st_size > 0:
                    break
                save_error = dismiss_save_error_prompt(desktop)
                if save_error:
                    print(f"Native GeoStudio PDF save attempt failed: {save_error}")
                if not wait_for_pdf_or_replace_prompt(output_pdf, desktop, timeout=5):
                    dismiss_replace_prompt_with_no(desktop)
                output_pdf = overwrite_pdf_output_path(layout_folder, analysis_name)
                print(f"Retrying native GeoStudio layout PDF path: {output_pdf}")
                save_dialog = wait_for_save_pdf_dialog(desktop, printer)
            wait_for_pdf(output_pdf)
            printed.append(output_pdf)

        return printed
    finally:
        try:
            win32print.SetDefaultPrinter(original_printer)
        except Exception as exc:
            print(f"WARNING: Could not restore default printer: {exc}")
        close_geostudio_process(app, project.name)


def print_geostudio_layout_pdfs_with_timeout(
    project: Path,
    output_folder: Path,
    analyses: list[str],
    printer: str,
    timeout_seconds: int,
) -> list[Path]:
    if timeout_seconds <= 0:
        return print_geostudio_layout_pdfs(project, output_folder, analyses, printer)

    before_process_ids = geostudio_process_ids()
    result_queue: queue.Queue[tuple[str, list[Path] | Exception]] = queue.Queue(maxsize=1)

    def worker() -> None:
        try:
            result_queue.put(("ok", print_geostudio_layout_pdfs(project, output_folder, analyses, printer)))
        except Exception as exc:
            result_queue.put(("error", exc))

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    try:
        status, payload = result_queue.get(timeout=timeout_seconds)
    except queue.Empty as exc:
        try:
            from pywinauto import Desktop

            for window in Desktop(backend="uia").windows():
                if "GeoStudio" in window.window_text() and project.name in window.window_text():
                    window.close()
        except Exception:
            pass
        stop_geostudio_processes(geostudio_process_ids() - before_process_ids)
        raise RuntimeError(f"Native GeoStudio print timed out after {timeout_seconds} seconds.") from exc

    if status == "error":
        raise payload
    return payload


def parse_rgb(raw: str | None, fallback: tuple[int, int, int]) -> tuple[int, int, int]:
    if not raw or "RGB=" not in raw:
        return fallback
    try:
        values = raw.split("RGB=", 1)[1].strip().strip("()")
        red, green, blue = [int(part.strip()) for part in values.split(",")[:3]]
        return red, green, blue
    except Exception:
        return fallback


def extract_project_root(project: Path) -> ET.Element | None:
    with zipfile.ZipFile(project, "r") as archive:
        xml_files = [name for name in archive.namelist() if name.lower().endswith(".xml")]
        root_xml_files = [name for name in xml_files if "/" not in name and "\\" not in name]
        candidates = root_xml_files + [name for name in xml_files if name not in root_xml_files]
        fallback = None
        for xml_file in candidates:
            try:
                root = ET.fromstring(archive.read(xml_file))
            except ET.ParseError:
                continue
            if fallback is None:
                fallback = root
            if analysis_elements(root):
                return root
        return fallback


def child_float(element: ET.Element, tag: str) -> float | None:
    value = element.findtext(tag)
    if value is None or value == "":
        return None
    try:
        return float(value)
    except ValueError:
        return None


def analysis_map_by_geometry(root: ET.Element) -> dict[str, list[ET.Element]]:
    analyses_by_id = {xml_child_text(analysis, "ID"): analysis for analysis in analysis_elements(root)}
    mapped = {}
    for analysis_id, analysis in analyses_by_id.items():
        geometry_id = xml_child_text(analysis, "GeometryId")
        if geometry_id:
            mapped.setdefault(geometry_id, []).append(analysis)
    return mapped


def stability_item_by_analysis(root: ET.Element) -> dict[str, ET.Element]:
    mapped = {}
    for item in root.findall(".//StabilityItem"):
        analysis_id = item.findtext("AnalysisID")
        if analysis_id:
            mapped[analysis_id] = item
    return mapped


def slip_paths_for_analysis(
    analysis: ET.Element,
    stability_items: dict[str, ET.Element],
) -> list[tuple[str, list[tuple[float, float]]]]:
    paths = []
    analysis_id = xml_child_text(analysis, "ID")
    item = stability_items.get(analysis_id or "")
    if item is None:
        return paths

    data_points = {}
    for point in item.findall(".//DataPoint"):
        number = point.get("Number")
        x_raw = point.get("X")
        y_raw = point.get("Y")
        if number and x_raw and y_raw:
            data_points[number] = (float(x_raw), float(y_raw))

    for slip in item.findall(".//FullySpecifiedSlip"):
        slip_id = xml_child_text(slip, "ID")
        point_ids = [
            (point.text or "").strip()
            for point in slip.findall("./DataPoints/DataPoint")
            if (point.text or "").strip()
        ]
        coordinates = [data_points[point_id] for point_id in point_ids if point_id in data_points]
        if len(coordinates) >= 2:
            paths.append((f"Slip {slip_id}".strip(), coordinates))
    return paths


def page_layout_settings(root: ET.Element) -> tuple[int, int, tuple[float, float], float]:
    layout = root.find("./PageLayout")
    if layout is None:
        return 1600, 1100, (0.0, 0.0), 1.0

    width = int(float(layout.findtext("PageWidth") or 2965))
    height = int(float(layout.findtext("PageHeight") or 2094))
    base = layout.find("BasePt")
    base_point = (
        float(base.get("X", "0")) if base is not None else 0.0,
        float(base.get("Y", "0")) if base is not None else 0.0,
    )
    zoom = float(layout.findtext("Zoom") or 1.0)
    return width, height, base_point, zoom


def project_page_size(project: Path) -> tuple[int, int] | None:
    root = extract_project_root(project)
    if root is None:
        return None
    page_width, page_height, _, _ = page_layout_settings(root)
    return page_width, page_height


def layout_viewport(root: ET.Element) -> tuple[tuple[float, float, float, float], tuple[float, float], float] | None:
    image = root.find("./LayoutSketchItems/Images/Image[FileName]")
    if image is None:
        return None
    file_name = image.findtext("FileName") or ""
    if "viewport" not in file_name:
        return None
    x1 = child_float(image, "X1")
    y1 = child_float(image, "Y1")
    x2 = child_float(image, "X2")
    y2 = child_float(image, "Y2")
    scale = child_float(image, "ModelScale") or 1.0
    origin = image.find("ModelOrigin")
    if None in {x1, y1, x2, y2} or origin is None:
        return None
    model_origin = (float(origin.get("X", "0")), float(origin.get("Y", "0")))
    return (x1, y1, x2, y2), model_origin, scale


def render_geometry_layout_pdf(project: Path, output_folder: Path, analyses_filter: list[str] | None = None) -> list[Path]:
    root = extract_project_root(project)
    if root is None:
        return []

    geometries = root.findall(".//Geometry")
    if not geometries:
        return []

    stability_items = stability_item_by_analysis(root)
    layout_folder = layout_pdf_output_folder(output_folder)
    layout_folder.mkdir(parents=True, exist_ok=True)
    pdf_files = []

    page_units_width, page_units_height, _, _ = page_layout_settings(root)
    page_width, page_height = int(page_units_width / 2), int(page_units_height / 2)
    viewport = layout_viewport(root)
    margin = 70
    palette = [
        (255, 246, 166),
        (206, 255, 179),
        (190, 237, 255),
        (230, 210, 255),
        (255, 220, 190),
        (220, 235, 220),
    ]
    try:
        font = ImageFont.truetype("arial.ttf", 24)
        small_font = ImageFont.truetype("arial.ttf", 18)
    except OSError:
        font = ImageFont.load_default()
        small_font = font

    geometry_by_id = {geometry.findtext("SourceId") or str(index): geometry for index, geometry in enumerate(geometries, start=1)}
    wanted = {name.lower() for name in analyses_filter or []}
    analyses = [
        analysis
        for analysis in analysis_elements(root)
        if not wanted or analysis_name(analysis).lower() in wanted
    ]
    for index, analysis in enumerate(analyses, start=1):
        current_analysis_name = analysis_name(analysis) or f"Analysis {index}"
        geometry_id = xml_child_text(analysis, "GeometryId") or str(index)
        geometry = geometry_by_id.get(geometry_id)
        if geometry is None:
            continue
        name = geometry.findtext("Name") or f"Geometry {index}"
        points = {}
        for point in geometry.findall("./Points/Point"):
            point_id = point.get("ID")
            x_raw = point.get("X")
            y_raw = point.get("Y")
            if point_id and x_raw and y_raw:
                points[point_id] = (float(x_raw), float(y_raw))
        if not points:
            continue

        sketch_lines = []
        for line in geometry.findall("./SketchItems/SkLines/SkLine"):
            try:
                sketch_lines.append(
                    (
                        (float(line.get("X1")), float(line.get("Y1"))),
                        (float(line.get("X2")), float(line.get("Y2"))),
                        max(int(float(line.get("Thickness", "1"))), 1),
                    )
                )
            except (TypeError, ValueError):
                continue

        sketch_texts = []
        for text in geometry.findall("./SketchItems/SkTexts/SkText"):
            x = child_float(text, "X")
            y = child_float(text, "Y")
            content = text.findtext("Text") or ""
            if x is not None and y is not None and content.strip() and not content.strip().startswith("{REPORT("):
                sketch_texts.append((x, y, content.strip(), parse_rgb(text.findtext("Color"), (0, 0, 0))))

        slip_paths = slip_paths_for_analysis(analysis, stability_items)

        if viewport is not None:
            (view_x1, view_y1, view_x2, view_y2), model_origin, model_scale = viewport
            scale = 2.6
            offset_x = (view_x1 / 2) - model_origin[0] * (model_scale / 2) + margin
            offset_y = (view_y2 / 2) + model_origin[1] * (model_scale / 2) - margin

            def map_point(point: tuple[float, float]) -> tuple[float, float]:
                x, y = point
                return offset_x + x * (model_scale / 2), offset_y - y * (model_scale / 2)
        else:
            all_coordinates = list(points.values())
            all_coordinates.extend(point for line in sketch_lines for point in line[:2])
            all_coordinates.extend((x, y) for x, y, _, _ in sketch_texts)
            all_coordinates.extend(point for _, path in slip_paths for point in path)

            xs = [point[0] for point in all_coordinates]
            ys = [point[1] for point in all_coordinates]
            min_x, max_x = min(xs), max(xs)
            min_y, max_y = min(ys), max(ys)
            span_x = max(max_x - min_x, 1.0)
            span_y = max(max_y - min_y, 1.0)
            scale = min((page_width - 2 * margin) / span_x, (page_height - 2 * margin) / span_y)
            drawing_width = span_x * scale
            drawing_height = span_y * scale
            offset_x = (page_width - drawing_width) / 2
            offset_y = (page_height - drawing_height) / 2 + 25

            def map_point(point: tuple[float, float]) -> tuple[float, float]:
                x, y = point
                return offset_x + (x - min_x) * scale, offset_y + (max_y - y) * scale

        image = Image.new("RGB", (page_width, page_height), "white")
        draw = ImageDraw.Draw(image)
        draw.rectangle((margin, margin, page_width - margin, page_height - margin), outline=(0, 0, 0), width=2)

        for region_index, region in enumerate(geometry.findall("./Regions/Region")):
            raw_ids = region.findtext("PointIDs") or ""
            polygon = [map_point(points[point_id]) for point_id in raw_ids.split(",") if point_id in points]
            if len(polygon) >= 3:
                draw.polygon(polygon, fill=palette[region_index % len(palette)], outline=(170, 170, 170))

        for line in geometry.findall("./Lines/Line"):
            point_1 = points.get(line.findtext("PointID1") or "")
            point_2 = points.get(line.findtext("PointID2") or "")
            if point_1 and point_2:
                draw.line([map_point(point_1), map_point(point_2)], fill=(25, 25, 25), width=4)

        for point_1, point_2, thickness in sketch_lines:
            draw.line(
                [map_point(point_1), map_point(point_2)],
                fill=(70, 70, 70),
                width=max(thickness * 2, 2),
            )

        for slip_index, (label, path) in enumerate(slip_paths, start=1):
            color = (210, 40, 40) if slip_index == 1 else (230, 110, 30)
            mapped_path = [map_point(point) for point in path]
            if len(mapped_path) >= 2:
                draw.line(mapped_path, fill=color, width=5)
                x, y = mapped_path[len(mapped_path) // 2]
                draw.text((x + 8, y - 18), label, fill=color, font=small_font)

        for point_id, point in points.items():
            x, y = map_point(point)
            radius = 4
            draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=(30, 80, 140))

        for x_model, y_model, content, color in sketch_texts:
            x, y = map_point((x_model, y_model))
            lines = content.splitlines()
            for line_index, line in enumerate(lines):
                y_line = y + line_index * 22
                bbox = draw.textbbox((x, y_line), line, font=small_font)
                draw.rectangle(
                    (bbox[0] - 3, bbox[1] - 2, bbox[2] + 3, bbox[3] + 2),
                    fill=(255, 255, 255),
                )
                draw.text((x, y_line), line, fill=color, font=small_font)

        output_pdf = overwrite_pdf_output_path(layout_folder, current_analysis_name)
        image.save(output_pdf, "PDF", resolution=150.0)
        pdf_files.append(output_pdf)

    return pdf_files


def run_geocmd(args: argparse.Namespace) -> int:
    project = Path(args.project.strip('"')).expanduser().resolve()
    if not project.exists() or not project.is_file():
        print(f"ERROR: GeoStudio project not found: {project}")
        return 1
    if project.suffix.lower() != ".gsz":
        print("ERROR: Project must be a .gsz file.")
        return 1

    output_folder_raw = args.output_folder.strip('"')
    base_output_folder = (
        Path(output_folder_raw).expanduser().resolve()
        if output_folder_raw
        else project.parent
    )
    base_output_folder.mkdir(parents=True, exist_ok=True)
    output_folder = create_run_output_folder(base_output_folder)

    geocmd = Path(args.geocmd.strip('"')).expanduser().resolve()
    if not geocmd.exists():
        print(f"ERROR: GeoCmd.exe not found: {geocmd}")
        return 1
    if args.print_pdf:
        args.layout_pdf = True

    work_project = unique_work_project_path(project, output_folder)
    shutil.copy2(project, work_project)

    try:
        command = [str(geocmd), str(work_project)]
        command.extend(parse_analyses(args.analyses))
        if args.solve:
            command.append("/solve")
        command.append("/report")
        if args.results:
            command.append("/results")

        html_search_folders = [output_folder, work_project.parent]
        before_html = snapshot_html_files(html_search_folders)

        print("\nGeoStudio report settings:")
        print(f"Project : {project}")
        print(f"Work copy: {work_project}")
        print(f"Output base: {base_output_folder}")
        print(f"Output  : {output_folder}")
        print(f"Analyses: {args.analyses or 'all'}")
        print(f"Solve   : {'yes' if args.solve else 'no'}")
        print(f"Results : {'yes' if args.results else 'no'}")
        print(f"PDF     : {'yes' if args.pdf else 'no'}")
        layout_requested = args.layout_pdf or args.layout_pdf_fallback
        print(f"Native GeoStudio layout PDF: {'yes' if args.layout_pdf else 'no'}")
        print(f"Approximate fallback layout PDF: {'yes' if args.layout_pdf_fallback else 'no'}")
        if args.layout_pdf:
            print(f"Layout printer: {args.layout_printer}")
        if args.pdf:
            print(f"N-up PDF: {args.pages_per_sheet} page(s) per sheet")
        print("\nRunning GeoCmd...")
        print(" ".join(f'"{part}"' if " " in part else part for part in command))

        try:
            result = subprocess.run(
                command,
                cwd=output_folder,
                capture_output=True,
                text=True,
                timeout=args.geocmd_timeout if args.geocmd_timeout > 0 else None,
            )
        except subprocess.TimeoutExpired as exc:
            print(f"ERROR: GeoCmd timed out after {args.geocmd_timeout} seconds.")
            if exc.stdout:
                print(str(exc.stdout).strip())
            if exc.stderr:
                print(str(exc.stderr).strip())
            return 1
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

        if layout_requested:
            analysis_names = selected_analysis_names(work_project, parse_analyses(args.analyses))
            if args.layout_pdf:
                print("\nPrinting GeoStudio page layout PDFs...")
                try:
                    layout_pdfs = print_geostudio_layout_pdfs_with_timeout(
                        work_project,
                        output_folder,
                        analysis_names,
                        args.layout_printer,
                        args.layout_print_timeout,
                    )
                except Exception as exc:
                    print(f"ERROR: Native GeoStudio print failed: {exc}")
                    if not args.layout_pdf_fallback:
                        return 1
                    print("Falling back to generated geometry layout PDFs...")
                    layout_pdfs = render_geometry_layout_pdf(work_project, output_folder, analysis_names)
                else:
                    missing_layouts = missing_or_empty(expected_layout_pdf_paths(output_folder, analysis_names))
                    if missing_layouts:
                        print("WARNING: Native GeoStudio print did not leave all expected layout PDFs:")
                        for missing_pdf in missing_layouts:
                            print(f"- Missing or empty: {missing_pdf}")
                        if not args.layout_pdf_fallback:
                            return 1
                        print("Regenerating layout PDFs with fallback renderer...")
                        layout_pdfs = render_geometry_layout_pdf(work_project, output_folder, analysis_names)
            else:
                print("\nGenerating GeoStudio page layout PDFs before HTML report PDF conversion...")
                layout_pdfs = render_geometry_layout_pdf(work_project, output_folder, analysis_names)
            if layout_pdfs:
                for pdf in layout_pdfs:
                    print(f"- {pdf}")
            else:
                print("- No geometry layouts found in the project.")

        if args.pdf:
            browser = find_browser()
            if browser is None:
                print("ERROR: Could not find Edge or Chrome for HTML-to-PDF conversion.")
                return 1

            pdf_folder = pdf_output_folder(output_folder)
            page_size = project_page_size(work_project)
            print("\nConverting HTML reports to PDF...")
            for html in html_files:
                output_pdf = pdf_folder / f"{html.stem}.pdf"
                if args.pages_per_sheet == 1:
                    html_to_pdf(browser, html, output_pdf, page_size)
                else:
                    raw_pdf = pdf_folder / f"{html.stem}.__raw.pdf"
                    html_to_pdf(browser, html, raw_pdf, page_size)
                    nup_pdf(raw_pdf, output_pdf, args.pages_per_sheet)
                    raw_pdf.unlink(missing_ok=True)
                print(f"- {output_pdf}")
    finally:
        for attempt in range(10):
            try:
                shutil.rmtree(work_project.parent, onerror=remove_readonly_path)
                print(f"\nRemoved work copy: {work_project.parent}")
                break
            except OSError as exc:
                if attempt == 9:
                    print(f"\nWARNING: Could not remove work copy folder: {exc}")
                else:
                    time.sleep(2)

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
