import json
import os
import queue
import subprocess
import sys
import threading
import time
import webbrowser
from datetime import datetime
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request
import win32print


SCRIPT_DIR = Path(__file__).resolve().parent
PYTHON_EXE = SCRIPT_DIR / "venv" / "Scripts" / "python.exe"
if not PYTHON_EXE.exists():
    PYTHON_EXE = Path(sys.executable)

PRINT_SCRIPT = SCRIPT_DIR / "Print code.py"
MERGE_SCRIPT = SCRIPT_DIR / "Merge PDF Tool.py"
SPLIT_SCRIPT = SCRIPT_DIR / "Split PDF Tool.py"
GEOSTUDIO_SCRIPT = SCRIPT_DIR / "GeoStudio Tool.py"
PDFXEDIT_PATH = r"C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe"
PRINT_EXTENSIONS = {
    "all": (".docx", ".doc", ".rtf", ".xlsx", ".xls", ".pdf"),
    "pdf": (".pdf",),
    "word": (".docx", ".doc", ".rtf"),
    "excel": (".xlsx", ".xls"),
}
MERGE_EXTENSIONS = (".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp")

app = Flask(__name__)
job_lock = threading.Lock()
jobs: dict[str, dict] = {}


def make_job_id(prefix: str) -> str:
    stamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
    return f"{prefix}-{stamp}"


def list_printers() -> list[str]:
    printers = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )
    return sorted(printer[2] for printer in printers)


def list_printers_with_timeout(timeout_seconds: float = 4.0) -> tuple[list[str], str]:
    result_queue: queue.Queue[tuple[list[str], str]] = queue.Queue(maxsize=1)

    def worker() -> None:
        try:
            result_queue.put((list_printers(), ""))
        except Exception as exc:
            result_queue.put(([], str(exc)))

    threading.Thread(target=worker, daemon=True).start()
    try:
        return result_queue.get(timeout=timeout_seconds)
    except queue.Empty:
        return [], "Printer list is taking too long. Try refreshing in a moment."


def pick_folder() -> str:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askdirectory(title="Choose folder")
    root.destroy()
    return str(Path(selected).resolve()) if selected else ""


def pick_merge_files() -> list[str]:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askopenfilenames(
        title="Choose PDFs/photos to merge",
        filetypes=[
            ("PDF and photo files", "*.pdf *.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp"),
            ("PDF files", "*.pdf"),
            ("Photo files", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp"),
            ("All files", "*.*"),
        ],
    )
    root.destroy()
    return [str(Path(path).resolve()) for path in selected]


def pick_print_file() -> str:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askopenfilename(
        title="Choose one file to print",
        filetypes=[
            ("Supported files", "*.pdf *.docx *.doc *.rtf *.xlsx *.xls"),
            ("PDF files", "*.pdf"),
            ("Word/RTF files", "*.docx *.doc *.rtf"),
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*"),
        ],
    )
    root.destroy()
    return str(Path(selected).resolve()) if selected else ""


def pick_pdf_file(title: str = "Choose PDF") -> str:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askopenfilename(
        title=title,
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    root.destroy()
    return str(Path(selected).resolve()) if selected else ""


def pick_geostudio_file() -> str:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askopenfilename(
        title="Choose GeoStudio project",
        filetypes=[("GeoStudio projects", "*.gsz"), ("All files", "*.*")],
    )
    root.destroy()
    return str(Path(selected).resolve()) if selected else ""


def parse_json_list(raw: str) -> list[str]:
    if not raw:
        return []
    try:
        return [str(item) for item in json.loads(raw) if item]
    except json.JSONDecodeError:
        return []


def relative_label(path: Path, folder: Path) -> str:
    try:
        return str(path.relative_to(folder))
    except ValueError:
        return str(path)


def preview_folder_files(folder_raw: str, mode: str, recursive: bool, extensions: tuple[str, ...]) -> list[dict]:
    folder = Path(folder_raw.strip('"')).expanduser().resolve()
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f"Folder not found: {folder}")

    entries = folder.rglob("*") if recursive else folder.iterdir()
    files = [
        path
        for path in entries
        if path.is_file()
        and path.suffix.lower() in extensions
        and not path.name.startswith("~$")
        and not path.name.endswith("_temp_print.pdf")
    ]

    if mode == "merge":
        files = [path for path in files if not path.name.lower().startswith("merged_") and "merged" not in path.stem.lower()]
    files = sorted(files, key=lambda path: str(path).lower())
    return [
        {
            "index": index,
            "name": path.name,
            "label": relative_label(path, folder),
            "path": str(path),
        }
        for index, path in enumerate(files, start=1)
    ]


def geostudio_analysis_names(project_raw: str) -> list[str]:
    import zipfile
    import xml.etree.ElementTree as ET

    project = Path(project_raw.strip('"')).expanduser().resolve()
    if not project.exists() or not project.is_file():
        raise FileNotFoundError(f"GeoStudio project not found: {project}")

    with zipfile.ZipFile(project, "r") as zip_ref:
        root_xml_files = [name for name in zip_ref.namelist() if name.endswith(".xml") and "/" not in name]
        if not root_xml_files:
            return []
        root = ET.fromstring(zip_ref.read(root_xml_files[0]))

    names = []
    analyses = root.find("Analyses")
    if analyses is not None:
        for analysis in analyses.findall("Analysis"):
            name_element = analysis.find("Name")
            if name_element is not None and name_element.text:
                names.append(name_element.text)
    return names


def write_file_list(files: list[str], prefix: str) -> Path:
    temp_list = SCRIPT_DIR / f".{prefix}_{int(time.time() * 1000)}.json"
    temp_list.write_text(json.dumps(files), encoding="utf-8")
    return temp_list


def cleanup_paths(paths: list[Path]) -> None:
    for path in paths:
        try:
            if path.exists():
                path.unlink()
        except OSError:
            pass


def launch_job(job_id: str, command: list[str], cleanup_after: list[Path] | None = None) -> None:
    def worker() -> None:
        with job_lock:
            jobs[job_id]["status"] = "running"
            jobs[job_id]["started_at"] = datetime.now().strftime("%H:%M:%S")

        cleanup_after_run = cleanup_after or []
        try:
            process = subprocess.Popen(
                command,
                cwd=SCRIPT_DIR,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
            )
            assert process.stdout is not None
            for line in process.stdout:
                with job_lock:
                    jobs[job_id]["lines"].append(line.rstrip())
            return_code = process.wait()
            with job_lock:
                jobs[job_id]["return_code"] = return_code
                jobs[job_id]["status"] = "done" if return_code == 0 else "failed"
                jobs[job_id]["finished_at"] = datetime.now().strftime("%H:%M:%S")
        except Exception as exc:
            with job_lock:
                jobs[job_id]["lines"].append(f"ERROR: {exc}")
                jobs[job_id]["status"] = "failed"
                jobs[job_id]["return_code"] = 1
                jobs[job_id]["finished_at"] = datetime.now().strftime("%H:%M:%S")
        finally:
            cleanup_paths(cleanup_after_run)

    threading.Thread(target=worker, daemon=True).start()


def create_job(kind: str, command: list[str], cleanup_after: list[Path] | None = None) -> str:
    job_id = make_job_id(kind)
    with job_lock:
        jobs[job_id] = {
            "kind": kind,
            "status": "queued",
            "lines": [],
            "return_code": None,
            "created_at": datetime.now().strftime("%H:%M:%S"),
            "started_at": "",
            "finished_at": "",
        }
    launch_job(job_id, command, cleanup_after)
    return job_id


def bool_from_form(name: str) -> bool:
    return request.form.get(name) == "on"


def is_pdf_printer(printer: str) -> bool:
    return "pdf" in printer.lower()


@app.get("/")
def index():
    return render_template("index.html", pdfxedit_path=PDFXEDIT_PATH)


@app.get("/api/printers")
def api_printers():
    printers, error = list_printers_with_timeout()
    return jsonify({"printers": printers, "error": error})


@app.post("/api/pick-folder")
def api_pick_folder():
    return jsonify({"path": pick_folder()})


@app.post("/api/pick-merge-files")
def api_pick_merge_files():
    return jsonify({"files": pick_merge_files()})


@app.post("/api/pick-print-file")
def api_pick_print_file():
    return jsonify({"path": pick_print_file()})


@app.post("/api/pick-pdf-file")
def api_pick_pdf_file():
    return jsonify({"path": pick_pdf_file("Choose PDF to split")})


@app.post("/api/pick-geostudio-file")
def api_pick_geostudio_file():
    return jsonify({"path": pick_geostudio_file()})


@app.post("/api/geostudio-analyses")
def api_geostudio_analyses():
    project = request.form.get("project", "").strip()
    if not project:
        return jsonify({"error": "Choose a GeoStudio project first."}), 400
    try:
        return jsonify({"analyses": geostudio_analysis_names(project)})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/api/preview-print")
def api_preview_print():
    folder = request.form.get("folder", "").strip()
    mode = request.form.get("mode", "all").strip()
    recursive = bool_from_form("recursive")
    if mode not in PRINT_EXTENSIONS:
        mode = "all"
    if not folder:
        return jsonify({"error": "Choose a folder first."}), 400
    try:
        return jsonify({"files": preview_folder_files(folder, "print", recursive, PRINT_EXTENSIONS[mode])})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/api/preview-merge")
def api_preview_merge():
    folder = request.form.get("folder", "").strip()
    recursive = bool_from_form("recursive")
    files = parse_json_list(request.form.get("files", "").strip())
    if files:
        folder_path = Path(os.path.commonpath([str(Path(path).resolve().parent) for path in files]))
        return jsonify(
            {
                "files": [
                    {
                        "index": index,
                        "name": Path(path).name,
                        "label": relative_label(Path(path), folder_path),
                        "path": path,
                    }
                    for index, path in enumerate(files, start=1)
                ]
            }
        )
    if not folder:
        return jsonify({"error": "Choose files or a folder first."}), 400
    try:
        return jsonify({"files": preview_folder_files(folder, "merge", recursive, MERGE_EXTENSIONS)})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/api/print")
def api_print():
    folder = request.form.get("folder", "").strip()
    printer = request.form.get("printer", "").strip()
    mode = request.form.get("mode", "all").strip()
    duplex = request.form.get("duplex", "none").strip()
    color = request.form.get("color", "yes").strip()
    pages = request.form.get("pages", "").strip()
    backend = request.form.get("backend", "pdfxchange").strip()
    delay = request.form.get("delay", "2").strip()
    exclude = request.form.get("exclude", "").strip()
    save_pdf_folder = request.form.get("save_pdf_folder", "").strip()

    if not folder:
        return jsonify({"error": "Choose a folder first."}), 400
    if not printer:
        return jsonify({"error": "Choose a printer first."}), 400
    if is_pdf_printer(printer) and not save_pdf_folder:
        return jsonify({"error": "Choose an output folder for PDF printing."}), 400

    command = [
        str(PYTHON_EXE),
        str(PRINT_SCRIPT),
        "--folder",
        folder,
        "--printer",
        printer,
        "--mode",
        mode,
        "--duplex",
        duplex,
        "--color",
        color,
        "--backend",
        backend,
        "--delay",
        delay or "2",
        "--yes",
    ]
    if is_pdf_printer(printer):
        command.extend(["--save-pdf-folder", save_pdf_folder])
    if pages:
        command.extend(["--pages", pages])
    if exclude:
        command.extend(["--exclude", exclude])
    if bool_from_form("recursive"):
        command.append("--recursive")
    if bool_from_form("print_dialog"):
        command.append("--print-dialog")
    if bool_from_form("keep_temp_pdfs"):
        command.append("--keep-temp-pdfs")
    if bool_from_form("dry_run"):
        command.append("--dry-run")

    job_id = create_job("print", command)
    return jsonify({"job_id": job_id})


@app.post("/api/print-one")
def api_print_one():
    file_path = request.form.get("file", "").strip()
    printer = request.form.get("printer", "").strip()
    duplex = request.form.get("duplex", "none").strip()
    color = request.form.get("color", "yes").strip()
    pages = request.form.get("pages", "").strip()
    backend = request.form.get("backend", "pdfxchange").strip()
    delay = request.form.get("delay", "2").strip()
    save_pdf_folder = request.form.get("save_pdf_folder", "").strip()

    if not file_path:
        return jsonify({"error": "Choose one file first."}), 400
    if not printer:
        return jsonify({"error": "Choose a printer first."}), 400
    if is_pdf_printer(printer) and not save_pdf_folder:
        return jsonify({"error": "Choose an output folder for PDF printing."}), 400
    path = Path(file_path.strip('"')).expanduser().resolve()
    if not path.exists() or not path.is_file():
        return jsonify({"error": f"File not found: {path}"}), 400

    temp_list = write_file_list([str(path)], "print_one_file")
    command = [
        str(PYTHON_EXE),
        str(PRINT_SCRIPT),
        "--file-list",
        str(temp_list),
        "--printer",
        printer,
        "--mode",
        "all",
        "--duplex",
        duplex,
        "--color",
        color,
        "--backend",
        backend,
        "--delay",
        delay or "2",
        "--yes",
    ]
    if is_pdf_printer(printer):
        command.extend(["--save-pdf-folder", save_pdf_folder])
    if pages:
        command.extend(["--pages", pages])
    if bool_from_form("print_dialog"):
        command.append("--print-dialog")
    if bool_from_form("keep_temp_pdfs"):
        command.append("--keep-temp-pdfs")
    if bool_from_form("dry_run"):
        command.append("--dry-run")

    job_id = create_job("print-one", command, [temp_list])
    return jsonify({"job_id": job_id})


@app.post("/api/merge")
def api_merge():
    folder = request.form.get("folder", "").strip()
    output = request.form.get("output", "").strip()
    exclude = request.form.get("exclude", "").strip()

    command = [str(PYTHON_EXE), str(MERGE_SCRIPT), "--yes"]
    merge_files = request.form.get("files", "").strip()
    if merge_files:
        files = [path for path in json.loads(merge_files) if path]
        if not files:
            return jsonify({"error": "Choose files or a folder first."}), 400
        temp_list = write_file_list(files, "merge_files")
        command.extend(["--file-list", str(temp_list)])
        cleanup_after = [temp_list]
    else:
        if not folder:
            return jsonify({"error": "Choose files or a folder first."}), 400
        command.extend(["--folder", folder])
        cleanup_after = []

    if output:
        command.extend(["--output", output])
    if exclude:
        command.extend(["--exclude", exclude])
    if bool_from_form("recursive"):
        command.append("--recursive")
    if bool_from_form("open_merged"):
        command.append("--open-merged")
    else:
        command.append("--no-open-merged")

    job_id = create_job("merge", command, cleanup_after)
    return jsonify({"job_id": job_id})


@app.post("/api/split")
def api_split():
    input_pdf = request.form.get("input_pdf", "").strip()
    output_folder = request.form.get("output_folder", "").strip()
    ranges = request.form.get("ranges", "").strip()

    if not input_pdf:
        return jsonify({"error": "Choose a PDF to split."}), 400
    if not output_folder:
        return jsonify({"error": "Choose an output folder."}), 400

    command = [
        str(PYTHON_EXE),
        str(SPLIT_SCRIPT),
        "--input",
        input_pdf,
        "--output-folder",
        output_folder,
        "--yes",
    ]
    if ranges:
        command.extend(["--ranges", ranges])
    if bool_from_form("open_folder"):
        command.append("--open-folder")

    job_id = create_job("split", command)
    return jsonify({"job_id": job_id})


@app.post("/api/geostudio")
def api_geostudio():
    project = request.form.get("project", "").strip()
    output_folder = request.form.get("output_folder", "").strip()
    analyses = request.form.get("analyses", "").strip()
    printer = request.form.get("printer", "").strip()
    print_pages = request.form.get("print_pages", "").strip()
    print_backend = request.form.get("print_backend", "pdfxchange").strip()
    pages_per_sheet = request.form.get("pages_per_sheet", "2").strip() or "2"
    print_pdf = bool_from_form("print_pdf")

    if not project:
        return jsonify({"error": "Choose a GeoStudio project first."}), 400
    if not output_folder:
        return jsonify({"error": "Choose an output folder."}), 400
    if print_pdf and not printer:
        return jsonify({"error": "Choose a printer for the generated GeoStudio PDFs."}), 400
    if pages_per_sheet not in {"1", "2", "4"}:
        return jsonify({"error": "Choose 1, 2, or 4 PDF pages per sheet."}), 400

    command = [
        str(PYTHON_EXE),
        str(GEOSTUDIO_SCRIPT),
        "--project",
        project,
        "--output-folder",
        output_folder,
    ]
    if analyses:
        command.extend(["--analyses", analyses])
    if bool_from_form("solve"):
        command.append("--solve")
    if bool_from_form("results"):
        command.append("--results")
    if bool_from_form("pdf") or print_pdf:
        command.append("--pdf")
        command.extend(["--pages-per-sheet", pages_per_sheet])
    if print_pdf:
        command.extend(["--print-pdf", "--printer", printer])
        if print_pages:
            command.extend(["--print-pages", print_pages])
        if print_backend:
            command.extend(["--print-backend", print_backend])
    if bool_from_form("open_folder"):
        command.append("--open-folder")

    job_id = create_job("geostudio", command)
    return jsonify({"job_id": job_id})


@app.get("/api/jobs/<job_id>")
def api_job(job_id: str):
    with job_lock:
        job = jobs.get(job_id)
        if job is None:
            return jsonify({"error": "Job not found."}), 404
        return jsonify(job)


@app.get("/api/jobs/<job_id>/events")
def api_job_events(job_id: str):
    def stream():
        last_line = 0
        while True:
            with job_lock:
                job = jobs.get(job_id)
                if job is None:
                    yield "event: error\ndata: Job not found.\n\n"
                    return
                lines = job["lines"][last_line:]
                last_line = len(job["lines"])
                payload = {
                    "status": job["status"],
                    "lines": lines,
                    "return_code": job["return_code"],
                    "finished_at": job["finished_at"],
                }
            yield f"data: {json.dumps(payload)}\n\n"
            if payload["status"] in {"done", "failed"}:
                return
            time.sleep(0.7)

    return Response(stream(), mimetype="text/event-stream")


def open_browser_later(port: int) -> None:
    def worker() -> None:
        time.sleep(1.0)
        webbrowser.open(f"http://127.0.0.1:{port}")

    threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    port = int(os.environ.get("PRINT_TOOL_PORT", "5000"))
    if "--no-browser" not in sys.argv:
        open_browser_later(port)
    print(f"Starting Local PDF Tools at http://127.0.0.1:{port}", flush=True)
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
