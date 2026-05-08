import ast
import importlib.util
import os
import re
import fnmatch
import struct
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

# ---- Optional dependency setup ----------------------------------------------
# pip install python-docx tqdm pypdf
if importlib.util.find_spec("docx") is not None:
    from docx import Document
else:
    Document = None

if importlib.util.find_spec("tqdm") is not None:
    from tqdm import tqdm
else:
    def tqdm(iterable, **_kwargs):
        return iterable

# ---- PDF extraction (extractable text only) ---------------------------------
# pip install pypdf
if importlib.util.find_spec("pypdf") is not None:
    from pypdf import PdfReader
else:
    PdfReader = None

# =============================================================================
# Repo-safe Chronological DOCX/PY/TXT/MD/PDF Log Extractor (Recursive, Threaded)
# =============================================================================

TIME_MODE = "ctime"  # "ctime" or "mtime"

# ---- Output splitting --------------------------------------------------------
MAX_OUTPUT_TXT_SIZE_MB = 4.5
OUTPUT_BASENAME = "Output_data"  # Output_data_part001.txt, part002, ...

# ---- Threading --------------------------------------------------------------
NUM_WORKERS = 12

# ---- File-type toggles ------------------------------------------------------
INCLUDE_DOCX_FILES = True
INCLUDE_PY_FILES   = True
INCLUDE_TXT_FILES  = True
INCLUDE_MD_FILES   = True
INCLUDE_PDF_FILES  = False
INCLUDE_NPZ_FILES  = False
INCLUDE_ZIP_FILES  = False

INCLUDE_DOCX_TEXT  = True
INCLUDE_PY_TEXT    = True
INCLUDE_TXT_TEXT   = True
INCLUDE_MD_TEXT    = True
INCLUDE_PDF_TEXT   = False
INCLUDE_NPZ_TEXT   = False
INCLUDE_ZIP_TEXT   = False

ONLY_FILENAME = True
REDACT_ABS_PATHS_IN_TEXT = True

# NOTE: We now do header/path cleanup *during writing/extraction*,
# so we do NOT do a post-read/post-rewrite pass anymore.
FORCE_FILENAME_ONLY_IN_HEADERS = True  # safe; implemented directly in header formatting

STRIP_LEAKY_MARKERS = True
LEAKY_MARKERS = ["OneDrive", "\\Desktop\\", "/Desktop/"]

# PDF behavior: include only if it yields meaningful extractable text
PDF_MIN_NONWS_CHARS = 30
PDF_MAX_PAGES = 0  # 0 = no limit; otherwise cap pages to extract (you said keep full, so keep 0)

FILE_EXCLUDES = [
    "*draft*", "*todo*", "*email*", "*notes*", "*journal*", "*brainstorm*",
    "*prompt*", "*conversation*", "*chat*", "*internal*", "*plan*",
    "*output_data*",
    "*.env*", "*secret*", "*token*", "*apikey*", "*password*",
]

FOLDER_EXCLUDES = [
    ".git", ".github", "__pycache__", ".idea", ".vscode", "venv", ".venv",
    "outputs", "output", "results", "dist", "build", "cache", "tmp", "temp",
    "archive", "archives", "backup", "backups", "old", "legacy", "deprecated",
    "notes", "drafts", "planning", "scratch",
    "Rotmod_LTG", "SPARC", "datasets", "data", "$RECYCLE.BIN", "System Volume Information",
]

# -----------------------------------------------------------------------------
# Precompiled regex (hot paths)
# -----------------------------------------------------------------------------
RE_WS = re.compile(r"\s+")
RE_WIN_ABS1 = re.compile(r"\b[A-Za-z]:\\[^\s\"\'<>|]+")
RE_WIN_ABS2 = re.compile(r"\b[A-Za-z]:/[^\s\"\'<>|]+")
RE_UNC = re.compile(r"\\\\[^\s\"\'<>|]+\\[^\s\"\'<>|]+(?:\\[^\s\"\'<>|]+)*")

# -----------------------------------------------------------------------------
# Extraction helpers
# -----------------------------------------------------------------------------
def extract_docx_text(path: str) -> str:
    if Document is None:
        return "[ERROR reading DOCX: missing dependency 'python-docx'. Install with: pip install python-docx]"

    try:
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as e:
        return f"[ERROR reading DOCX: {e}]"

def extract_py_text(path: str) -> str:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except UnicodeDecodeError:
        try:
            with open(path, "r", encoding="latin-1") as f:
                return f.read()
        except Exception as e:
            return f"[ERROR reading PY (latin-1 fallback failed): {e}]"
    except Exception as e:
        return f"[ERROR reading PY: {e}]"

def extract_txt_text(path: str) -> str:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except UnicodeDecodeError:
        try:
            with open(path, "r", encoding="latin-1") as f:
                return f.read()
        except Exception as e:
            return f"[ERROR reading TXT (latin-1 fallback failed): {e}]"
    except Exception as e:
        return f"[ERROR reading TXT: {e}]"


def extract_md_text(path: str) -> str:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except UnicodeDecodeError:
        try:
            with open(path, "r", encoding="latin-1") as f:
                return f.read()
        except Exception as e:
            return f"[ERROR reading MD (latin-1 fallback failed): {e}]"
    except Exception as e:
        return f"[ERROR reading MD: {e}]"

def extract_pdf_text(path: str) -> str:
    """
    Extracts full text from PDFs *only if* the PDF has extractable text.
    Skips scanned/image-only PDFs by returning "" (empty string).
    (No truncation; extracts all pages unless PDF_MAX_PAGES != 0)
    """
    if PdfReader is None:
        return "[ERROR reading PDF: missing dependency 'pypdf'. Install with: pip install pypdf]"

    try:
        reader = PdfReader(path)
        total_pages = len(reader.pages)
        page_limit = total_pages if (PDF_MAX_PAGES == 0) else min(total_pages, PDF_MAX_PAGES)

        texts = []
        for i in range(page_limit):
            try:
                t = reader.pages[i].extract_text() or ""
            except Exception:
                t = ""
            texts.append(t)

        full = "\n".join(texts).strip()

        nonws = RE_WS.sub("", full)
        if len(nonws) < PDF_MIN_NONWS_CHARS:
            return ""  # no meaningful extractable text

        return full
    except Exception as e:
        return f"[ERROR reading PDF: {e}]"


def parse_npy_header_from_npz_member(member) -> dict:
    magic = member.read(6)
    if magic != b"\x93NUMPY":
        return {}

    major, minor = member.read(2)
    if (major, minor) == (1, 0):
        header_len = struct.unpack("<H", member.read(2))[0]
    else:
        header_len = struct.unpack("<I", member.read(4))[0]

    header_text = member.read(header_len).decode("latin-1").strip()
    return ast.literal_eval(header_text)


def format_shape(shape) -> str:
    if isinstance(shape, tuple):
        return str(shape)
    return "unknown"


def shape_size(shape) -> str:
    if not isinstance(shape, tuple):
        return "unknown"

    size = 1
    for dim in shape:
        size *= dim
    return str(size)


def extract_npz_text(path: str) -> str:
    """
    Build a lightweight text summary of NPZ contents (arrays + shapes/dtypes).
    """
    try:
        with zipfile.ZipFile(path, "r") as zf:
            infos = [info for info in zf.infolist() if not info.is_dir()]
            if not infos:
                return "[EMPTY NPZ: no arrays found]"

            lines = ["NPZ CONTENT SUMMARY:"]
            for info in infos:
                key = info.filename[:-4] if info.filename.endswith(".npy") else info.filename
                with zf.open(info, "r") as member:
                    header = parse_npy_header_from_npz_member(member)

                if header:
                    shape = header.get("shape")
                    dtype = header.get("descr", "unknown")
                    lines.append(
                        f"- {key}: shape={format_shape(shape)}, dtype={dtype}, "
                        f"size={shape_size(shape)}"
                    )
                else:
                    lines.append(f"- {key}: {info.file_size} bytes")

            return "\n".join(lines)
    except Exception as e:
        return f"[ERROR reading NPZ: {e}]"


def extract_zip_text(path: str) -> str:
    """
    Build a lightweight text summary of ZIP archive entries.
    """
    try:
        with zipfile.ZipFile(path, "r") as zf:
            infos = zf.infolist()
            if not infos:
                return "[EMPTY ZIP: no files found]"

            lines = ["ZIP CONTENT SUMMARY:"]
            for info in infos:
                kind = "dir" if info.is_dir() else "file"
                lines.append(f"- {info.filename} ({kind}, {info.file_size} bytes)")
            return "\n".join(lines)
    except Exception as e:
        return f"[ERROR reading ZIP: {e}]"

# -----------------------------------------------------------------------------
# Matching + redaction helpers
# -----------------------------------------------------------------------------
def matches_any_pattern(text: str, patterns: list[str]) -> bool:
    t = text.lower()
    for pat in patterns:
        p = pat.lower()
        if ("*" in p) or ("?" in p):
            if fnmatch.fnmatch(t, p):
                return True
        else:
            if p in t:
                return True
    return False

def folder_is_excluded(folder_name: str, patterns: list[str]) -> bool:
    fn = folder_name.lower()
    for pat in patterns:
        p = pat.lower()
        if ("*" in p) or ("?" in p):
            if fnmatch.fnmatch(fn, p):
                return True
        else:
            if fn == p:
                return True
    return False

def get_sort_timestamp(path: str, mode: str) -> float:
    try:
        return os.path.getmtime(path) if mode == "mtime" else os.path.getctime(path)
    except Exception:
        return 0.0

def safe_label(path: str) -> str:
    # We also enforce filename-only in headers below if FORCE_FILENAME_ONLY_IN_HEADERS
    return os.path.basename(path) if ONLY_FILENAME else path

def redact_abs_paths_text(s: str) -> str:
    if not s:
        return s
    s = RE_WIN_ABS1.sub("[REDACTED_PATH]", s)
    s = RE_WIN_ABS2.sub("[REDACTED_PATH]", s)
    s = RE_UNC.sub("[REDACTED_PATH]", s)
    return s

def strip_leaky_markers(s: str) -> str:
    if not s:
        return s
    if STRIP_LEAKY_MARKERS:
        for marker in LEAKY_MARKERS:
            s = s.replace(marker, "[REDACTED_MARKER]")
    return s

# -----------------------------------------------------------------------------
# Thread worker
# -----------------------------------------------------------------------------
def process_one(index: int, ts: float, path: str):
    """
    Returns a tuple:
      (index, ts, path, type_str, content_str, status)

    status: "ok" | "skipped_no_text" | "omitted"
    """
    name = os.path.basename(path)
    lower = name.lower()

    def finalize(txt: str) -> str:
        if REDACT_ABS_PATHS_IN_TEXT:
            txt = redact_abs_paths_text(txt)
        txt = strip_leaky_markers(txt)
        return txt

    if lower.endswith(".docx"):
        if not INCLUDE_DOCX_TEXT:
            return (index, ts, path, "DOCX", "[DOCX CONTENT OMITTED]\n", "omitted")
        txt = finalize(extract_docx_text(path))
        return (index, ts, path, "DOCX", txt + "\n", "ok")

    if lower.endswith(".py"):
        if not INCLUDE_PY_TEXT:
            return (index, ts, path, "PYTHON", "[PY CONTENT OMITTED]\n", "omitted")
        txt = finalize(extract_py_text(path))
        return (index, ts, path, "PYTHON", txt + "\n", "ok")

    if lower.endswith(".txt"):
        if not INCLUDE_TXT_TEXT:
            return (index, ts, path, "TEXT", "[TXT CONTENT OMITTED]\n", "omitted")
        txt = finalize(extract_txt_text(path))
        return (index, ts, path, "TEXT", txt + "\n", "ok")

    if lower.endswith(".md"):
        if not INCLUDE_MD_TEXT:
            return (index, ts, path, "MARKDOWN", "[MD CONTENT OMITTED]\n", "omitted")
        txt = finalize(extract_md_text(path))
        return (index, ts, path, "MARKDOWN", txt + "\n", "ok")

    if lower.endswith(".pdf"):
        if not INCLUDE_PDF_TEXT:
            return (index, ts, path, "PDF", "[PDF CONTENT OMITTED]\n", "omitted")
        txt = extract_pdf_text(path)
        if txt == "":
            return (index, ts, path, "PDF", "[PDF SKIPPED: no meaningful extractable text detected]\n", "skipped_no_text")
        txt = finalize(txt)
        return (index, ts, path, "PDF", txt + "\n", "ok")

    if lower.endswith(".npz"):
        if not INCLUDE_NPZ_TEXT:
            return (index, ts, path, "NPZ", "[NPZ CONTENT OMITTED]\n", "omitted")
        txt = finalize(extract_npz_text(path))
        return (index, ts, path, "NPZ", txt + "\n", "ok")

    if lower.endswith(".zip"):
        if not INCLUDE_ZIP_TEXT:
            return (index, ts, path, "ZIP", "[ZIP CONTENT OMITTED]\n", "omitted")
        txt = finalize(extract_zip_text(path))
        return (index, ts, path, "ZIP", txt + "\n", "ok")

    return (index, ts, path, "UNKNOWN", "[UNSUPPORTED FILE TYPE]\n", "omitted")

# -----------------------------------------------------------------------------
# Output splitting helpers
# -----------------------------------------------------------------------------
def make_part_path(output_dir: str, part_index: int) -> str:
    return os.path.join(output_dir, f"{OUTPUT_BASENAME}_part{part_index:03d}.txt")

def build_log_settings_text() -> str:
    return (
        "LOG_SETTINGS\n"
        f"include_docx_files = {INCLUDE_DOCX_FILES}\n"
        f"include_py_files   = {INCLUDE_PY_FILES}\n"
        f"include_txt_files  = {INCLUDE_TXT_FILES}\n"
        f"include_md_files   = {INCLUDE_MD_FILES}\n"
        f"include_pdf_files  = {INCLUDE_PDF_FILES}\n"
        f"include_npz_files  = {INCLUDE_NPZ_FILES}\n"
        f"include_zip_files  = {INCLUDE_ZIP_FILES}\n"
        f"include_docx_text  = {INCLUDE_DOCX_TEXT}\n"
        f"include_py_text    = {INCLUDE_PY_TEXT}\n"
        f"include_txt_text   = {INCLUDE_TXT_TEXT}\n"
        f"include_md_text    = {INCLUDE_MD_TEXT}\n"
        f"include_pdf_text   = {INCLUDE_PDF_TEXT}\n"
        f"include_npz_text   = {INCLUDE_NPZ_TEXT}\n"
        f"include_zip_text   = {INCLUDE_ZIP_TEXT}\n"
        f"only_filename      = {ONLY_FILENAME}\n"
        f"time_mode          = {TIME_MODE}\n"
        f"num_workers        = {NUM_WORKERS}\n"
        f"max_output_mb      = {MAX_OUTPUT_TXT_SIZE_MB}\n"
        "================================================================================\n\n"
    )

def format_entry(ts2: float, path2: str, type_str: str, content: str) -> str:
    # Enforce filename-only in headers if requested (fast; avoids later cleanup pass)
    label = os.path.basename(path2) if FORCE_FILENAME_ONLY_IN_HEADERS else safe_label(path2)
    return (
        "=" * 80 + "\n"
        f"TIMESTAMP ({TIME_MODE}): {ts2}\n"
        f"FILE: {label}\n"
        f"TYPE: {type_str}\n"
        + "=" * 80 + "\n\n"
        + content.rstrip("\n") + "\n\n"
    )

# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
def main():
    VERBOSE_SCAN_FOLDERS = False   # printing folders is slow
    USE_PROGRESS_BAR = True

    input_dir = input("Directory to scan: ").strip()
    if not os.path.isdir(input_dir):
        print("Invalid directory path.")
        return

    max_bytes = int(MAX_OUTPUT_TXT_SIZE_MB * 1024 * 1024)
    output_dir = input_dir

    files_to_process: list[tuple[float, str]] = []

    print("\n--- PHASE 1: Scanning directory tree ---")
    for root, dirs, files in os.walk(input_dir, onerror=print):
        if VERBOSE_SCAN_FOLDERS:
            print(f"Scanning folder: {root}")

        dirs[:] = [d for d in dirs if not folder_is_excluded(d, FOLDER_EXCLUDES)]

        for name in files:
            # Exclude any output parts (or old output) from scanning
            lower_name = name.lower()
            if lower_name.startswith(OUTPUT_BASENAME.lower()) and lower_name.endswith(".txt"):
                continue

            path = os.path.join(root, name)

            lower = lower_name
            if not (
                (INCLUDE_DOCX_FILES and lower.endswith(".docx")) or
                (INCLUDE_PY_FILES   and lower.endswith(".py"))   or
                (INCLUDE_TXT_FILES  and lower.endswith(".txt"))  or
                (INCLUDE_MD_FILES   and lower.endswith(".md"))   or
                (INCLUDE_PDF_FILES  and lower.endswith(".pdf"))  or
                (INCLUDE_NPZ_FILES  and lower.endswith(".npz"))  or
                (INCLUDE_ZIP_FILES  and lower.endswith(".zip"))
            ):
                continue

            if matches_any_pattern(name, FILE_EXCLUDES):
                continue

            ts = get_sort_timestamp(path, TIME_MODE)
            files_to_process.append((ts, path))

    files_to_process.sort(key=lambda x: x[0])
    print(f"Found {len(files_to_process)} files to process.")
    print(f"Using NUM_WORKERS = {NUM_WORKERS}")
    print("--- PHASE 2: Extracting in parallel + writing in-order (lower memory) ---\n")

    # Prepare output part 001
    written_parts: list[str] = []
    part_index = 1
    current_path = make_part_path(output_dir, part_index)
    out = open(current_path, "w", encoding="utf-8")
    written_parts.append(current_path)

    header_text = build_log_settings_text()
    header_bytes = len(header_text.encode("utf-8", errors="replace"))
    out.write(header_text)
    current_bytes = header_bytes

    def rotate_part():
        nonlocal out, part_index, current_path, current_bytes
        out.close()
        part_index += 1
        current_path = make_part_path(output_dir, part_index)
        out = open(current_path, "w", encoding="utf-8")
        written_parts.append(current_path)
        out.write(header_text)
        current_bytes = header_bytes

    # Counts
    docx_count = py_count = txt_count = md_count = pdf_count = npz_count = zip_count = 0
    pdf_skipped_no_text = 0

    # Submit work with stable indices so we can write in chronological order
    futures = {}
    with ThreadPoolExecutor(max_workers=NUM_WORKERS) as ex:
        for idx, (ts, path) in enumerate(files_to_process):
            fut = ex.submit(process_one, idx, ts, path)
            futures[fut] = idx

        iterator = as_completed(futures)
        if USE_PROGRESS_BAR:
            iterator = tqdm(iterator, total=len(futures), desc="Extracting", unit="file")

        # Buffer completed results until we can write them in order
        ready: dict[int, tuple[int, float, str, str, str, str]] = {}
        next_idx_to_write = 0

        for fut in iterator:
            index, ts2, path2, type_str, content, status = fut.result()
            ready[index] = (index, ts2, path2, type_str, content, status)

            # Write any contiguous ready results (chronological) immediately
            while next_idx_to_write in ready:
                _, ts_w, path_w, type_w, content_w, status_w = ready.pop(next_idx_to_write)

                entry = format_entry(ts_w, path_w, type_w, content_w)
                entry_bytes = len(entry.encode("utf-8", errors="replace"))

                # Roll over BEFORE writing to avoid splitting blocks
                if (current_bytes + entry_bytes) > max_bytes and current_bytes > header_bytes:
                    rotate_part()

                out.write(entry)
                current_bytes += entry_bytes

                # Update counts
                if type_w == "DOCX":
                    docx_count += 1
                elif type_w == "PYTHON":
                    py_count += 1
                elif type_w == "TEXT":
                    txt_count += 1
                elif type_w == "MARKDOWN":
                    md_count += 1
                elif type_w == "PDF":
                    if status_w == "ok":
                        pdf_count += 1
                    elif status_w == "skipped_no_text":
                        pdf_skipped_no_text += 1
                elif type_w == "NPZ":
                    npz_count += 1
                elif type_w == "ZIP":
                    zip_count += 1

                next_idx_to_write += 1

    out.close()

    print(
        f"Done! Wrote {docx_count} DOCX, {py_count} PY, {txt_count} TXT, "
        f"{md_count} MD, {pdf_count} PDF (text), {npz_count} NPZ, "
        f"and {zip_count} ZIP entries."
    )
    if pdf_skipped_no_text:
        print(f"Skipped {pdf_skipped_no_text} PDF(s) with no extractable text.")

    print("\nOutput parts written:")
    for p in written_parts:
        print(" - " + p)

    print("Changes applied: no post-cleanup reread; in-order writing while extracting; regex precompiled.")

if __name__ == "__main__":
    main()
