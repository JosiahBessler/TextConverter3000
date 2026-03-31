# TextConverter3000

TextConverter3000 is an open-source data compiler for development history.

It scans a directory recursively, sorts files by timestamp, and compiles a chronological work log with:
- file timestamps,
- file names,
- file types,
- extracted content (or optional metadata-only output).

This is useful for documenting the evolution of projects and creating auditable process logs.

## What it supports
The main script (`Docx to Text converter V2.py`) can include/extract from:
- `.docx`
- `.py`
- `.txt`
- `.pdf` (extractable text PDFs)
- `.npz` (NumPy archive summary)
- `.zip` (archive file listing summary)

## Quick start

## 1) Install dependencies
```bash
pip install python-docx tqdm pypdf numpy
```

## 2) Run the compiler
```bash
python "Docx to Text converter V2.py"
```

When prompted, enter the directory you want to scan.

## 3) Review outputs
The tool writes one or more output files into the scanned directory:
- `Output_data_part001.txt`
- `Output_data_part002.txt`
- etc.

Output is automatically split by size using `MAX_OUTPUT_TXT_SIZE_MB`.

## Configuration switches (top of script)
Set these booleans near the top of the script:

### File-type include switches
- `INCLUDE_DOCX_FILES`
- `INCLUDE_PY_FILES`
- `INCLUDE_TXT_FILES`
- `INCLUDE_PDF_FILES`
- `INCLUDE_NPZ_FILES`
- `INCLUDE_ZIP_FILES`

### Content extraction switches
- `INCLUDE_DOCX_TEXT`
- `INCLUDE_PY_TEXT`
- `INCLUDE_TXT_TEXT`
- `INCLUDE_PDF_TEXT`
- `INCLUDE_NPZ_TEXT`
- `INCLUDE_ZIP_TEXT`

If a file type is enabled but text extraction is disabled, the file is still logged with a `CONTENT OMITTED` placeholder.

## Important options
- `TIME_MODE`: choose `"ctime"` or `"mtime"` for sorting chronology.
- `MAX_OUTPUT_TXT_SIZE_MB`: max size per output part.
- `NUM_WORKERS`: parallel extraction worker count.
- `ONLY_FILENAME`: hides directory path and logs only file names.
- `REDACT_ABS_PATHS_IN_TEXT`: redacts Windows/UNC absolute paths in extracted text.
- `FILE_EXCLUDES` / `FOLDER_EXCLUDES`: pattern-based filtering.

## Notes on NPZ and ZIP handling
- **NPZ**: the tool logs an array summary (array key, shape, dtype, size), not raw binary payload.
- **ZIP**: the tool logs archive entry names and sizes, not file contents.

These defaults are intended to keep logs readable and avoid dumping large binary data into text output.

## Typical workflow
1. Enable the file types you want.
2. Choose whether each file type should include full text or metadata-only placeholders.
3. Run the script after a coding session.
4. Keep generated output parts as your timestamped development trail.

## License / open-source use
This project is suitable for open-source use as a lightweight project process logger. If you plan to share logs publicly, review redaction settings before generating output.
