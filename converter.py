"""
converter.py — File Changer core conversion engine
Security hardened: macro blocking, path traversal protection, file size limits, magic byte validation
"""

import os
import struct
import comtypes.client

# ── Allowed extensions ─────────────────────────────────────────────────────────
WORD_EXTENSIONS = {".doc", ".docx"}
PPT_EXTENSIONS  = {".ppt", ".pptx", ".pps", ".ppsx"}
CONVERTIBLE_EXTENSIONS = WORD_EXTENSIONS | PPT_EXTENSIONS

# ── Security constants ─────────────────────────────────────────────────────────
MAX_FILE_SIZE_MB = 100                        # reject files larger than 100 MB
MAX_FILE_SIZE    = MAX_FILE_SIZE_MB * 1024 * 1024

# Magic bytes (file signatures) for supported formats
# OLE2 compound document: .doc .ppt .xls (old Office)
OLE2_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
# ZIP/OOXML: .docx .pptx .xlsx (modern Office — all are ZIP archives)
ZIP_MAGIC  = b"PK\x03\x04"

VALID_MAGIC = {OLE2_MAGIC, ZIP_MAGIC}

OUTPUT_FORMATS = {
    "PDF (.pdf)": {
        "ext":  ".pdf",
        "word": 17,
        "ppt":  32,
    },
    "Word (.docx)": {
        "ext":  ".docx",
        "word": 16,
        "ppt":  None,
    },
    "Word (.doc)": {
        "ext":  ".doc",
        "word": 0,
        "ppt":  None,
    },
    "PowerPoint (.pptx)": {
        "ext":  ".pptx",
        "word": None,
        "ppt":  24,
    },
    "PowerPoint (.ppt)": {
        "ext":  ".ppt",
        "word": None,
        "ppt":  1,
    },
}


# ── Security helpers ───────────────────────────────────────────────────────────

def _check_file_size(file_path):
    """Returns (ok, reason). Rejects files exceeding MAX_FILE_SIZE."""
    try:
        size = os.path.getsize(file_path)
    except OSError:
        return False, "Cannot read file size"
    if size > MAX_FILE_SIZE:
        mb = size / (1024 * 1024)
        return False, f"File too large ({mb:.1f} MB, limit {MAX_FILE_SIZE_MB} MB)"
    if size == 0:
        return False, "File is empty"
    return True, None


def _check_magic_bytes(file_path):
    """Returns (ok, reason). Validates file signature matches expected Office format."""
    try:
        with open(file_path, "rb") as fh:
            header = fh.read(8)
    except OSError as exc:
        return False, f"Cannot read file: {exc}"

    for magic in VALID_MAGIC:
        if header.startswith(magic):
            return True, None

    return False, "File signature does not match a valid Office document (possible renamed or corrupt file)"


def _check_path_safety(file_path, base_dirs=None):
    """
    Returns (ok, reason).
    Blocks path traversal attempts and optionally restricts to allowed base dirs.
    """
    abs_path = os.path.realpath(file_path)          # resolves symlinks + ..
    original = os.path.abspath(file_path)

    # Detect path traversal: realpath differs from abspath after resolving symlinks
    if abs_path != original:
        return False, "Path traversal attempt detected"

    if base_dirs:
        for base in base_dirs:
            base = os.path.realpath(base)
            if abs_path.startswith(base + os.sep) or abs_path == base:
                return True, None
        return False, "File is outside allowed directories"

    return True, None


def validate_file(file_path):
    """
    Full security validation pipeline.
    Returns (ok: bool, reason: str | None).
    """
    ok, reason = _check_path_safety(file_path)
    if not ok:
        return False, reason

    ok, reason = _check_file_size(file_path)
    if not ok:
        return False, reason

    ok, reason = _check_magic_bytes(file_path)
    if not ok:
        return False, reason

    return True, None


# ── Public helpers ─────────────────────────────────────────────────────────────

def get_output_settings(format_label):
    return OUTPUT_FORMATS.get(format_label, OUTPUT_FORMATS["PDF (.pdf)"])


def collect_files(target):
    if isinstance(target, list):
        return [os.path.abspath(p) for p in target if os.path.isfile(p)]

    target = os.path.abspath(target)
    if os.path.isdir(target):
        files = []
        for name in sorted(os.listdir(target)):
            full_path = os.path.join(target, name)
            if os.path.isfile(full_path):
                files.append(full_path)
        return files

    if os.path.isfile(target):
        return [target]

    return []


def can_convert(file_path, output_ext):
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()

    if ext not in CONVERTIBLE_EXTENSIONS:
        return False
    if ext in WORD_EXTENSIONS:
        return output_ext in {".pdf", ".docx", ".doc"}
    if ext in PPT_EXTENSIONS:
        return output_ext in {".pdf", ".pptx", ".ppt"}
    return False


def preview_files(file_paths, output_ext):
    summary = {"convertible": 0, "skipped": 0, "unsupported": 0, "total": 0}

    for file_path in collect_files(file_paths):
        summary["total"] += 1
        if not can_convert(file_path, output_ext):
            summary["unsupported"] += 1
            continue
        output_path = f"{os.path.splitext(file_path)[0]}{output_ext}"
        if os.path.exists(output_path):
            summary["skipped"] += 1
        else:
            summary["convertible"] += 1

    return summary


# ── Conversion helpers ─────────────────────────────────────────────────────────

def convert_word_file(word_app, input_path, output_path, file_format):
    try:
        # AddToRecentFiles=False  — don't pollute MRU list
        doc = word_app.Documents.Open(
            input_path,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        doc.SaveAs(output_path, FileFormat=file_format)
        doc.Close(SaveChanges=False)
        return True, os.path.basename(input_path)
    except Exception as exc:
        return False, f"{os.path.basename(input_path)}: {exc}"


def convert_ppt_file(ppt_app, input_path, output_path, file_format):
    try:
        deck = ppt_app.Presentations.Open(input_path, WithWindow=False, ReadOnly=True)
        deck.SaveAs(output_path, FileFormat=file_format)
        deck.Close()
        return True, os.path.basename(input_path)
    except Exception as exc:
        return False, f"{os.path.basename(input_path)}: {exc}"


def process_file(file_path, word_app, ppt_app, stats, output_settings, log=None):
    def write(msg):
        if log:
            log(msg)

    abs_path   = os.path.abspath(file_path)
    output_ext = output_settings["ext"]

    if not os.path.exists(abs_path):
        write(f"[WARN] File not found: {file_path}")
        stats["not_found"] += 1
        return

    # ── Security validation ────────────────────────────────────────────────
    ok, reason = validate_file(abs_path)
    if not ok:
        write(f"[SECURITY] Blocked {os.path.basename(abs_path)}: {reason}")
        stats["failed"] += 1
        return

    if not can_convert(abs_path, output_ext):
        write(f"[SKIP] Cannot convert {os.path.basename(file_path)} to {output_ext}")
        stats["unsupported"] += 1
        return

    output_path = f"{os.path.splitext(abs_path)[0]}{output_ext}"
    if os.path.exists(output_path):
        write(f"[SKIP] Output already exists: {os.path.basename(output_path)}")
        stats["skipped"] += 1
        return

    _, ext = os.path.splitext(abs_path)
    ext = ext.lower()

    if ext in WORD_EXTENSIONS:
        ok, detail = convert_word_file(word_app, abs_path, output_path, output_settings["word"])
    else:
        ok, detail = convert_ppt_file(ppt_app, abs_path, output_path, output_settings["ppt"])

    if ok:
        write(f"[OK] Converted: {detail} -> {output_ext}")
        stats["converted"] += 1
    else:
        write(f"[ERROR] {detail}")
        stats["failed"] += 1


# ── Main batch runner ──────────────────────────────────────────────────────────

def batch_pdf_converter(
    target,
    output_format="PDF (.pdf)",
    log=None,
    progress=None,
    cancel_check=None,
):
    def write(msg):
        if log:
            log(msg)

    files           = collect_files(target)
    output_settings = get_output_settings(output_format)

    if not files:
        write("[ERROR] No files selected.")
        return {"converted": 0, "skipped": 0, "unsupported": 0, "not_found": 0, "failed": 0}

    write("Initializing Microsoft Word and PowerPoint...")

    word_app = comtypes.client.CreateObject("Word.Application")
    word_app.Visible = False
    # ── SECURITY: Disable all macros in Word ──────────────────────────────
    word_app.AutomationSecurity = 3   # msoAutomationSecurityForceDisable

    ppt_app = comtypes.client.CreateObject("Powerpoint.Application")
    ppt_app.Visible = False           # keep PowerPoint hidden too
    # ── SECURITY: Disable all macros in PowerPoint ────────────────────────
    ppt_app.AutomationSecurity = 3    # msoAutomationSecurityForceDisable

    write("[SECURITY] Macros disabled for this session.")

    stats = {"converted": 0, "skipped": 0, "unsupported": 0, "not_found": 0, "failed": 0}
    total = len(files)

    try:
        for index, file_path in enumerate(files, start=1):
            if cancel_check and cancel_check():
                write("[STOP] Conversion cancelled.")
                break
            process_file(file_path, word_app, ppt_app, stats, output_settings, log=log)
            if progress:
                progress(index, total)
    finally:
        write("Closing Office applications...")
        try:
            word_app.Quit()
        except Exception:
            pass
        try:
            ppt_app.Quit()
        except Exception:
            pass
        write(
            f"Finished.  Converted: {stats['converted']},  "
            f"Skipped: {stats['skipped']},  Failed: {stats['failed']},  "
            f"Unsupported: {stats['unsupported']}"
        )

    return stats
