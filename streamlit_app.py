"""
File Type Changer — Streamlit Web App
Converts Word/PowerPoint files using LibreOffice (no MS Office needed)
"""

import os
import subprocess
import tempfile
import zipfile
import shutil
from pathlib import Path

import streamlit as st

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="File Type Changer",
    page_icon="📄",
    layout="centered",
)

# ── Styling ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Main background */
    .stApp { background-color: #0D1117; color: #F0F6FC; }

    /* Cards */
    .upload-card {
        background: #161B22;
        border: 1px solid #30363D;
        border-radius: 14px;
        padding: 28px 24px;
        margin-bottom: 16px;
    }

    /* Section eyebrow */
    .eyebrow {
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        color: #8B949E;
        margin-bottom: 6px;
    }

    /* Stat boxes */
    .stat-row { display: flex; gap: 12px; margin: 16px 0; }
    .stat-box {
        flex: 1;
        background: #21262D;
        border: 1px solid #30363D;
        border-radius: 10px;
        padding: 14px 10px;
        text-align: center;
    }
    .stat-num  { font-size: 26px; font-weight: 700; }
    .stat-label { font-size: 12px; color: #8B949E; margin-top: 2px; }
    .blue  { color: #3B82F6; }
    .green { color: #10B981; }
    .amber { color: #F59E0B; }
    .red   { color: #EF4444; }

    /* Hide Streamlit branding */
    #MainMenu, footer, header { visibility: hidden; }

    /* Button overrides */
    .stButton > button {
        background: #3B82F6;
        color: white;
        border: none;
        border-radius: 10px;
        font-weight: 600;
        padding: 10px 24px;
        width: 100%;
    }
    .stButton > button:hover { background: #2563EB; }

    /* File uploader */
    .stFileUploader > div {
        background: #161B22;
        border: 2px dashed #30363D;
        border-radius: 12px;
    }

    /* Select box */
    .stSelectbox > div > div {
        background: #21262D;
        border: 1px solid #30363D;
        border-radius: 10px;
        color: #F0F6FC;
    }

    /* Log box */
    .log-box {
        background: #21262D;
        border: 1px solid #30363D;
        border-radius: 10px;
        padding: 16px;
        font-family: 'Consolas', monospace;
        font-size: 13px;
        max-height: 280px;
        overflow-y: auto;
        white-space: pre-wrap;
    }
    .log-ok    { color: #10B981; }
    .log-error { color: #EF4444; }
    .log-skip  { color: #F59E0B; }
    .log-info  { color: #8B949E; }
</style>
""", unsafe_allow_html=True)

# ── Constants ──────────────────────────────────────────────────────────────────
WORD_EXTS = {".doc", ".docx"}
PPT_EXTS  = {".ppt", ".pptx", ".pps", ".ppsx"}
ALL_EXTS  = WORD_EXTS | PPT_EXTS

OUTPUT_FORMATS = {
    "PDF (.pdf)":         {"ext": ".pdf",  "lo_filter": "writer_pdf_Export", "ppt_filter": "impress_pdf_Export"},
    "Word (.docx)":       {"ext": ".docx", "lo_filter": "MS Word 2007 XML",  "ppt_filter": None},
    "PowerPoint (.pptx)": {"ext": ".pptx", "lo_filter": None,                "ppt_filter": "Impress MS PowerPoint 2007 XML"},
}

MAX_MB        = 50
MAX_BYTES     = MAX_MB * 1024 * 1024
OLE2_MAGIC    = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
ZIP_MAGIC     = b"PK\x03\x04"

# ── Security helpers ───────────────────────────────────────────────────────────

def validate_bytes(data: bytes, filename: str):
    """Returns (ok, reason)."""
    if len(data) == 0:
        return False, "File is empty"
    if len(data) > MAX_BYTES:
        return False, f"File exceeds {MAX_MB} MB limit"
    if not (data[:8].startswith(OLE2_MAGIC) or data[:4].startswith(ZIP_MAGIC)):
        return False, "Invalid file format — not a real Office document"
    ext = Path(filename).suffix.lower()
    if ext not in ALL_EXTS:
        return False, f"Unsupported extension: {ext}"
    return True, None

# ── LibreOffice converter ──────────────────────────────────────────────────────

def libreoffice_available():
    return shutil.which("libreoffice") is not None or shutil.which("soffice") is not None

def get_lo_cmd():
    return shutil.which("libreoffice") or shutil.which("soffice") or "libreoffice"

def convert_file(input_path: Path, output_dir: Path, output_ext: str) -> tuple[bool, str, bytes | None]:
    """
    Convert a file using LibreOffice headless.
    Returns (success, message, file_bytes).
    """
    cmd = get_lo_cmd()
    ext = input_path.suffix.lower()

    # Choose LibreOffice filter
    if output_ext == ".pdf":
        lo_args = ["--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(input_path)]
    elif output_ext == ".docx":
        lo_args = ["--headless", "--convert-to", "docx:'MS Word 2007 XML'", "--outdir", str(output_dir), str(input_path)]
    elif output_ext == ".pptx":
        lo_args = ["--headless", "--convert-to", "pptx:'Impress MS PowerPoint 2007 XML'", "--outdir", str(output_dir), str(input_path)]
    else:
        return False, f"Unsupported output format: {output_ext}", None

    try:
        result = subprocess.run(
            [cmd] + lo_args,
            capture_output=True, text=True, timeout=60
        )
        # Find output file
        stem = input_path.stem
        out_file = output_dir / f"{stem}{output_ext}"

        # LibreOffice sometimes uses .pdf regardless — find any matching stem
        if not out_file.exists():
            candidates = list(output_dir.glob(f"{stem}.*"))
            if candidates:
                out_file = candidates[0]

        if out_file.exists():
            data = out_file.read_bytes()
            return True, f"✔ Converted: {input_path.name}", data
        else:
            err = result.stderr.strip() or result.stdout.strip() or "Unknown error"
            return False, f"✘ Failed: {input_path.name} — {err}", None

    except subprocess.TimeoutExpired:
        return False, f"✘ Timeout: {input_path.name} took too long", None
    except FileNotFoundError:
        return False, "LibreOffice not found on this server", None
    except Exception as exc:
        return False, f"✘ Error: {input_path.name} — {exc}", None

# ── UI ─────────────────────────────────────────────────────────────────────────

# Header
st.markdown("""
<div style="display:flex;align-items:center;gap:14px;margin-bottom:6px;">
  <div style="width:6px;height:40px;background:#3B82F6;border-radius:3px;"></div>
  <div>
    <div style="font-size:26px;font-weight:700;color:#F0F6FC;">File Type Changer</div>
    <div style="font-size:14px;color:#8B949E;">Batch convert Word & PowerPoint files — free, no install needed</div>
  </div>
</div>
<hr style="border-color:#30363D;margin:16px 0 20px 0;">
""", unsafe_allow_html=True)

# Check LibreOffice
if not libreoffice_available():
    st.error("⚠️ LibreOffice is not installed on this server. Conversion will not work.")
    st.info("If running locally: install LibreOffice from https://www.libreoffice.org/download/")
    st.stop()

# ── Upload section ─────────────────────────────────────────────────────────────
st.markdown('<div class="eyebrow">Input Files</div>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    label="Upload files",
    type=["doc", "docx", "ppt", "pptx", "pps", "ppsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help=f"Supported: .doc .docx .ppt .pptx .pps .ppsx — Max {MAX_MB} MB each",
)

# ── Format section ─────────────────────────────────────────────────────────────
st.markdown('<div class="eyebrow" style="margin-top:20px;">Output Format</div>', unsafe_allow_html=True)

col1, col2 = st.columns([2, 1])
with col1:
    format_choice = st.selectbox(
        "Convert to",
        options=list(OUTPUT_FORMATS.keys()),
        label_visibility="collapsed",
    )

output_ext = OUTPUT_FORMATS[format_choice]["ext"]

# ── File preview stats ─────────────────────────────────────────────────────────
if uploaded_files:
    valid_files   = []
    invalid_files = []

    for f in uploaded_files:
        data = f.read()
        f.seek(0)
        ok, reason = validate_bytes(data, f.name)
        if ok:
            valid_files.append(f)
        else:
            invalid_files.append((f.name, reason))

    total    = len(uploaded_files)
    ready    = len(valid_files)
    rejected = len(invalid_files)

    st.markdown(f"""
    <div class="stat-row">
        <div class="stat-box">
            <div class="stat-num blue">{total}</div>
            <div class="stat-label">Selected</div>
        </div>
        <div class="stat-box">
            <div class="stat-num green">{ready}</div>
            <div class="stat-label">Ready</div>
        </div>
        <div class="stat-box">
            <div class="stat-num red">{rejected}</div>
            <div class="stat-label">Rejected</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if invalid_files:
        with st.expander(f"⚠️ {rejected} file(s) rejected"):
            for name, reason in invalid_files:
                st.markdown(f"**{name}** — {reason}")

# ── Convert button ─────────────────────────────────────────────────────────────
    if ready > 0:
        if st.button(f"Convert {ready} file(s) to {output_ext}  →", type="primary"):

            log_lines  = []
            converted  = []
            failed     = 0

            progress_bar = st.progress(0, text="Starting conversion...")
            log_placeholder = st.empty()

            with tempfile.TemporaryDirectory() as tmpdir:
                tmp = Path(tmpdir)

                for i, f in enumerate(valid_files):
                    f.seek(0)
                    data = f.read()

                    # Write uploaded file to temp dir
                    input_path = tmp / f.name
                    input_path.write_bytes(data)

                    # Convert
                    ok, msg, out_bytes = convert_file(input_path, tmp, output_ext)

                    if ok:
                        out_name = Path(f.name).stem + output_ext
                        converted.append((out_name, out_bytes))
                        log_lines.append(f'<span class="log-ok">[OK]    {msg}</span>')
                    else:
                        failed += 1
                        log_lines.append(f'<span class="log-error">[ERROR] {msg}</span>')

                    pct = int((i + 1) / ready * 100)
                    progress_bar.progress(pct, text=f"Converting {i+1} / {ready}...")

                    log_placeholder.markdown(
                        f'<div class="log-box">' + "<br>".join(log_lines) + "</div>",
                        unsafe_allow_html=True,
                    )

            progress_bar.progress(100, text="Done!")

            # ── Download ───────────────────────────────────────────────────────
            if converted:
                if len(converted) == 1:
                    name, data = converted[0]
                    st.download_button(
                        label=f"⬇ Download {name}",
                        data=data,
                        file_name=name,
                        mime="application/octet-stream",
                        use_container_width=True,
                    )
                else:
                    # Multiple files — zip them
                    import io
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for name, data in converted:
                            zf.writestr(name, data)
                    zip_buf.seek(0)

                    st.download_button(
                        label=f"⬇ Download all {len(converted)} files (.zip)",
                        data=zip_buf,
                        file_name="converted_files.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

                if failed:
                    st.warning(f"{failed} file(s) failed — check log above.")
                else:
                    st.success(f"All {len(converted)} file(s) converted successfully!")

else:
    st.markdown("""
    <div style="text-align:center;padding:40px 20px;color:#8B949E;
                background:#161B22;border:2px dashed #30363D;border-radius:12px;">
        <div style="font-size:36px;">⬆</div>
        <div style="font-size:15px;margin-top:8px;">Upload files above to get started</div>
        <div style="font-size:12px;margin-top:4px;">Supports .doc .docx .ppt .pptx .pps .ppsx</div>
    </div>
    """, unsafe_allow_html=True)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("""
<hr style="border-color:#30363D;margin:32px 0 12px 0;">
<div style="text-align:center;font-size:12px;color:#8B949E;">
    File Changer — Files are processed in memory and never stored on the server
</div>
""", unsafe_allow_html=True)
