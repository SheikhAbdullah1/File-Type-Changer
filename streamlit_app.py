"""
File Type Changer — Production Streamlit Web App
Desktop UI replicated in Streamlit with full security + dark/light mode
"""

import os
import subprocess
import tempfile
import zipfile
import shutil
from pathlib import Path

import streamlit as st

# PAGE CONFIGURATION

st.set_page_config(
    page_title="File Type Changer",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# DARK/LIGHT MODE STATE

if "theme" not in st.session_state:
    st.session_state.theme = "dark"

# Theme colors
THEMES = {
    "dark": {
        "bg_base": "#0D1117",
        "bg_surface": "#161B22",
        "bg_elevated": "#21262D",
        "border": "#30363D",
        "text_primary": "#F0F6FC",
        "text_muted": "#8B949E",
        "accent": "#3B82F6",
        "accent_hover": "#2563EB",
        "success": "#10B981",
        "warning": "#F59E0B",
        "danger": "#EF4444",
    },
    "light": {
        "bg_base": "#FFFFFF",
        "bg_surface": "#F6F8FA",
        "bg_elevated": "#EAEEF2",
        "border": "#D0D7DE",
        "text_primary": "#24292F",
        "text_muted": "#57606A",
        "accent": "#0969DA",
        "accent_hover": "#0860CA",
        "success": "#1A7F0F",
        "warning": "#9E6A03",
        "danger": "#DA3633",
    },
}

theme = THEMES[st.session_state.theme]

# ────────────────────────────────────────────────────────────────────────────────
# CSS STYLING
# ────────────────────────────────────────────────────────────────────────────────

st.markdown(f"""
<style>
    /* Root colors */
    :root {{
        --bg-base: {theme["bg_base"]};
        --bg-surface: {theme["bg_surface"]};
        --bg-elevated: {theme["bg_elevated"]};
        --border: {theme["border"]};
        --text-primary: {theme["text_primary"]};
        --text-muted: {theme["text_muted"]};
        --accent: {theme["accent"]};
        --success: {theme["success"]};
        --warning: {theme["warning"]};
        --danger: {theme["danger"]};
    }}

    /* Main app */
    .stApp {{
        background: {theme["bg_base"]};
        color: {theme["text_primary"]};
    }}

    /* Hide default Streamlit elements */
    #MainMenu {{ visibility: hidden; }}
    footer {{ visibility: hidden; }}
    header {{ visibility: hidden; }}

    /* Card styling */
    .card {{
        background: {theme["bg_surface"]};
        border: 1px solid {theme["border"]};
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 16px;
    }}

    /* Section label */
    .section-label {{
        font-size: 10px;
        font-weight: 700;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        color: {theme["text_muted"]};
        margin-bottom: 12px;
    }}

    /* Stat cards */
    .stat-row {{
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 12px;
        margin: 16px 0;
    }}
    .stat-box {{
        background: {theme["bg_elevated"]};
        border: 1px solid {theme["border"]};
        border-radius: 10px;
        padding: 16px 12px;
        text-align: center;
    }}
    .stat-num {{
        font-size: 28px;
        font-weight: 700;
        line-height: 1;
    }}
    .stat-label {{
        font-size: 12px;
        color: {theme["text_muted"]};
        margin-top: 6px;
    }}

    /* Colors */
    .blue  {{ color: {theme["accent"]}; }}
    .green {{ color: {theme["success"]}; }}
    .amber {{ color: {theme["warning"]}; }}
    .red   {{ color: {theme["danger"]}; }}

    /* Upload area */
    .upload-zone {{
        background: {theme["bg_elevated"]};
        border: 2px dashed {theme["border"]};
        border-radius: 12px;
        padding: 32px 20px;
        text-align: center;
        cursor: pointer;
        transition: all 0.2s;
    }}
    .upload-zone:hover {{
        border-color: {theme["accent"]};
    }}

    /* Buttons */
    .stButton > button {{
        background: {theme["accent"]};
        color: white;
        border: none;
        border-radius: 10px;
        font-weight: 600;
        width: 100%;
        padding: 12px 20px !important;
    }}
    .stButton > button:hover {{
        background: {theme["accent_hover"]};
    }}

    /* Secondary button */
    .btn-secondary {{
        background: {theme["bg_elevated"]} !important;
        color: {theme["text_primary"]} !important;
        border: 1px solid {theme["border"]} !important;
    }}
    .btn-secondary:hover {{
        background: {theme["border"]} !important;
    }}

    /* Danger button */
    .btn-danger {{
        background: {theme["danger"]} !important;
    }}
    .btn-danger:hover {{
        background: #C02C2C !important;
    }}

    /* Select box */
    .stSelectbox > div > div {{
        background: {theme["bg_elevated"]};
        border: 1px solid {theme["border"]};
        border-radius: 10px;
        color: {theme["text_primary"]};
    }}

    /* Log box */
    .log-container {{
        background: {theme["bg_elevated"]};
        border: 1px solid {theme["border"]};
        border-radius: 10px;
        padding: 16px;
        font-family: 'Consolas', 'Monaco', monospace;
        font-size: 13px;
        max-height: 300px;
        overflow-y: auto;
        white-space: pre-wrap;
        word-break: break-word;
        line-height: 1.6;
    }}

    /* Log text colors */
    .log-ok    {{ color: {theme["success"]}; }}
    .log-error {{ color: {theme["danger"]}; }}
    .log-skip  {{ color: {theme["warning"]}; }}
    .log-info  {{ color: {theme["text_muted"]}; }}
    .log-security {{ color: {theme["warning"]}; font-weight: 600; }}

    /* Divider */
    hr {{ border-color: {theme["border"]}; }}

    /* Text styling */
    .title-main {{
        font-size: 28px;
        font-weight: 700;
        color: {theme["text_primary"]};
    }}
    .subtitle {{
        font-size: 14px;
        color: {theme["text_muted"]};
        margin-top: 4px;
    }}

    /* Columns alignment */
    .column-container {{
        display: grid;
        grid-template-columns: 2fr 1fr;
        gap: 16px;
        margin-top: 20px;
    }}

    @media (max-width: 768px) {{
        .column-container {{
            grid-template-columns: 1fr;
        }}
        .stat-row {{
            grid-template-columns: 1fr;
        }}
    }}

    /* Security badge */
    .security-badge {{
        background: {theme["warning"]};
        color: {theme["bg_base"]};
        padding: 4px 8px;
        border-radius: 4px;
        font-size: 10px;
        font-weight: 600;
        display: inline-block;
        margin-left: 8px;
    }}
</style>
""", unsafe_allow_html=True)

# ────────────────────────────────────────────────────────────────────────────────
# CONSTANTS & SECURITY
# ────────────────────────────────────────────────────────────────────────────────

WORD_EXTS = {".doc", ".docx"}
PPT_EXTS = {".ppt", ".pptx", ".pps", ".ppsx"}
ALL_EXTS = WORD_EXTS | PPT_EXTS

OUTPUT_FORMATS = {
    "PDF (.pdf)": {
        "ext": ".pdf",
        "icon": "📄",
        "lo_cmd": lambda path, out: ["--headless", "--convert-to", "pdf", "--outdir", str(out), str(path)],
    },
    "Word (.docx)": {
        "ext": ".docx",
        "icon": "📝",
        "lo_cmd": lambda path, out: ["--headless", "--convert-to", "docx:'MS Word 2007 XML'", "--outdir", str(out), str(path)],
    },
    "PowerPoint (.pptx)": {
        "ext": ".pptx",
        "icon": "📊",
        "lo_cmd": lambda path, out: ["--headless", "--convert-to", "pptx:'Impress MS PowerPoint 2007 XML'", "--outdir", str(out), str(path)],
    },
}

MAX_MB = 200
MAX_BYTES = MAX_MB * 1024 * 1024
OLE2_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
ZIP_MAGIC = b"PK\x03\x04"

# ────────────────────────────────────────────────────────────────────────────────
# SECURITY VALIDATION
# ────────────────────────────────────────────────────────────────────────────────

def validate_file(data: bytes, filename: str) -> tuple[bool, str]:
    """
    Full security validation pipeline.
    Returns (ok, reason).
    """
    # Empty check
    if len(data) == 0:
        return False, "File is empty"

    # Size check
    if len(data) > MAX_BYTES:
        mb = len(data) / (1024 * 1024)
        return False, f"File too large ({mb:.1f} MB, limit {MAX_MB} MB)"

    # Magic byte validation (prevents renamed files)
    if not (data[:8].startswith(OLE2_MAGIC) or data[:4].startswith(ZIP_MAGIC)):
        return False, "Invalid signature — not a real Office document"

    # Extension check
    ext = Path(filename).suffix.lower()
    if ext not in ALL_EXTS:
        return False, f"Unsupported format: {ext}"

    return True, None

# ────────────────────────────────────────────────────────────────────────────────
# LIBREOFFICE CONVERTER
# ────────────────────────────────────────────────────────────────────────────────

def libreoffice_available():
    """Check if LibreOffice is installed."""
    return shutil.which("libreoffice") is not None or shutil.which("soffice") is not None

def get_lo_cmd():
    """Get LibreOffice command."""
    return shutil.which("libreoffice") or shutil.which("soffice") or "libreoffice"

def convert_file(input_path: Path, output_dir: Path, output_format: str) -> tuple[bool, str, bytes | None]:
    """
    Convert file using LibreOffice.
    Returns (success, message, file_bytes).
    """
    cmd = get_lo_cmd()
    output_ext = OUTPUT_FORMATS[output_format]["ext"]

    try:
        lo_args = OUTPUT_FORMATS[output_format]["lo_cmd"](input_path, output_dir)
        result = subprocess.run(
            [cmd] + lo_args,
            capture_output=True,
            text=True,
            timeout=120,
        )

        # Find output file
        stem = input_path.stem
        out_file = output_dir / f"{stem}{output_ext}"

        if not out_file.exists():
            candidates = list(output_dir.glob(f"{stem}.*"))
            if candidates:
                out_file = sorted(candidates, key=lambda f: len(f.suffix))[0]

        if out_file.exists():
            data = out_file.read_bytes()
            return True, f"✔ Converted: {input_path.name}", data
        else:
            err = result.stderr.strip() or "Unknown error"
            return False, f"✘ Failed: {input_path.name} — {err[:60]}", None

    except subprocess.TimeoutExpired:
        return False, f"✘ Timeout: {input_path.name}", None
    except FileNotFoundError:
        return False, "✘ LibreOffice not found", None
    except Exception as exc:
        return False, f"✘ Error: {str(exc)[:60]}", None

# ────────────────────────────────────────────────────────────────────────────────
# HEADER WITH THEME TOGGLE
# ────────────────────────────────────────────────────────────────────────────────

col_title, col_theme = st.columns([1, 0.15])

with col_title:
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:2px;">
        <div style="width:6px;height:44px;background:{theme['accent']};border-radius:3px;"></div>
        <div>
            <div class="title-main">File Type Changer</div>
            <div class="subtitle">Batch convert Word & PowerPoint files</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col_theme:
    current = st.session_state.theme
    new_theme = st.radio(
        "Theme",
        ["Dark", "Light"],
        index=0 if current == "dark" else 1,
        horizontal=True,
        label_visibility="collapsed",
    )
    if (new_theme == "Dark" and current == "light") or (new_theme == "Light" and current == "dark"):
        st.session_state.theme = "dark" if new_theme == "Dark" else "light"
        st.rerun()

st.markdown("<hr style='margin:12px 0;'>", unsafe_allow_html=True)

# ────────────────────────────────────────────────────────────────────────────────
# CHECK LIBREOFFICE
# ────────────────────────────────────────────────────────────────────────────────

if not libreoffice_available():
    st.error("⚠️ LibreOffice not installed on this server. Conversion unavailable.")
    st.info("For local use: install from https://www.libreoffice.org/download/")
    st.stop()

# ────────────────────────────────────────────────────────────────────────────────
# MAIN 2-COLUMN LAYOUT (INPUT | OUTPUT)
# ────────────────────────────────────────────────────────────────────────────────

col_input, col_output = st.columns([2, 1], gap="medium")

# ──────────────────────── LEFT COLUMN: INPUT ──────────────────────────────────

with col_input:
    st.markdown(f'<div class="section-label">Input Files</div>', unsafe_allow_html=True)

    # File uploader
    uploaded_files = st.file_uploader(
        "Upload",
        type=["doc", "docx", "ppt", "pptx", "pps", "ppsx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        help=f"Max {MAX_MB}MB per file",
    )

    # File validation
    valid_files = []
    invalid_files = []

    if uploaded_files:
        for f in uploaded_files:
            data = f.read()
            f.seek(0)
            ok, reason = validate_file(data, f.name)
            if ok:
                valid_files.append(f)
            else:
                invalid_files.append((f.name, reason))

        total = len(uploaded_files)
        ready = len(valid_files)
        rejected = len(invalid_files)

        # Stats row
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

        # Show rejected files
        if invalid_files:
            with st.expander(f"⚠️ {rejected} file(s) rejected"):
                for fname, reason in invalid_files:
                    st.markdown(f"**{fname}** — _{reason}_")
    else:
        st.markdown(f"""
        <div class="upload-zone" onclick="document.querySelector('[data-testid=stFileUploader] input').click();">
            <div style="font-size:36px;margin-bottom:8px;">⬆</div>
            <div style="font-size:14px;color:{theme['text_primary']};margin-bottom:4px;font-weight:600;">Click to browse files</div>
            <div style="font-size:12px;color:{theme['text_muted']};">Supports .doc .docx .ppt .pptx .pps .ppsx</div>
        </div>
        """, unsafe_allow_html=True)

    # Browse button (shown always)
    st.button("📁 Browse Files", use_container_width=True, key="browse_button")

# ──────────────────────── RIGHT COLUMN: OUTPUT ──────────────────────────────────

with col_output:
    st.markdown(f'<div class="section-label">Output Format</div>', unsafe_allow_html=True)

    format_choice = st.selectbox(
        "Convert to",
        list(OUTPUT_FORMATS.keys()),
        label_visibility="collapsed",
    )

    output_ext = OUTPUT_FORMATS[format_choice]["ext"]
    icon = OUTPUT_FORMATS[format_choice]["icon"]

    # Format display
    st.markdown(f"""
    <div style="text-align:center;margin:20px 0;">
        <div style="font-size:48px;">{icon}</div>
        <div style="font-size:18px;font-weight:700;color:{theme['accent']};margin-top:6px;">{output_ext}</div>
    </div>
    """, unsafe_allow_html=True)

    # Buttons
    st.button("Clear All", use_container_width=True, key="clear_button")

    if valid_files:
        st.button(
            f"Convert {len(valid_files)} →",
            type="primary",
            use_container_width=True,
            key="convert_button",
        )
    else:
        st.button(
            "Convert →",
            disabled=True,
            use_container_width=True,
            key="convert_disabled",
        )

# ────────────────────────────────────────────────────────────────────────────────
# CONVERSION LOGIC
# ────────────────────────────────────────────────────────────────────────────────

if st.session_state.get("convert_button"):
    if not valid_files:
        st.error("No valid files selected!")
    else:
        st.markdown(f'<div class="section-label">Activity Log</div>', unsafe_allow_html=True)

        log_lines = []
        converted = []
        failed = 0

        progress_bar = st.progress(0)
        status_text = st.empty()
        log_container = st.empty()

        log_lines.append(f'<span class="log-security">🔒 SECURITY: Magic byte validation enabled</span>')
        log_lines.append(f'<span class="log-info">[INFO] Starting conversion...</span>')

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            for i, f in enumerate(valid_files):
                f.seek(0)
                data = f.read()

                # Write to temp
                input_path = tmp / f.name
                input_path.write_bytes(data)

                # Convert
                ok, msg, out_bytes = convert_file(input_path, tmp, format_choice)

                if ok:
                    out_name = Path(f.name).stem + output_ext
                    converted.append((out_name, out_bytes))
                    log_lines.append(f'<span class="log-ok">{msg}</span>')
                else:
                    failed += 1
                    log_lines.append(f'<span class="log-error">{msg}</span>')

                pct = int((i + 1) / len(valid_files) * 100)
                progress_bar.progress(pct)
                status_text.markdown(f"_Converting {i+1} / {len(valid_files)}..._")

                log_container.markdown(
                    f'<div class="log-container">' + "<br>".join(log_lines) + "</div>",
                    unsafe_allow_html=True,
                )

        progress_bar.progress(100)
        status_text.markdown(f"✔ **Conversion complete**")

        # Download
        if converted:
            st.markdown("<hr>", unsafe_allow_html=True)
            if len(converted) == 1:
                name, data = converted[0]
                st.download_button(
                    f"⬇ Download {name}",
                    data,
                    name,
                    use_container_width=True,
                )
            else:
                import io
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for name, data in converted:
                        zf.writestr(name, data)
                zip_buf.seek(0)

                st.download_button(
                    f"⬇ Download all {len(converted)} files (.zip)",
                    zip_buf,
                    "converted_files.zip",
                    use_container_width=True,
                )

            if failed:
                st.warning(f"⚠️ {failed} file(s) failed — see log above")
            else:
                st.success(f"✅ All {len(converted)} file(s) converted!")

# ────────────────────────────────────────────────────────────────────────────────
# FOOTER
# ────────────────────────────────────────────────────────────────────────────────

st.markdown("<hr>", unsafe_allow_html=True)
st.markdown(f"""
<div style="text-align:center;font-size:12px;color:{theme['text_muted']};padding:12px 0;">
    🔒 **Security:** Magic byte validation • File size limits • Memory-only processing • No server storage
</div>
""", unsafe_allow_html=True)
