# File Type Changer 📄

[![Live Demo](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://file-type-changer.streamlit.app)

Batch convert Word and PowerPoint files to PDF, DOCX, PPTX — free, no installation needed.

---

## 🚀 Live Demo

👉 **[Click here to use the app](https://file-type-changer.streamlit.app)**

---

## Features

- Convert `.doc`, `.docx`, `.ppt`, `.pptx`, `.pps`, `.ppsx` files
- Output formats: PDF, Word (docx), PowerPoint (pptx)
- Batch conversion — download all as ZIP
- Files processed in memory, never stored on server
- Dark UI, mobile friendly

## Security

- **Magic byte validation** — real format check, not just extension
- **File size limit** — 50 MB max per file
- **In-memory processing** — files deleted after conversion
- **No storage** — nothing saved on server

---

## Run Locally

```bash
git clone https://github.com/SheikhAbdullah1/File-Type-Changer.git
cd File-Type-Changer

pip install -r requirements.txt
streamlit run streamlit_app.py
```

> LibreOffice must be installed: https://www.libreoffice.org/download/

---

## Project Structure

```
File-Type-Changer/
├── streamlit_app.py   # Web app (Streamlit)
├── app.py             # Desktop app (CustomTkinter, Windows only)
├── converter.py       # Desktop converter logic
├── main.py            # Desktop entry point
├── launch.bat         # Desktop one-click launcher
├── requirements.txt   # Python dependencies
├── packages.txt       # System packages (LibreOffice for Streamlit Cloud)
└── README.md
```

## Dependencies

| Package | Purpose |
|---------|---------|
| `streamlit` | Web UI framework |
| `libreoffice` | File conversion engine (server-side) |
| `customtkinter` | Desktop UI (Windows only) |
| `comtypes` | Windows COM automation (Desktop only) |

---

## License

MIT
