# File Changer 📄

Batch convert Word and PowerPoint files to PDF, DOCX, PPTX and more — with a clean desktop UI.

---

## Features

- Convert `.doc`, `.docx`, `.ppt`, `.pptx`, `.pps`, `.ppsx` files
- Output formats: PDF, Word (docx/doc), PowerPoint (pptx/ppt)
- Batch conversion with live progress bar and activity log
- Color-coded log: OK (green), skipped (amber), errors (red)
- Cancel mid-conversion
- Dark / Light theme toggle

## Security

- **Macros disabled** — Word and PowerPoint macros are force-disabled during conversion
- **Magic byte validation** — files are checked against their real format signature, not just extension
- **File size limit** — files over 100 MB are rejected
- **Path traversal protection** — symlink and `..` attacks are blocked
- **Read-only open** — source files are never modified

---

## Requirements

- Windows 10 / 11
- Microsoft Office installed (Word + PowerPoint)
- Python 3.10+

## Installation

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/file-changer.git
cd file-changer

# 2. Create virtual environment
python -m venv .venv
.venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
python main.py
```

Or just double-click **`launch.bat`** after setup.

## Build a standalone .exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "FileChanger" main.py
```

The `.exe` will be in the `dist/` folder. No Python required on the target machine.

---

## Project Structure

```
file-changer/
├── app.py            # UI — CustomTkinter window, layout, events
├── converter.py      # Core logic — Office COM automation + security checks
├── main.py           # Entry point
├── launch.bat        # One-click launcher (Windows)
└── requirements.txt  # Dependencies
```

## Dependencies

| Package | Purpose |
|---------|---------|
| `customtkinter` | Modern Tkinter UI framework |
| `comtypes` | Windows COM automation (controls Word/PowerPoint) |
| `pillow` | Image support for CustomTkinter |

---

## License

MIT
