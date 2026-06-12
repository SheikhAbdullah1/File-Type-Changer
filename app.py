import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from converter import (
    CONVERTIBLE_EXTENSIONS,
    OUTPUT_FORMATS,
    batch_pdf_converter,
    get_output_settings,
    preview_files,
)

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SUPPORTED_TEXT = ", ".join(sorted(CONVERTIBLE_EXTENSIONS))

# ── Design Tokens ──────────────────────────────────────────────────────────────
ACCENT       = "#3B82F6"          # vivid blue — primary actions
ACCENT_HOVER = "#2563EB"
SUCCESS      = "#10B981"          # emerald — already exists / done
WARNING      = "#F59E0B"          # amber — unsupported
DANGER       = "#EF4444"          # red — errors / cancel

BG_BASE      = "#0D1117"          # near-black canvas
BG_SURFACE   = "#161B22"          # cards / panels
BG_ELEVATED  = "#21262D"          # inner boxes / log
BORDER       = "#30363D"          # subtle dividers

TEXT_PRIMARY  = "#F0F6FC"
TEXT_MUTED    = "#8B949E"

# NOTE: CTkFont objects are created inside FileChangerApp.__init__()
# because Tk root window must exist first.

# Format icon mapping
FORMAT_ICONS = {
    "PDF (.pdf)":         "📄",
    "Word (.docx)":       "📝",
    "Word (.doc)":        "📝",
    "PowerPoint (.pptx)": "📊",
    "PowerPoint (.ppt)":  "📊",
}
class SectionLabel(ctk.CTkLabel):
    """Small uppercase eyebrow label for sections."""
    def __init__(self, parent, text, **kwargs):
        super().__init__(
            parent,
            text=text.upper(),
            font=("Segoe UI", 10, "bold"),
            text_color=TEXT_MUTED,
            **kwargs,
        )

class StatCard(ctk.CTkFrame):
    def __init__(self, parent, title, color, **kwargs):
        super().__init__(
            parent,
            corner_radius=12,
            fg_color=BG_ELEVATED,
            border_width=1,
            border_color=BORDER,
            **kwargs,
        )
        self.value_label = ctk.CTkLabel(
            self, text="0", font=("Segoe UI", 30, "bold"), text_color=color
        )
        self.value_label.pack(padx=20, pady=(18, 2))

        ctk.CTkLabel(
            self, text=title, font=("Segoe UI", 12), text_color=TEXT_MUTED
        ).pack(padx=20, pady=(0, 16))

    def set(self, value):
        self.value_label.configure(text=str(value))

class FileChangerApp(ctk.CTk):
    def __init__(self):
        super().__init__()  # Tk root created here — safe to make CTkFont now

        # Fonts must be created after super().__init__()
        global FONT_DISPLAY, FONT_HEADING, FONT_BODY, FONT_SMALL
        global FONT_MONO, FONT_STAT, FONT_BTN
        FONT_DISPLAY = ctk.CTkFont(family="Segoe UI", size=26, weight="bold")
        FONT_HEADING = ctk.CTkFont(family="Segoe UI", size=15, weight="bold")
        FONT_BODY    = ctk.CTkFont(family="Segoe UI", size=13)
        FONT_SMALL   = ctk.CTkFont(family="Segoe UI", size=12)
        FONT_MONO    = ctk.CTkFont(family="Consolas",  size=12)
        FONT_STAT    = ctk.CTkFont(family="Segoe UI",  size=30, weight="bold")
        FONT_BTN     = ctk.CTkFont(family="Segoe UI",  size=13, weight="bold")

        self.title("File Changer")
        self.geometry("960x720")
        self.minsize(860, 660)
        self.configure(fg_color=BG_BASE)

        self.status_text     = tk.StringVar(value="Select files to get started.")
        self.selected_files  = []
        self.output_format   = tk.StringVar(value="PDF (.pdf)")
        self.is_running      = False
        self.cancel_requested = False
        self.log_queue       = queue.Queue()

        self._build_layout()
        self.after(100, self._poll_log_queue)

    # ── Layout ─────────────────────────────────────────────────────────────────

    def _build_layout(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)  # log area expands

        self._build_header()
        self._build_work_area()
        self._build_log_panel()
        self._build_footer()

    # ── Header ─────────────────────────────────────────────────────────────────

    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=32, pady=(28, 0))
        header.grid_columnconfigure(0, weight=1)

        left = ctk.CTkFrame(header, fg_color="transparent")
        left.grid(row=0, column=0, sticky="w")

        # pill badge
        badge = ctk.CTkFrame(left, corner_radius=6, fg_color=ACCENT, width=6, height=28)
        badge.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(0, 12))
        badge.grid_propagate(False)

        ctk.CTkLabel(left, text="File Changer", font=FONT_DISPLAY,
                     text_color=TEXT_PRIMARY).grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(left, text="Batch convert Word & PowerPoint files",
                     font=FONT_BODY, text_color=TEXT_MUTED).grid(row=1, column=1, sticky="w", pady=(2, 0))

        # theme toggle
        theme_btn = ctk.CTkSegmentedButton(
            header,
            values=["Dark", "Light"],
            command=self._toggle_theme,
            width=130,
            height=32,
            font=FONT_SMALL,
        )
        theme_btn.set("Dark")
        theme_btn.grid(row=0, column=1, sticky="e")

        # divider
        divider = ctk.CTkFrame(self, height=1, fg_color=BORDER)
        divider.grid(row=0, column=0, sticky="sew", padx=32, pady=(72, 0))

    # ── Work Area (upload + format side by side) ───────────────────────────────

    def _build_work_area(self):
        work = ctk.CTkFrame(self, fg_color="transparent")
        work.grid(row=1, column=0, sticky="ew", padx=32, pady=(20, 0))
        work.grid_columnconfigure(0, weight=3)
        work.grid_columnconfigure(1, weight=2)

        self._build_upload_panel(work)
        self._build_format_panel(work)

    def _build_upload_panel(self, parent):
        panel = ctk.CTkFrame(parent, corner_radius=16,
                             fg_color=BG_SURFACE, border_width=1, border_color=BORDER)
        panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        panel.grid_columnconfigure(0, weight=1)

        SectionLabel(panel, "Input Files").grid(row=0, column=0, sticky="w", padx=20, pady=(18, 0))

        # drop zone
        self.drop_zone = ctk.CTkFrame(
            panel, corner_radius=12,
            fg_color=BG_ELEVATED,
            border_width=2, border_color=BORDER,
        )
        self.drop_zone.grid(row=1, column=0, sticky="ew", padx=16, pady=(10, 0))
        self.drop_zone.grid_columnconfigure(0, weight=1)
        for widget in [self.drop_zone]:
            widget.bind("<Button-1>", lambda _e: self.browse_files())
            widget.bind("<Enter>",    lambda _e: self.drop_zone.configure(border_color=ACCENT))
            widget.bind("<Leave>",    lambda _e: self.drop_zone.configure(border_color=BORDER))

        self.drop_icon = ctk.CTkLabel(
            self.drop_zone, text="⬆", font=ctk.CTkFont(size=32),
            text_color=TEXT_MUTED,
        )
        self.drop_icon.grid(row=0, column=0, pady=(20, 6))
        self.drop_icon.bind("<Button-1>", lambda _e: self.browse_files())

        self.drop_title = ctk.CTkLabel(
            self.drop_zone, text="Click to browse files",
            font=FONT_HEADING, text_color=TEXT_PRIMARY,
        )
        self.drop_title.grid(row=1, column=0, pady=(0, 4))
        self.drop_title.bind("<Button-1>", lambda _e: self.browse_files())

        self.drop_hint = ctk.CTkLabel(
            self.drop_zone,
            text=f"Supports: {SUPPORTED_TEXT}",
            font=FONT_SMALL, text_color=TEXT_MUTED,
        )
        self.drop_hint.grid(row=2, column=0, pady=(0, 20))
        self.drop_hint.bind("<Button-1>", lambda _e: self.browse_files())

        # file list label
        self.files_label = ctk.CTkLabel(
            panel, text="No files selected",
            font=FONT_SMALL, text_color=TEXT_MUTED,
            wraplength=360, justify="left", anchor="w",
        )
        self.files_label.grid(row=2, column=0, sticky="w", padx=20, pady=(12, 0))

        # stats row
        stats_row = ctk.CTkFrame(panel, fg_color="transparent")
        stats_row.grid(row=3, column=0, sticky="ew", padx=16, pady=(12, 0))
        for col in range(3):
            stats_row.grid_columnconfigure(col, weight=1)

        stat_defs = [
            ("convertible", "Ready",         ACCENT),
            ("skipped",     "Already Done",  SUCCESS),
            ("unsupported", "Unsupported",   WARNING),
        ]
        self.stat_cards = {}
        for i, (key, title, color) in enumerate(stat_defs):
            card = StatCard(stats_row, title, color)
            card.grid(row=0, column=i, sticky="nsew",
                      padx=(0 if i == 0 else 6, 6 if i < 2 else 0))
            self.stat_cards[key] = card

        # browse button
        self.browse_button = ctk.CTkButton(
            panel, text="Browse Files",
            height=40, corner_radius=10,
            font=FONT_BTN,
            fg_color=BG_ELEVATED, hover_color=BORDER,
            border_width=1, border_color=BORDER,
            command=self.browse_files,
        )
        self.browse_button.grid(row=4, column=0, sticky="ew", padx=16, pady=(14, 16))

    def _build_format_panel(self, parent):
        panel = ctk.CTkFrame(parent, corner_radius=16,
                             fg_color=BG_SURFACE, border_width=1, border_color=BORDER)
        panel.grid(row=0, column=1, sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)

        SectionLabel(panel, "Output Format").grid(row=0, column=0, sticky="w", padx=20, pady=(18, 0))

        ctk.CTkLabel(panel, text="Convert to", font=FONT_HEADING,
                     text_color=TEXT_PRIMARY).grid(row=1, column=0, sticky="w", padx=20, pady=(8, 0))

        self.format_dropdown = ctk.CTkComboBox(
            panel,
            values=list(OUTPUT_FORMATS.keys()),
            variable=self.output_format,
            height=44,
            corner_radius=10,
            border_color=BORDER,
            button_color=ACCENT,
            button_hover_color=ACCENT_HOVER,
            font=FONT_BODY,
            command=self.on_format_change,
        )
        self.format_dropdown.grid(row=2, column=0, sticky="ew", padx=20, pady=(10, 0))

        # format icon display
        self.format_icon_label = ctk.CTkLabel(
            panel, text="📄",
            font=ctk.CTkFont(size=48),
        )
        self.format_icon_label.grid(row=3, column=0, pady=(18, 0))

        self.format_ext_label = ctk.CTkLabel(
            panel, text=".pdf",
            font=ctk.CTkFont(family="Consolas", size=18, weight="bold"),
            text_color=ACCENT,
        )
        self.format_ext_label.grid(row=4, column=0, pady=(4, 0))

        # spacer
        ctk.CTkFrame(panel, fg_color="transparent", height=16).grid(row=5, column=0)

        self.clear_button = ctk.CTkButton(
            panel, text="Clear All",
            height=38, corner_radius=10,
            font=FONT_BTN,
            fg_color=BG_ELEVATED, hover_color=BORDER,
            border_width=1, border_color=BORDER,
            text_color=TEXT_MUTED,
            command=self.clear_files,
        )
        self.clear_button.grid(row=6, column=0, sticky="ew", padx=20, pady=(0, 8))

        self.convert_button = ctk.CTkButton(
            panel, text="Convert  →",
            height=48, corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self.start_conversion,
        )
        self.convert_button.grid(row=7, column=0, sticky="ew", padx=20, pady=(0, 20))

    # ── Log Panel ──────────────────────────────────────────────────────────────

    def _build_log_panel(self):
        log_card = ctk.CTkFrame(self, corner_radius=16,
                                fg_color=BG_SURFACE, border_width=1, border_color=BORDER)
        log_card.grid(row=2, column=0, sticky="nsew", padx=32, pady=(16, 0))
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 8))
        log_header.grid_columnconfigure(0, weight=1)

        SectionLabel(log_header, "Activity Log").grid(row=0, column=0, sticky="w")

        self.clear_log_btn = ctk.CTkButton(
            log_header, text="Clear Log",
            width=80, height=26, corner_radius=6,
            font=FONT_SMALL,
            fg_color=BG_ELEVATED, hover_color=BORDER,
            border_width=1, border_color=BORDER,
            text_color=TEXT_MUTED,
            command=self._clear_log,
        )
        self.clear_log_btn.grid(row=0, column=1, sticky="e")

        self.log_box = ctk.CTkTextbox(
            log_card, corner_radius=10,
            font=FONT_MONO,
            fg_color=BG_ELEVATED,
            border_width=1, border_color=BORDER,
            text_color=TEXT_PRIMARY,
            scrollbar_button_color=BORDER,
        )
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 16))
        self.log_box.configure(state="disabled")

    # ── Footer ─────────────────────────────────────────────────────────────────

    def _build_footer(self):
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=3, column=0, sticky="ew", padx=32, pady=(12, 24))
        footer.grid_columnconfigure(0, weight=1)

        self.progress = ctk.CTkProgressBar(
            footer, height=6, corner_radius=3,
            fg_color=BG_ELEVATED,
            progress_color=ACCENT,
        )
        self.progress.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        self.progress.set(0)

        ctk.CTkLabel(
            footer, textvariable=self.status_text,
            anchor="w", font=FONT_SMALL, text_color=TEXT_MUTED,
        ).grid(row=1, column=0, sticky="w")

        self.cancel_button = ctk.CTkButton(
            footer, text="Cancel",
            width=100, height=38, corner_radius=10,
            font=FONT_BTN,
            fg_color=DANGER, hover_color="#DC2626",
            command=self.request_cancel,
            state="disabled",
        )
        self.cancel_button.grid(row=1, column=2, sticky="e")

    # ── Actions ────────────────────────────────────────────────────────────────

    def _toggle_theme(self, value):
        ctk.set_appearance_mode("Light" if value == "Light" else "Dark")

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def browse_files(self):
        if self.is_running:
            return
        filetypes = [
            ("Supported files", "*.doc;*.docx;*.ppt;*.pptx;*.pps;*.ppsx"),
            ("All files", "*.*"),
        ]
        paths = filedialog.askopenfilenames(
            initialdir=SCRIPT_DIR,
            title="Select files to convert",
            filetypes=filetypes,
        )
        if paths:
            self.selected_files = list(paths)
            self.refresh_preview()

    def clear_files(self):
        if self.is_running:
            return
        self.selected_files = []
        self.refresh_preview()

    def on_format_change(self, choice):
        icon = FORMAT_ICONS.get(choice, "📄")
        ext  = get_output_settings(choice)["ext"]
        self.format_icon_label.configure(text=icon)
        self.format_ext_label.configure(text=ext)
        self.refresh_preview()

    def refresh_preview(self):
        if not self.selected_files:
            self.files_label.configure(text="No files selected")
            self.drop_title.configure(text="Click to browse files")
            self.drop_icon.configure(text="⬆", text_color=TEXT_MUTED)
            for card in self.stat_cards.values():
                card.set(0)
            self.status_text.set("Select files to get started.")
            return

        names = [os.path.basename(p) for p in self.selected_files]
        if len(names) == 1:
            files_text = f"✔  {names[0]}"
            drop_title = names[0]
        elif len(names) <= 3:
            files_text = "\n".join(f"✔  {n}" for n in names)
            drop_title = f"{len(names)} files selected"
        else:
            files_text = "\n".join(f"✔  {n}" for n in names[:3]) + f"\n   … and {len(names)-3} more"
            drop_title = f"{len(names)} files selected"

        self.files_label.configure(text=files_text)
        self.drop_title.configure(text=drop_title)
        self.drop_icon.configure(text="✔", text_color=SUCCESS)

        output_ext = get_output_settings(self.output_format.get())["ext"]
        summary    = preview_files(self.selected_files, output_ext)

        for key, card in self.stat_cards.items():
            card.set(summary[key])

        self.status_text.set(
            f"{summary['total']} file(s) selected  •  {summary['convertible']} ready to convert"
        )

    # ── Logging ────────────────────────────────────────────────────────────────

    def append_log(self, message):
        # color-code log lines
        self.log_box.configure(state="normal")
        tag = None
        if message.startswith("[OK]"):
            tag = "ok"
        elif message.startswith("[ERROR]"):
            tag = "error"
        elif message.startswith("[SKIP]") or message.startswith("[WARN]"):
            tag = "warn"

        self.log_box.insert("end", message + "\n", tag)

        # configure tags once
        self.log_box.tag_config("ok",    foreground=SUCCESS)
        self.log_box.tag_config("error", foreground=DANGER)
        self.log_box.tag_config("warn",  foreground=WARNING)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _poll_log_queue(self):
        while True:
            try:
                item = self.log_queue.get_nowait()
            except queue.Empty:
                break

            kind = item[0]
            if kind == "log":
                self.append_log(item[1])
            elif kind == "progress":
                current, total = item[1], item[2]
                self.progress.set(current / total if total else 0)
                self.status_text.set(f"Converting  {current} / {total} …")
            elif kind == "done":
                self._finish_conversion(item[1])

        self.after(100, self._poll_log_queue)

    # ── Conversion ─────────────────────────────────────────────────────────────

    def set_running_state(self, running):
        self.is_running = running
        state = "disabled" if running else "normal"
        self.convert_button.configure(state=state)
        self.browse_button.configure(state=state)
        self.clear_button.configure(state=state)
        self.format_dropdown.configure(state=state)
        self.cancel_button.configure(state="normal" if running else "disabled")

    def request_cancel(self):
        if self.is_running:
            self.cancel_requested = True
            self.status_text.set("Cancelling after current file …")

    def start_conversion(self):
        if not self.selected_files:
            messagebox.showinfo("No Files", "Please select at least one file first.")
            return

        output_ext = get_output_settings(self.output_format.get())["ext"]
        summary    = preview_files(self.selected_files, output_ext)

        if summary["convertible"] == 0:
            messagebox.showinfo(
                "Nothing to Convert",
                f"No files can be converted to {output_ext}.\n"
                "They may already exist or are unsupported for this format.",
            )
            return

        if not messagebox.askyesno(
            "Start Conversion",
            f"Convert {summary['convertible']} file(s) to {output_ext}?",
        ):
            return

        self._clear_log()
        self.progress.set(0)
        self.cancel_requested = False
        self.set_running_state(True)
        self.status_text.set("Starting conversion …")

        threading.Thread(target=self._run_conversion, daemon=True).start()

    def _run_conversion(self):
        stats = batch_pdf_converter(
            self.selected_files,
            output_format=self.output_format.get(),
            log    = lambda msg: self.log_queue.put(("log",      msg)),
            progress = lambda c, t: self.log_queue.put(("progress", c, t)),
            cancel_check = lambda: self.cancel_requested,
        )
        self.log_queue.put(("done", stats))

    def _finish_conversion(self, stats):
        self.set_running_state(False)
        self.progress.set(1)
        self.refresh_preview()
        output_ext = get_output_settings(self.output_format.get())["ext"]
        self.status_text.set(
            f"Done — {stats['converted']} converted, {stats['failed']} failed."
        )

        if stats["failed"]:
            messagebox.showwarning(
                "Conversion Complete",
                f"Converted: {stats['converted']}\nFailed: {stats['failed']}\n"
                "Check the activity log for details.",
            )
        else:
            messagebox.showinfo(
                "All Done! 🎉",
                f"Successfully converted {stats['converted']} file(s) to {output_ext}.",
            )


def main():
    app = FileChangerApp()
    app.mainloop()


if __name__ == "__main__":
    main()