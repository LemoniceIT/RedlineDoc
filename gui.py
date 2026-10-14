"""
CustomTkinter Graphical User Interface for RedlineDoc.

Features:
- 1024x768 window, maximizable/fullscreen-capable, Cyberpunk/Matrix dark theme.
- Two file-load buttons (Original Document / Modified Document) for
  .doc, .docx, .pdf formats with extension-match validation.
- Configuration panel modal with four checkboxes:
    * Ignore formatting  (forced-on, disabled)
    * Ignore images       (forced-on, disabled)
    * Ignore header       (user-toggleable)
    * Ignore footer       (user-toggleable)
- Compare button: extracts text via pywin32 / python-docx / pdfplumber,
  runs difflib.HtmlDiff, writes a temporary HTML report and opens it
  in the default browser.
- Read-only Log Terminal panel (replicated from reference BotMyPC style)
  polling messages from a thread-safe queue.Queue.
- Credits modal with developer attribution, library licenses, and the
- Credits modal with developer attribution, library licenses, and MIT license
  notice (attribution required, no warranty).
- Multi-language support (EN, ES, FR, IT, DE) with live switching via header
  dropdown – no restart required.
"""

from __future__ import annotations

import logging
import queue
import threading
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

from config import (
    APP_AUTHOR,
    APP_DISPLAY_TITLE,
    APP_NAME,
    APP_VERSION,
    GUI_ACCENT_COLOR,
    GUI_ACCENT_HOVER,
    GUI_BG_COLOR,
    GUI_BORDER_COLOR,
    GUI_DANGER_COLOR,
    GUI_DANGER_HOVER,
    GUI_HEIGHT,
    GUI_MUTED_TEXT,
    GUI_PANEL_BG,
    GUI_QUEUE_POLL_INTERVAL_MS,
    GUI_SECONDARY_BTN,
    GUI_SECONDARY_BTN_BORDER,
    GUI_SECONDARY_BTN_BORDER_HOVER,
    GUI_SECONDARY_BTN_HOVER,
    GUI_TERMINAL_BAR,
    GUI_TERMINAL_BG,
    GUI_TEXT_COLOR,
    GUI_WARNING_COLOR,
    GUI_WIDTH,
    LICENSE_TYPE,
    get_supported_filetypes,
)
from document_processor import DocumentProcessor, ExtractOptions
from i18n import LANGUAGES, get_language, set_language, t

logger = logging.getLogger(APP_NAME)

# ---------------------------------------------------------------------------
# Global appearance & theme (matches reference)
# ---------------------------------------------------------------------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")


class ConfigDialog(ctk.CTkToplevel):
    """Modal dialog with four comparison checkboxes."""

    def __init__(
        self,
        parent: ctk.CTk,
        var_ignore_header: ctk.BooleanVar,
        var_ignore_footer: ctk.BooleanVar,
    ) -> None:
        super().__init__(parent)
        self.title(t("config.title"))
        self.geometry("480x320")
        self.resizable(False, False)
        self.configure(fg_color=GUI_BG_COLOR)

        self.transient(parent)
        self.grab_set()

        # Header
        title_lbl = ctk.CTkLabel(
            self,
            text=t("config.header"),
            font=ctk.CTkFont(family="Consolas", size=16, weight="bold"),
            text_color=GUI_ACCENT_COLOR,
        )
        title_lbl.pack(pady=(24, 6))

        info_lbl = ctk.CTkLabel(
            self,
            text=t("config.description"),
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=GUI_TEXT_COLOR,
            justify="center",
        )
        info_lbl.pack(pady=(0, 18))

        # Checkbox container
        chk_frame = ctk.CTkFrame(
            self,
            fg_color=GUI_PANEL_BG,
            corner_radius=8,
            border_width=1,
            border_color=GUI_BORDER_COLOR,
        )
        chk_frame.pack(fill="x", padx=24, pady=(0, 18))

        # Ignore formatting — forced-on, disabled (non-deselectable)
        self._var_format = ctk.BooleanVar(value=True)
        self.chk_format = ctk.CTkCheckBox(
            chk_frame,
            text=t("config.ignore_formatting"),
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=GUI_TEXT_COLOR,
            fg_color=GUI_ACCENT_COLOR,
            hover_color=GUI_ACCENT_HOVER,
            border_color=GUI_BORDER_COLOR,
            variable=self._var_format,
            state="disabled",
        )
        self.chk_format.pack(anchor="w", padx=16, pady=(12, 4))

        # Ignore images — forced-on, disabled (non-deselectable)
        self._var_images = ctk.BooleanVar(value=True)
        self.chk_images = ctk.CTkCheckBox(
            chk_frame,
            text=t("config.ignore_images"),
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=GUI_TEXT_COLOR,
            fg_color=GUI_ACCENT_COLOR,
            hover_color=GUI_ACCENT_HOVER,
            border_color=GUI_BORDER_COLOR,
            variable=self._var_images,
            state="disabled",
        )
        self.chk_images.pack(anchor="w", padx=16, pady=(4, 4))

        # Ignore header — user-toggleable
        self.chk_header = ctk.CTkCheckBox(
            chk_frame,
            text=t("config.ignore_header"),
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=GUI_TEXT_COLOR,
            fg_color=GUI_ACCENT_COLOR,
            hover_color=GUI_ACCENT_HOVER,
            border_color=GUI_BORDER_COLOR,
            variable=var_ignore_header,
        )
        self.chk_header.pack(anchor="w", padx=16, pady=(4, 4))

        # Ignore footer — user-toggleable
        self.chk_footer = ctk.CTkCheckBox(
            chk_frame,
            text=t("config.ignore_footer"),
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=GUI_TEXT_COLOR,
            fg_color=GUI_ACCENT_COLOR,
            hover_color=GUI_ACCENT_HOVER,
            border_color=GUI_BORDER_COLOR,
            variable=var_ignore_footer,
        )
        self.chk_footer.pack(anchor="w", padx=16, pady=(4, 12))

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=(0, 12))

        cancel_btn = ctk.CTkButton(
            btn_frame,
            text=t("btn.cancel"),
            width=120,
            height=34,
            fg_color=GUI_SECONDARY_BTN_BORDER,
            hover_color=GUI_SECONDARY_BTN_BORDER_HOVER,
            text_color=GUI_TEXT_COLOR,
            font=ctk.CTkFont(weight="bold"),
            command=self.destroy,
        )
        cancel_btn.pack(side="left", padx=10)

        save_btn = ctk.CTkButton(
            btn_frame,
            text=t("btn.apply"),
            width=140,
            height=34,
            fg_color=GUI_ACCENT_COLOR,
            hover_color=GUI_ACCENT_HOVER,
            text_color="#000000",
            font=ctk.CTkFont(weight="bold"),
            command=self.destroy,
        )
        save_btn.pack(side="left", padx=10)


class CreditsDialog(ctk.CTkToplevel):
    """Modal dialog displaying developer info, library licenses, and GPLv3 grant."""

    def __init__(self, parent: ctk.CTk) -> None:
        super().__init__(parent)
        self.title(t("credits.title"))
        self.geometry("560x520")
        self.resizable(False, False)
        self.configure(fg_color=GUI_BG_COLOR)

        self.transient(parent)
        self.grab_set()

        title_lbl = ctk.CTkLabel(
            self,
            text=f"\u26a1 {APP_NAME} v{APP_VERSION}",
            font=ctk.CTkFont(family="Consolas", size=18, weight="bold"),
            text_color=GUI_ACCENT_COLOR,
        )
        title_lbl.pack(pady=(24, 6))

        author_lbl = ctk.CTkLabel(
            self,
            text=t("credits.developed_by").format(author=APP_AUTHOR),
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=GUI_ACCENT_COLOR,
        )
        author_lbl.pack(pady=(0, 4))

        sub_lbl = ctk.CTkLabel(
            self,
            text=t("credits.description"),
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=GUI_TEXT_COLOR,
        )
        sub_lbl.pack(pady=(0, 14))

        # License block
        license_frame = ctk.CTkFrame(
            self,
            fg_color=GUI_PANEL_BG,
            corner_radius=8,
            border_width=1,
            border_color=GUI_BORDER_COLOR,
        )
        license_frame.pack(fill="x", padx=24, pady=(0, 14))

        license_lbl = ctk.CTkLabel(
            license_frame,
            text=t("credits.license_info").format(license=LICENSE_TYPE),
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=GUI_WARNING_COLOR,
            justify="center",
        )
        license_lbl.pack(pady=14)

        # Libraries requiring license attribution
        tech_text = (
            t("credits.tech_header") + "\n"
            "\u2022 Python 3.11+ (Python Software Foundation License)\n"
            f"\u2022 {t('credits.tech_customtkinter')}\n"
            f"\u2022 {t('credits.tech_python_docx')}\n"
            f"\u2022 {t('credits.tech_pywin32')}\n"
            f"\u2022 {t('credits.tech_pdfplumber')}\n"
            f"\u2022 {t('credits.tech_lxml')}\n"
            f"\u2022 {t('credits.tech_difflib')}\n"
            f"\u2022 {t('credits.tech_tkinter')}"
        )

        tech_lbl = ctk.CTkLabel(
            self,
            text=tech_text,
            font=ctk.CTkFont(family="Consolas", size=11),
            text_color="#B0B0B0",
            justify="left",
        )
        tech_lbl.pack(pady=(0, 16), padx=34)

        close_btn = ctk.CTkButton(
            self,
            text=t("btn.close"),
            width=120,
            height=34,
            fg_color=GUI_ACCENT_COLOR,
            hover_color=GUI_ACCENT_HOVER,
            text_color="#000000",
            font=ctk.CTkFont(weight="bold"),
            command=self.destroy,
        )
        close_btn.pack(pady=(0, 18))


class MainWindow(ctk.CTk):
    """
    Main CustomTkinter application window.

    Owns all UI widgets and runs exclusively on the main thread.
    Long-running work (extraction + comparison) is delegated to a daemon
    thread; progress and errors are communicated back via a thread-safe
    queue.Queue that the main loop polls.
    """

    def __init__(self) -> None:
        super().__init__()

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.processor = DocumentProcessor()

        # File paths
        self.original_path: Path | None = None
        self.modified_path: Path | None = None

        # Extraction / comparison options
        self.var_ignore_header = ctk.BooleanVar(value=False)
        self.var_ignore_footer = ctk.BooleanVar(value=False)

        # Window settings — 1024x768, maximizable to fullscreen
        self.title(APP_DISPLAY_TITLE)
        self.geometry(f"{GUI_WIDTH}x{GUI_HEIGHT}")
        self.resizable(True, True)
        self.configure(fg_color=GUI_BG_COLOR)

        # Build UI layout
        self._build_header()
        self._build_controls()
        self._build_terminal()
        self._build_footer()

        # Connect window close protocol
        self.protocol("WM_DELETE_WINDOW", self._on_window_close)

        # Start log queue polling on main thread
        self.after(GUI_QUEUE_POLL_INTERVAL_MS, self._poll_queue)

        self._log(
            t("app.startup").format(name=APP_NAME, version=APP_VERSION)
        )
        self._log(t("app.ready"))

    # ===================================================================
    # UI Construction
    # ===================================================================

    def _build_header(self) -> None:
        header_frame = ctk.CTkFrame(
            self,
            fg_color=GUI_PANEL_BG,
            corner_radius=8,
            border_width=1,
            border_color=GUI_BORDER_COLOR,
        )
        header_frame.pack(fill="x", padx=16, pady=(16, 8))

        lbl_lightning = ctk.CTkLabel(
            header_frame,
            text="\u26a1",
            font=ctk.CTkFont(family="Segoe UI Emoji", size=22),
            text_color=GUI_ACCENT_COLOR,
        )
        lbl_lightning.pack(side="left", padx=(16, 6), pady=12)

        app_title = ctk.CTkLabel(
            header_frame,
            text=APP_NAME,
            font=ctk.CTkFont(family="Consolas", size=20, weight="bold"),
            text_color=GUI_ACCENT_COLOR,
        )
        app_title.pack(side="left", padx=(0, 16), pady=12)

        # Language selector (right side of header)
        lang_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        lang_frame.pack(side="right", padx=(0, 12), pady=12)

        self.lbl_lang = ctk.CTkLabel(
            lang_frame,
            text=t("lang.label"),
            font=ctk.CTkFont(family="Consolas", size=11),
            text_color=GUI_MUTED_TEXT,
        )
        self.lbl_lang.pack(side="left", padx=(0, 4))

        # Build display name → code lookup
        self._lang_display_to_code = {v: k for k, v in LANGUAGES.items()}
        current_display = LANGUAGES[get_language()]
        display_names = list(LANGUAGES.values())

        self._lang_var = ctk.StringVar(value=current_display)
        self.lang_menu = ctk.CTkOptionMenu(
            lang_frame,
            variable=self._lang_var,
            values=display_names,
            width=120,
            height=26,
            fg_color=GUI_SECONDARY_BTN,
            button_color=GUI_SECONDARY_BTN_HOVER,
            button_hover_color=GUI_ACCENT_HOVER,
            text_color=GUI_TEXT_COLOR,
            dropdown_fg_color=GUI_PANEL_BG,
            dropdown_text_color=GUI_TEXT_COLOR,
            dropdown_hover_color=GUI_ACCENT_COLOR,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            command=self._on_language_changed,
        )
        self.lang_menu.pack(side="left")

        version_lbl = ctk.CTkLabel(
            header_frame,
            text=f"v{APP_VERSION}",
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color=GUI_MUTED_TEXT,
        )
        version_lbl.pack(side="right", padx=(0, 12), pady=12)

    def _build_controls(self) -> None:
        ctrl_frame = ctk.CTkFrame(
            self,
            fg_color=GUI_PANEL_BG,
            corner_radius=8,
            border_width=1,
            border_color=GUI_BORDER_COLOR,
        )
        ctrl_frame.pack(fill="x", padx=16, pady=4)

        # --- File selection row ---
        file_row = ctk.CTkFrame(ctrl_frame, fg_color="transparent")
        file_row.pack(fill="x", padx=12, pady=12)

        # Original Document
        self.btn_original = ctk.CTkButton(
            file_row,
            text=t("btn.original_doc"),
            width=180,
            height=34,
            fg_color=GUI_ACCENT_COLOR,
            hover_color=GUI_ACCENT_HOVER,
            text_color="#000000",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=self._load_original,
        )
        self.btn_original.pack(side="left", padx=(4, 4))

        self.lbl_original_path = ctk.CTkLabel(
            file_row,
            text=t("file.no_selection"),
            font=ctk.CTkFont(family="Consolas", size=10),
            text_color=GUI_MUTED_TEXT,
            width=300,
            justify="left",
        )
        self.lbl_original_path.pack(side="left", padx=(0, 12))

        # Modified Document
        self.btn_modified = ctk.CTkButton(
            file_row,
            text=t("btn.modified_doc"),
            width=180,
            height=34,
            fg_color=GUI_ACCENT_COLOR,
            hover_color=GUI_ACCENT_HOVER,
            text_color="#000000",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=self._load_modified,
        )
        self.btn_modified.pack(side="left", padx=(4, 4))

        self.lbl_modified_path = ctk.CTkLabel(
            file_row,
            text=t("file.no_selection"),
            font=ctk.CTkFont(family="Consolas", size=10),
            text_color=GUI_MUTED_TEXT,
            width=300,
            justify="left",
        )
        self.lbl_modified_path.pack(side="left", padx=(0, 4))

        # --- Action row ---
        action_row = ctk.CTkFrame(ctrl_frame, fg_color="transparent")
        action_row.pack(fill="x", padx=12, pady=(0, 12))

        self.btn_config = ctk.CTkButton(
            action_row,
            text=t("btn.config"),
            width=150,
            height=34,
            fg_color=GUI_SECONDARY_BTN,
            hover_color=GUI_SECONDARY_BTN_HOVER,
            border_width=1,
            border_color=GUI_BORDER_COLOR,
            text_color=GUI_TEXT_COLOR,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=self._open_config_dialog,
        )
        self.btn_config.pack(side="left", padx=(4, 6))

        self.btn_compare = ctk.CTkButton(
            action_row,
            text=t("btn.compare"),
            width=140,
            height=34,
            fg_color=GUI_DANGER_COLOR,
            hover_color=GUI_DANGER_HOVER,
            text_color="#000000",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=self._start_comparison,
        )
        self.btn_compare.pack(side="left", padx=(0, 6))

        self.btn_credits = ctk.CTkButton(
            action_row,
            text=t("btn.credits"),
            width=100,
            height=34,
            fg_color=GUI_SECONDARY_BTN,
            hover_color=GUI_SECONDARY_BTN_HOVER,
            border_width=1,
            border_color=GUI_BORDER_COLOR,
            text_color=GUI_TEXT_COLOR,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=self._open_credits_dialog,
        )
        self.btn_credits.pack(side="right", padx=(6, 4))

    def _build_terminal(self) -> None:
        """Replicates the read-only Log Terminal panel from the reference GUI."""
        term_container = ctk.CTkFrame(
            self,
            fg_color=GUI_TERMINAL_BG,
            corner_radius=8,
            border_width=1,
            border_color=GUI_BORDER_COLOR,
        )
        term_container.pack(fill="both", expand=True, padx=16, pady=8)

        # Terminal header bar
        bar = ctk.CTkFrame(
            term_container, fg_color=GUI_TERMINAL_BAR, height=28, corner_radius=6
        )
        bar.pack(fill="x", padx=2, pady=2)

        self.lbl_term = ctk.CTkLabel(
            bar,
            text=t("terminal.header"),
            font=ctk.CTkFont(family="Consolas", size=11, weight="bold"),
            text_color=GUI_ACCENT_COLOR,
        )
        self.lbl_term.pack(side="left", padx=6, pady=4)

        self.btn_clear = ctk.CTkButton(
            bar,
            text=t("btn.clear"),
            width=60,
            height=20,
            fg_color="transparent",
            hover_color=GUI_SECONDARY_BTN_HOVER,
            text_color=GUI_MUTED_TEXT,
            font=ctk.CTkFont(family="Consolas", size=10),
            command=self._clear_terminal,
        )
        self.btn_clear.pack(side="right", padx=6, pady=4)

        # Scrollable read-only Text widget
        self.term_textbox = ctk.CTkTextbox(
            term_container,
            fg_color=GUI_TERMINAL_BG,
            text_color=GUI_ACCENT_COLOR,
            font=ctk.CTkFont(family="Consolas", size=11),
            corner_radius=4,
            wrap="char",
            activate_scrollbars=True,
        )
        self.term_textbox.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.term_textbox.configure(state="disabled")

    def _build_footer(self) -> None:
        footer = ctk.CTkFrame(self, fg_color="transparent", height=26)
        footer.pack(fill="x", padx=16, pady=(0, 8))

        self.lbl_status = ctk.CTkLabel(
            footer,
            text=t("status.ready"),
            font=ctk.CTkFont(family="Consolas", size=10),
            text_color=GUI_MUTED_TEXT,
        )
        self.lbl_status.pack(side="left")

        lbl_ver = ctk.CTkLabel(
            footer,
            text=f"v{APP_VERSION} | {LICENSE_TYPE} \u00a9 {APP_AUTHOR}",
            font=ctk.CTkFont(family="Consolas", size=10),
            text_color=GUI_MUTED_TEXT,
        )
        lbl_ver.pack(side="right")

    # ===================================================================
    # Language switching
    # ===================================================================

    def _on_language_changed(self, display_name: str) -> None:
        """Handle the language dropdown selection."""
        code = self._lang_display_to_code.get(display_name)
        if code and code != get_language():
            set_language(code)
            self._refresh_ui_texts()

    def _refresh_ui_texts(self) -> None:
        """Update every visible widget text to the current language."""
        # Buttons
        self.btn_original.configure(text=t("btn.original_doc"))
        self.btn_modified.configure(text=t("btn.modified_doc"))
        self.btn_config.configure(text=t("btn.config"))
        self.btn_compare.configure(text=t("btn.compare"))
        self.btn_credits.configure(text=t("btn.credits"))
        self.btn_clear.configure(text=t("btn.clear"))

        # File path labels (reset placeholder if no file loaded)
        if self.original_path is None:
            self.lbl_original_path.configure(text=t("file.no_selection"))
        if self.modified_path is None:
            self.lbl_modified_path.configure(text=t("file.no_selection"))

        # Terminal header
        self.lbl_term.configure(text=t("terminal.header"))

        # Status bar (only update if showing READY/PROCESSING, not mid-operation)
        current_status = self.lbl_status.cget("text")
        # Check against all known "ready" translations
        ready_texts = {
            "PRONTO", "READY", "LISTO", "PRÊT", "BEREIT",
        }
        if current_status in ready_texts:
            self.lbl_status.configure(text=t("status.ready"))

        # Language label
        self.lbl_lang.configure(text=t("lang.label"))

    # ===================================================================
    # Logging (thread-safe via queue)
    # ===================================================================

    def _log(self, message: str) -> None:
        """Enqueue a log message (safe to call from any thread)."""
        self.log_queue.put(message)

    def _poll_queue(self) -> None:
        """Pull log messages from the queue and insert into the terminal."""
        has_new = False
        self.term_textbox.configure(state="normal")
        while True:
            try:
                msg = self.log_queue.get_nowait()
                self.term_textbox.insert("end", msg + "\n")
                has_new = True
            except queue.Empty:
                break
        if has_new:
            self.term_textbox.see("end")
        self.term_textbox.configure(state="disabled")
        self.after(GUI_QUEUE_POLL_INTERVAL_MS, self._poll_queue)

    def _clear_terminal(self) -> None:
        self.term_textbox.configure(state="normal")
        self.term_textbox.delete("1.0", "end")
        self.term_textbox.configure(state="disabled")

    # ===================================================================
    # Event Handlers & Actions
    # ===================================================================

    def _update_file_label(
        self, label: ctk.CTkLabel, path: Path | None, placeholder: str
    ) -> None:
        if path:
            display = str(path)
            if len(display) > 52:
                display = str(path.parent / ("\u2026" + path.name))
                if len(display) > 52:
                    display = display[:49] + "\u2026"
            label.configure(text=display, text_color=GUI_TEXT_COLOR)
        else:
            label.configure(text=placeholder, text_color=GUI_MUTED_TEXT)

    def _load_original(self) -> None:
        file_path = filedialog.askopenfilename(
            parent=self,
            title=t("file.dialog_original"),
            filetypes=get_supported_filetypes(),
        )
        if file_path:
            self.original_path = Path(file_path)
            self._update_file_label(
                self.lbl_original_path,
                self.original_path,
                t("file.no_selection"),
            )
            self._log(
                t("file.loaded_original").format(path=self.original_path)
            )

    def _load_modified(self) -> None:
        file_path = filedialog.askopenfilename(
            parent=self,
            title=t("file.dialog_modified"),
            filetypes=get_supported_filetypes(),
        )
        if file_path:
            self.modified_path = Path(file_path)
            self._update_file_label(
                self.lbl_modified_path,
                self.modified_path,
                t("file.no_selection"),
            )
            self._log(
                t("file.loaded_modified").format(path=self.modified_path)
            )

    def _open_config_dialog(self) -> None:
        ConfigDialog(
            self,
            var_ignore_header=self.var_ignore_header,
            var_ignore_footer=self.var_ignore_footer,
        )

    def _open_credits_dialog(self) -> None:
        CreditsDialog(self)

    def _get_extract_options(self) -> ExtractOptions:
        return ExtractOptions(
            ignore_formatting=True,
            ignore_images=True,
            ignore_header=self.var_ignore_header.get(),
            ignore_footer=self.var_ignore_footer.get(),
        )

    def _start_comparison(self) -> None:
        """Validate inputs and launch the comparison in a background thread."""
        if not self.original_path or not self.modified_path:
            self._log(t("compare.both_required"))
            return

        ext_orig = self.original_path.suffix.lower()
        ext_mod = self.modified_path.suffix.lower()

        if ext_orig != ext_mod:
            self._log(
                t("compare.ext_mismatch").format(
                    ext1=ext_orig, ext2=ext_mod
                )
            )
            return

        options = self._get_extract_options()
        self._log(
            t("compare.options").format(
                header=t("compare.yes") if options.ignore_header else t("compare.no"),
                footer=t("compare.yes") if options.ignore_footer else t("compare.no"),
            )
        )

        # Disable compare button during processing
        self.btn_compare.configure(
            state="disabled", text=t("compare.in_progress")
        )
        self.lbl_status.configure(
            text=t("status.processing"), text_color=GUI_WARNING_COLOR
        )

        thread = threading.Thread(
            target=self._run_comparison, args=(options,), daemon=True
        )
        thread.start()

    def _run_comparison(self, options: ExtractOptions) -> None:
        """Background worker: extract text, compare, generate report, open browser."""
        try:
            self._log(t("compare.start_original"))
            orig_text = DocumentProcessor.extract_text(self.original_path, options)
            self._log(
                t("compare.done_chars").format(count=len(orig_text))
            )

            self._log(t("compare.start_modified"))
            mod_text = DocumentProcessor.extract_text(self.modified_path, options)
            self._log(
                t("compare.done_chars").format(count=len(mod_text))
            )

            self._log(t("compare.generating"))
            report_path = DocumentProcessor.compare_documents(orig_text, mod_text)
            self._log(
                t("compare.report_generated").format(path=report_path)
            )

            self._log(t("compare.opening_browser"))
            DocumentProcessor.open_in_browser(report_path)

            self._log(t("compare.completed"))
        except Exception as exc:
            self._log(t("compare.error").format(error=exc))
            logger.exception("Comparison failed")
        finally:
            self._on_comparison_done()

    def _on_comparison_done(self) -> None:
        """Re-enable UI controls after the background thread finishes."""
        self.btn_compare.configure(state="normal", text=t("btn.compare"))
        self.lbl_status.configure(text=t("status.ready"), text_color=GUI_MUTED_TEXT)

    def _on_window_close(self) -> None:
        """Intercepts the window close button for orderly shutdown."""
        self._log(t("app.closing"))
        self.destroy()


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
