"""
Configuration module for RedlineDoc.

Centralizes all non-sensitive configuration settings, styling constants,
file-format constraints, and application directory paths.

Palette follows the Cyber/Matrix dark theme of the reference BotMyPC project.
"""

import os
import sys
from pathlib import Path

from i18n import t

# Application Metadata
APP_NAME: str = "RedlineDoc"
APP_DISPLAY_TITLE: str = "RedlineDoc"
APP_VERSION: str = "1.0.0"
APP_AUTHOR: str = "Stefano D'Angelo"

# License information
LICENSE_TYPE: str = "MIT"

# Application tagline/claim
APP_CLAIM: str = "Local PDF, DOCX, and DOC text revision and diff analyzer."

# Application Directories and Paths
if getattr(sys, "frozen", False):
    BASE_DIR: Path = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    EXE_DIR: Path = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
    EXE_DIR = BASE_DIR

ASSETS_DIR: Path = BASE_DIR / "assets"
ICON_PATH: Path = ASSETS_DIR / "app_icon.ico"

# Local user data directory: %LOCALAPPDATA%\RedlineDoc
LOCAL_DATA_DIR: Path = Path(os.environ.get("LOCALAPPDATA", Path.home())) / APP_NAME
LOCAL_DATA_DIR.mkdir(parents=True, exist_ok=True)

LOG_FILE_PATH: Path = LOCAL_DATA_DIR / "redline_doc.log"

# GUI Dimensions & Theme Colors (Matrix/Cyberpunk Palette)
GUI_WIDTH: int = 1024
GUI_HEIGHT: int = 768
GUI_BG_COLOR: str = "#0A0A0A"
GUI_PANEL_BG: str = "#121212"
GUI_TERMINAL_BG: str = "#000000"
GUI_ACCENT_COLOR: str = "#00FF66"
GUI_ACCENT_HOVER: str = "#00CC52"
GUI_TEXT_COLOR: str = "#E0E0E0"
GUI_MUTED_TEXT: str = "#7A7A7A"
GUI_DANGER_COLOR: str = "#FF3366"
GUI_DANGER_HOVER: str = "#CC2952"
GUI_WARNING_COLOR: str = "#FFB300"
GUI_BORDER_COLOR: str = "#1F3D2B"

# Terminal bar color (matches reference: #0F1A13)
GUI_TERMINAL_BAR: str = "#0F1A13"

# Secondary button palette (matches reference: #1F2A22 / #2A3D2F)
GUI_SECONDARY_BTN: str = "#1F2A22"
GUI_SECONDARY_BTN_HOVER: str = "#2A3D2F"
GUI_SECONDARY_BTN_BORDER: str = "#2A2A2A"
GUI_SECONDARY_BTN_BORDER_HOVER: str = "#3A3A3A"

# Polling Interval (ms) for GUI Queue Consumer
GUI_QUEUE_POLL_INTERVAL_MS: int = 100

# File Format Constraints
SUPPORTED_EXTENSIONS: tuple[str, ...] = (".doc", ".docx", ".pdf")


def get_supported_filetypes() -> list[tuple[str, str]]:
    """Return translated file-type filter list for file dialogs."""
    return [
        (t("filetype.doc"), "*.doc"),
        (t("filetype.docx"), "*.docx"),
        (t("filetype.pdf"), "*.pdf"),
    ]


# Extraction library mapping (for attribution & credits display)
def get_extraction_libraries() -> dict[str, str]:
    """Return translated extraction-library descriptions."""
    return {
        "pywin32": t("lib.pywin32"),
        "python-docx": t("lib.python_docx"),
        "pdfplumber": t("lib.pdfplumber"),
        "difflib": t("lib.difflib"),
    }
