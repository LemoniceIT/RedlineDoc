# RedlineDoc

**Local PDF, DOCX, and DOC text revision and diff analyzer.**

RedlineDoc extracts text from two versions of the same document, computes differences using `difflib.HtmlDiff`, and generates a navigable HTML report highlighting changes relative to the original document. The CustomTkinter interface adopts a dark Cyber/Matrix theme with an integrated log terminal.

---

## Table of Contents

- [Features](#features)
- [System Requirements](#system-requirements)
- [Installation](#installation)
- [Launch](#launch)
- [Usage](#usage)
- [Architecture](#architecture)
- [Diff Engine](#diff-engine)
- [Internationalization](#internationalization)
- [Privacy & Data Handling](#privacy--data-handling)
- [Project Structure](#project-structure)
- [Third-Party Libraries](#third-party-libraries)
- [License](#license)

---

## Features

- **Textual comparison**: side-by-side analysis of two documents with HTML report generated via `difflib.HtmlDiff` (context, additions, and deletions highlighted).
- **Multi-format extraction**:
  - `.docx` → `python-docx`
  - `.doc` → `pywin32` (Microsoft Word COM automation)
  - `.pdf` → `pdfplumber`
- **Extraction options**: ability to exclude headers and footers from comparison; formatting and images are always ignored (comparison is purely textual).
- **Input validation**: verifies both documents are loaded and extensions match before processing.
- **Non-blocking execution**: extraction and comparison run in a background thread; the UI remains responsive with progress visible in the log terminal.
- **Read-only log terminal**, cyber/terminal style, fed by a thread-safe queue.
- **Multilingual**: 5 languages (EN, ES, FR, IT, DE) with hot-swappable language switching, no restart required.
- **Dark Matrix/Cyberpunk theme** consistent throughout the application.

## System Requirements

| Requirement | Detail |
|---|---|
| Operating System | Windows 10/11 |
| Python | 3.11 or higher |
| Microsoft Word | Required **only** for `.doc` file extraction (COM automation) |
| Browser | Default browser for report viewing |

## Installation

From the project directory:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install customtkinter python-docx pdfplumber pywin32
```

> Dependencies `python-docx`, `pdfplumber`, and `pywin32` are imported lazily (on demand): the application starts even if one is missing, reporting an error only when extracting that specific format.

Runtime dependencies:

- `customtkinter` — graphical interface
- `python-docx` (+ `lxml`) — `.docx` text extraction
- `pdfplumber` — `.pdf` text extraction
- `pywin32` — `.doc` text extraction via COM
- `tkinter.filedialog`, `difflib`, `tempfile`, `webbrowser` — standard library

## Launch

```bash
python gui.py
```

Or via the included script:

```bat
run.bat
```

## Usage

1. Press **Original Document** (or **Documento Originale** if language is set to Italian) and select the reference document.
2. Press **Modified Document** (or **Documento Modificato**) and select the modified version (must have the same extension as the original).
3. (Optional) Press **Configuration** (**Configurazione**) to enable/disable header and footer exclusion.
4. Press **Compare** (**Confronta**): the HTML report is generated in a temporary folder and opened automatically in the default browser.

Extraction options:

| Option | State | Effect |
|---|---|---|
| Ignore formatting | Always on | Plain text extraction only |
| Ignore images | Always on | Non-textual content excluded |
| Ignore header | User-toggleable | Headers excluded from comparison |
| Ignore footer | User-toggleable | Footers excluded from comparison |

## Architecture

The code is organized into four modules with well-separated responsibilities:

| Module | Responsibility |
|---|---|
| `config.py` | Application metadata, paths, theme palette, format constraints, GUI parameters |
| `i18n.py` | Dependency-free translation system with tables for 5 languages and chained fallback (active language → English → key) |
| `document_processor.py` | Per-format text extraction logic and comparison report generation |
| `gui.py` | CustomTkinter interface, modal dialogs, worker thread, and log queue polling |

**Concurrency model.** The main thread handles UI exclusively. Extraction and comparison are delegated to a *daemon* thread; progress logs and errors are communicated via `queue.Queue`, consumed with `after()` every 100 ms. This guarantees no Tkinter widget is ever updated from a secondary thread.

**PDF extraction pipeline.** When exclusion options are active, `pdfplumber` operates on a cropped area (crop box) corresponding to the top and/or bottom 10% of the page, effectively discarding physical headers and footers.

**DOC extraction pipeline.** The document is opened read-only via `Word.Application` COM; headers/footers are read per section (primary, first page, even pages) with per-section exception handling.

## Diff Engine

`DocumentProcessor.compare_documents()` uses `difflib.HtmlDiff(wrapcolumn=88)` with `context=True`, producing a self-contained HTML report with color-coded difference legend. The report is written to a temporary file (`sd_diff_report_*.html`) and opened in the default browser via `webbrowser.open`.

## Internationalization

Available languages: **English** (default), **Español**, **Français**, **Italiano**, **Deutsch**. Language is selected from the dropdown in the header; all interface strings update immediately without restart. The system falls back to English for missing keys in the selected language.

Strings are centralized in `i18n.py`, serving as the single maintenance point for adding new languages.

## Privacy & Data Handling

The application is designed with a *privacy by design* approach:

- **No personal data is stored in source code**, except the author attribution string (`config.py`), required for the MIT license copyright notice.
- **No credentials, tokens, or user identifiers** are present in the code.
- **Document paths** selected are displayed only in the runtime log terminal; they are not written to disk or transmitted.
- **Document content** never leaves the machine: extraction and comparison are entirely local.
- Temporary HTML reports contain the processed document text and reside in the user's temp folder; deletion after analysis is recommended.

## Project Structure

```text
RedlineDoc/
├── gui.py                  # CustomTkinter interface (MainWindow, modal dialogs)
├── document_processor.py   # Text extraction and report generation
├── config.py               # Configuration, theme, paths
├── i18n.py                 # Translations (EN, ES, FR, IT, DE)
├── assets/
│   └── app_icon.ico        # Application icon
├── dist/                   # Precompiled standalone executable (.exe)
├── README.md               # Technical project documentation
└── run.bat                 # Quick launch script
```

## Third-Party Libraries

| Library | License | Usage |
|---|---|---|
| CustomTkinter | MIT | Dark graphical interface |
| python-docx | MIT | `.docx` text extraction |
| lxml | BSD | python-docx dependency |
| pdfplumber | MIT | `.pdf` text extraction |
| pywin32 | Microsoft Public License | `.doc` text extraction (COM) |
| difflib / tempfile / webbrowser / tkinter | PSF (Python stdlib) | Comparison, temp files, browser, dialog |

## License

Released under the **MIT** license. Author attribution is required; the software is provided "as is", without warranty of any kind.

```text
MIT License

Copyright (c) 2026 Stefano D'Angelo

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```