# RedlineDoc 1.3

RedlineDoc compares the text of two local documents in DOC, DOCX or PDF format. Either input can use any supported format, including DOCX versus PDF. The desktop interface offers a side-by-side review and exports the comparison to HTML, PDF or JSON.

## What's new in 1.3

- Compare any pair of supported formats with word-level highlights, a change summary and detection of unambiguous paragraph moves.
- Review changes in the desktop app with synchronized scrolling, search, filters and drag and drop.
- Export the full comparison to HTML, PDF or JSON. PDF source page references are included when available.

## Installation

Requirements are Windows 10 or 11 and Python 3.11 or later. Microsoft Word is required to read legacy DOC files.

From the project directory in PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe gui.py
```

## Usage

1. Pick the reference and revised documents using **Original Document** and **Modified Document**, or drop one file onto each corresponding path field.
2. Use **Configuration** to exclude headers or footers when needed.
3. Select **Compare** to open the review. Use the arrows to jump between differences, the dropdown to filter by change type, and **Find** to search across both visible panes. Vertical scrolling is synchronized.
4. Select **Export** to save the full comparison, including the summary and PDF page references, as HTML, PDF or JSON. JSON includes versioned structured rows and highlighted spans.

## Local data and privacy

The interface supports English, Español, Français, Italiano and Deutsch. The preferences file at `%LOCALAPPDATA%\RedlineDoc\preferences.json` stores the language and header/footer options. Document selection and comparison run locally. Exported reports contain text from both input documents and are saved to the location you choose.

## How comparison works

- DOCX extraction walks paragraphs and tables in document order. Legacy DOC uses a dedicated Word COM instance and opens the source read-only with macros disabled.
- PDF extraction reconstructs paragraphs using line spacing and detects tables with visible rules. The page number accompanies extracted blocks. Headers and footers are identified by repeated text within the outer 10% of pages. Unique or varying margin text remains in the comparison.
- The comparison normalizes whitespace, Unicode composition and common typographic ligatures. It aligns blocks, highlights word-level edits and identifies unambiguous moved blocks. Formatting and images are excluded.
- PDFs need a selectable text layer. Reading order for multi-column pages and extraction for borderless PDF tables depend on the source layout. The review panes show normalized text.

## Development

| Module | Responsibility |
|---|---|
| `document_processor.py` | DOC, DOCX and PDF extraction |
| `comparison.py` | Block alignment, word differences and change counts |
| `gui.py` and `review_window.py` | CustomTkinter interface, worker queue and review controls |
| `reports.py` | HTML, PDF and JSON exports |
| `preferences.py` | Local settings |
| `config.py`, `i18n.py` and `version.py` | Application settings, translations and version |

Run the automated tests from the project directory:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Runtime dependencies are listed in `requirements.txt`. RedlineDoc uses CustomTkinter and tkinterdnd2 for the desktop interface, python-docx, pdfplumber and pywin32 for extraction, and ReportLab for PDF export.

Set `$env:REDLINEDOC_TEST_WORD = '1'` before running the suite to include the optional Word COM integration test.

## License

RedlineDoc is released under the [MIT license](LICENSE). Copyright © 2026 Stefano D'Angelo.
