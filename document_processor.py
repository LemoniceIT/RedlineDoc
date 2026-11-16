"""Extract comparable text blocks from Word and PDF documents."""

from __future__ import annotations

import difflib
import os
import tempfile
import webbrowser
from collections import Counter
from dataclasses import dataclass
from html import escape
from pathlib import Path

from comparison import Block, Comparison, compare
from i18n import t


@dataclass
class ExtractOptions:
    ignore_formatting: bool = True
    ignore_images: bool = True
    ignore_header: bool = False
    ignore_footer: bool = False


class DocumentProcessor:
    def __init__(self, log_callback=None) -> None:
        self._log = log_callback or (lambda _msg: None)

    @staticmethod
    def extract_blocks(file_path: str | Path, options: ExtractOptions) -> list[Block]:
        path = Path(file_path)
        extension = path.suffix.lower()
        extractors = {
            ".docx": DocumentProcessor._extract_docx,
            ".doc": DocumentProcessor._extract_doc,
            ".pdf": DocumentProcessor._extract_pdf,
        }
        if extension not in extractors:
            raise ValueError(t("doc.unsupported_format").format(ext=extension))
        return extractors[extension](path, options)

    @staticmethod
    def extract_text(file_path: str | Path, options: ExtractOptions) -> str:
        return "\n".join(
            block.text for block in DocumentProcessor.extract_blocks(file_path, options)
        )

    @staticmethod
    def review(original: Path, modified: Path, options: ExtractOptions) -> Comparison:
        left = DocumentProcessor.extract_blocks(original, options)
        right = DocumentProcessor.extract_blocks(modified, options)
        return compare(left, right, original.name, modified.name)

    @staticmethod
    def compare_documents(
        original_text: str,
        modified_text: str,
        from_desc: str | None = None,
        to_desc: str | None = None,
        output_path: str | None = None,
    ) -> str:
        """Generate the legacy HTML line diff for callers that supply plain text."""
        content = difflib.HtmlDiff(wrapcolumn=88).make_file(
            original_text.splitlines(),
            modified_text.splitlines(),
            fromdesc=escape(t("doc.original") if from_desc is None else from_desc),
            todesc=escape(t("doc.modified") if to_desc is None else to_desc),
            context=True,
        )
        if output_path is None:
            fd, output_path = tempfile.mkstemp(suffix=".html", prefix="sd_diff_report_")
            os.close(fd)
        Path(output_path).write_text(content, encoding="utf-8")
        return output_path

    @staticmethod
    def open_in_browser(file_path: str | Path) -> None:
        webbrowser.open(Path(file_path).resolve().as_uri())

    @staticmethod
    def _extract_docx(path: Path, options: ExtractOptions) -> list[Block]:
        from docx import Document
        from docx.table import Table

        doc = Document(str(path))

        def blocks(container, kind="paragraph"):
            for item in container.iter_inner_content():
                if isinstance(item, Table):
                    for row in item.rows:
                        cells = {cell._tc: cell for cell in row.cells}
                        text = " | ".join(
                            " ".join(block.text for block in blocks(cell))
                            for cell in cells.values()
                        )
                        if text.strip(" |"):
                            yield Block(text, "table" if kind == "paragraph" else kind)
                elif item.text.strip():
                    yield Block(item.text, kind)

        def margins(kind):
            seen = set()
            for section in doc.sections:
                for prefix in ("", "first_page_", "even_page_"):
                    story = getattr(section, prefix + kind)
                    if story.is_linked_to_previous or story.part in seen:
                        continue
                    seen.add(story.part)
                    yield from blocks(story, kind)

        headers = [] if options.ignore_header else list(margins("header"))
        footers = [] if options.ignore_footer else list(margins("footer"))
        return [*headers, *blocks(doc), *footers]

    @staticmethod
    def _extract_doc(path: Path, options: ExtractOptions) -> list[Block]:
        import pythoncom
        from win32com.client import DispatchEx

        def margins(doc, collection, kind):
            seen = set()
            for section in doc.Sections:
                for index in (1, 2, 3):
                    story = getattr(section, collection)(index)
                    text = story.Range.Text.strip() if story.Exists else ""
                    if text and text not in seen:
                        seen.add(text)
                        yield Block(text, kind)

        pythoncom.CoInitialize()
        try:
            word = DispatchEx("Word.Application")
            update_links = None
            try:
                word.Visible = False
                word.DisplayAlerts = 0
                word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
                update_links = word.Options.UpdateLinksAtOpen
                word.Options.UpdateLinksAtOpen = False
                doc = word.Documents.Open(
                    str(path.resolve()), ReadOnly=True, AddToRecentFiles=False,
                    Visible=False, ConfirmConversions=False, NoEncodingDialog=True,
                )
                try:
                    body = []
                    row_end = -1
                    for paragraph in doc.Content.Paragraphs:
                        rng = paragraph.Range
                        if rng.Start < row_end:
                            continue
                        if rng.Information(12):  # wdWithInTable
                            try:
                                row = rng.Rows(1)
                                row_end = row.Range.End
                                cells = [cell.Range.Text.rstrip("\r\x07") for cell in row.Cells]
                            except pythoncom.com_error:
                                # Word disallows row indexing in some merged tables.
                                row_end = -1
                                cells = [rng.Text.rstrip("\r\x07")]
                            if any(cell.strip() for cell in cells):
                                body.append(Block(" | ".join(cells), "table"))
                        else:
                            text = rng.Text.rstrip("\r")
                            if text.strip():
                                body.append(Block(text))

                    headers = [] if options.ignore_header else list(margins(doc, "Headers", "header"))
                    footers = [] if options.ignore_footer else list(margins(doc, "Footers", "footer"))
                    return [*headers, *body, *footers]
                finally:
                    doc.Close(SaveChanges=0)
            finally:
                try:
                    if update_links is not None:
                        word.Options.UpdateLinksAtOpen = update_links
                finally:
                    word.Quit(SaveChanges=0)
        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _extract_pdf(path: Path, options: ExtractOptions) -> list[Block]:
        import pdfplumber

        pages = []
        repeated = Counter()
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                tables = page.find_tables()
                boxes = [table.bbox for table in tables]
                lines = page.filter(
                    lambda obj, boxes=boxes: obj["object_type"] != "char" or not any(
                        x0 <= (obj["x0"] + obj["x1"]) / 2 <= x1
                        and top <= (obj["top"] + obj["bottom"]) / 2 <= bottom
                        for x0, top, x1, bottom in boxes
                    )
                ).extract_text_lines(return_chars=False)
                items = []
                page_margins = set()
                for line in lines:
                    text = line["text"]
                    kind = ""
                    if line["bottom"] <= page.bbox[1] + page.height * 0.10:
                        kind = "header"
                    elif line["top"] >= page.bbox[3] - page.height * 0.10:
                        kind = "footer"
                    if kind:
                        page_margins.add((kind, " ".join(text.split())))
                    items.append((line["top"], line["x0"], line["bottom"], text, kind))
                repeated.update(page_margins)
                for table in tables:
                    x0, top, _, bottom = table.bbox
                    for row in table.extract():
                        text = " | ".join(" ".join((cell or "").split()) for cell in row)
                        if text.strip(" |"):
                            items.append((top, x0, bottom, text, "table"))
                pages.append(sorted(items, key=lambda item: item[:2]))

        headers, body, footers = [], [], []
        emitted_margins = set()
        for number, items in enumerate(pages, start=1):
            paragraph = []
            previous_bottom = None
            for top, _, bottom, text, kind in items:
                key = (kind, " ".join(text.split()))
                if kind in {"header", "footer"} and repeated[key] >= 2:
                    ignore = options.ignore_header if kind == "header" else options.ignore_footer
                    if not ignore and key not in emitted_margins:
                        (headers if kind == "header" else footers).append(Block(text, kind, number))
                        emitted_margins.add(key)
                    continue
                if paragraph and (kind == "table" or top - previous_bottom > (bottom - top) * 0.9):
                    body.append(Block("\n".join(paragraph), page=number))
                    paragraph.clear()
                if kind == "table":
                    body.append(Block(text, kind, number))
                else:
                    paragraph.append(text)
                previous_bottom = bottom
            if paragraph:
                body.append(Block("\n".join(paragraph), page=number))

        if not (headers or body or footers):
            raise ValueError(t("review.empty_pdf").format(name=path.name))
        return [*headers, *body, *footers]
