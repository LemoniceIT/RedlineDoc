import json
import os
import tempfile
import unittest
from pathlib import Path

import pdfplumber
from docx import Document
from reportlab.pdfgen import canvas

from comparison import Block, compare
from document_processor import DocumentProcessor, ExtractOptions
from preferences import Preferences
from reports import export_report


class ComparisonTests(unittest.TestCase):
    def test_layout_changes_are_ignored(self):
        result = compare(
            [Block("Café with efficient service"), Block("Next paragraph")],
            [Block("Cafe\u0301 with eﬃcient\nservice\nNext paragraph", page=1)],
            "a.docx", "b.pdf",
        )
        self.assertEqual(sum(result.summary.values()), 0)

    def test_marks_changed_words_and_counts_moves_once(self):
        original = list(map(Block, ["Move this paragraph", "Terms apply", "Due in 30 days.", "Footer"]))
        modified = list(map(Block, ["Terms apply", "Due in 90 days.", "Footer", "Move this paragraph"]))
        result = compare(original, modified, "a.doc", "b.docx")
        self.assertEqual(result.summary, {"inserted": 0, "deleted": 0, "modified": 1, "moved": 1})
        modified_row = next(row for row in result.rows if row.kind == "modified")
        self.assertEqual([span.text for span in modified_row.before if span.changed], ["30"])
        self.assertEqual([span.text for span in modified_row.after if span.changed], ["90"])
        self.assertEqual([row.move for row in result.rows if row.kind == "moved"], [1, 1])

    def test_structural_differences_are_preserved(self):
        result = compare([Block("A | B", "table")], [Block("A | B")], "a.docx", "b.pdf")
        self.assertEqual(result.summary["modified"], 1)

    def test_insertions_and_deletions(self):
        result = compare(
            [Block("Removed"), Block("Shared")], [Block("Shared"), Block("Added")],
            "a.pdf", "b.pdf",
        )
        self.assertEqual(result.summary, {"inserted": 1, "deleted": 1, "modified": 0, "moved": 0})


class FileTests(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.folder = Path(folder.name)

    def test_docx_and_pdf_compare_with_table_in_document_order(self):
        docx_path = self.folder / "original.docx"
        docx = Document()
        docx.add_paragraph("The payment is due within 30 days.")
        cells = docx.add_table(rows=1, cols=2).rows[0].cells
        cells[0].text, cells[1].text = "Price", "100"
        docx.add_paragraph("All sales are final.")
        docx.save(docx_path)

        pdf_path = self.folder / "modified.pdf"
        pdf = canvas.Canvas(str(pdf_path))
        pdf.drawString(72, 720, "The payment is due")
        pdf.drawString(72, 706, "within 90 days.")
        pdf.rect(70, 620, 220, 40)
        pdf.line(180, 620, 180, 660)
        pdf.drawString(82, 636, "Price")
        pdf.drawString(192, 636, "100")
        pdf.drawString(72, 585, "All sales are final.")
        pdf.save()

        options = ExtractOptions()
        for path in (docx_path, pdf_path):
            blocks = DocumentProcessor.extract_blocks(path, options)
            self.assertEqual([block.kind for block in blocks], ["paragraph", "table", "paragraph"])
            self.assertEqual(blocks[1].text, "Price | 100")

        result = DocumentProcessor.review(docx_path, pdf_path, options)
        self.assertEqual(result.summary["modified"], 1)
        self.assertEqual(result.rows[0].modified[0].page, 1)
        self.assertEqual([span.text for span in result.rows[0].after if span.changed], ["90"])
        self.assertEqual(result.rows[1].kind, "equal")

    @unittest.skipUnless(os.environ.get("REDLINEDOC_TEST_WORD") == "1", "requires Microsoft Word")
    def test_legacy_doc_and_docx_match(self):
        from win32com.client import DispatchEx

        legacy, modern = self.folder / "test.doc", self.folder / "test.docx"
        word = DispatchEx("Word.Application")
        try:
            word.DisplayAlerts = 0
            doc = word.Documents.Add()
            try:
                doc.Range().Text = "Opening paragraph\rClosing paragraph"
                start = doc.Paragraphs(2).Range.Start
                table = doc.Tables.Add(doc.Range(start, start), 1, 2)
                table.Cell(1, 1).Range.Text = "Price"
                table.Cell(1, 2).Range.Text = "100"
                doc.Sections(1).Headers(1).Range.Text = "Header"
                doc.SaveAs2(str(legacy), FileFormat=0)
                doc.SaveAs2(str(modern), FileFormat=16)
            finally:
                doc.Close(SaveChanges=0)
        finally:
            word.Quit(SaveChanges=0)

        options = ExtractOptions()
        self.assertEqual(
            [(block.kind, block.text) for block in DocumentProcessor.extract_blocks(legacy, options)],
            [("header", "Header"), ("paragraph", "Opening paragraph"),
             ("table", "Price | 100"), ("paragraph", "Closing paragraph")],
        )
        self.assertEqual(sum(DocumentProcessor.review(legacy, modern, options).summary.values()), 0)

    def test_pdf_removes_repeated_margins_and_preserves_unique_body_text(self):
        path = self.folder / "multipage.pdf"
        pdf = canvas.Canvas(str(path))
        for number in (1, 2):
            pdf.drawString(72, 815, "Repeated header")
            pdf.drawString(72, 780, f"Body paragraph {number}")
            pdf.drawString(72, 25, "Repeated footer")
            pdf.showPage()
        pdf.save()

        all_blocks = DocumentProcessor.extract_blocks(path, ExtractOptions())
        self.assertEqual([block.kind for block in all_blocks], ["header", "paragraph", "paragraph", "footer"])
        body = DocumentProcessor.extract_blocks(path, ExtractOptions(ignore_header=True, ignore_footer=True))
        self.assertEqual([block.text for block in body], ["Body paragraph 1", "Body paragraph 2"])
        self.assertEqual([block.page for block in body], [1, 2])

    def test_pdf_without_text_is_rejected(self):
        path = self.folder / "blank.pdf"
        pdf = canvas.Canvas(str(path))
        pdf.showPage()
        pdf.save()
        with self.assertRaisesRegex(ValueError, "no comparable text"):
            DocumentProcessor.extract_blocks(path, ExtractOptions())

    def test_exports_and_escapes_untrusted_text(self):
        result = compare(
            [Block("Café & montant 30")], [Block("Café & montant 90", page=4)],
            "<b>reference</b>", "changed.pdf",
        )
        html, data, pdf = (self.folder / ("report." + ext) for ext in ("html", "json", "pdf"))
        for path in (html, data, pdf):
            export_report(result, path)
        markup = html.read_text(encoding="utf-8")
        self.assertIn("&lt;b&gt;reference&lt;/b&gt;", markup)
        self.assertIn("<mark>90</mark>", markup)
        self.assertIn("Page 4", markup)
        parsed = json.loads(data.read_text(encoding="utf-8"))
        self.assertEqual(parsed["schema_version"], 1)
        self.assertEqual(parsed["summary"]["modified"], 1)
        with pdfplumber.open(pdf) as report:
            self.assertIn("Café & montant 90", report.pages[0].extract_text())

    def test_pdf_export_splits_rows_across_pages(self):
        pdf = self.folder / "long.pdf"
        text = " ".join(f"mot{index}" for index in range(1200))
        export_report(compare([Block(text)], [Block(text)], "a.docx", "b.pdf"), pdf)
        with pdfplumber.open(pdf) as report:
            self.assertGreater(len(report.pages), 1)
            self.assertIn("mot1199", report.pages[-1].extract_text())

    def test_legacy_html_escapes_document_labels(self):
        path = self.folder / "legacy.html"
        DocumentProcessor.compare_documents("30 days", "90 days", "<i>original</i>", "modified", str(path))
        self.assertIn("&lt;i&gt;original&lt;/i&gt;", path.read_text(encoding="utf-8"))

    def test_preferences_round_trip_and_recover_from_invalid_json(self):
        path = self.folder / "preferences.json"
        expected = Preferences("it", True, False)
        self.assertEqual(Preferences.load(path), Preferences())
        expected.save(path)
        self.assertEqual(Preferences.load(path), expected)
        path.write_text('{', encoding="utf-8")
        self.assertEqual(Preferences.load(path), Preferences())


if __name__ == "__main__":
    unittest.main()
