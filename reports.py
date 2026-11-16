"""Write self-contained local reports for a comparison."""

import json
import os
from dataclasses import asdict
from html import escape
from pathlib import Path
from tempfile import NamedTemporaryFile
from xml.sax.saxutils import escape as xml_escape

from comparison import Change, Comparison, Span
from i18n import get_language, t


def summary_text(comparison: Comparison) -> str:
    return " · ".join(
        f"{t('review.' + kind)}: {count}" for kind, count in comparison.summary.items()
    )


def location_text(row: Change, side: str) -> str:
    blocks = getattr(row, side)
    pages = sorted({block.page for block in blocks if block.page is not None})
    return ", ".join(t("review.page").format(page=page) for page in pages)


def status_text(row: Change) -> str:
    label = t("review." + row.kind)
    return f"{label} #{row.move}" if row.move else label


def _html_cell(spans: tuple[Span, ...], location: str) -> str:
    text = "".join(
        f"<mark>{escape(span.text)}</mark>" if span.changed else escape(span.text)
        for span in spans
    )
    source = f"<small>{escape(location)}</small>" if location else ""
    return f"<td>{source}<div>{text}</div></td>"


def _html(comparison: Comparison) -> str:
    rows = []
    for row in comparison.rows:
        before = _html_cell(row.before, location_text(row, "original"))
        after = _html_cell(row.after, location_text(row, "modified"))
        rows.append(f'<tr class="{row.kind}"><th scope="row">{escape(status_text(row))}</th>{before}{after}</tr>')
    return f"""<!doctype html>
<html lang="{get_language()}">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>RedlineDoc — {escape(t("review.title"))}</title>
<style>
body {{ font: 16px system-ui; margin: 2rem; color: #152117; }}
table {{ border-collapse: collapse; width: 100%; table-layout: fixed; }}
th, td {{ border: 1px solid #b4c3b6; padding: .65rem; vertical-align: top; overflow-wrap: anywhere; }}
th:first-child {{ width: 18%; }}
td div {{ white-space: pre-wrap; }}
small {{ color: #465a49; display: block; margin-bottom: .3rem; }}
mark {{ background: #dcfce7; color: inherit; }}
td:nth-child(2) mark {{ background: #fee2e2; }}
.moved {{ background: #dbeafe; }}
</style>
<h1>{escape(t("review.title"))}</h1>
<p>{escape(summary_text(comparison))}</p>
<table>
<thead><tr><th>{escape(t("review.filter"))}</th>
<th>{escape(comparison.original_name)}</th><th>{escape(comparison.modified_name)}</th></tr></thead>
<tbody>{"".join(rows)}</tbody>
</table>
</html>
"""


def _pdf(comparison: Comparison, path: Path) -> None:
    from reportlab import rl_config
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import (
        LongTable,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        TableStyle,
    )

    windows_fonts = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    candidates = [windows_fonts / "segoeui.ttf"]
    candidates.extend(Path(folder) / "Vera.ttf" for folder in rl_config.TTFSearchPath)
    font_path = next((candidate for candidate in candidates if candidate.is_file()), None)
    if font_path:
        pdfmetrics.registerFont(TTFont("RedlineDoc", str(font_path)))
    font = "RedlineDoc" if font_path else "Helvetica"
    styles = getSampleStyleSheet()
    body = ParagraphStyle("Revision", parent=styles["BodyText"], fontName=font)
    title = ParagraphStyle("RevisionTitle", parent=styles["Heading1"], fontName=font)
    doc = SimpleDocTemplate(str(path), title=f'RedlineDoc — {t("review.title")}')

    def markup(text):
        safe = "".join(char for char in text if char == "\n" or ord(char) >= 32)
        return xml_escape(safe).replace("\n", "<br/>")

    def cell(spans, location, highlight):
        text = "".join(
            f'<font backColor="{highlight}">{markup(span.text)}</font>'
            if span.changed else markup(span.text)
            for span in spans
        )
        if location:
            text = markup(location) + "<br/>" + text
        return Paragraph(text, body)

    rows = [[
        Paragraph(markup(t("review.filter")), body),
        Paragraph(markup(comparison.original_name), body),
        Paragraph(markup(comparison.modified_name), body),
    ]]
    for row in comparison.rows:
        rows.append([
            Paragraph(markup(status_text(row)), body),
            cell(row.before, location_text(row, "original"), "#fee2e2"),
            cell(row.after, location_text(row, "modified"), "#dcfce7"),
        ])
    table = LongTable(
        rows, colWidths=[doc.width * fraction for fraction in (0.18, 0.41, 0.41)],
        repeatRows=1, splitInRow=1,
    )
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dcefe0")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#b4c3b6")),
    ]))
    doc.build([
        Paragraph(markup(t("review.title")), title),
        Paragraph(markup(summary_text(comparison)), body),
        Spacer(1, 12),
        table,
    ])


def export_report(comparison: Comparison, output_path: Path) -> None:
    extension = output_path.suffix.lower()
    if extension not in {".html", ".json", ".pdf"}:
        raise ValueError(t("doc.unsupported_format").format(ext=extension))
    with NamedTemporaryFile(dir=output_path.parent, delete=False) as file:
        temporary_path = Path(file.name)
    try:
        if extension == ".html":
            temporary_path.write_text(_html(comparison), encoding="utf-8")
        elif extension == ".json":
            payload = {"schema_version": 1, "summary": comparison.summary, **asdict(comparison)}
            temporary_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        else:
            _pdf(comparison, temporary_path)
        temporary_path.replace(output_path)
    finally:
        temporary_path.unlink(missing_ok=True)
