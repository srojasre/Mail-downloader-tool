"""Conversores de cuerpo de correo y archivos adjuntos a formato PDF."""

from __future__ import annotations

import csv
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import textwrap
import zipfile
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape as xml_escape

import pymupdf as fitz
from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document as DocxDocument
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from lxml import etree
from openpyxl import load_workbook
from PIL import Image, ImageOps, ImageSequence
from pptx import Presentation
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    Image as RLImage,
    PageBreak,
    Paragraph,
    Preformatted,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from src.imap_client import EmailMessageData
from src.utils import get_unique_path

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".tiff", ".tif", ".bmp", ".gif"}
TEXT_EXTENSIONS = {".txt", ".log", ".md", ".json", ".ics", ".vcf", ".yml", ".yaml", ".ini", ".cfg"}
CSV_EXTENSIONS = {".csv", ".tsv"}
HTML_EXTENSIONS = {".html", ".htm"}
OFFICE_WORD_EXTENSIONS = {".docx", ".doc"}
OFFICE_EXCEL_EXTENSIONS = {".xlsx", ".xls"}
OFFICE_PPT_EXTENSIONS = {".pptx", ".ppt"}
OFFICE_EXTENSIONS = OFFICE_WORD_EXTENSIONS | OFFICE_EXCEL_EXTENSIONS | OFFICE_PPT_EXTENSIONS

_FONTS_INITIALIZED = False
_FONT_SANS = "Helvetica"
_FONT_SANS_BOLD = "Helvetica-Bold"
_FONT_SANS_ITALIC = "Helvetica-Oblique"
_FONT_MONO = "Courier"
_HAS_UNICODE_FONTS = False


def _init_fonts() -> None:
    """Registra fuentes TrueType de Windows si están disponibles para soporte Unicode completo."""
    global _FONTS_INITIALIZED, _FONT_SANS, _FONT_SANS_BOLD, _FONT_SANS_ITALIC, _FONT_MONO, _HAS_UNICODE_FONTS
    if _FONTS_INITIALIZED:
        return
    _FONTS_INITIALIZED = True

    win_fonts = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    arial = win_fonts / "arial.ttf"
    arial_bd = win_fonts / "arialbd.ttf"
    arial_it = win_fonts / "ariali.ttf"
    arial_bi = win_fonts / "arialbi.ttf"
    consola = win_fonts / "consola.ttf"
    cour = win_fonts / "cour.ttf"

    try:
        if arial.exists() and arial_bd.exists():
            pdfmetrics.registerFont(TTFont("AppSans", str(arial)))
            pdfmetrics.registerFont(TTFont("AppSans-Bold", str(arial_bd)))
            italic_path = arial_it if arial_it.exists() else arial
            bold_italic_path = arial_bi if arial_bi.exists() else arial_bd
            pdfmetrics.registerFont(TTFont("AppSans-Italic", str(italic_path)))
            pdfmetrics.registerFont(TTFont("AppSans-BoldItalic", str(bold_italic_path)))
            registerFontFamily(
                "AppSans",
                normal="AppSans",
                bold="AppSans-Bold",
                italic="AppSans-Italic",
                boldItalic="AppSans-BoldItalic",
            )
            _FONT_SANS = "AppSans"
            _FONT_SANS_BOLD = "AppSans-Bold"
            _FONT_SANS_ITALIC = "AppSans-Italic"
            _HAS_UNICODE_FONTS = True

        mono_path = consola if consola.exists() else (cour if cour.exists() else None)
        if mono_path is not None:
            pdfmetrics.registerFont(TTFont("AppMono", str(mono_path)))
            _FONT_MONO = "AppMono"
    except Exception:
        # Mantiene Helvetica/Courier si alguna fuente falla al cargarse
        pass


def _clean_text_for_pdf(text: str | None) -> str:
    """Limpia caracteres no soportados por ReportLab/fuentes estándar manteniendo tildes y símbolos."""
    _init_fonts()
    if not text:
        return ""

    cleaned_chars: list[str] = []
    for ch in str(text):
        code = ord(ch)
        if ch in ("\n", "\r", "\t"):
            cleaned_chars.append("    " if ch == "\t" else ch)
        elif code < 32 or code == 127:
            continue
        elif code > 0xFFFF:
            # Emojis suplementarios fuera del BMP no suelen estar en Arial TTF
            continue
        elif not _HAS_UNICODE_FONTS and code > 255:
            # Reemplazos comunes para fuentes WinAnsi fallback
            replacements = {
                "€": "EUR",
                "•": "*",
                "–": "-",
                "—": "-",
                "‘": "'",
                "’": "'",
                "“": '"',
                "”": '"',
                "…": "...",
            }
            cleaned_chars.append(replacements.get(ch, "?"))
        else:
            cleaned_chars.append(ch)
    return "".join(cleaned_chars)


def _p_text(text: str | None) -> str:
    """Escapa texto para usarlo de forma segura dentro de reportlab.platypus.Paragraph."""
    cleaned = _clean_text_for_pdf(text)
    escaped = xml_escape(cleaned)
    return escaped.replace("\n", "<br/>")


def _build_styles() -> dict[str, ParagraphStyle]:
    """Construye la hoja de estilos estándar para documentos PDF generados."""
    _init_fonts()
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "AppTitle",
            parent=base["Normal"],
            fontName=_FONT_SANS_BOLD,
            fontSize=15,
            leading=19,
            textColor=colors.HexColor("#111827"),
            spaceAfter=8,
        ),
        "h1": ParagraphStyle(
            "AppH1",
            parent=base["Normal"],
            fontName=_FONT_SANS_BOLD,
            fontSize=13,
            leading=17,
            textColor=colors.HexColor("#1F2937"),
            spaceBefore=10,
            spaceAfter=5,
        ),
        "h2": ParagraphStyle(
            "AppH2",
            parent=base["Normal"],
            fontName=_FONT_SANS_BOLD,
            fontSize=11,
            leading=15,
            textColor=colors.HexColor("#374151"),
            spaceBefore=8,
            spaceAfter=4,
        ),
        "body": ParagraphStyle(
            "AppBody",
            parent=base["Normal"],
            fontName=_FONT_SANS,
            fontSize=9.5,
            leading=13.5,
            textColor=colors.HexColor("#1F2937"),
            spaceAfter=5,
            alignment=TA_LEFT,
        ),
        "small": ParagraphStyle(
            "AppSmall",
            parent=base["Normal"],
            fontName=_FONT_SANS,
            fontSize=8,
            leading=10.5,
            textColor=colors.HexColor("#4B5563"),
            spaceAfter=3,
        ),
        "small_bold": ParagraphStyle(
            "AppSmallBold",
            parent=base["Normal"],
            fontName=_FONT_SANS_BOLD,
            fontSize=8,
            leading=10.5,
            textColor=colors.HexColor("#111827"),
        ),
        "small_right": ParagraphStyle(
            "AppSmallRight",
            parent=base["Normal"],
            fontName=_FONT_SANS,
            fontSize=8,
            leading=10.5,
            textColor=colors.HexColor("#111827"),
            alignment=TA_RIGHT,
        ),
        "mono": ParagraphStyle(
            "AppMonoStyle",
            parent=base["Normal"],
            fontName=_FONT_MONO,
            fontSize=8,
            leading=10.5,
            textColor=colors.HexColor("#111827"),
        ),
        "quote": ParagraphStyle(
            "AppQuote",
            parent=base["Normal"],
            fontName=_FONT_SANS_ITALIC,
            fontSize=9,
            leading=12.5,
            leftIndent=14,
            textColor=colors.HexColor("#4B5563"),
            spaceBefore=4,
            spaceAfter=6,
        ),
    }


def _create_doc(output_path: Path, is_landscape: bool = False, margins_mm: float = 15.0) -> SimpleDocTemplate:
    pagesize = landscape(A4) if is_landscape else A4
    margin = margins_mm * mm
    return SimpleDocTemplate(
        str(output_path),
        pagesize=pagesize,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin,
    )


def _decode_bytes(payload: bytes) -> str:
    """Decodifica bytes a texto probando codificaciones comunes."""
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return payload.decode(enc)
        except UnicodeDecodeError:
            continue
    return payload.decode("utf-8", errors="replace")


def _extract_inline_html_markup(node: Tag | NavigableString) -> str:
    """Convierte nodos inline HTML en markup compatible con ReportLab Paragraph."""
    if isinstance(node, NavigableString):
        text = re.sub(r"\s+", " ", str(node))
        return xml_escape(_clean_text_for_pdf(text))

    if not isinstance(node, Tag):
        return ""

    name = node.name.lower()
    if name in {"script", "style", "head", "noscript", "iframe", "svg"}:
        return ""
    if name == "br":
        return "<br/>"
    if name == "img":
        alt = (node.get("alt") or "").strip()
        return xml_escape(_clean_text_for_pdf(f"[Imagen: {alt}]" if alt else "[Imagen]"))

    inner = "".join(_extract_inline_html_markup(child) for child in node.children)
    if name in {"b", "strong"}:
        return f"<b>{inner}</b>" if inner.strip() else ""
    if name in {"i", "em"}:
        return f"<i>{inner}</i>" if inner.strip() else ""
    if name == "u":
        return f"<u>{inner}</u>" if inner.strip() else ""
    if name == "code":
        return f'<font face="{_FONT_MONO}">{inner}</font>' if inner.strip() else ""
    if name == "a":
        href = (node.get("href") or "").strip()
        if href and href.startswith(("http://", "https://", "mailto:")) and inner.strip():
            safe_href = xml_escape(_clean_text_for_pdf(href), {'"': "&quot;"})
            return f'<link href="{safe_href}" color="#1D4ED8"><u>{inner}</u></link>'
        return inner

    return inner


def _html_table_to_flowable(table_tag: Tag, styles: dict[str, ParagraphStyle], max_width: float) -> Table | None:
    """Convierte una etiqueta <table> HTML en una Table de ReportLab."""
    rows_data: list[list[Paragraph]] = []
    for tr in table_tag.find_all("tr", recursive=True):
        # Evita procesar filas de tablas profundamente anidadas en este nivel si ya pertenecen a otra subtabla
        parent_table = tr.find_parent("table")
        if parent_table is not table_tag:
            continue
        cells = tr.find_all(["th", "td"], recursive=False)
        if not cells:
            continue
        row: list[Paragraph] = []
        for cell in cells:
            is_header = cell.name.lower() == "th"
            cell_text = re.sub(r"\s+", " ", cell.get_text(" ", strip=True))
            style = styles["small_bold"] if is_header else styles["small"]
            row.append(Paragraph(_p_text(cell_text), style))
        if row:
            rows_data.append(row)

    if not rows_data:
        return None

    max_cols = max(len(r) for r in rows_data)
    if max_cols == 0:
        return None

    # Si la tabla de maquetación de correo tiene 1 celda o columnas muy dispares y texto extenso,
    # es mejor dejarla pasar o normalizar el número de columnas.
    for r in rows_data:
        while len(r) < max_cols:
            r.append(Paragraph("", styles["small"]))

    # Limita columnas extremas de maquetaciones de email
    if max_cols > 10:
        rows_data = [r[:10] for r in rows_data]
        max_cols = 10

    col_width = max_width / max_cols
    tbl = Table(rows_data, colWidths=[col_width] * max_cols, repeatRows=1 if len(rows_data) > 1 else 0)
    tbl.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return tbl


def _html_to_flowables(raw_html: str, styles: dict[str, ParagraphStyle], max_width: float) -> list[Any]:
    """Convierte HTML en una lista de elementos Flowable de ReportLab."""
    try:
        soup = BeautifulSoup(raw_html, "lxml")
    except Exception:
        soup = BeautifulSoup(raw_html, "html.parser")

    for tag in soup(["script", "style", "head", "title", "meta", "link", "noscript", "iframe", "svg"]):
        tag.decompose()

    root = soup.body if soup.body else soup
    flowables: list[Any] = []

    def _walk_block(node: Tag | NavigableString) -> None:
        if isinstance(node, NavigableString):
            txt = re.sub(r"\s+", " ", str(node)).strip()
            if txt:
                flowables.append(Paragraph(_p_text(txt), styles["body"]))
            return

        if not isinstance(node, Tag):
            return

        tag_name = node.name.lower()
        if tag_name in {"h1", "h2"}:
            txt = node.get_text(" ", strip=True)
            if txt:
                flowables.append(Paragraph(_p_text(txt), styles["h1"]))
        elif tag_name in {"h3", "h4", "h5", "h6"}:
            txt = node.get_text(" ", strip=True)
            if txt:
                flowables.append(Paragraph(_p_text(txt), styles["h2"]))
        elif tag_name == "hr":
            flowables.append(HRFlowable(width="100%", thickness=0.6, color=colors.HexColor("#D1D5DB")))
            flowables.append(Spacer(1, 4))
        elif tag_name == "pre":
            txt = _clean_text_for_pdf(node.get_text())
            if txt.strip():
                wrapped = "\n".join(textwrap.fill(line, width=95) for line in txt.splitlines())
                flowables.append(Preformatted(wrapped, styles["mono"]))
                flowables.append(Spacer(1, 4))
        elif tag_name == "blockquote":
            txt = node.get_text("\n", strip=True)
            if txt:
                flowables.append(Paragraph(_p_text(txt), styles["quote"]))
        elif tag_name in {"ul", "ol"}:
            items = node.find_all("li", recursive=False)
            for idx, li in enumerate(items, start=1):
                prefix = f"{idx}. " if tag_name == "ol" else "• "
                markup = _extract_inline_html_markup(li).strip()
                if markup:
                    flowables.append(Paragraph(f"{xml_escape(_clean_text_for_pdf(prefix))}{markup}", styles["body"]))
            flowables.append(Spacer(1, 3))
        elif tag_name == "table":
            # Muchas newsletters usan tablas de 1 celda solo como layout; en ese caso recorremos sus hijos
            direct_trs = [tr for tr in node.find_all("tr", recursive=True) if tr.find_parent("table") is node]
            max_cells = max((len(tr.find_all(["th", "td"], recursive=False)) for tr in direct_trs), default=0)
            nested_tables = node.find_all("table")
            if max_cells <= 1 or nested_tables:
                for child in node.children:
                    _walk_block(child)
            else:
                tbl = _html_table_to_flowable(node, styles, max_width)
                if tbl is not None:
                    flowables.append(tbl)
                    flowables.append(Spacer(1, 6))
        elif tag_name == "p":
            markup = _extract_inline_html_markup(node).strip()
            if markup:
                flowables.append(Paragraph(markup, styles["body"]))
        elif tag_name in {"div", "section", "article", "main", "header", "footer", "tbody", "thead", "tr", "td", "th", "center"}:
            has_block_child = any(
                isinstance(c, Tag)
                and c.name.lower()
                in {"p", "div", "table", "ul", "ol", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre", "section", "article"}
                for c in node.children
            )
            if has_block_child:
                for child in node.children:
                    _walk_block(child)
            else:
                markup = _extract_inline_html_markup(node).strip()
                if markup:
                    flowables.append(Paragraph(markup, styles["body"]))
        else:
            markup = _extract_inline_html_markup(node).strip()
            if markup:
                flowables.append(Paragraph(markup, styles["body"]))

    for child in root.children:
        _walk_block(child)

    return flowables


def render_email_body_to_pdf(email_data: EmailMessageData, output_pdf_path: Path) -> Path:
    """Genera el archivo PDF del cuerpo y metadatos del correo."""
    output_pdf_path.parent.mkdir(parents=True, exist_ok=True)
    styles = _build_styles()
    doc = _create_doc(output_pdf_path, is_landscape=False, margins_mm=16.0)
    usable_width = A4[0] - 32.0 * mm

    story: list[Any] = []
    story.append(Paragraph(_p_text(email_data.subject or "Sin_Asunto"), styles["title"]))
    story.append(Spacer(1, 4))

    att_names = ", ".join(att.filename for att in email_data.attachments) if email_data.attachments else "Ninguno"
    meta_rows = [
        ("De:", email_data.sender or "-"),
        ("Para:", email_data.recipients or "-"),
    ]
    if email_data.cc:
        meta_rows.append(("CC:", email_data.cc))
    meta_rows.extend(
        [
            ("Fecha:", email_data.date_str or "-"),
            ("Adjuntos:", att_names),
        ]
    )

    table_data = [
        [Paragraph(_p_text(label), styles["small_bold"]), Paragraph(_p_text(val), styles["small"])]
        for label, val in meta_rows
    ]
    meta_table = Table(table_data, colWidths=[2.6 * cm, usable_width - 2.6 * cm])
    meta_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F3F4F6")),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D1D5DB")),
                ("LINEBELOW", (0, 0), (-1, -2), 0.3, colors.HexColor("#E5E7EB")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(meta_table)
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor("#9CA3AF")))
    story.append(Spacer(1, 10))

    body_added = False
    if email_data.body_html.strip():
        html_flowables = _html_to_flowables(email_data.body_html, styles, usable_width)
        if html_flowables:
            story.extend(html_flowables)
            body_added = True

    if not body_added and email_data.body_text.strip():
        for block in re.split(r"\r?\n\s*\r?\n", email_data.body_text.strip()):
            if block.strip():
                story.append(Paragraph(_p_text(block.strip()), styles["body"]))
        body_added = True

    if not body_added:
        story.append(Paragraph(_p_text("(Este correo no contiene texto en el cuerpo del mensaje.)"), styles["quote"]))

    try:
        doc.build(story)
    except Exception:
        # Fallback ultra seguro en texto plano si un párrafo HTML raro provoca error en ReportLab
        fallback_doc = _create_doc(output_pdf_path, is_landscape=False, margins_mm=16.0)
        plain_body = email_data.body_text.strip()
        if not plain_body and email_data.body_html.strip():
            plain_body = BeautifulSoup(email_data.body_html, "html.parser").get_text("\n", strip=True)
        fallback_story = [
            Paragraph(_p_text(email_data.subject or "Sin_Asunto"), styles["title"]),
            meta_table,
            Spacer(1, 10),
            Paragraph(_p_text(plain_body or "(Sin contenido textual)"), styles["body"]),
        ]
        fallback_doc.build(fallback_story)

    return output_pdf_path


def _convert_image_to_pdf(payload: bytes, output_pdf_path: Path) -> Path:
    """Convierte imágenes (incluso multipágina TIFF/GIF) a un PDF bien encuadrado en páginas A4."""
    pdf_doc = fitz.open()
    try:
        with Image.open(io.BytesIO(payload)) as img:
            for frame in ImageSequence.Iterator(img):
                oriented = ImageOps.exif_transpose(frame)
                if oriented.mode in ("RGBA", "LA", "P"):
                    rgba = oriented.convert("RGBA")
                    bg = Image.new("RGB", rgba.size, (255, 255, 255))
                    bg.paste(rgba, mask=rgba.split()[3])
                    rgb = bg
                elif oriented.mode != "RGB":
                    rgb = oriented.convert("RGB")
                else:
                    rgb = oriented

                w, h = rgb.size
                is_land = w > h
                page_w, page_h = (842, 595) if is_land else (595, 842)
                margin = 24
                max_w = page_w - 2 * margin
                max_h = page_h - 2 * margin
                scale = min(max_w / max(w, 1), max_h / max(h, 1))
                draw_w = w * scale
                draw_h = h * scale
                x0 = (page_w - draw_w) / 2
                y0 = (page_h - draw_h) / 2

                img_buf = io.BytesIO()
                rgb.save(img_buf, format="PNG")
                page = pdf_doc.new_page(width=page_w, height=page_h)
                page.insert_image(fitz.Rect(x0, y0, x0 + draw_w, y0 + draw_h), stream=img_buf.getvalue())

        pdf_doc.save(str(output_pdf_path))
    finally:
        pdf_doc.close()
    return output_pdf_path


def _render_text_lines_to_pdf(title: str, text: str, output_pdf_path: Path, mono: bool = True) -> Path:
    """Renderiza texto plano o código formateado en un PDF A4 paginado."""
    styles = _build_styles()
    doc = _create_doc(output_pdf_path, is_landscape=False, margins_mm=15.0)
    story: list[Any] = [
        Paragraph(_p_text(title), styles["h1"]),
        HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#D1D5DB")),
        Spacer(1, 8),
    ]

    cleaned = _clean_text_for_pdf(text)
    if not cleaned.strip():
        story.append(Paragraph(_p_text("(Archivo vacío)"), styles["quote"]))
    elif mono:
        wrapped_lines: list[str] = []
        for raw_line in cleaned.splitlines():
            if not raw_line:
                wrapped_lines.append("")
            else:
                wrapped_lines.extend(textwrap.wrap(raw_line, width=96, replace_whitespace=False, drop_whitespace=False))
        # Dividir en bloques para no saturar un único Flowable en archivos largos
        chunk_size = 45
        for i in range(0, len(wrapped_lines), chunk_size):
            chunk = "\n".join(wrapped_lines[i : i + chunk_size])
            story.append(Preformatted(chunk, styles["mono"]))
    else:
        for para in re.split(r"\r?\n\s*\r?\n", cleaned):
            if para.strip():
                story.append(Paragraph(_p_text(para.strip()), styles["body"]))

    doc.build(story)
    return output_pdf_path


def _convert_text_to_pdf(payload: bytes, output_pdf_path: Path, filename: str, is_json: bool = False) -> Path:
    """Convierte archivos de texto plano, logs, markdown o JSON a PDF."""
    text = _decode_bytes(payload)
    if is_json:
        try:
            parsed = json.loads(text)
            text = json.dumps(parsed, indent=2, ensure_ascii=False)
        except Exception:
            pass
    return _render_text_lines_to_pdf(filename, text, output_pdf_path, mono=True)


def _render_tabular_rows_to_pdf(
    title: str,
    sections: list[tuple[str, list[list[str]]]],
    output_pdf_path: Path,
) -> Path:
    """Renderiza una o más tablas de datos (CSV o hojas Excel) a PDF."""
    styles = _build_styles()
    max_cols = max((max((len(r) for r in rows), default=0) for _, rows in sections), default=1)
    is_landscape = max_cols > 5
    page_w = (A4[1] if is_landscape else A4[0]) - 24.0 * mm
    doc = _create_doc(output_pdf_path, is_landscape=is_landscape, margins_mm=12.0)

    story: list[Any] = [Paragraph(_p_text(title), styles["h1"]), Spacer(1, 6)]

    for sec_idx, (section_name, rows) in enumerate(sections):
        if sec_idx > 0:
            story.append(PageBreak())
        if section_name:
            story.append(Paragraph(_p_text(section_name), styles["h2"]))
            story.append(Spacer(1, 4))

        if not rows:
            story.append(Paragraph(_p_text("(Sin datos)"), styles["quote"]))
            continue

        col_count = max(len(r) for r in rows)
        if col_count == 0:
            story.append(Paragraph(_p_text("(Tabla vacía)"), styles["quote"]))
            continue

        # Limita a 14 columnas por ancho de página legible, agregando aviso si se recorta
        truncated_cols = col_count > 14
        visible_cols = min(col_count, 14)

        table_rows: list[list[Paragraph]] = []
        for r_idx, row in enumerate(rows):
            padded = [(row[c] if c < len(row) else "") for c in range(visible_cols)]
            cell_style = styles["small_bold"] if r_idx == 0 else styles["small"]
            table_rows.append([Paragraph(_p_text(val), cell_style) for val in padded])

        col_w = page_w / visible_cols
        tbl = Table(table_rows, colWidths=[col_w] * visible_cols, repeatRows=1)
        tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E5E7EB")),
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(tbl)
        if truncated_cols:
            story.append(Spacer(1, 4))
            story.append(
                Paragraph(
                    _p_text(f"Nota: La tabla tiene {col_count} columnas; se muestran las primeras 14 por ajuste de página."),
                    styles["small"],
                )
            )

    doc.build(story)
    return output_pdf_path


def _convert_csv_to_pdf(payload: bytes, output_pdf_path: Path, filename: str, delimiter_hint: str | None = None) -> Path:
    """Convierte un archivo CSV o TSV en una tabla PDF formateada."""
    text = _decode_bytes(payload)
    sample = text[:4096]
    delimiter = delimiter_hint or ","
    if delimiter_hint is None:
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            delimiter = dialect.delimiter
        except csv.Error:
            delimiter = ";" if sample.count(";") > sample.count(",") else ","

    rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    return _render_tabular_rows_to_pdf(filename, [("", rows)], output_pdf_path)


def _convert_html_file_to_pdf(payload: bytes, output_pdf_path: Path, filename: str) -> Path:
    """Convierte un archivo adjunto HTML/HTM a PDF usando PyMuPDF Story o fallback ReportLab."""
    raw_html = _decode_bytes(payload)
    try:
        soup = BeautifulSoup(raw_html, "lxml")
    except Exception:
        soup = BeautifulSoup(raw_html, "html.parser")

    for tag in soup(["script", "noscript", "iframe", "object", "embed"]):
        tag.decompose()
    # Evita llamadas de red al renderizar imágenes remotas en Story
    for img in soup.find_all("img"):
        src = (img.get("src") or "").strip()
        if src.startswith(("http://", "https://")):
            alt = (img.get("alt") or "Imagen externa").strip()
            img.replace_with(f"[{alt}]")

    sanitized_html = str(soup)
    try:
        story = fitz.Story(html=sanitized_html)
        writer = fitz.DocumentWriter(str(output_pdf_path))
        mediabox = fitz.paper_rect("a4")
        where = mediabox + (36, 36, -36, -36)
        more = True
        pages = 0
        while more and pages < 200:
            dev = writer.begin_page(mediabox)
            more, _ = story.place(where)
            story.draw(dev, None)
            writer.end_page()
            pages += 1
        writer.close()
        if pages > 0 and output_pdf_path.exists() and output_pdf_path.stat().st_size > 0:
            return output_pdf_path
    except Exception:
        pass

    styles = _build_styles()
    doc = _create_doc(output_pdf_path, is_landscape=False, margins_mm=15.0)
    usable_width = A4[0] - 30.0 * mm
    flowables = _html_to_flowables(sanitized_html, styles, usable_width)
    if not flowables:
        flowables = [Paragraph(_p_text(soup.get_text("\n", strip=True) or "(HTML vacío)"), styles["body"])]
    doc.build(flowables)
    return output_pdf_path


def _localname(tag: str) -> str:
    """Extrae el nombre local de una etiqueta XML ignorando su namespace."""
    if not isinstance(tag, str):
        return ""
    if "}" in tag:
        return tag.rsplit("}", 1)[-1]
    if ":" in tag:
        return tag.rsplit(":", 1)[-1]
    return tag


def _find_first_text(node: etree._Element | None, names: tuple[str, ...]) -> str:
    """Busca recursivamente el primer subelemento cuyo localname esté en names y retorna su texto."""
    if node is None:
        return ""
    for el in node.iter():
        if _localname(el.tag) in names and el.text and el.text.strip():
            return el.text.strip()
    return ""


def _find_direct_or_child_text(node: etree._Element | None, names: tuple[str, ...]) -> str:
    return _find_first_text(node, names)


def _find_first_element(node: etree._Element | None, names: tuple[str, ...]) -> etree._Element | None:
    if node is None:
        return None
    for el in node.iter():
        if _localname(el.tag) in names:
            return el
    return None


def _find_all_elements(node: etree._Element | None, names: tuple[str, ...]) -> list[etree._Element]:
    if node is None:
        return []
    return [el for el in node.iter() if _localname(el.tag) in names]


def _unwrap_attached_document(root: etree._Element) -> tuple[etree._Element, str]:
    """Si el XML es un AttachedDocument (UBL/DIAN), extrae la factura embebida en CDATA."""
    root_name = _localname(root.tag)
    parent_cufe = ""
    if root_name == "AttachedDocument":
        parent_cufe = _find_first_text(root, ("UUID",))
        for desc in _find_all_elements(root, ("Description",)):
            inner_text = (desc.text or "").strip()
            if "<" in inner_text and any(k in inner_text for k in ("Invoice", "CreditNote", "DebitNote")):
                try:
                    parser = etree.XMLParser(recover=True)
                    inner_root = etree.fromstring(inner_text.encode("utf-8"), parser=parser)
                    if inner_root is not None:
                        return inner_root, parent_cufe
                except Exception:
                    continue
    return root, parent_cufe


def _extract_xml_flat_nodes(root: etree._Element, max_rows: int = 250) -> list[tuple[str, str]]:
    """Extrae pares (ruta/etiqueta, valor) de los nodos hoja y atributos relevantes del XML."""
    rows: list[tuple[str, str]] = []

    def _visit(el: etree._Element, path: str) -> None:
        if len(rows) >= max_rows:
            return
        lname = _localname(el.tag)
        if not lname:
            return
        current_path = f"{path}/{lname}" if path else lname

        for attr_k, attr_v in el.attrib.items():
            if len(rows) >= max_rows:
                return
            attr_name = _localname(attr_k)
            if attr_name not in ("schemaLocation",) and attr_v.strip():
                rows.append((f"{current_path} [@{attr_name}]", attr_v.strip()))

        text_val = (el.text or "").strip()
        if text_val and len(text_val) < 1000 and not text_val.startswith("<"):
            rows.append((current_path, text_val))

        for child in el:
            _visit(child, current_path)

    _visit(root, "")
    return rows


def _convert_xml_to_pdf(payload: bytes, output_pdf_path: Path, filename: str) -> Path:
    """Convierte facturas electrónicas (UBL/DIAN, CFDI) o XML general a un PDF legible."""
    parser = etree.XMLParser(recover=True, resolve_entities=False)
    try:
        raw_root = etree.fromstring(payload, parser=parser)
    except Exception:
        return _convert_text_to_pdf(payload, output_pdf_path, filename, is_json=False)

    if raw_root is None:
        return _convert_text_to_pdf(payload, output_pdf_path, filename, is_json=False)

    root, fallback_cufe = _unwrap_attached_document(raw_root)
    doc_type = _localname(root.tag)
    styles = _build_styles()
    doc = _create_doc(output_pdf_path, is_landscape=False, margins_mm=14.0)
    usable_width = A4[0] - 28.0 * mm
    story: list[Any] = []

    is_ubl_invoice = doc_type in {"Invoice", "CreditNote", "DebitNote"}
    is_cfdi = doc_type == "Comprobante"

    if is_ubl_invoice:
        doc_title_map = {
            "Invoice": "Factura Electrónica (XML UBL)",
            "CreditNote": "Nota Crédito Electrónica (XML UBL)",
            "DebitNote": "Nota Débito Electrónica (XML UBL)",
        }
        story.append(Paragraph(_p_text(doc_title_map.get(doc_type, doc_type)), styles["title"]))
        story.append(Paragraph(_p_text(f"Archivo fuente: {filename}"), styles["small"]))
        story.append(Spacer(1, 6))

        inv_id = _find_direct_or_child_text(root, ("ID",))
        issue_date = _find_direct_or_child_text(root, ("IssueDate",))
        issue_time = _find_direct_or_child_text(root, ("IssueTime",))
        currency = _find_direct_or_child_text(root, ("DocumentCurrencyCode",))
        cufe = _find_direct_or_child_text(root, ("UUID",)) or fallback_cufe

        supplier = _find_first_element(root, ("AccountingSupplierParty",))
        customer = _find_first_element(root, ("AccountingCustomerParty",))

        sup_name = _find_first_text(supplier, ("RegistrationName", "Name"))
        sup_id = _find_first_text(supplier, ("CompanyID", "ID"))
        sup_addr = _find_first_text(supplier, ("Line", "StreetName"))
        sup_city = _find_first_text(supplier, ("CityName",))
        sup_email = _find_first_text(supplier, ("ElectronicMail",))

        cus_name = _find_first_text(customer, ("RegistrationName", "Name"))
        cus_id = _find_first_text(customer, ("CompanyID", "ID"))
        cus_addr = _find_first_text(customer, ("Line", "StreetName"))
        cus_city = _find_first_text(customer, ("CityName",))
        cus_email = _find_first_text(customer, ("ElectronicMail",))

        summary_rows = [
            [
                Paragraph(_p_text("Número de Documento:"), styles["small_bold"]),
                Paragraph(_p_text(inv_id or "-"), styles["small"]),
                Paragraph(_p_text("Fecha / Hora Emisión:"), styles["small_bold"]),
                Paragraph(_p_text(f"{issue_date} {issue_time}".strip() or "-"), styles["small"]),
            ],
            [
                Paragraph(_p_text("Moneda:"), styles["small_bold"]),
                Paragraph(_p_text(currency or "-"), styles["small"]),
                Paragraph(_p_text("CUFE / UUID:"), styles["small_bold"]),
                Paragraph(_p_text(cufe or "-"), styles["small"]),
            ],
        ]
        hdr_tbl = Table(summary_rows, colWidths=[3.5 * cm, 4.8 * cm, 3.6 * cm, usable_width - 11.9 * cm])
        hdr_tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F9FAFB")),
                    ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D1D5DB")),
                    ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#E5E7EB")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(hdr_tbl)
        story.append(Spacer(1, 8))

        party_rows = [
            [
                Paragraph(_p_text("EMISOR / PROVEEDOR"), styles["small_bold"]),
                Paragraph(_p_text("CLIENTE / ADQUIRIENTE"), styles["small_bold"]),
            ],
            [
                Paragraph(
                    _p_text(
                        "\n".join(
                            line
                            for line in [
                                sup_name or "-",
                                f"ID / NIT: {sup_id}" if sup_id else "",
                                f"Dirección: {sup_addr} {sup_city}".strip() if (sup_addr or sup_city) else "",
                                f"Correo: {sup_email}" if sup_email else "",
                            ]
                            if line
                        )
                    ),
                    styles["small"],
                ),
                Paragraph(
                    _p_text(
                        "\n".join(
                            line
                            for line in [
                                cus_name or "-",
                                f"ID / NIT: {cus_id}" if cus_id else "",
                                f"Dirección: {cus_addr} {cus_city}".strip() if (cus_addr or cus_city) else "",
                                f"Correo: {cus_email}" if cus_email else "",
                            ]
                            if line
                        )
                    ),
                    styles["small"],
                ),
            ],
        ]
        party_tbl = Table(party_rows, colWidths=[usable_width / 2, usable_width / 2])
        party_tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E5E7EB")),
                    ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D1D5DB")),
                    ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#D1D5DB")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(party_tbl)
        story.append(Spacer(1, 10))

        # Líneas de detalle
        line_els = _find_all_elements(root, ("InvoiceLine", "CreditNoteLine", "DebitNoteLine"))
        if line_els:
            story.append(Paragraph(_p_text("Detalle de Ítems / Conceptos"), styles["h2"]))
            inv_hdr_style = ParagraphStyle("InvHdr", parent=styles["small_bold"], textColor=colors.white)
            item_rows = [
                [
                    Paragraph(_p_text("#"), inv_hdr_style),
                    Paragraph(_p_text("Descripción"), inv_hdr_style),
                    Paragraph(_p_text("Cant."), inv_hdr_style),
                    Paragraph(_p_text("Valor Unit."), inv_hdr_style),
                    Paragraph(_p_text("Total Línea"), inv_hdr_style),
                ]
            ]
            for idx, line_el in enumerate(line_els, start=1):
                line_num = _find_first_text(line_el, ("ID",)) or str(idx)
                item_node = _find_first_element(line_el, ("Item",))
                desc = _find_first_text(item_node if item_node is not None else line_el, ("Description", "Name"))
                qty = _find_first_text(line_el, ("InvoicedQuantity", "CreditedQuantity", "DebitedQuantity", "Quantity"))
                price_node = _find_first_element(line_el, ("Price",))
                unit_price = _find_first_text(price_node, ("PriceAmount",))
                line_total = _find_first_text(line_el, ("LineExtensionAmount",))
                item_rows.append(
                    [
                        Paragraph(_p_text(line_num), styles["small"]),
                        Paragraph(_p_text(desc or "-"), styles["small"]),
                        Paragraph(_p_text(qty or "-"), styles["small_right"]),
                        Paragraph(_p_text(unit_price or "-"), styles["small_right"]),
                        Paragraph(_p_text(line_total or "-"), styles["small_right"]),
                    ]
                )
            items_tbl = Table(
                item_rows,
                colWidths=[1.1 * cm, usable_width - 8.1 * cm, 1.8 * cm, 2.5 * cm, 2.7 * cm],
                repeatRows=1,
            )
            items_tbl.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F2937")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("TOPPADDING", (0, 0), (-1, -1), 3),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ]
                )
            )
            story.append(items_tbl)
            story.append(Spacer(1, 10))

        # Totales
        monetary = _find_first_element(root, ("LegalMonetaryTotal", "RequestedMonetaryTotal"))
        tax_total = _find_first_element(root, ("TaxTotal",))
        subtotal = _find_first_text(monetary, ("LineExtensionAmount",))
        tax_amt = _find_first_text(tax_total, ("TaxAmount",))
        payable = _find_first_text(monetary, ("PayableAmount",))
        if subtotal or tax_amt or payable:
            totals_rows = [
                [Paragraph(_p_text("Subtotal:"), styles["small_bold"]), Paragraph(_p_text(f"{currency} {subtotal}".strip()), styles["small_right"])],
                [Paragraph(_p_text("Impuestos:"), styles["small_bold"]), Paragraph(_p_text(f"{currency} {tax_amt}".strip()), styles["small_right"])],
                [Paragraph(_p_text("TOTAL A PAGAR:"), styles["small_bold"]), Paragraph(_p_text(f"{currency} {payable}".strip()), styles["small_right"])],
            ]
            tot_tbl = Table(totals_rows, colWidths=[3.5 * cm, 4.2 * cm], hAlign="RIGHT")
            tot_tbl.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 2), (-1, 2), colors.HexColor("#F3F4F6")),
                        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#9CA3AF")),
                        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#D1D5DB")),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ]
                )
            )
            story.append(tot_tbl)
            story.append(Spacer(1, 12))

    elif is_cfdi:
        story.append(Paragraph(_p_text("Comprobante Fiscal Digital (CFDI XML)"), styles["title"]))
        story.append(Paragraph(_p_text(f"Archivo fuente: {filename}"), styles["small"]))
        story.append(Spacer(1, 6))

        emisor = _find_first_element(root, ("Emisor",))
        receptor = _find_first_element(root, ("Receptor",))
        tfd = _find_first_element(root, ("TimbreFiscalDigital",))

        meta_rows = [
            ("Serie / Folio:", f"{root.attrib.get('Serie', '')} {root.attrib.get('Folio', '')}".strip() or "-"),
            ("Fecha:", root.attrib.get("Fecha", "-")),
            ("Moneda:", root.attrib.get("Moneda", "-")),
            ("UUID Fiscal:", tfd.attrib.get("UUID", "-") if tfd is not None else "-"),
            ("Emisor:", f"{emisor.attrib.get('Nombre', '')} (RFC: {emisor.attrib.get('Rfc', '')})" if emisor is not None else "-"),
            ("Receptor:", f"{receptor.attrib.get('Nombre', '')} (RFC: {receptor.attrib.get('Rfc', '')})" if receptor is not None else "-"),
            ("Subtotal:", root.attrib.get("SubTotal", "-")),
            ("Total:", root.attrib.get("Total", "-")),
        ]
        cfdi_tbl = Table(
            [[Paragraph(_p_text(k), styles["small_bold"]), Paragraph(_p_text(v), styles["small"])] for k, v in meta_rows],
            colWidths=[3.5 * cm, usable_width - 3.5 * cm],
        )
        cfdi_tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F9FAFB")),
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(cfdi_tbl)
        story.append(Spacer(1, 10))

        conceptos = _find_all_elements(root, ("Concepto",))
        if conceptos:
            story.append(Paragraph(_p_text("Conceptos"), styles["h2"]))
            c_rows = [
                [
                    Paragraph(_p_text("Descripción"), styles["small_bold"]),
                    Paragraph(_p_text("Cant."), styles["small_bold"]),
                    Paragraph(_p_text("V. Unitario"), styles["small_bold"]),
                    Paragraph(_p_text("Importe"), styles["small_bold"]),
                ]
            ]
            for c in conceptos:
                c_rows.append(
                    [
                        Paragraph(_p_text(c.attrib.get("Descripcion", "-")), styles["small"]),
                        Paragraph(_p_text(c.attrib.get("Cantidad", "-")), styles["small_right"]),
                        Paragraph(_p_text(c.attrib.get("ValorUnitario", "-")), styles["small_right"]),
                        Paragraph(_p_text(c.attrib.get("Importe", "-")), styles["small_right"]),
                    ]
                )
            c_tbl = Table(c_rows, colWidths=[usable_width - 7.0 * cm, 1.8 * cm, 2.6 * cm, 2.6 * cm], repeatRows=1)
            c_tbl.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E5E7EB")),
                        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
                        ("TOPPADDING", (0, 0), (-1, -1), 3),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ]
                )
            )
            story.append(c_tbl)
            story.append(Spacer(1, 10))
    else:
        story.append(Paragraph(_p_text(f"Documento XML: {filename}"), styles["title"]))
        story.append(Paragraph(_p_text(f"Elemento raíz: <{doc_type}>"), styles["small"]))
        story.append(Spacer(1, 8))

    # Tabla resumen de nodos XML completos (garantiza que cualquier campo adicional del XML sea legible)
    flat_nodes = _extract_xml_flat_nodes(root)
    if flat_nodes:
        story.append(Paragraph(_p_text("Estructura y Datos del XML"), styles["h2"]))
        xml_rows = [
            [Paragraph(_p_text("Nodo / Ruta"), styles["small_bold"]), Paragraph(_p_text("Valor"), styles["small_bold"])]
        ]
        for node_path, val in flat_nodes:
            xml_rows.append([Paragraph(_p_text(node_path), styles["small"]), Paragraph(_p_text(val), styles["small"])])

        xml_tbl = Table(xml_rows, colWidths=[usable_width * 0.45, usable_width * 0.55], repeatRows=1)
        xml_tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E5E7EB")),
                    ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#D1D5DB")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
                ]
            )
        )
        story.append(xml_tbl)

    doc.build(story)
    return output_pdf_path


def _convert_office_via_com(input_path: Path, ext: str, output_pdf_path: Path) -> bool:
    """Nivel 1: Convierte Office a PDF usando Microsoft Office COM en Windows."""
    if os.name != "nt":
        return False

    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError:
        return False

    pythoncom.CoInitialize()
    app = None
    doc_obj = None
    try:
        in_str = str(input_path.resolve())
        temp_out = input_path.parent / "com_output.pdf"
        out_str = str(temp_out.resolve())

        if ext in OFFICE_WORD_EXTENSIONS:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            doc_obj = app.Documents.Open(
                in_str,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                NoEncodingDialog=True,
            )
            doc_obj.SaveAs2(out_str, FileFormat=17)  # wdFormatPDF
        elif ext in OFFICE_EXCEL_EXTENSIONS:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            doc_obj = app.Workbooks.Open(in_str, ReadOnly=True, IgnoreReadOnlyRecommended=True)
            doc_obj.ExportAsFixedFormat(0, out_str)  # xlTypePDF
        elif ext in OFFICE_PPT_EXTENSIONS:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            doc_obj = app.Presentations.Open(in_str, ReadOnly=True, Untitled=False, WithWindow=False)
            doc_obj.SaveAs(out_str, 32)  # ppSaveAsPDF
        else:
            return False

        if temp_out.exists() and temp_out.stat().st_size > 0:
            shutil.copyfile(temp_out, output_pdf_path)
            return True
        return False
    except Exception:
        return False
    finally:
        if doc_obj is not None:
            try:
                if ext in OFFICE_WORD_EXTENSIONS:
                    doc_obj.Close(SaveChanges=False)
                elif ext in OFFICE_EXCEL_EXTENSIONS:
                    doc_obj.Close(SaveChanges=False)
                else:
                    doc_obj.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _find_libreoffice_binary() -> str | None:
    """Busca el ejecutable de LibreOffice (soffice) si está instalado."""
    in_path = shutil.which("soffice") or shutil.which("libreoffice")
    if in_path:
        return in_path

    candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]
    for cand in candidates:
        if cand.exists():
            return str(cand)
    return None


def _convert_office_via_libreoffice(input_path: Path, output_pdf_path: Path) -> bool:
    """Nivel 2: Convierte documentos Office a PDF mediante LibreOffice headless."""
    soffice = _find_libreoffice_binary()
    if not soffice:
        return False

    try:
        with tempfile.TemporaryDirectory() as out_dir:
            res = subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, str(input_path)],
                capture_output=True,
                text=True,
                timeout=60,
                check=False,
            )
            generated = Path(out_dir) / f"{input_path.stem}.pdf"
            if res.returncode == 0 and generated.exists() and generated.stat().st_size > 0:
                shutil.copyfile(generated, output_pdf_path)
                return True
    except Exception:
        return False
    return False


def _convert_docx_pure_python(input_path: Path, output_pdf_path: Path, filename: str) -> Path:
    """Nivel 3: Convierte .docx a PDF usando python-docx + ReportLab sin Office instalado."""
    styles = _build_styles()
    doc = _create_doc(output_pdf_path, is_landscape=False, margins_mm=15.0)
    usable_width = A4[0] - 30.0 * mm
    docx_obj = DocxDocument(str(input_path))
    story: list[Any] = []

    for child in docx_obj.element.body.iterchildren():
        if isinstance(child, CT_P):
            para = DocxParagraph(child, docx_obj)
            text = para.text.strip()
            if not text:
                continue
            style_name = (para.style.name or "").lower() if para.style else ""
            if "heading 1" in style_name or "title" in style_name or "título" in style_name:
                story.append(Paragraph(_p_text(text), styles["h1"]))
            elif "heading" in style_name or "subtítulo" in style_name:
                story.append(Paragraph(_p_text(text), styles["h2"]))
            else:
                story.append(Paragraph(_p_text(text), styles["body"]))
        elif isinstance(child, CT_Tbl):
            tbl = DocxTable(child, docx_obj)
            rows_data: list[list[Paragraph]] = []
            for r_idx, row in enumerate(tbl.rows):
                p_style = styles["small_bold"] if r_idx == 0 else styles["small"]
                rows_data.append([Paragraph(_p_text(cell.text.strip()), p_style) for cell in row.cells])
            if rows_data:
                col_count = max(len(r) for r in rows_data)
                if col_count > 0:
                    for r in rows_data:
                        while len(r) < col_count:
                            r.append(Paragraph("", styles["small"]))
                    rl_tbl = Table(rows_data, colWidths=[usable_width / col_count] * col_count)
                    rl_tbl.setStyle(
                        TableStyle(
                            [
                                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F3F4F6")),
                                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
                                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                ("TOPPADDING", (0, 0), (-1, -1), 3),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                            ]
                        )
                    )
                    story.append(rl_tbl)
                    story.append(Spacer(1, 6))

    if not story:
        story.append(Paragraph(_p_text(f"Documento Word: {filename} (sin texto extraíble)"), styles["quote"]))

    doc.build(story)
    return output_pdf_path


def _convert_xlsx_pure_python(input_path: Path, output_pdf_path: Path, filename: str) -> Path:
    """Nivel 3: Convierte .xlsx a PDF usando openpyxl + ReportLab."""
    wb = load_workbook(filename=str(input_path), data_only=True, read_only=True)
    sections: list[tuple[str, list[list[str]]]] = []
    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows: list[list[str]] = []
            for row in ws.iter_rows(values_only=True):
                str_row = ["" if cell is None else str(cell) for cell in row]
                if any(val.strip() for val in str_row):
                    # Recorta celdas vacías al final de cada fila
                    while str_row and not str_row[-1].strip():
                        str_row.pop()
                    rows.append(str_row)
            sections.append((f"Hoja: {sheet_name}", rows))
    finally:
        wb.close()

    if not sections:
        sections = [("", [])]
    return _render_tabular_rows_to_pdf(filename, sections, output_pdf_path)


def _convert_pptx_pure_python(input_path: Path, output_pdf_path: Path, filename: str) -> Path:
    """Nivel 3: Convierte .pptx a PDF usando python-pptx + ReportLab en formato horizontal."""
    styles = _build_styles()
    doc = _create_doc(output_pdf_path, is_landscape=True, margins_mm=15.0)
    usable_width = A4[1] - 30.0 * mm
    prs = Presentation(str(input_path))
    story: list[Any] = []

    for idx, slide in enumerate(prs.slides, start=1):
        if idx > 1:
            story.append(PageBreak())

        slide_title = ""
        if slide.shapes.title and slide.shapes.title.text:
            slide_title = slide.shapes.title.text.strip()

        header_text = f"Diapositiva {idx}" + (f": {slide_title}" if slide_title else "")
        story.append(Paragraph(_p_text(header_text), styles["h1"]))
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#D1D5DB")))
        story.append(Spacer(1, 6))

        content_added = False
        for shape in slide.shapes:
            if slide.shapes.title is not None and shape == slide.shapes.title:
                continue

            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    txt = p.text.strip()
                    if txt:
                        story.append(Paragraph(_p_text(txt), styles["body"]))
                        content_added = True

            if shape.has_table:
                tbl_rows: list[list[Paragraph]] = []
                for r_i, row in enumerate(shape.table.rows):
                    p_style = styles["small_bold"] if r_i == 0 else styles["small"]
                    tbl_rows.append([Paragraph(_p_text(c.text.strip()), p_style) for c in row.cells])
                if tbl_rows:
                    cols = len(tbl_rows[0])
                    rl_tbl = Table(tbl_rows, colWidths=[usable_width / max(cols, 1)] * max(cols, 1))
                    rl_tbl.setStyle(
                        TableStyle(
                            [
                                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F3F4F6")),
                                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
                                ("TOPPADDING", (0, 0), (-1, -1), 3),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                            ]
                        )
                    )
                    story.append(rl_tbl)
                    story.append(Spacer(1, 6))
                    content_added = True

            if shape.shape_type == 13:  # PICTURE
                try:
                    img_bytes = shape.image.blob
                    with Image.open(io.BytesIO(img_bytes)) as pil_img:
                        w, h = pil_img.size
                    max_w = min(usable_width * 0.7, 14 * cm)
                    max_h = 8 * cm
                    scale = min(max_w / max(w, 1), max_h / max(h, 1))
                    story.append(RLImage(io.BytesIO(img_bytes), width=w * scale, height=h * scale))
                    story.append(Spacer(1, 6))
                    content_added = True
                except Exception:
                    pass

        if not content_added and not slide_title:
            story.append(Paragraph(_p_text("(Diapositiva sin elementos textuales)"), styles["quote"]))

    if not story:
        story.append(Paragraph(_p_text(f"Presentación: {filename} (vacía)"), styles["quote"]))

    doc.build(story)
    return output_pdf_path


def _convert_office_to_pdf(
    payload: bytes,
    ext: str,
    filename: str,
    output_pdf_path: Path,
    use_office_com: bool = True,
) -> Path:
    """Conversión híbrida de 3 niveles para documentos Office: MS Office COM -> LibreOffice -> Python puro."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        safe_input_name = f"input{ext}"
        in_path = Path(tmp_dir) / safe_input_name
        in_path.write_bytes(payload)

        if use_office_com and _convert_office_via_com(in_path, ext, output_pdf_path):
            return output_pdf_path

        if _convert_office_via_libreoffice(in_path, output_pdf_path):
            return output_pdf_path

        if ext == ".docx":
            return _convert_docx_pure_python(in_path, output_pdf_path, filename)
        if ext == ".xlsx":
            return _convert_xlsx_pure_python(in_path, output_pdf_path, filename)
        if ext == ".pptx":
            return _convert_pptx_pure_python(in_path, output_pdf_path, filename)

        # Formatos binarios heredados (.doc, .xls, .ppt) cuando no hay MS Office ni LibreOffice
        return _render_text_lines_to_pdf(
            filename,
            (
                f"El archivo '{filename}' utiliza un formato binario antiguo de Microsoft Office ({ext}).\n"
                "No se detectó Microsoft Office ni LibreOffice en este equipo para convertirlo nativamente.\n\n"
                f"Tamaño del archivo: {len(payload):,} bytes\n"
                f"SHA-256: {hashlib.sha256(payload).hexdigest()}"
            ),
            output_pdf_path,
            mono=False,
        )


def _convert_fallback_to_pdf(payload: bytes, output_pdf_path: Path, filename: str, content_type: str = "") -> Path:
    """Conversor de respaldo: renderiza texto si es legible o ficha informativa si es binario."""
    # Heurística simple: si no tiene bytes nulos y decodifica bien con pocos caracteres de control
    if b"\x00" not in payload[:4096]:
        decoded = _decode_bytes(payload)
        printable_ratio = sum(1 for c in decoded[:2000] if c.isprintable() or c in "\n\r\t") / max(
            len(decoded[:2000]), 1
        )
        if printable_ratio > 0.85 and decoded.strip():
            return _render_text_lines_to_pdf(filename, decoded, output_pdf_path, mono=True)

    info = (
        f"Archivo adjunto no convertible directamente a contenido visual.\n\n"
        f"Nombre original: {filename}\n"
        f"Tipo MIME: {content_type or 'Desconocido'}\n"
        f"Tamaño: {len(payload):,} bytes\n"
        f"SHA-256: {hashlib.sha256(payload).hexdigest()}"
    )
    return _render_text_lines_to_pdf(filename, info, output_pdf_path, mono=False)


def _convert_zip_to_pdf(
    payload: bytes,
    dest_dir: Path,
    zip_filename: str,
    use_office_com: bool = True,
) -> list[Path]:
    """Descomprime un ZIP de forma segura y convierte cada archivo interno a PDF."""
    generated: list[Path] = []
    try:
        with zipfile.ZipFile(io.BytesIO(payload)) as zf:
            entries = [info for info in zf.infolist() if not info.is_dir()]
            if not entries:
                out_pdf = get_unique_path(dest_dir, Path(zip_filename).stem, is_folder=False, ext=".pdf", fallback="zip_vacio")
                generated.append(
                    _render_text_lines_to_pdf(zip_filename, "El archivo ZIP está vacío.", out_pdf, mono=False)
                )
                return generated

            for info in entries:
                # Ignora metadatos de macOS o rutas peligrosas
                parts = Path(info.filename).parts
                if any(p in ("__MACOSX", "..") for p in parts) or Path(info.filename).name == ".DS_Store":
                    continue
                inner_name = Path(info.filename).name
                if not inner_name:
                    continue
                inner_bytes = zf.read(info)
                generated.extend(
                    convert_attachment_to_pdf(
                        inner_name,
                        inner_bytes,
                        dest_dir,
                        content_type="",
                        use_office_com=use_office_com,
                    )
                )
    except Exception as exc:
        out_pdf = get_unique_path(dest_dir, Path(zip_filename).stem, is_folder=False, ext=".pdf", fallback="zip_error")
        generated.append(
            _render_text_lines_to_pdf(
                zip_filename,
                f"No se pudo descomprimir '{zip_filename}' (posiblemente protegido con contraseña o dañado).\nDetalle: {exc}",
                out_pdf,
                mono=False,
            )
        )
    return generated


def convert_attachment_to_pdf(
    filename: str,
    payload: bytes,
    dest_dir: Path,
    content_type: str = "",
    use_office_com: bool = True,
) -> list[Path]:
    """Convierte cualquier adjunto soportado a uno o más archivos PDF dentro de dest_dir."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    path_obj = Path(filename)
    ext = path_obj.suffix.lower()
    stem = path_obj.stem or "adjunto"

    if ext == ".zip":
        return _convert_zip_to_pdf(payload, dest_dir, filename, use_office_com=use_office_com)

    out_pdf = get_unique_path(dest_dir, stem, is_folder=False, ext=".pdf", fallback="adjunto")

    if ext == ".pdf":
        out_pdf.write_bytes(payload)
        return [out_pdf]

    if ext in IMAGE_EXTENSIONS:
        return [_convert_image_to_pdf(payload, out_pdf)]

    if ext in CSV_EXTENSIONS:
        delimiter = "\t" if ext == ".tsv" else None
        return [_convert_csv_to_pdf(payload, out_pdf, filename, delimiter_hint=delimiter)]

    if ext in HTML_EXTENSIONS:
        return [_convert_html_file_to_pdf(payload, out_pdf, filename)]

    if ext == ".xml":
        return [_convert_xml_to_pdf(payload, out_pdf, filename)]

    if ext in TEXT_EXTENSIONS:
        return [_convert_text_to_pdf(payload, out_pdf, filename, is_json=(ext == ".json"))]

    if ext in OFFICE_EXTENSIONS:
        return [_convert_office_to_pdf(payload, ext, filename, out_pdf, use_office_com=use_office_com)]

    return [_convert_fallback_to_pdf(payload, out_pdf, filename, content_type=content_type)]
