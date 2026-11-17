"""Pruebas automatizadas de extremo a extremo para el pipeline de descarga/conversión a PDF."""

from __future__ import annotations

import io
import os
import sys
import tempfile
import unittest
import zipfile
from datetime import date, datetime
from email.message import EmailMessage as StdEmailMessage
from pathlib import Path

import pymupdf as fitz
from docx import Document as DocxDocument
from openpyxl import Workbook
from PIL import Image
from pptx import Presentation

# Asegura importación desde la raíz del proyecto
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from main import process_email_batch  # noqa: E402
from src.imap_client import (  # noqa: E402
    EmailAttachment,
    EmailMessageData,
    format_imap_date,
    parse_raw_email,
    resolve_imap_server,
)
from src.utils import decode_mime_text, get_unique_path, sanitize_filename  # noqa: E402


def _make_sample_pdf_bytes(text: str = "PDF Nativo de Prueba") -> bytes:
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), text)
    data = doc.tobytes()
    doc.close()
    return data


def _make_sample_png_bytes() -> bytes:
    img = Image.new("RGB", (320, 180), color=(30, 120, 200))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _make_sample_docx_bytes() -> bytes:
    doc = DocxDocument()
    doc.add_heading("Informe Mensual", level=1)
    doc.add_paragraph("Este es un documento Word generado para pruebas portables.")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Concepto"
    table.cell(0, 1).text = "Valor"
    table.cell(1, 0).text = "Servicio Cloud"
    table.cell(1, 1).text = "$ 150.000"
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _make_sample_xlsx_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Ventas"
    ws.append(["Fecha", "Cliente", "Total"])
    ws.append(["2026-11-10", "ACME S.A.S", 1250000])
    ws.append(["2026-11-12", "Tech Corp", 890000])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _make_sample_pptx_bytes() -> bytes:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Presentación Trimestral"
    slide.placeholders[1].text = "Resumen de resultados operativos\nMeta cumplida al 100%"
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _make_sample_ubl_attached_xml_bytes() -> bytes:
    invoice_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
         xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
         xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">
  <cbc:ID>FE-98765</cbc:ID>
  <cbc:UUID>cufe-abc1234567890-test</cbc:UUID>
  <cbc:IssueDate>2026-11-15</cbc:IssueDate>
  <cbc:IssueTime>14:30:00-05:00</cbc:IssueTime>
  <cbc:DocumentCurrencyCode>COP</cbc:DocumentCurrencyCode>
  <cac:AccountingSupplierParty>
    <cac:Party>
      <cac:PartyTaxScheme>
        <cbc:RegistrationName>Proveedor Ejemplo S.A.S</cbc:RegistrationName>
        <cbc:CompanyID>900123456</cbc:CompanyID>
      </cac:PartyTaxScheme>
    </cac:Party>
  </cac:AccountingSupplierParty>
  <cac:AccountingCustomerParty>
    <cac:Party>
      <cac:PartyTaxScheme>
        <cbc:RegistrationName>Cliente Final Ltda</cbc:RegistrationName>
        <cbc:CompanyID>800987654</cbc:CompanyID>
      </cac:PartyTaxScheme>
    </cac:Party>
  </cac:AccountingCustomerParty>
  <cac:TaxTotal>
    <cbc:TaxAmount currencyID="COP">19000.00</cbc:TaxAmount>
  </cac:TaxTotal>
  <cac:LegalMonetaryTotal>
    <cbc:LineExtensionAmount currencyID="COP">100000.00</cbc:LineExtensionAmount>
    <cbc:PayableAmount currencyID="COP">119000.00</cbc:PayableAmount>
  </cac:LegalMonetaryTotal>
  <cac:InvoiceLine>
    <cbc:ID>1</cbc:ID>
    <cbc:InvoicedQuantity unitCode="94">2</cbc:InvoicedQuantity>
    <cbc:LineExtensionAmount currencyID="COP">100000.00</cbc:LineExtensionAmount>
    <cac:Item>
      <cbc:Description>Licencia Anual Software Checker</cbc:Description>
    </cac:Item>
    <cac:Price>
      <cbc:PriceAmount currencyID="COP">50000.00</cbc:PriceAmount>
    </cac:Price>
  </cac:InvoiceLine>
</Invoice>"""
    attached_doc = f"""<?xml version="1.0" encoding="UTF-8"?>
<AttachedDocument xmlns="urn:oasis:names:specification:ubl:schema:xsd:AttachedDocument-2"
                  xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
                  xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">
  <cbc:ID>AD-1</cbc:ID>
  <cac:Attachment>
    <cac:ExternalReference>
      <cbc:Description><![CDATA[{invoice_xml}]]></cbc:Description>
    </cac:ExternalReference>
  </cac:Attachment>
</AttachedDocument>"""
    return attached_doc.encode("utf-8")


def _make_sample_zip_bytes() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("nota_interna.txt", "Contenido dentro de archivo ZIP\nSegunda línea.")
        zf.writestr("sello_interno.png", _make_sample_png_bytes())
    return buf.getvalue()


class PipelineTests(unittest.TestCase):
    def test_utils_and_imap_resolution(self) -> None:
        self.assertEqual(sanitize_filename('Factura: #102/2026? "Urgente"*'), "Factura_ #102_2026_ _Urgente__")
        self.assertEqual(sanitize_filename("CON"), "CON_")
        self.assertEqual(sanitize_filename("   "), "Sin_Asunto")
        self.assertEqual(decode_mime_text("=?utf-8?B?RmFjdHVyYSBFbGVjdHLDs25pY2E=?="), "Factura Electrónica")
        self.assertEqual(resolve_imap_server("usuario@gmail.com"), "imap.gmail.com")
        self.assertEqual(resolve_imap_server("usuario@outlook.com"), "outlook.office365.com")
        self.assertEqual(resolve_imap_server("persona@miempresa.com"), "imap.miempresa.com")
        self.assertEqual(resolve_imap_server("persona@miempresa.com", "mail.miempresa.com"), "mail.miempresa.com")
        self.assertEqual(format_imap_date(date(2026, 11, 5)), "05-Nov-2026")

    def test_parse_raw_email_mime(self) -> None:
        msg = StdEmailMessage()
        msg["Subject"] = "Prueba MIME Español: Árbol"
        msg["From"] = "Remitente <remitente@gmail.com>"
        msg["To"] = "destinatario@outlook.com"
        msg["Date"] = "Sun, 15 Nov 2026 10:15:00 -0500"
        msg.set_content("Texto plano del mensaje.")
        msg.add_alternative("<html><body><h1>Hola</h1><p>Correo en HTML.</p></body></html>", subtype="html")
        msg.add_attachment(b"contenido txt", maintype="text", subtype="plain", filename="resumen.txt")

        parsed = parse_raw_email("101", msg.as_bytes())
        self.assertEqual(parsed.subject, "Prueba MIME Español: Árbol")
        self.assertIn("Hola", parsed.body_html)
        self.assertIn("Texto plano", parsed.body_text)
        self.assertEqual(len(parsed.attachments), 1)
        self.assertEqual(parsed.attachments[0].filename, "resumen.txt")

    def test_end_to_end_export_and_pdf_validation(self) -> None:
        start_d = date(2026, 11, 1)
        end_d = date(2026, 11, 17)

        email_1 = EmailMessageData(
            uid="1",
            subject="Factura Electrónica: #102/2026?",
            sender="facturacion@proveedor.com",
            recipients="cliente@gmail.com",
            cc="contabilidad@gmail.com",
            date_str="2026-11-15 14:35:00",
            parsed_date=datetime(2026, 11, 15, 14, 35, 0),
            body_text="Adjuntamos la factura electrónica y los soportes del mes.",
            body_html=(
                "<html><body><h2>Estimado cliente</h2>"
                "<p>Adjuntamos su <b>factura electrónica</b> y reportes.</p>"
                "<ul><li>Factura XML + PDF</li><li>Detalle en Excel/CSV</li></ul>"
                "</body></html>"
            ),
            attachments=[
                EmailAttachment("original.pdf", "application/pdf", _make_sample_pdf_bytes()),
                EmailAttachment("comprobante.png", "image/png", _make_sample_png_bytes()),
                EmailAttachment("notas.txt", "text/plain", "Línea 1: Prueba\nLínea 2: Éxito".encode("utf-8")),
                EmailAttachment("tabla.csv", "text/csv", "Producto,Cantidad,Total\nServicio,2,100000\n".encode("utf-8")),
                EmailAttachment("resumen.html", "text/html", b"<html><body><h3>Resumen</h3><p>OK</p></body></html>"),
                EmailAttachment("factura_dian.xml", "application/xml", _make_sample_ubl_attached_xml_bytes()),
                EmailAttachment("informe.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", _make_sample_docx_bytes()),
                EmailAttachment("ventas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", _make_sample_xlsx_bytes()),
                EmailAttachment("presentacion.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", _make_sample_pptx_bytes()),
                EmailAttachment("paquete_soportes.zip", "application/zip", _make_sample_zip_bytes()),
            ],
        )

        # Segundo correo con exactamente el mismo asunto para verificar desambiguación de carpetas
        email_2 = EmailMessageData(
            uid="2",
            subject="Factura Electrónica: #102/2026?",
            sender="otro@proveedor.com",
            recipients="cliente@gmail.com",
            cc="",
            date_str="2026-11-16 09:00:00",
            parsed_date=datetime(2026, 11, 16, 9, 0, 0),
            body_text="Segundo correo con asunto idéntico.",
            body_html="",
            attachments=[],
        )

        # Tercer correo fuera del rango de fechas (debe omitirse)
        email_out = EmailMessageData(
            uid="3",
            subject="Correo Fuera de Rango",
            sender="spam@ejemplo.com",
            recipients="cliente@gmail.com",
            cc="",
            date_str="2026-10-30 23:00:00",
            parsed_date=datetime(2026, 10, 30, 23, 0, 0),
            body_text="No debería exportarse.",
            body_html="",
            attachments=[],
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            base_out = Path(tmp_dir)
            summary = process_email_batch(
                [email_1, email_2, email_out],
                start_date=start_d,
                end_date=end_d,
                output_base_dir=base_out,
                use_office_com=False,
            )

            self.assertEqual(summary.errors, [])
            self.assertEqual(summary.emails_processed, 2)
            self.assertEqual(summary.emails_skipped_out_of_range, 1)
            # email_1: 1 cuerpo + 9 archivos individuales + 2 dentro del ZIP = 12 PDFs
            # email_2: 1 cuerpo = 1 PDF -> Total = 13 PDFs
            self.assertEqual(summary.pdfs_generated, 13)

            expected_range_dir = base_out / "2026-11-01_a_2026-11-17"
            self.assertTrue(expected_range_dir.is_dir())

            mail_dir_1 = expected_range_dir / "Factura Electrónica_ #102_2026_"
            mail_dir_2 = expected_range_dir / "Factura Electrónica_ #102_2026_ (2)"
            self.assertTrue(mail_dir_1.is_dir())
            self.assertTrue(mail_dir_2.is_dir())

            expected_pdfs_1 = {
                "Factura Electrónica_ #102_2026_.pdf",
                "original.pdf",
                "comprobante.pdf",
                "notas.pdf",
                "tabla.pdf",
                "resumen.pdf",
                "factura_dian.pdf",
                "informe.pdf",
                "ventas.pdf",
                "presentacion.pdf",
                "nota_interna.pdf",
                "sello_interno.pdf",
            }
            actual_pdfs_1 = {p.name for p in mail_dir_1.iterdir() if p.is_file()}
            self.assertEqual(actual_pdfs_1, expected_pdfs_1)

            # Verifica que TODOS los PDFs generados sean válidos con PyMuPDF y tengan páginas
            all_pdfs = list(expected_range_dir.rglob("*.pdf"))
            self.assertEqual(len(all_pdfs), 13)
            for pdf_path in all_pdfs:
                with fitz.open(str(pdf_path)) as doc:
                    self.assertGreaterEqual(doc.page_count, 1, f"PDF sin páginas: {pdf_path}")

            # Verifica contenido clave extraído en cuerpo, XML factura, docx y xlsx
            with fitz.open(str(mail_dir_1 / "Factura Electrónica_ #102_2026_.pdf")) as doc:
                body_pdf_text = "\n".join(page.get_text() for page in doc)
                self.assertIn("facturacion@proveedor.com", body_pdf_text)
                self.assertIn("factura electrónica", body_pdf_text)

            with fitz.open(str(mail_dir_1 / "factura_dian.pdf")) as doc:
                xml_pdf_text = "\n".join(page.get_text() for page in doc)
                self.assertIn("FE-98765", xml_pdf_text)
                self.assertIn("Proveedor Ejemplo S.A.S", xml_pdf_text)
                self.assertIn("Licencia Anual Software Checker", xml_pdf_text)
                self.assertIn("119000.00", xml_pdf_text)

            with fitz.open(str(mail_dir_1 / "informe.pdf")) as doc:
                docx_pdf_text = "\n".join(page.get_text() for page in doc)
                self.assertIn("Informe Mensual", docx_pdf_text)
                self.assertIn("Servicio Cloud", docx_pdf_text)

            with fitz.open(str(mail_dir_1 / "ventas.pdf")) as doc:
                xlsx_pdf_text = "\n".join(page.get_text() for page in doc)
                self.assertIn("ACME S.A.S", xlsx_pdf_text)

            with fitz.open(str(mail_dir_1 / "presentacion.pdf")) as doc:
                pptx_pdf_text = "\n".join(page.get_text() for page in doc)
                self.assertIn("Presentación Trimestral", pptx_pdf_text)


if __name__ == "__main__":
    unittest.main(verbosity=2)
