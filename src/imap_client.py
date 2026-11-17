"""Cliente IMAP multi-plataforma para descargar correos y adjuntos por rango de fechas."""

from __future__ import annotations

import base64
import email
import imaplib
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from email.message import Message
from email.utils import parsedate_to_datetime
from typing import Callable

from src.utils import decode_mime_text

IMAP_SERVERS: dict[str, str] = {
    "gmail.com": "imap.gmail.com",
    "googlemail.com": "imap.gmail.com",
    "outlook.com": "outlook.office365.com",
    "hotmail.com": "outlook.office365.com",
    "live.com": "outlook.office365.com",
    "msn.com": "outlook.office365.com",
    "office365.com": "outlook.office365.com",
    "yahoo.com": "imap.mail.yahoo.com",
    "yahoo.es": "imap.mail.yahoo.com",
    "icloud.com": "imap.mail.me.com",
    "me.com": "imap.mail.me.com",
    "zoho.com": "imap.zoho.com",
    "zohomail.com": "imap.zoho.com",
}

IMAP_MONTHS = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
]

LIST_RESPONSE_RE = re.compile(r'\((?P<flags>[^)]*)\)\s+"(?P<delimiter>[^"]*)"\s+(?P<name>.+)')


@dataclass
class EmailAttachment:
    """Representa un archivo adjunto extraído de un correo."""

    filename: str
    content_type: str
    payload: bytes


@dataclass
class EmailMessageData:
    """Representa un correo descargado y listo para exportar a PDF."""

    uid: str
    subject: str
    sender: str
    recipients: str
    cc: str
    date_str: str
    parsed_date: datetime | None
    body_text: str
    body_html: str
    attachments: list[EmailAttachment] = field(default_factory=list)


def resolve_imap_server(email_address: str, custom_server: str | None = None) -> str:
    """Resuelve el servidor IMAP a partir del dominio del correo o servidor personalizado."""
    if custom_server and custom_server.strip():
        return custom_server.strip()

    if "@" not in email_address:
        raise ValueError(f"Dirección de correo inválida: '{email_address}'")

    domain = email_address.strip().rsplit("@", 1)[-1].lower()
    if domain in IMAP_SERVERS:
        return IMAP_SERVERS[domain]

    return f"imap.{domain}"


def format_imap_date(d: date) -> str:
    """Formatea una fecha en formato IMAP RFC3501 (dd-Mon-YYYY en inglés)."""
    return f"{d.day:02d}-{IMAP_MONTHS[d.month - 1]}-{d.year}"


def _decode_imap_utf7(folder_name: str) -> str:
    """Decodifica nombres de carpeta IMAP en UTF-7 modificado (RFC 3501)."""
    result: list[str] = []
    i = 0
    while i < len(folder_name):
        ch = folder_name[i]
        if ch == "&":
            end = folder_name.find("-", i + 1)
            if end == -1:
                result.append(ch)
                i += 1
                continue
            encoded = folder_name[i + 1 : end]
            if not encoded:
                result.append("&")
            else:
                b64 = encoded.replace(",", "/")
                try:
                    pad_len = (-len(b64)) % 4
                    raw = base64.b64decode((b64 + "=" * pad_len).encode("ascii"))
                    result.append(raw.decode("utf-16-be", errors="replace"))
                except Exception:
                    result.append(folder_name[i : end + 1])
            i = end + 1
        else:
            result.append(ch)
            i += 1
    return "".join(result)


def _decode_part_payload(part: Message) -> str:
    """Decodifica el payload textual de una parte MIME probando charsets comunes."""
    payload = part.get_payload(decode=True)
    if payload is None:
        raw = part.get_payload()
        return str(raw) if raw else ""

    charsets = []
    declared = part.get_content_charset()
    if declared:
        charsets.append(declared)
    charsets.extend(["utf-8", "latin-1", "cp1252"])

    for cs in charsets:
        try:
            return payload.decode(cs)
        except (UnicodeDecodeError, LookupError):
            continue

    return payload.decode("utf-8", errors="replace")


def _parse_email_datetime(date_header: str) -> datetime | None:
    """Convierte el encabezado Date del correo a datetime en zona horaria local."""
    if not date_header:
        return None
    try:
        dt = parsedate_to_datetime(date_header)
        if dt.tzinfo is not None:
            return dt.astimezone()
        return dt
    except Exception:
        return None


def parse_raw_email(uid: str, raw_bytes: bytes) -> EmailMessageData:
    """Parsea un mensaje RFC822 en bytes a EmailMessageData."""
    msg = email.message_from_bytes(raw_bytes)

    subject = decode_mime_text(msg.get("Subject")) or "Sin_Asunto"
    sender = decode_mime_text(msg.get("From"))
    recipients = decode_mime_text(msg.get("To"))
    cc = decode_mime_text(msg.get("Cc"))
    date_header = decode_mime_text(msg.get("Date"))
    parsed_dt = _parse_email_datetime(date_header)

    if parsed_dt is not None:
        date_str = parsed_dt.strftime("%Y-%m-%d %H:%M:%S")
    else:
        date_str = date_header or "Fecha desconocida"

    text_parts: list[str] = []
    html_parts: list[str] = []
    attachments: list[EmailAttachment] = []

    if msg.is_multipart():
        for part in msg.walk():
            content_maintype = part.get_content_maintype()
            if content_maintype == "multipart":
                continue

            content_type = (part.get_content_type() or "application/octet-stream").lower()
            disposition = (part.get("Content-Disposition") or "").lower()
            raw_filename = part.get_filename()
            filename = decode_mime_text(raw_filename) if raw_filename else ""

            is_attachment = (
                "attachment" in disposition
                or bool(filename)
                or (content_maintype not in ("text",) and "inline" not in disposition)
            )

            # Evita confundir el cuerpo text/plain o text/html sin filename con un adjunto
            if content_type in ("text/plain", "text/html") and "attachment" not in disposition and not filename:
                is_attachment = False

            if is_attachment:
                payload = part.get_payload(decode=True)
                if payload is None:
                    raw_payload = part.get_payload()
                    if isinstance(raw_payload, list):
                        # p. ej. message/rfc822 anidado
                        payload = b"\n".join(item.as_bytes() for item in raw_payload if hasattr(item, "as_bytes"))
                    elif isinstance(raw_payload, str):
                        payload = raw_payload.encode("utf-8", errors="replace")
                    else:
                        payload = b""

                if not filename:
                    ext = ".bin"
                    if "/" in content_type:
                        sub = content_type.split("/", 1)[1].split(";")[0].strip()
                        if sub and len(sub) <= 10:
                            ext = f".{sub}"
                    filename = f"adjunto_{len(attachments) + 1}{ext}"

                attachments.append(
                    EmailAttachment(
                        filename=filename,
                        content_type=content_type,
                        payload=payload,
                    )
                )
            elif content_type == "text/plain":
                text_parts.append(_decode_part_payload(part))
            elif content_type == "text/html":
                html_parts.append(_decode_part_payload(part))
    else:
        content_type = (msg.get_content_type() or "text/plain").lower()
        if content_type == "text/html":
            html_parts.append(_decode_part_payload(msg))
        else:
            text_parts.append(_decode_part_payload(msg))

    return EmailMessageData(
        uid=uid,
        subject=subject,
        sender=sender,
        recipients=recipients,
        cc=cc,
        date_str=date_str,
        parsed_date=parsed_dt,
        body_text="\n\n".join(p for p in text_parts if p.strip()),
        body_html="\n".join(p for p in html_parts if p.strip()),
        attachments=attachments,
    )


class EmailFetcher:
    """Cliente IMAP SSL para buscar y descargar correos en un rango de fechas."""

    def __init__(
        self,
        email_address: str,
        password: str,
        imap_server: str | None = None,
        imap_port: int = 993,
    ) -> None:
        self.email_address = email_address.strip()
        self.imap_server = resolve_imap_server(self.email_address, imap_server)
        clean_pw = password.strip() if self.imap_server.lower() == "imap.gmail.com" else password
        # Las contraseñas de aplicación de Google vienen como 4 bloques de 4 letras separados por espacios
        if self.imap_server.lower() == "imap.gmail.com":
            no_spaces = re.sub(r"\s+", "", clean_pw)
            if len(no_spaces) == 16 and no_spaces.isalpha():
                clean_pw = no_spaces
        self.password = clean_pw
        self.imap_port = imap_port
        self.conn: imaplib.IMAP4_SSL | None = None

    def __enter__(self) -> EmailFetcher:
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.disconnect()

    def connect(self) -> None:
        """Abre conexión IMAP SSL y autentica al usuario."""
        self.conn = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
        status, data = self.conn.login(self.email_address, self.password)
        if status != "OK":
            raise ConnectionError(f"Error al autenticar en {self.imap_server}: {data}")

    def disconnect(self) -> None:
        """Cierra la conexión IMAP de forma segura."""
        if self.conn is not None:
            try:
                self.conn.close()
            except Exception:
                pass
            try:
                self.conn.logout()
            except Exception:
                pass
            self.conn = None

    def _require_conn(self) -> imaplib.IMAP4_SSL:
        if self.conn is None:
            raise RuntimeError("No hay una conexión IMAP activa. Llama a connect() primero.")
        return self.conn

    def _list_folders_detailed(self) -> list[tuple[str, str, str]]:
        """Retorna una lista de tuplas (flags, nombre_imap_raw, nombre_legible) de carpetas."""
        conn = self._require_conn()
        status, items = conn.list()
        if status != "OK" or not items:
            return []

        folders: list[tuple[str, str, str]] = []
        for item in items:
            if not item:
                continue
            line = item.decode("utf-8", errors="replace") if isinstance(item, bytes) else str(item)
            match = LIST_RESPONSE_RE.match(line)
            if match:
                flags = match.group("flags") or ""
                raw_name = match.group("name").strip()
                unquoted = raw_name[1:-1] if raw_name.startswith('"') and raw_name.endswith('"') else raw_name
                readable = _decode_imap_utf7(unquoted)
                folders.append((flags, unquoted, readable))
        return folders

    def list_folders(self) -> list[tuple[str, str]]:
        """Retorna una lista de tuplas (nombre_imap_raw, nombre_legible) de carpetas."""
        return [(raw_name, readable) for _, raw_name, readable in self._list_folders_detailed()]

    def _try_select(self, folder_name: str) -> tuple[str, list]:
        conn = self._require_conn()
        quoted_folder = (
            f'"{folder_name}"' if not (folder_name.startswith('"') and folder_name.endswith('"')) else folder_name
        )
        status, data = conn.select(quoted_folder, readonly=True)
        if status != "OK":
            status, data = conn.select(folder_name, readonly=True)
        return status, data

    def select_folder(self, folder: str = "INBOX") -> int:
        """Selecciona una carpeta IMAP en modo solo lectura y retorna cantidad de mensajes."""
        status, data = self._try_select(folder)
        if status != "OK":
            target_lower = folder.strip().lower()
            all_aliases = {
                "todos",
                "all",
                "all mail",
                "[gmail]/todos",
                "[gmail]/all mail",
                "[gmail]/todos los mensajes",
            }
            detailed = self._list_folders_detailed()
            candidate: str | None = None

            if target_lower in all_aliases:
                for flags, raw_name, _ in detailed:
                    if "\\all" in flags.lower():
                        candidate = raw_name
                        break

            if candidate is None:
                for _, raw_name, readable in detailed:
                    if raw_name.lower() == target_lower or readable.lower() == target_lower:
                        candidate = raw_name
                        break

            if candidate is None:
                for _, raw_name, readable in detailed:
                    if readable.lower().endswith(f"/{target_lower}") or raw_name.lower().endswith(f"/{target_lower}"):
                        candidate = raw_name
                        break

            if candidate is not None:
                status, data = self._try_select(candidate)

        if status != "OK":
            raise ValueError(f"No se pudo abrir la carpeta IMAP '{folder}': {data}")

        try:
            return int(data[0])
        except (TypeError, ValueError, IndexError):
            return 0

    def search_uids_in_range(
        self,
        start_date: date,
        end_date: date,
        folder: str = "INBOX",
    ) -> list[bytes]:
        """Busca UIDs de correos entre start_date y end_date (ambos inclusive)."""
        if start_date > end_date:
            raise ValueError("La fecha inicial no puede ser posterior a la fecha final.")

        conn = self._require_conn()
        self.select_folder(folder)

        since_str = format_imap_date(start_date)
        # BEFORE en IMAP es exclusivo y evalúa solo fecha; +1 día incluye todo end_date
        before_str = format_imap_date(end_date + timedelta(days=1))

        status, data = conn.uid("SEARCH", None, "SINCE", since_str, "BEFORE", before_str)
        if status != "OK":
            raise RuntimeError(f"Error ejecutando búsqueda IMAP: {data}")

        if not data or not data[0]:
            return []

        return data[0].split()

    def fetch_email_by_uid(self, uid: bytes) -> EmailMessageData | None:
        """Descarga y parsea un correo dado su UID."""
        conn = self._require_conn()
        status, msg_data = conn.uid("FETCH", uid, "(RFC822)")
        if status != "OK" or not msg_data:
            return None

        raw_bytes: bytes | None = None
        for item in msg_data:
            if isinstance(item, tuple) and len(item) >= 2 and isinstance(item[1], (bytes, bytearray)):
                raw_bytes = bytes(item[1])
                break

        if raw_bytes is None:
            return None

        uid_str = uid.decode("ascii", errors="replace") if isinstance(uid, bytes) else str(uid)
        return parse_raw_email(uid_str, raw_bytes)

    def fetch_emails_in_range(
        self,
        start_date: date,
        end_date: date,
        folder: str = "INBOX",
        progress_callback: Callable[[int, int, EmailMessageData | None], None] | None = None,
    ) -> list[EmailMessageData]:
        """Descarga todos los correos cuyo timestamp esté dentro de [start_date, end_date]."""
        uids = self.search_uids_in_range(start_date, end_date, folder=folder)
        total = len(uids)
        results: list[EmailMessageData] = []

        for idx, uid in enumerate(uids, start=1):
            email_obj = self.fetch_email_by_uid(uid)
            if email_obj is not None:
                # Filtro exacto por fecha del encabezado Date en zona local (evita desfases UTC del servidor)
                if email_obj.parsed_date is not None:
                    msg_day = email_obj.parsed_date.date()
                    if start_date <= msg_day <= end_date:
                        results.append(email_obj)
                    else:
                        email_obj = None
                else:
                    results.append(email_obj)

            if progress_callback is not None:
                progress_callback(idx, total, email_obj)

        return results
