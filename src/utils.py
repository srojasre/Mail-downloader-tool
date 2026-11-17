"""Utilidades de sanitización de rutas, nombres de archivo y decodificación MIME."""

from __future__ import annotations

import re
from datetime import date
from email.header import decode_header
from pathlib import Path

WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

INVALID_FILENAME_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
WHITESPACE_RE = re.compile(r"\s+")


def decode_mime_text(value: str | bytes | None) -> str:
    """Decodifica encabezados MIME (Subject, From, filename, etc.) a texto Unicode."""
    if value is None:
        return ""
    if isinstance(value, bytes):
        for enc in ("utf-8", "latin-1"):
            try:
                return value.decode(enc)
            except UnicodeDecodeError:
                continue
        return value.decode("utf-8", errors="replace")

    raw_text = str(value)
    if not raw_text.strip():
        return ""

    parts: list[str] = []
    try:
        decoded_chunks = decode_header(raw_text)
    except Exception:
        return raw_text.strip()

    for chunk, charset in decoded_chunks:
        if isinstance(chunk, bytes):
            encoding = charset or "utf-8"
            try:
                parts.append(chunk.decode(encoding, errors="replace"))
            except LookupError:
                parts.append(chunk.decode("utf-8", errors="replace"))
        else:
            parts.append(chunk)

    return "".join(parts).strip()


def sanitize_filename(name: str | None, max_length: int = 80, fallback: str = "Sin_Asunto") -> str:
    """Limpia una cadena para que sea válida como nombre de carpeta o archivo en Windows."""
    text = decode_mime_text(name)
    text = INVALID_FILENAME_CHARS_RE.sub("_", text)
    text = WHITESPACE_RE.sub(" ", text).strip()
    text = text.strip(". ")

    if not text:
        text = fallback

    if len(text) > max_length:
        text = text[:max_length].rstrip(". ")

    if not text:
        text = fallback

    stem_upper = text.split(".")[0].upper()
    if stem_upper in WINDOWS_RESERVED_NAMES:
        text = f"{text}_"

    return text


def get_unique_path(
    base_dir: Path,
    name: str,
    is_folder: bool = True,
    ext: str = "",
    max_length: int = 80,
    fallback: str = "Sin_Asunto",
) -> Path:
    """Genera una ruta única dentro de base_dir evitando colisiones con ' (2)', ' (3)', etc."""
    clean_ext = ext
    if clean_ext and not clean_ext.startswith("."):
        clean_ext = f".{clean_ext}"

    # Protege contra MAX_PATH en Windows manteniendo margen suficiente para subarchivos
    available = max(20, min(max_length, 220 - len(str(base_dir)) - len(clean_ext) - 10))
    clean_name = sanitize_filename(name, max_length=available, fallback=fallback)

    def _build_candidate(suffix: str = "") -> Path:
        if is_folder:
            return base_dir / f"{clean_name}{suffix}"
        return base_dir / f"{clean_name}{suffix}{clean_ext}"

    candidate = _build_candidate()
    if not candidate.exists():
        return candidate

    counter = 2
    while True:
        candidate = _build_candidate(f" ({counter})")
        if not candidate.exists():
            return candidate
        counter += 1


def format_date_range_folder(start_date: date, end_date: date) -> str:
    """Genera el nombre estándar de carpeta para el rango de fechas."""
    return f"{start_date.strftime('%Y-%m-%d')}_a_{end_date.strftime('%Y-%m-%d')}"
