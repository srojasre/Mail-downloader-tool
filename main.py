"""CLI y asistente interactivo para descargar correos vía IMAP y exportarlos a PDF."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from dotenv import load_dotenv, set_key
from rich.console import Console
from rich.panel import Panel
from rich.progress import BarColumn, MofNCompleteColumn, Progress, SpinnerColumn, TextColumn, TimeElapsedColumn
from rich.prompt import Confirm, Prompt
from rich.table import Table

from src.imap_client import EmailFetcher, EmailMessageData, resolve_imap_server
from src.pdf_converters import convert_attachment_to_pdf, render_email_body_to_pdf
from src.utils import format_date_range_folder, get_unique_path

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

console = Console()


@dataclass
class ExportSummary:
    """Resumen de resultados de una exportación de correos a PDF."""

    range_dir: Path
    emails_processed: int = 0
    emails_skipped_out_of_range: int = 0
    pdfs_generated: int = 0
    attachments_converted: int = 0
    errors: list[str] = field(default_factory=list)


def _get_app_base_dir() -> Path:
    """Retorna el directorio del ejecutable si está congelado o el directorio actual."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()


def _resolve_env_path() -> Path:
    """Ubica el archivo .env en el cwd o junto al script/ejecutable."""
    cwd_env = Path.cwd() / ".env"
    if cwd_env.exists():
        return cwd_env
    base_env = _get_app_base_dir() / ".env"
    if base_env.exists():
        return base_env
    return cwd_env


def parse_iso_date(value: str) -> date:
    """Parsea una fecha en formato YYYY-MM-DD."""
    return datetime.strptime(value.strip(), "%Y-%m-%d").date()


def export_single_email(
    email_obj: EmailMessageData,
    range_dir: Path,
    use_office_com: bool = True,
) -> tuple[Path, int, list[str]]:
    """Crea la carpeta del correo, genera <Asunto>.pdf y convierte todos sus adjuntos a PDF."""
    errors: list[str] = []
    mail_dir = get_unique_path(range_dir, email_obj.subject, is_folder=True, fallback="Sin_Asunto")
    mail_dir.mkdir(parents=True, exist_ok=True)

    pdfs_count = 0
    body_pdf_path = get_unique_path(mail_dir, email_obj.subject, is_folder=False, ext=".pdf", fallback="Sin_Asunto")
    try:
        render_email_body_to_pdf(email_obj, body_pdf_path)
        pdfs_count += 1
    except Exception as exc:
        errors.append(f"[{email_obj.subject}] Error generando PDF del cuerpo: {exc}")

    att_pdfs_count = 0
    for att in email_obj.attachments:
        try:
            generated = convert_attachment_to_pdf(
                filename=att.filename,
                payload=att.payload,
                dest_dir=mail_dir,
                content_type=att.content_type,
                use_office_com=use_office_com,
            )
            att_pdfs_count += len(generated)
            pdfs_count += len(generated)
        except Exception as exc:
            errors.append(f"[{email_obj.subject}] Error convirtiendo adjunto '{att.filename}': {exc}")

    return mail_dir, pdfs_count, errors


def process_email_batch(
    emails: list[EmailMessageData],
    start_date: date,
    end_date: date,
    output_base_dir: Path,
    use_office_com: bool = True,
) -> ExportSummary:
    """Procesa un lote ya descargado de correos dentro de la carpeta del rango de fechas."""
    range_folder_name = format_date_range_folder(start_date, end_date)
    range_dir = output_base_dir / range_folder_name
    range_dir.mkdir(parents=True, exist_ok=True)

    summary = ExportSummary(range_dir=range_dir)

    for email_obj in emails:
        if email_obj.parsed_date is not None:
            msg_day = email_obj.parsed_date.date()
            if not (start_date <= msg_day <= end_date):
                summary.emails_skipped_out_of_range += 1
                continue

        mail_dir, pdfs_count, errs = export_single_email(
            email_obj,
            range_dir=range_dir,
            use_office_com=use_office_com,
        )
        summary.emails_processed += 1
        summary.pdfs_generated += pdfs_count
        # Resta el PDF del cuerpo si se generó correctamente (cuando hubo al menos 1 PDF y sin error de cuerpo)
        body_failed = any("PDF del cuerpo" in e for e in errs)
        summary.attachments_converted += max(0, pdfs_count - (0 if body_failed else 1))
        summary.errors.extend(errs)

    return summary


def _save_env_config(
    env_path: Path,
    email_addr: str,
    password: str,
    imap_server: str,
    imap_port: int,
    folder: str,
    output_dir: str,
) -> None:
    """Guarda o actualiza la configuración en .env."""
    if not env_path.exists():
        env_path.write_text("", encoding="utf-8")
    set_key(str(env_path), "EMAIL_ADDRESS", email_addr)
    set_key(str(env_path), "EMAIL_PASSWORD", password)
    set_key(str(env_path), "IMAP_SERVER", imap_server)
    set_key(str(env_path), "IMAP_PORT", str(imap_port))
    set_key(str(env_path), "IMAP_FOLDER", folder)
    set_key(str(env_path), "OUTPUT_DIR", output_dir)


def _prompt_date(label: str, default_str: str) -> date:
    """Solicita una fecha YYYY-MM-DD validando el formato."""
    while True:
        raw = Prompt.ask(label, default=default_str)
        try:
            return parse_iso_date(raw)
        except ValueError:
            console.print("[red]Formato inválido. Usa YYYY-MM-DD (ej. 2026-11-01).[/red]")


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Descarga correos por rango de fechas vía IMAP y exporta cuerpo y adjuntos a PDF.",
    )
    parser.add_argument("--email", help="Dirección de correo electrónico (Gmail, Outlook, corporativo, etc.).")
    parser.add_argument("--password", help="Contraseña o Contraseña de Aplicación del correo.")
    parser.add_argument("--imap-server", help="Servidor IMAP personalizado (opcional, se autodetecta por dominio).")
    parser.add_argument("--imap-port", type=int, default=None, help="Puerto IMAP SSL (por defecto: 993).")
    parser.add_argument(
        "--folder",
        default=None,
        help='Carpeta IMAP a leer (por defecto: INBOX). En Gmail puedes usar "[Gmail]/Todos" o "[Gmail]/All Mail".',
    )
    parser.add_argument("--start", help="Fecha inicial inclusiva en formato YYYY-MM-DD.")
    parser.add_argument("--end", help="Fecha final inclusiva en formato YYYY-MM-DD.")
    parser.add_argument("--output", default=None, help="Directorio base de salida (por defecto: ./descargas).")
    parser.add_argument(
        "--test-connection",
        action="store_true",
        help="Solo prueba la conexión IMAP, lista las carpetas disponibles y finaliza.",
    )
    parser.add_argument(
        "--no-office-com",
        action="store_true",
        help="Desactiva Microsoft Office COM y usa LibreOffice o el motor portable Python para Office.",
    )
    return parser


def _run_test_connection(email_addr: str, password: str, imap_server: str, imap_port: int) -> int:
    console.print(f"\n[cyan]Probando conexión IMAP SSL a[/cyan] [bold]{imap_server}:{imap_port}[/bold]...")
    try:
        with EmailFetcher(email_addr, password, imap_server=imap_server, imap_port=imap_port) as fetcher:
            folders = fetcher.list_folders()
    except Exception as exc:
        console.print(f"[bold red]Error de conexión/autenticación:[/bold red] {exc}")
        return 1

    console.print("[bold green]Conexión exitosa.[/bold green]\n")
    if folders:
        table = Table(title="Carpetas IMAP Disponibles")
        table.add_column("Nombre para --folder", style="cyan")
        table.add_column("Nombre legible", style="white")
        for raw_name, readable in folders:
            table.add_row(raw_name, readable)
        console.print(table)
    return 0


def run_cli(argv: list[str] | None = None) -> int:
    """Punto de entrada principal de la aplicación CLI."""
    env_path = _resolve_env_path()
    load_dotenv(dotenv_path=env_path)

    import os

    parser = _build_arg_parser()
    args = parser.parse_args(argv)
    interactive_mode = len(sys.argv) == 1 if argv is None else False

    env_email = os.getenv("EMAIL_ADDRESS", "").strip()
    env_password = os.getenv("EMAIL_PASSWORD", "").strip()
    env_server = os.getenv("IMAP_SERVER", "").strip()
    env_port_str = os.getenv("IMAP_PORT", "993").strip()
    env_folder = os.getenv("IMAP_FOLDER", "INBOX").strip() or "INBOX"
    env_output = os.getenv("OUTPUT_DIR", "descargas").strip() or "descargas"

    try:
        default_port = int(env_port_str)
    except ValueError:
        default_port = 993

    email_addr = (args.email or "").strip()
    password = (args.password or "").strip() if args.password else ""
    custom_server = args.imap_server if args.imap_server is not None else (env_server or None)
    imap_port = args.imap_port if args.imap_port is not None else default_port
    folder = args.folder or env_folder
    output_dir_str = args.output or env_output

    need_wizard = (
        not email_addr
        or not password
        or (not args.test_connection and (not args.start or not args.end))
    )

    verified_folders: list[tuple[str, str]] | None = None

    if need_wizard:
        console.print(
            Panel.fit(
                "[bold cyan]Checker - Exportador Universal de Correos y Adjuntos a PDF[/bold cyan]\n"
                "Compatible con Gmail, Outlook/Office 365, Yahoo, iCloud, Zoho y servidores IMAP corporativos.",
                border_style="cyan",
            )
        )

        entered_new_creds = False
        while True:
            if not email_addr:
                if env_email:
                    email_addr = Prompt.ask("Correo electrónico", default=env_email).strip()
                else:
                    email_addr = Prompt.ask("Correo electrónico").strip()

            detected_server = resolve_imap_server(email_addr, custom_server)
            if args.imap_server is None:
                imap_server = Prompt.ask("Servidor IMAP", default=detected_server).strip()
            else:
                imap_server = detected_server

            if not password:
                if env_password and email_addr == env_email:
                    use_saved = Confirm.ask("¿Usar la contraseña guardada en .env?", default=True)
                    if use_saved:
                        password = env_password

                while not password.strip():
                    if imap_server.lower() == "imap.gmail.com":
                        console.print(
                            "[dim]Nota Gmail: debes ingresar una Contraseña de Aplicación de 16 letras "
                            "(https://myaccount.google.com/apppasswords), no tu contraseña normal.[/dim]"
                        )
                    password = Prompt.ask("Contraseña / Contraseña de Aplicación", password=False)
                    if not password.strip():
                        console.print("[red]La contraseña no puede quedar vacía.[/red]")
                    else:
                        entered_new_creds = True

            console.print(f"[cyan]Verificando acceso IMAP a {imap_server}:{imap_port}...[/cyan]")
            try:
                with EmailFetcher(email_addr, password, imap_server=imap_server, imap_port=imap_port) as fetcher:
                    verified_folders = fetcher.list_folders()
                console.print("[bold green]Conexión IMAP verificada correctamente.[/bold green]")
                break
            except Exception as exc:
                console.print(f"[bold red]Error de autenticación/conexión:[/bold red] {exc}")
                if imap_server.lower() == "imap.gmail.com":
                    console.print(
                        "[yellow]Sugerencia Gmail: verifica que hayas generado una Contraseña de Aplicación "
                        "(16 letras) en https://myaccount.google.com/apppasswords con la Verificación en 2 pasos "
                        "activa, y que no hayas escrito tu contraseña normal ni dejado espacios extra.[/yellow]"
                    )
                if not Confirm.ask("¿Deseas reintentar ingresando las credenciales nuevamente?", default=True):
                    if interactive_mode:
                        input("\nPresiona Enter para salir...")
                    return 1
                email_addr = ""
                password = ""
                custom_server = None
                env_password = ""

        if not args.test_connection:
            if args.folder is None:
                folder = Prompt.ask(
                    'Carpeta IMAP (ej. INBOX, "todos" o "[Gmail]/Todos")',
                    default=folder,
                ).strip() or "INBOX"
            if args.output is None:
                output_dir_str = Prompt.ask("Carpeta de destino", default=output_dir_str).strip() or "descargas"

        if entered_new_creds or not env_path.exists():
            if Confirm.ask(f"¿Guardar configuración en {env_path.name} para futuras ejecuciones?", default=True):
                _save_env_config(env_path, email_addr, password, imap_server, imap_port, folder, output_dir_str)
                console.print(f"[green]Configuración guardada en {env_path}[/green]")
    else:
        imap_server = resolve_imap_server(email_addr, custom_server)

    if args.test_connection:
        if verified_folders is not None:
            table = Table(title="Carpetas IMAP Disponibles")
            table.add_column("Nombre para --folder", style="cyan")
            table.add_column("Nombre legible", style="white")
            for raw_name, readable in verified_folders:
                table.add_row(raw_name, readable)
            console.print(table)
            return 0
        return _run_test_connection(email_addr, password, imap_server, imap_port)

    # Validación de fechas
    today = date.today()
    first_of_month = today.replace(day=1)

    if args.start:
        try:
            start_date = parse_iso_date(args.start)
        except ValueError:
            console.print(f"[bold red]Fecha inicial inválida '{args.start}'. Usa YYYY-MM-DD.[/bold red]")
            return 2
    else:
        start_date = _prompt_date("Fecha inicial (YYYY-MM-DD)", first_of_month.strftime("%Y-%m-%d"))

    if args.end:
        try:
            end_date = parse_iso_date(args.end)
        except ValueError:
            console.print(f"[bold red]Fecha final inválida '{args.end}'. Usa YYYY-MM-DD.[/bold red]")
            return 2
    else:
        end_date = _prompt_date("Fecha final (YYYY-MM-DD)", today.strftime("%Y-%m-%d"))

    while start_date > end_date:
        console.print("[red]La fecha inicial no puede ser mayor que la fecha final.[/red]")
        if args.start and args.end:
            return 2
        start_date = _prompt_date("Fecha inicial (YYYY-MM-DD)", first_of_month.strftime("%Y-%m-%d"))
        end_date = _prompt_date("Fecha final (YYYY-MM-DD)", today.strftime("%Y-%m-%d"))

    output_base_dir = Path(output_dir_str)
    if not output_base_dir.is_absolute():
        output_base_dir = Path.cwd() / output_base_dir

    range_dir = output_base_dir / format_date_range_folder(start_date, end_date)
    range_dir.mkdir(parents=True, exist_ok=True)
    use_office_com = not args.no_office_com

    console.print(
        f"\n[cyan]Conectando a[/cyan] [bold]{imap_server}:{imap_port}[/bold] "
        f"[cyan]como[/cyan] [bold]{email_addr}[/bold] (carpeta: [bold]{folder}[/bold])..."
    )

    summary = ExportSummary(range_dir=range_dir)

    try:
        with EmailFetcher(email_addr, password, imap_server=imap_server, imap_port=imap_port) as fetcher:
            uids = fetcher.search_uids_in_range(start_date, end_date, folder=folder)
            total = len(uids)

            if total == 0:
                console.print(
                    f"[yellow]No se encontraron correos en '{folder}' entre {start_date} y {end_date}.[/yellow]"
                )
                if interactive_mode:
                    input("\nPresiona Enter para salir...")
                return 0

            console.print(f"[green]Se encontraron {total} correo(s) candidatos. Procesando...[/green]\n")

            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                MofNCompleteColumn(),
                TimeElapsedColumn(),
                console=console,
            ) as progress:
                task_id = progress.add_task("Descargando y convirtiendo a PDF...", total=total)

                for uid in uids:
                    try:
                        email_obj = fetcher.fetch_email_by_uid(uid)
                        if email_obj is None:
                            summary.errors.append(f"UID {uid!r}: No se pudo leer el contenido del mensaje.")
                            continue

                        if email_obj.parsed_date is not None:
                            msg_day = email_obj.parsed_date.date()
                            if not (start_date <= msg_day <= end_date):
                                summary.emails_skipped_out_of_range += 1
                                continue

                        short_subj = (email_obj.subject[:42] + "...") if len(email_obj.subject) > 45 else email_obj.subject
                        progress.update(task_id, description=f"Procesando: {short_subj}")

                        _, pdfs_count, errs = export_single_email(
                            email_obj,
                            range_dir=range_dir,
                            use_office_com=use_office_com,
                        )
                        summary.emails_processed += 1
                        summary.pdfs_generated += pdfs_count
                        body_failed = any("PDF del cuerpo" in e for e in errs)
                        summary.attachments_converted += max(0, pdfs_count - (0 if body_failed else 1))
                        summary.errors.extend(errs)
                    except Exception as exc:
                        summary.errors.append(f"UID {uid!r}: Error inesperado: {exc}")
                    finally:
                        progress.advance(task_id, 1)

    except Exception as exc:
        console.print(f"\n[bold red]Error durante la descarga IMAP:[/bold red] {exc}")
        if interactive_mode:
            input("\nPresiona Enter para salir...")
        return 1

    # Tabla resumen final
    res_table = Table(title="Resumen de Exportación", border_style="green")
    res_table.add_column("Métrica", style="bold cyan")
    res_table.add_column("Valor", style="white")
    res_table.add_row("Rango de fechas", f"{start_date} a {end_date}")
    res_table.add_row("Carpeta generada", str(summary.range_dir.resolve()))
    res_table.add_row("Correos exportados", str(summary.emails_processed))
    res_table.add_row("PDFs totales generados", str(summary.pdfs_generated))
    res_table.add_row("Adjuntos convertidos a PDF", str(summary.attachments_converted))
    if summary.emails_skipped_out_of_range:
        res_table.add_row("Omitidos (fuera de fecha local)", str(summary.emails_skipped_out_of_range))
    res_table.add_row("Errores", str(len(summary.errors)))

    console.print()
    console.print(res_table)

    if summary.errors:
        console.print("\n[bold yellow]Detalle de advertencias/errores:[/bold yellow]")
        for err in summary.errors:
            console.print(f" - [yellow]{err}[/yellow]")

    if interactive_mode:
        input("\nPresiona Enter para salir...")
    return 0


if __name__ == "__main__":
    raise SystemExit(run_cli())
