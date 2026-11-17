# Checker - Exportador Universal de Correos y Adjuntos a PDF

Herramienta CLI para Windows (y portable a `.exe`) que conecta vía **IMAP SSL** a cualquier plataforma de correo (**Gmail, Outlook/Office 365, Yahoo, iCloud, Zoho o servidores corporativos**), descarga todos los correos de un rango de fechas y genera la siguiente estructura de carpetas:

```text
descargas/
└── 2026-11-01_a_2026-11-17/
    ├── Asunto del correo 1/
    │   ├── Asunto del correo 1.pdf
    │   ├── factura_electronica.pdf
    │   └── imagen_adjunta.pdf
    └── Asunto del correo 2/
        ├── Asunto del correo 2.pdf
        └── reporte_mensual.pdf
```

## Capacidades de Conversión a PDF

- **Cuerpo del correo**: Se guarda como `<Asunto>.pdf` con encabezado de metadatos (De, Para, CC, Fecha, lista de adjuntos) y el contenido HTML o texto plano formateado.
- **PDF nativos**: Copia directa.
- **Imágenes** (`.jpg`, `.jpeg`, `.png`, `.webp`, `.tiff`, `.bmp`, `.gif`): Conversión encuadrada a páginas A4 (incluye soporte multipágina y orientación EXIF).
- **Textos y Datos** (`.txt`, `.csv`, `.tsv`, `.html`, `.htm`, `.json`, `.md`, `.log`): Tablas formateadas para CSV/TSV, renderizado HTML y código/texto paginado.
- **Facturas Electrónicas / XML** (`.xml`): Soporte especial para facturas UBL 2.1 / DIAN (`Invoice`, `CreditNote`, `DebitNote`, `AttachedDocument` con XML embebido en CDATA), CFDI y XML genérico.
- **Documentos de Office** (`.docx`, `.doc`, `.xlsx`, `.xls`, `.pptx`, `.ppt`): Conversión híbrida en 3 niveles para funcionar en cualquier máquina Windows:
  1. Microsoft Office COM (si Word/Excel/PowerPoint están instalados en Windows).
  2. LibreOffice headless (si está instalado).
  3. Motor puro en Python (`python-docx`, `openpyxl`, `python-pptx` + `reportlab`), sin requerir Office instalado.
- **Archivos `.zip`**: Descomprime automáticamente y convierte cada archivo interno a PDF.

---

## Requisitos Previos para la Primera Prueba con Gmail

Para conectar **Gmail** mediante IMAP estándar de forma segura, Google requiere una **Contraseña de Aplicación** (de 16 caracteres) en lugar de tu contraseña normal:

1. Entra a tu cuenta de Google en [myaccount.google.com/security](https://myaccount.google.com/security).
2. Asegúrate de tener activada la **Verificación en 2 pasos** (*2-Step Verification*).
3. En el buscador superior de la cuenta de Google, escribe **Contraseñas de aplicaciones** (*App passwords*) o ve a [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords).
4. Ponle un nombre (ej. `Checker PDF`) y haz clic en **Crear**.
5. Google te mostrará una contraseña de 16 letras (ej. `abcd efgh ijkl mnop`). Cópiala (puedes escribirla con o sin espacios).
6. Verifica que IMAP esté habilitado en Gmail (por defecto suele estar activo, o revísalo en Gmail -> *Configuración* -> *Reenvío y correo POP/IMAP*).

> **Nota para otras plataformas:**
> - **Outlook / Hotmail / Office 365**: Usa `outlook.office365.com` (autodetectado). Algunas cuentas corporativas de Microsoft 365 requieren habilitar IMAP o usar contraseña de aplicación si tienen MFA.
> - **Corporativos / cPanel**: El programa intenta `imap.<tudominio>` automáticamente, o puedes especificar `--imap-server mail.tudominio.com`.

---

## Cómo Ejecutar

### Opción 1: Doble clic o consola con `ejecutar.bat` (Recomendado en Windows)

Ejecuta `ejecutar.bat`. La primera vez creará automáticamente el entorno virtual `.venv`, instalará las dependencias y lanzará el asistente interactivo en consola.

### Opción 2: Línea de comandos directa

1. Activar o crear entorno virtual e instalar dependencias:
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   ```

2. Probar la conexión y listar tus carpetas de Gmail:
   ```powershell
   python main.py --email tu_correo@gmail.com --password "tu_app_password" --test-connection
   ```

3. Descargar correos de un rango de fechas:
   ```powershell
   python main.py --email tu_correo@gmail.com --password "tu_app_password" --start 2026-11-01 --end 2026-11-17
   ```

   Para descargar incluyendo correos archivados o enviados en Gmail, puedes usar la carpeta `"All Mail"` (en inglés suele ser `"[Gmail]/All Mail"`, en español `"[Gmail]/Todos"` — verifícalo primero con `--test-connection`):
   ```powershell
   python main.py --email tu_correo@gmail.com --password "tu_app_password" --folder "[Gmail]/Todos" --start 2026-11-01 --end 2026-11-17
   ```

4. También puedes copiar `.env.example` a `.env` con tus datos y ejecutar simplemente:
   ```powershell
   python main.py
   ```

### Argumentos disponibles

- `--email`: Dirección de correo.
- `--password`: Contraseña o Contraseña de Aplicación.
- `--imap-server`: Servidor IMAP manual (opcional).
- `--imap-port`: Puerto IMAP SSL (por defecto `993`).
- `--folder`: Carpeta IMAP (por defecto `INBOX`).
- `--start`: Fecha inicial `YYYY-MM-DD`.
- `--end`: Fecha final `YYYY-MM-DD`.
- `--output`: Carpeta raíz donde se guardarán las descargas (por defecto `descargas`).
- `--test-connection`: Prueba el acceso IMAP y muestra las carpetas disponibles.
- `--no-office-com`: Omite Microsoft Office COM y usa LibreOffice o el motor portable Python.

---

## Portabilidad a Cualquier Máquina Windows (`.exe`)

Para generar un único archivo ejecutable autónomo que funcione en cualquier PC con Windows sin instalar Python ni librerías:

1. Haz doble clic en `compilar_exe.bat`.
2. Se generará `dist\CorreoPDFExporter.exe`.
3. Puedes copiar `CorreoPDFExporter.exe` (y opcionalmente tu `.env`) a una memoria USB o cualquier otro equipo Windows y ejecutarlo directamente.
