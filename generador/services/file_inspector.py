"""
Filtro Inteligente de Archivos (0 Tokens) y Sanitizador de Textos para Reportes y Documentos.

Regla de Oro:
- Jamás enviar archivos crudos pesados al modelo de IA.
- Si el contratista adjunta un Excel (incluso de millones de filas), extrae únicamente
  el nombre del archivo, los nombres de las hojas y los primeros encabezados de la Fila 1
  usando openpyxl en modo read_only=True.
- Sanitiza textos para erradicar etiquetas como <para> o caracteres XML sin escapar
  que causan errores en procesadores de documentos y motores de reporte.
"""

import os
import re
import csv
import io
import zipfile
import openpyxl
from django.core.files.base import ContentFile

_OCR_ENGINE = None


def _obtener_motor_ocr():
    """Inicializa de forma perezosa el motor local RapidOCR para leer capturas y evidencias visuales."""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        try:
            from rapidocr_onnxruntime import RapidOCR

            _OCR_ENGINE = RapidOCR()
        except Exception:
            _OCR_ENGINE = False
    return _OCR_ENGINE


ARCHIVOS_BASURA_NOMBRES = {
    "thumbs.db",
    "ehthumbs.db",
    "ehthumbs_vista.db",
    "desktop.ini",
    ".ds_store",
    "_.ds_store",
    "icon\r",
    ".gitkeep",
    ".gitignore",
}

CARPETAS_BASURA_SEGMENTOS = {
    "__macosx",
    "__pycache__",
    ".git",
    ".svn",
    ".hg",
    ".idea",
    ".vscode",
    "node_modules",
    ".venv",
    "venv",
}

EXTENSIONES_BASURA = {
    ".tmp",
    ".temp",
    ".bak",
    ".swp",
    ".pyc",
    ".pyo",
    ".pyd",
    ".db-journal",
    ".lnk",
}

MAX_BYTES_POR_ARCHIVO = 25 * 1024 * 1024  # 25 MB máximo por archivo soporte
EXTENSIONES_PESADAS_IGNORAR = {".mp4", ".mkv", ".avi", ".mov", ".iso", ".exe", ".dll", ".bak"}


def es_archivo_basura_sistema(ruta_o_nombre: str) -> bool:
    """
    Detecta archivos basura del sistema operativo o cachés temporales que jamás deben
    incluirse como evidencia contractual (ej. Thumbs.db, .DS_Store, __MACOSX, ~$archivo.docx, .pyc).
    """
    if not ruta_o_nombre:
        return True
    normalizada = ruta_o_nombre.replace("\\", "/").strip("/")
    partes = [p.strip() for p in normalizada.split("/") if p.strip()]
    if not partes:
        return True

    for seg in partes[:-1]:
        if seg.lower() in CARPETAS_BASURA_SEGMENTOS or seg.startswith("."):
            return True

    nombre_base = partes[-1]
    nombre_low = nombre_base.lower()

    if nombre_low in ARCHIVOS_BASURA_NOMBRES:
        return True
    if nombre_base.startswith("~$") or nombre_base.startswith(".~") or nombre_base.startswith("._"):
        return True

    _, ext = os.path.splitext(nombre_low)
    if ext in EXTENSIONES_BASURA:
        return True

    return False


def _decodificar_nombre_zip(zinfo: zipfile.ZipInfo) -> str:
    """
    Corrige la codificación de nombres de archivo dentro de archivos .ZIP creados en Windows
    (donde el flag UTF-8 0x800 está apagado y Python decodifica como cp437 generando 'Jess' o '┼ô').
    """
    nombre = zinfo.filename
    if not (zinfo.flag_bits & 0x800):
        try:
            raw_bytes = nombre.encode("cp437")
            try:
                nombre = raw_bytes.decode("utf-8")
            except UnicodeDecodeError:
                nombre = raw_bytes.decode("cp850", errors="replace")
        except Exception:
            pass
    # Limpiar caracteres de reemplazo si venían truncados
    nombre = nombre.replace("\ufffd", "ú").replace("─", "-")
    return nombre


def desempacar_archivos_subidos(archivos_subidos) -> list:
    """
    Recibe una lista de archivos subidos desde Django (`request.FILES.getlist(...)`).
    - Si un archivo es un `.zip`, lo abre automáticamente, ignora `Thumbs.db`, `__MACOSX`, `__pycache__`,
      archivos temporales `~$` y carpetas vacías, y devuelve cada archivo real listo para guardarse
      y clasificarse en su obligación correspondiente.
    - Si es un archivo individual normal, verifica que no sea `Thumbs.db` u otro archivo basura y lo devuelve.
    Retorna lista de dicts:
      `{"file_obj": ContentFile|UploadedFile, "nombre_original": str, "ruta_relativa": str, "tamano_bytes": int, "origen_zip": str}`
    """
    items_validos = []

    for f_up in archivos_subidos:
        nombre_up = getattr(f_up, "name", "archivo")
        if es_archivo_basura_sistema(nombre_up):
            continue

        ext_up = os.path.splitext(nombre_up.lower())[1]
        if ext_up == ".zip":
            try:
                if hasattr(f_up, "seek"):
                    f_up.seek(0)
                with zipfile.ZipFile(f_up, "r") as zf:
                    for zinfo in zf.infolist():
                        if zinfo.is_dir() or zinfo.file_size <= 0 or zinfo.file_size > MAX_BYTES_POR_ARCHIVO:
                            continue
                        _, ext_interna = os.path.splitext(zinfo.filename.lower())
                        if ext_interna in EXTENSIONES_PESADAS_IGNORAR:
                            continue
                        ruta_interna = _decodificar_nombre_zip(zinfo).replace("\\", "/").strip("/")
                        if es_archivo_basura_sistema(ruta_interna):
                            continue
                        nombre_limpio = os.path.basename(ruta_interna)
                        if not nombre_limpio or nombre_limpio.lower().endswith(".zip"):
                            continue

                        # Si ya había otro archivo con el mismo nombre en otra carpeta del ZIP, prefijar su carpeta
                        nombres_ya_vistos = {it["nombre_original"].lower() for it in items_validos}
                        if nombre_limpio.lower() in nombres_ya_vistos:
                            carpeta_padre = os.path.basename(os.path.dirname(ruta_interna))
                            if carpeta_padre:
                                nombre_limpio = f"{carpeta_padre}_{nombre_limpio}"

                        datos_bytes = zf.read(zinfo.filename)
                        cf = ContentFile(datos_bytes, name=nombre_limpio)
                        items_validos.append({
                            "file_obj": cf,
                            "nombre_original": nombre_limpio,
                            "ruta_relativa": ruta_interna,
                            "tamano_bytes": len(datos_bytes),
                            "origen_zip": nombre_up,
                        })
                continue
            except Exception:
                # Si el zip está corrupto o no se pudo abrir como zip, continuar ignorándolo para no meter binarios rotos
                continue

        items_validos.append({
            "file_obj": f_up,
            "nombre_original": os.path.basename(nombre_up),
            "ruta_relativa": nombre_up,
            "tamano_bytes": getattr(f_up, "size", 0),
            "origen_zip": "",
        })

    return items_validos


def clasificar_tipo_archivo(nombre_archivo: str) -> str:
    """Clasifica un archivo según su extensión para decidir su tratamiento en el árbol de anexos."""
    ext = os.path.splitext(nombre_archivo.lower())[1]
    if ext in (".png", ".jpg", ".jpeg", ".bmp", ".webp", ".gif"):
        return "IMAGEN"
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".csv"):
        return "EXCEL"
    if ext == ".pdf":
        return "PDF"
    if ext in (".docx", ".doc"):
        return "WORD"
    if ext in (".py", ".js", ".ts", ".tsx", ".jsx", ".sql", ".sh", ".ps1", ".html", ".css", ".json", ".yaml", ".yml", ".cs", ".java", ".php"):
        return "CODIGO"
    if ext in (".txt", ".log", ".md"):
        return "TEXTO"
    return "OTRO"


def extraer_metadatos_ligeros_archivo(ruta_o_fileobj, nombre_archivo: str, ruta_relativa: str = "") -> dict:
    """
    Inspecciona en milisegundos el archivo subido sin cargar filas pesadas en memoria:
    - Excel (.xlsx / .xlsm): abre con openpyxl(read_only=True, data_only=True) y lee solo pestañas y Fila 1.
    - CSV: lee únicamente la primera línea.
    - Word (.docx): extrae los primeros 6 párrafos/títulos.
    - PDF (.pdf): extrae palabras clave semánticas del nombre/ruta y metadatos ligeros de título si están visibles.
    - Código (.py, .sql, .js, .ts, .html, etc.): extrae docstrings iniciales y firmas de funciones/clases/tablas (`def`, `class`, `CREATE`, `SELECT`) en los primeros 3 KB.
    - Texto/Logs (.txt, .log, .md): extrae las primeras 15 líneas no vacías.
    """
    ext = os.path.splitext(nombre_archivo.lower())[1]
    metadatos = {
        "nombre_archivo": nombre_archivo,
        "extension": ext,
        "hojas": [],
        "encabezados_fila_1": {},
    }
    if ruta_relativa and ruta_relativa != nombre_archivo:
        metadatos["ruta_en_zip"] = ruta_relativa

    try:
        if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
            if hasattr(ruta_o_fileobj, "seek"):
                ruta_o_fileobj.seek(0)
            wb = openpyxl.load_workbook(ruta_o_fileobj, read_only=True, data_only=True)
            hojas = wb.sheetnames[:10]
            metadatos["hojas"] = hojas
            encabezados = {}
            for nombre_hoja in hojas[:5]:
                ws = wb[nombre_hoja]
                primera_fila = []
                for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                    primera_fila = [
                        str(celda).strip()[:60]
                        for celda in row[:20]
                        if celda is not None and str(celda).strip() != ""
                    ]
                    break
                encabezados[nombre_hoja] = primera_fila
            metadatos["encabezados_fila_1"] = encabezados
            wb.close()
            if hasattr(ruta_o_fileobj, "seek"):
                ruta_o_fileobj.seek(0)
        elif ext == ".csv":
            if hasattr(ruta_o_fileobj, "seek"):
                ruta_o_fileobj.seek(0)
                linea = ruta_o_fileobj.readline()
                if isinstance(linea, bytes):
                    linea = linea.decode("utf-8", errors="ignore")
                ruta_o_fileobj.seek(0)
            else:
                with open(ruta_o_fileobj, "r", encoding="utf-8", errors="ignore") as f:
                    linea = f.readline()
            lector = next(csv.reader([linea]), [])
            metadatos["hojas"] = ["CSV_Principal"]
            metadatos["encabezados_fila_1"] = {
                "CSV_Principal": [str(col).strip()[:60] for col in lector[:20] if str(col).strip()]
            }
        elif ext in (".py", ".js", ".ts", ".tsx", ".jsx", ".sql", ".sh", ".ps1", ".html", ".css", ".json", ".yaml", ".yml", ".cs", ".java", ".php"):
            if hasattr(ruta_o_fileobj, "seek"):
                ruta_o_fileobj.seek(0)
                raw = ruta_o_fileobj.read(4096)
                if isinstance(raw, bytes):
                    raw = raw.decode("utf-8", errors="ignore")
                ruta_o_fileobj.seek(0)
            else:
                with open(ruta_o_fileobj, "r", encoding="utf-8", errors="ignore") as f:
                    raw = f.read(4096)
            simbolos = []
            for ln in raw.splitlines():
                s = ln.strip()
                if not s:
                    continue
                if re.match(r"^(def |class |async def |function |export |const |CREATE |ALTER |SELECT |INSERT |UPDATE |#|--|\"\"\"|/\*)", s, flags=re.IGNORECASE):
                    simbolos.append(s[:110])
                if len(simbolos) >= 12:
                    break
            if not simbolos:
                simbolos = [ln.strip()[:100] for ln in raw.splitlines() if ln.strip()][:8]
            metadatos["resumen_lineas"] = simbolos
        elif ext in (".txt", ".log", ".md"):
            if hasattr(ruta_o_fileobj, "seek"):
                ruta_o_fileobj.seek(0)
                raw = ruta_o_fileobj.read(3072)
                if isinstance(raw, bytes):
                    raw = raw.decode("utf-8", errors="ignore")
                ruta_o_fileobj.seek(0)
            else:
                with open(ruta_o_fileobj, "r", encoding="utf-8", errors="ignore") as f:
                    raw = f.read(3072)
            lineas = [ln.strip()[:100] for ln in raw.splitlines() if ln.strip()][:15]
            metadatos["resumen_lineas"] = lineas
        elif ext == ".docx":
            from docx import Document

            if hasattr(ruta_o_fileobj, "seek"):
                ruta_o_fileobj.seek(0)
                raw_bytes = ruta_o_fileobj.read()
                ruta_o_fileobj.seek(0)
                doc_ins = Document(io.BytesIO(raw_bytes))
            else:
                doc_ins = Document(ruta_o_fileobj)
            primeros_parrafos = [p.text.strip()[:180] for p in doc_ins.paragraphs if p.text.strip()][:16]
            for tabla in doc_ins.tables[:4]:
                for fila in tabla.rows[:4]:
                    texto_fila = " | ".join(celda.text.strip() for celda in fila.cells if celda.text.strip())
                    if texto_fila:
                        primeros_parrafos.append(texto_fila[:180])
            metadatos["resumen_lineas"] = primeros_parrafos
            txt_docx = " ".join(primeros_parrafos).lower()
            nom_docx_low = nombre_archivo.lower()
            if any(k in nom_docx_low for k in ("das_", "das-", "hld", "arquitectura", "diseño", "diseno")):
                metadatos["tag_semantico"] = "ARQUITECTURA_DISENO"
            elif any(k in txt_docx for k in ("seguridad", "inmutabilidad", "sha-256", "control de evidencias")):
                metadatos["tag_semantico"] = "SEGURIDAD_INMUTABILIDAD"
            elif any(k in txt_docx for k in ("arquitectura", "das", "hld", "diseño")):
                metadatos["tag_semantico"] = "ARQUITECTURA_DISENO"
        elif ext == ".pdf":
            try:
                import pypdf

                if hasattr(ruta_o_fileobj, "read"):
                    if hasattr(ruta_o_fileobj, "seek"):
                        ruta_o_fileobj.seek(0)
                    reader = pypdf.PdfReader(ruta_o_fileobj)
                    if hasattr(ruta_o_fileobj, "seek"):
                        ruta_o_fileobj.seek(0)
                else:
                    reader = pypdf.PdfReader(ruta_o_fileobj)

                texto_pdf = []
                for page in reader.pages[:8]:
                    txt_p = page.extract_text() or ""
                    if txt_p.strip():
                        texto_pdf.extend([ln.strip()[:140] for ln in txt_p.splitlines() if ln.strip()])

                metadatos["resumen_lineas"] = texto_pdf[:30]
                txt_pdf_low = " ".join(texto_pdf).lower()
                metadatos["texto_pdf"] = txt_pdf_low[:2000]

                nom_low = nombre_archivo.lower()
                txt_eval = f"{nom_low} {txt_pdf_low}"
                if any(k in nom_low for k in ("cronograma", "plan", "gantt")) or any(k in txt_pdf_low for k in ("predecesoras", "diagrama de gantt")):
                    metadatos["tag_semantico"] = "CRONOGRAMA_PLANEACION"
                elif any(k in nom_low for k in ("das_", "das-", "hld", "arquitectura", "diseño", "diseno")):
                    metadatos["tag_semantico"] = "ARQUITECTURA_DISENO"
                elif any(k in txt_eval for k in ("permisos de administrador", "cooperapp", "solicitud de permisos", "asignación de permisos")):
                    metadatos["tag_semantico"] = "PERMISOS_SISTEMAS"
                elif any(k in txt_eval for k in ("seguridad", "inmutabilidad", "sha-256", "control de evidencias")):
                    metadatos["tag_semantico"] = "SEGURIDAD_INMUTABILIDAD"
                elif any(k in txt_eval for k in ("arquitectura", "diseño de alto nivel", "diseño de software")):
                    metadatos["tag_semantico"] = "ARQUITECTURA_DISENO"
            except Exception as exc:
                nombre_humano = re.sub(r"[_\-]+", " ", os.path.splitext(nombre_archivo)[0]).strip()
                metadatos["resumen_lineas"] = [f"Documento PDF: {nombre_humano}"]
        elif ext in (".png", ".jpg", ".jpeg", ".bmp", ".webp"):
            # Reconocimiento Óptico de Caracteres (OCR) con motor RapidOCR
            ocr = _obtener_motor_ocr()
            if ocr:
                try:
                    from PIL import Image
                    import numpy as np

                    if hasattr(ruta_o_fileobj, "read"):
                        if hasattr(ruta_o_fileobj, "seek"):
                            ruta_o_fileobj.seek(0)
                        img_bytes = ruta_o_fileobj.read()
                        if hasattr(ruta_o_fileobj, "seek"):
                            ruta_o_fileobj.seek(0)
                        img_pil = Image.open(io.BytesIO(img_bytes)).convert("RGB")
                    else:
                        img_pil = Image.open(ruta_o_fileobj).convert("RGB")

                    # OPTIMIZACIÓN CRÍTICA DE VELOCIDAD Y RAM: Reducir a máx 1100px antes del OCR
                    img_pil.thumbnail((1100, 1100))
                    img_arr = np.array(img_pil)
                    res_ocr, _ = ocr(img_arr)

                    if res_ocr:
                        lineas_ocr = [line[1].strip()[:140] for line in res_ocr if line[1].strip()]
                        metadatos["lineas_ocr"] = lineas_ocr[:25]
                        metadatos["resumen_lineas"] = lineas_ocr[:12]
                        txt_ocr_low = " ".join(lineas_ocr).lower()
                        metadatos["texto_ocr"] = txt_ocr_low[:1800]

                        # Clasificación semántica de la captura
                        if any(k in txt_ocr_low for k in ("asistencia", "participantes", "resumen de la reunión", "reunión", "unirse", "teams", "salas para sesión", "convocatoria")):
                            metadatos["tag_semantico"] = "REUNION_SEGUIMIENTO"
                        elif any(k in txt_ocr_low for k in (
                            "merge requests", "merge request", "hotfix", "commits", "commit",
                            "gitlab", "branch", "repositorio", "feat(", "fix(",
                        )):
                            metadatos["tag_semantico"] = "SOFTWARE_GIT"
                        elif any(k in txt_ocr_low for k in ("kanban", "closed", "in progress", "sprint", "tablero", "backlog", "to do")):
                            metadatos["tag_semantico"] = "TABLERO_AGIL"
                        elif any(k in txt_ocr_low for k in ("cronograma", "gantt", "duración", "comienzo", "fin", "predecesoras")):
                            metadatos["tag_semantico"] = "CRONOGRAMA_PLANEACION"
                        elif any(k in txt_ocr_low for k in ("arquitectura", "das", "hld", "diseño de software")):
                            metadatos["tag_semantico"] = "ARQUITECTURA_DISENO"
                        elif any(k in txt_ocr_low for k in ("seguridad", "inmutabilidad", "sha-256", "control de evidencias", "hashing")):
                            metadatos["tag_semantico"] = "SEGURIDAD_INMUTABILIDAD"
                        elif any(k in txt_ocr_low for k in ("incidente", "régimen tributario", "retención", "rol consulta", "permisos")):
                            metadatos["tag_semantico"] = "SOPORTE_INCIDENTES"
                except Exception as exc:
                    metadatos["nota_ocr"] = f"OCR omitido ({exc})"
    except Exception as exc:
        metadatos["nota_inspeccion"] = f"Archivo registrado ({type(exc).__name__})"

    return metadatos


def sanitizar_texto_para_reporte(texto: str) -> str:
    """
    Garantiza compatibilidad 100% con procesadores de texto y motores de reporte.
    Reemplaza secuencias de etiquetas XML/HTML peligrosas como <para>, </para>, <br>, <font>,
    y evita corchetes angulares crudos o caracteres de control que rompan analizadores XML/Word,
    manteniendo el texto legible y limpio para insertar en Word o pegar en plataformas.
    """
    if not texto:
        return ""
    limpio = str(texto)
    # Erradicar bytes nulos y caracteres de control ASCII que hacen crashear a python-docx (lxml)
    limpio = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", limpio)
    limpio = re.sub(r"<\s*/?\s*para\s*>", "'para'", limpio, flags=re.IGNORECASE)
    limpio = re.sub(r"<\s*/?\s*(font|b|i|u|br|span|div|p|table|tr|td)\b[^>]*>", "", limpio, flags=re.IGNORECASE)
    limpio = limpio.replace("<", "‹").replace(">", "›")
    return limpio.strip()

# Alias de compatibilidad
sanitizar_texto_para_reporte = sanitizar_texto_para_reporte
