"""
Filtro Inteligente de Archivos (0 Tokens) y Sanitizador de Textos para Reportes y Documentos.
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
    """Inicializa de forma perezosa el motor local RapidOCR para leer capturas visuales."""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
            _OCR_ENGINE = RapidOCR()
        except Exception:
            _OCR_ENGINE = False
    return _OCR_ENGINE


ARCHIVOS_BASURA_NOMBRES = {
    "thumbs.db", "ehthumbs.db", "desktop.ini", ".ds_store", "_.ds_store", ".gitkeep", ".gitignore"
}
CARPETAS_BASURA_SEGMENTOS = {
    "__macosx", "__pycache__", ".git", ".svn", "node_modules", ".venv", "venv"
}
EXTENSIONES_BASURA = {".tmp", ".temp", ".bak", ".swp", ".pyc", ".pyo", ".pyd", ".db-journal", ".lnk"}
MAX_BYTES_POR_ARCHIVO = 25 * 1024 * 1024
EXTENSIONES_PESADAS_IGNORAR = {".mp4", ".mkv", ".avi", ".mov", ".iso", ".exe", ".dll"}


def es_archivo_basura_sistema(ruta_o_nombre: str) -> bool:
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
    if nombre_low in ARCHIVOS_BASURA_NOMBRES or nombre_base.startswith("~$") or nombre_base.startswith("._"):
        return True
    _, ext = os.path.splitext(nombre_low)
    return ext in EXTENSIONES_BASURA


def _decodificar_nombre_zip(zinfo: zipfile.ZipInfo) -> str:
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
    return nombre.replace("\ufffd", "ú").replace("─", "-")


def desempacar_archivos_subidos(archivos_subidos) -> list:
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
    ext = os.path.splitext(nombre_archivo.lower())[1]
    if ext in (".png", ".jpg", ".jpeg", ".bmp", ".webp", ".gif"):
        return "IMAGEN"
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".csv"):
        return "EXCEL"
    if ext == ".pdf":
        return "PDF"
    if ext in (".docx", ".doc"):
        return "WORD"
    if ext in (".py", ".js", ".ts", ".tsx", ".jsx", ".sql", ".sh", ".ps1", ".html", ".css", ".json", ".yaml", ".yml"):
        return "CODIGO"
    if ext in (".txt", ".log", ".md"):
        return "TEXTO"
    return "OTRO"


def extraer_metadatos_ligeros_archivo(ruta_o_fileobj, nombre_archivo: str, ruta_relativa: str = "") -> dict:
    ext = os.path.splitext(nombre_archivo.lower())[1]
    nom_low = nombre_archivo.lower()
    metadatos = {
        "nombre_archivo": nombre_archivo,
        "extension": ext,
        "hojas": [],
        "encabezados_fila_1": {},
        "resumen_lineas": [],
    }
    if ruta_relativa and ruta_relativa != nombre_archivo:
        metadatos["ruta_en_zip"] = ruta_relativa

    try:
        if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
            if hasattr(ruta_o_fileobj, "seek"):
                ruta_o_fileobj.seek(0)
            wb = openpyxl.load_workbook(ruta_o_fileobj, read_only=True, data_only=True)
            hojas = wb.sheetnames[:8]
            metadatos["hojas"] = hojas
            encabezados = {}
            for nombre_hoja in hojas[:4]:
                ws = wb[nombre_hoja]
                for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                    encabezados[nombre_hoja] = [
                        str(celda).strip()[:50] for celda in row[:15] if celda and str(celda).strip()
                    ]
                    break
            metadatos["encabezados_fila_1"] = encabezados
            wb.close()
            if hasattr(ruta_o_fileobj, "seek"):
                ruta_o_fileobj.seek(0)

        elif ext == ".docx":
            from docx import Document
            if hasattr(ruta_o_fileobj, "seek"):
                ruta_o_fileobj.seek(0)
                raw_bytes = ruta_o_fileobj.read()
                ruta_o_fileobj.seek(0)
                doc_ins = Document(io.BytesIO(raw_bytes))
            else:
                doc_ins = Document(ruta_o_fileobj)

            lineas = []
            for p in doc_ins.paragraphs[:15]:
                txt_p = p.text.strip()
                if txt_p and not re.match(r"^\d+(\.\d+)*\s*$", txt_p):
                    lineas.append(txt_p[:180])
            metadatos["resumen_lineas"] = lineas
            txt_all = " ".join(lineas).lower()

            if any(k in nom_low or k in txt_all for k in ("seguridad", "inmutabilidad", "sha-256", "control de evidencias")):
                metadatos["tag_semantico"] = "SEGURIDAD_INMUTABILIDAD"
            elif any(k in nom_low or k in txt_all for k in ("das", "hld", "arquitectura", "diseño", "diseno")):
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

                lineas_pdf = []
                for page in reader.pages[:6]:
                    txt_p = page.extract_text() or ""
                    for ln in txt_p.splitlines():
                        limpia = re.sub(r"\s+", " ", ln).strip()
                        # Omitir correos personales y líneas de puntos de índice
                        if "@" in limpia or re.search(r"\.{4,}", limpia):
                            continue
                        if len(limpia) > 12:
                            lineas_pdf.append(limpia[:150])
                    if len(lineas_pdf) > 25:
                        break

                metadatos["resumen_lineas"] = lineas_pdf[:15]
                txt_eval = f"{nom_low} {' '.join(lineas_pdf).lower()}"

                if any(k in nom_low for k in ("cronograma", "plant120", "gantt")):
                    metadatos["tag_semantico"] = "CRONOGRAMA_PLANEACION"
                elif any(k in nom_low for k in ("das_", "hld", "arquitectura", "diseño de software", "diseno de software")):
                    metadatos["tag_semantico"] = "ARQUITECTURA_DISENO"
                elif any(k in nom_low for k in ("cooperapp", "permisos")):
                    metadatos["tag_semantico"] = "PERMISOS_SISTEMAS"
                elif "seguridad y control" in nom_low or "it-sigaf" in nom_low:
                    metadatos["tag_semantico"] = "SEGURIDAD_INMUTABILIDAD"
            except Exception:
                pass

        elif ext in (".png", ".jpg", ".jpeg", ".bmp", ".webp"):
            ocr = _obtener_motor_ocr()
            if ocr:
                try:
                    from PIL import Image
                    import numpy as np
                    if hasattr(ruta_o_fileobj, "read"):
                        if hasattr(ruta_o_fileobj, "seek"):
                            ruta_o_fileobj.seek(0)
                        raw_img = ruta_o_fileobj.read()
                        if hasattr(ruta_o_fileobj, "seek"):
                            ruta_o_fileobj.seek(0)
                        img_pil = Image.open(io.BytesIO(raw_img)).convert("RGB")
                    else:
                        img_pil = Image.open(ruta_o_fileobj).convert("RGB")

                    img_pil.thumbnail((1200, 1200))
                    res_ocr, _ = ocr(np.array(img_pil))
                    if res_ocr:
                        lineas_ocr = [line[1].strip()[:140] for line in res_ocr if line[1].strip()]
                        metadatos["resumen_lineas"] = lineas_ocr[:10]
                        txt_ocr = " ".join(lineas_ocr).lower()

                        if any(k in txt_ocr for k in ("merge request", "gitlab", "commit", "branch", "pull request")):
                            metadatos["tag_semantico"] = "SOFTWARE_GIT"
                        elif any(k in txt_ocr for k in ("kanban", "sprint", "tablero", "in progress", "backlog")):
                            metadatos["tag_semantico"] = "TABLERO_AGIL"
                        elif any(k in txt_ocr for k in ("teams", "asistentes", "reunión", "unirse")):
                            metadatos["tag_semantico"] = "REUNION_SEGUIMIENTO"
                        elif any(k in txt_ocr for k in ("cronograma", "gantt", "duración")):
                            metadatos["tag_semantico"] = "CRONOGRAMA_PLANEACION"
                except Exception:
                    pass
    except Exception as exc:
        metadatos["nota_inspeccion"] = f"Lectura finalizada ({type(exc).__name__})"

    return metadatos


def sanitizar_texto_para_reporte(texto: str) -> str:
    """Garantiza textos 100% limpios de etiquetas XML/HTML y corchetes angulares sin escapar."""
    if not texto:
        return ""
    limpio = str(texto)
    limpio = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", limpio)
    limpio = re.sub(r"<\s*/?\s*para\s*>", "'para'", limpio, flags=re.IGNORECASE)
    limpio = re.sub(r"<\s*/?\s*(font|b|i|u|br|span|div|p|table|tr|td)\b[^>]*>", "", limpio, flags=re.IGNORECASE)
    # Reemplazar corchetes angulares para compatibilidad con ReportLab y Word
    limpio = limpio.replace("<", "‹").replace(">", "›")
    return limpio.strip()