"""
Conversor de los documentos base reales de plantillas_base/ a plantillas dinámicas docxtpl (.docx)
preservando al 100% los encabezados institucionales UAEGRTD, logos, tablas, bordes, sombreados (#C5E0B3) y fuentes Arial.

Genera en plantillas_dinamicas/:
1. plantilla_informe_principal.docx   (basada en plantillas_base/CTO_1031_JOHN_VEGA_INFORME_10.docx)
2. plantilla_reporte_actividad_oti.docx (basada en plantillas_base/PLANTILLA_REPORTE_INDIVIDUAL_ACTIVIDAD_OTI.docx)
3. plantilla_anexo_obligacion.docx    (basada en plantillas_base/PLANTILLA_ANEXO_EVIDENCIAS_OBLIGACION.docx)
4. plantilla_ruta_evidencias.docx     (basada en plantillas_base/Evidencia_Pago_10.docx)
"""

import os
from copy import deepcopy
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.table import _Cell
from django.conf import settings


def _reemplazar_texto_celda_preservando_estilo(cell, nuevo_texto: str, bold=None, font_size_pt=None, align=None):
    """Limpia los párrafos adicionales de una celda y asigna nuevo_texto preservando la fuente Arial y el estilo."""
    while len(cell.paragraphs) > 1:
        p_del = cell.paragraphs[-1]._element
        p_del.getparent().remove(p_del)

    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align

    if p.runs:
        r0 = p.runs[0]
        r0.text = nuevo_texto
        if bold is not None:
            r0.bold = bold
        if font_size_pt is not None:
            r0.font.size = Pt(font_size_pt)
        r0.font.name = "Arial"
        for r_extra in p.runs[1:]:
            r_extra.text = ""
    else:
        r0 = p.add_run(nuevo_texto)
        r0.font.name = "Arial"
        if bold is not None:
            r0.bold = bold
        if font_size_pt is not None:
            r0.font.size = Pt(font_size_pt)


def _desenvolver_sdt_en_fila(tr_element):
    """
    Si una fila <w:tr> contiene un control de contenido <w:sdt> envolviendo un <w:tc>
    (como el desplegable de NÚMERO DE PAGO en la Tabla 0 de GC-FO-15), extrae el <w:tc>
    directamente dentro de <w:tr> para que python-docx y docxtpl puedan escribir el número de pago.
    """
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for sdt in tr_element.findall("w:sdt", ns):
        sdt_content = sdt.find("w:sdtContent", ns)
        if sdt_content is not None:
            tcs = sdt_content.findall("w:tc", ns)
            idx = list(tr_element).index(sdt)
            for offset, tc in enumerate(tcs):
                tr_element.insert(idx + offset, tc)
        tr_element.remove(sdt)


def construir_plantilla_informe_principal(origen_path: str, destino_path: str):
    """
    Convierte CTO_1031_JOHN_VEGA_INFORME_10.docx en plantilla_informe_principal.docx
    con marcadores Jinja2 compatibles con docxtpl y estructura intacta para python-docx.
    """
    doc = Document(origen_path)

    # Tabla 0: Condiciones del Contrato
    t0 = doc.tables[0]
    # Desenvolver <w:sdt> en la fila 5 (Número de Pago)
    _desenvolver_sdt_en_fila(t0.rows[5]._tr)

    # Fila 0: Nombre del contratista y Cédula
    tcs_r0 = [_Cell(tc, t0) for tc in t0.rows[0]._tr.tc_lst]
    _reemplazar_texto_celda_preservando_estilo(tcs_r0[1], "{{ nombre_contratista }}")
    _reemplazar_texto_celda_preservando_estilo(tcs_r0[3], "{{ cedula_contratista }}")

    # Fila 1: Número contrato y Año contrato
    tcs_r1 = [_Cell(tc, t0) for tc in t0.rows[1]._tr.tc_lst]
    _reemplazar_texto_celda_preservando_estilo(tcs_r1[1], "{{ numero_contrato }}")
    _reemplazar_texto_celda_preservando_estilo(tcs_r1[3], "{{ anio_contrato }}")

    # Fila 2: Fecha inicio y Fecha terminación actual
    tcs_r2 = [_Cell(tc, t0) for tc in t0.rows[2]._tr.tc_lst]
    _reemplazar_texto_celda_preservando_estilo(tcs_r2[1], "{{ fecha_inicio }}")
    _reemplazar_texto_celda_preservando_estilo(tcs_r2[3], "{{ fecha_terminacion_actual }}")

    # Fila 4: Objeto del contrato
    tcs_r4 = [_Cell(tc, t0) for tc in t0.rows[4]._tr.tc_lst]
    _reemplazar_texto_celda_preservando_estilo(tcs_r4[0], "{{ objeto_contrato }}")

    # Fila 5: Número de pago y Supervisor
    tcs_r5 = [_Cell(tc, t0) for tc in t0.rows[5]._tr.tc_lst]
    if len(tcs_r5) >= 4:
        _reemplazar_texto_celda_preservando_estilo(tcs_r5[1], "{{ numero_pago }}")
        _reemplazar_texto_celda_preservando_estilo(tcs_r5[3], "{{ nombre_supervisor }}")

    # Fila 6: Cargo supervisor
    tcs_r6 = [_Cell(tc, t0) for tc in t0.rows[6]._tr.tc_lst]
    _reemplazar_texto_celda_preservando_estilo(tcs_r6[1], "{{ cargo_supervisor }}")

    # Fila 7: Dirección u oficina
    tcs_r7 = [_Cell(tc, t0) for tc in t0.rows[7]._tr.tc_lst]
    _reemplazar_texto_celda_preservando_estilo(tcs_r7[1], "{{ direccion_oficina }}")

    # Fila 8: Dirección territorial
    tcs_r8 = [_Cell(tc, t0) for tc in t0.rows[8]._tr.tc_lst]
    _reemplazar_texto_celda_preservando_estilo(tcs_r8[1], "{{ direccion_territorial }}")

    # Párrafo 30: Observaciones adicionales y Párrafo 60: Firma del contratista
    for p in doc.paragraphs:
        if "En este espacio el contratista podrá incluir observaciones" in p.text:
            if p.runs:
                p.runs[0].text = "{{ observaciones_adicionales }}"
                for r in p.runs[1:]:
                    r.text = ""
        elif "John Felipe Vega Ot" in p.text:
            if p.runs:
                p.runs[0].text = "{{ nombre_contratista }}"
                for r in p.runs[1:]:
                    r.text = ""

    # Tabla 3: Bloque de firma del supervisor al final
    if len(doc.tables) >= 4:
        t3 = doc.tables[3]
        _reemplazar_texto_celda_preservando_estilo(t3.rows[0].cells[0], "{{ nombre_supervisor }}", bold=True, font_size_pt=8)
        _reemplazar_texto_celda_preservando_estilo(t3.rows[1].cells[0], "{{ cargo_supervisor }}\nSUPERVISOR", bold=False, font_size_pt=8)

    doc.save(destino_path)


def construir_plantilla_reporte_actividad_oti(origen_path: str, destino_path: str):
    """
    Convierte PLANTILLA_REPORTE_INDIVIDUAL_ACTIVIDAD_OTI.docx en plantilla_reporte_actividad_oti.docx
    preservando su tabla de Información General, sus estilos de títulos y el bloque de firma.
    """
    doc = Document(origen_path)

    # Tabla 0: Información General (4 filas x 2 columnas)
    t0 = doc.tables[0]
    _reemplazar_texto_celda_preservando_estilo(t0.rows[0].cells[1], "{{ nombre_contratista }}")
    _reemplazar_texto_celda_preservando_estilo(t0.rows[1].cells[1], "{{ numero_contrato }}")
    _reemplazar_texto_celda_preservando_estilo(t0.rows[2].cells[1], "{{ periodo_reportado }}")
    _reemplazar_texto_celda_preservando_estilo(t0.rows[3].cells[1], "{{ fecha_reporte }}")

    # Actualizar párrafo de Obligación Contractual y limpiar cualquier firma incrustada del molde base
    for idx, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("Obligación Contractual:"):
            p.clear()
            r_label = p.add_run("Obligación Contractual: ")
            r_label.bold = True
            r_label.font.name = "Arial"
            r_label.font.size = Pt(11)
            r_val = p.add_run("{{ descripcion_obligacion }}")
            r_val.bold = False
            r_val.font.name = "Arial"
            r_val.font.size = Pt(11)
        elif p.text.strip().startswith("Firma:") and idx + 1 < len(doc.paragraphs):
            doc.paragraphs[idx + 1].clear()

    doc.save(destino_path)


def construir_plantilla_anexo_obligacion(origen_path: str, destino_path: str):
    """
    Convierte PLANTILLA_ANEXO_EVIDENCIAS_OBLIGACION.docx en plantilla_anexo_obligacion.docx
    preservando márgenes, estilos Arial centrados de encabezado y subtítulo.
    """
    doc = Document(origen_path)
    if len(doc.paragraphs) >= 2:
        p0 = doc.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p0.runs:
            p0.runs[0].text = "EVIDENCIAS Y ANEXOS TÉCNICOS — OBLIGACIÓN {{ numero_obligacion }} ({{ mes_mayus }} {{ anio }})"
            p0.runs[0].bold = True
            p0.runs[0].font.name = "Arial"
            for r in p0.runs[1:]:
                r.text = ""

        p1 = doc.paragraphs[1]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p1.runs:
            p1.runs[0].text = "Contrato N° {{ numero_contrato }} de {{ anio_contrato }} | Contratista: {{ nombre_contratista }} | {{ proyecto_area }}"
            p1.runs[0].font.name = "Arial"
            p1.runs[0].font.size = Pt(10)
            for r in p1.runs[1:]:
                r.text = ""

    doc.save(destino_path)


def construir_plantilla_ruta_evidencias(origen_path: str, destino_path: str):
    """
    Convierte Evidencia_Pago_10.docx en plantilla_ruta_evidencias.docx
    preservando el formato de la Circular 0009 de 2024 y la captura del explorador de archivos NAS.
    """
    doc = Document(origen_path)
    if len(doc.paragraphs) >= 4:
        p0 = doc.paragraphs[0]
        if p0.runs:
            p0.runs[0].text = "RUTA EVIDENCIAS CUENTAS DE COBRO {{ mes_mayus }} - {{ anio }}"
            p0.runs[0].bold = True
            for r in p0.runs[1:]:
                r.text = ""
        p3 = doc.paragraphs[3]
        if p3.runs:
            p3.runs[0].text = "{{ ruta_pago_nas }}"
            for r in p3.runs[1:]:
                r.text = ""
    doc.save(destino_path)


def asegurar_plantillas_dinamicas() -> dict:
    """
    Ejecuta la conversión de los 4 moldes maestros de plantillas_base/ hacia plantillas_dinamicas/
    si no existen o si se solicita reconstruirlos. Retorna un diccionario con las rutas generadas.
    """
    base_dir = settings.PLANTILLAS_BASE_DIR
    din_dir = settings.PLANTILLAS_DINAMICAS_DIR
    os.makedirs(din_dir, exist_ok=True)

    rutas = {
        "informe_principal": os.path.join(din_dir, "plantilla_informe_principal.docx"),
        "reporte_actividad_oti": os.path.join(din_dir, "plantilla_reporte_actividad_oti.docx"),
        "anexo_obligacion": os.path.join(din_dir, "plantilla_anexo_obligacion.docx"),
        "ruta_evidencias": os.path.join(din_dir, "plantilla_ruta_evidencias.docx"),
    }

    construir_plantilla_informe_principal(
        os.path.join(base_dir, "CTO_1031_JOHN_VEGA_INFORME_10.docx"),
        rutas["informe_principal"],
    )
    construir_plantilla_reporte_actividad_oti(
        os.path.join(base_dir, "PLANTILLA_REPORTE_INDIVIDUAL_ACTIVIDAD_OTI.docx"),
        rutas["reporte_actividad_oti"],
    )
    construir_plantilla_anexo_obligacion(
        os.path.join(base_dir, "PLANTILLA_ANEXO_EVIDENCIAS_OBLIGACION.docx"),
        rutas["anexo_obligacion"],
    )
    construir_plantilla_ruta_evidencias(
        os.path.join(base_dir, "Evidencia_Pago_10.docx"),
        rutas["ruta_evidencias"],
    )

    return rutas
