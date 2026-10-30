"""
Motor Constructor del Árbol Completo de Carpetas y los 2N+2 Documentos Word (`services/package_builder.py`).

La salida de cada pago mensual para un contratista con N obligaciones genera y empaqueta automáticamente
(en carpeta local y en un archivo `.zip` descargable) TODO el árbol completo idéntico a los Pagos 9 y 10:

PAQUETE_COMPLETO_PAGO_<XX>_CTO_<NUM>/
├── 1_DOCUMENTOS_CUENTA_DE_COBRO/
│   ├── CTO_<NUM>_<NOMBRE>_INFORME_<XX>.docx
│   └── Evidencia_Pago_<XX>.docx
└── 2_CARPETA_NAS_EVIDENCIAS/
    └── Pago <XX>/
        ├── Actividad_1/
        │   ├── REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx
        │   └── Anexos/
        │       ├── Obligacion 1.docx
        │       └── <archivos soporte originales de la obligación 1>
        ├── ...
        └── Actividad_N/
            ├── REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx
            └── Anexos/
                ├── Obligacion N.docx
                └── <archivos soporte originales de la obligación N>
"""

import os
import io
import re
import shutil
import zipfile
from copy import deepcopy
from typing import Dict, Any, List
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.table import _Cell
from docx.oxml import OxmlElement
from django.conf import settings

from .template_converter import asegurar_plantillas_dinamicas, _desenvolver_sdt_en_fila
from .file_inspector import sanitizar_texto_para_reporte


def _set_cell_text_formatted(
    cell,
    texto: str,
    font_name: str = "Arial",
    font_size_pt: float = 8.0,
    bold: bool = False,
    align=None,
):
    """
    Limpia una celda y escribe uno o múltiples párrafos (separados por \\n)
    conservando la tipografía oficial Arial y el tamaño exacto de la plantilla.
    """
    while len(cell.paragraphs) > 1:
        p_del = cell.paragraphs[-1]._element
        p_del.getparent().remove(p_del)

    lineas = str(texto or "").split("\n")
    p0 = cell.paragraphs[0]
    p0.clear()
    if align is not None:
        p0.alignment = align
    r0 = p0.add_run(lineas[0] if lineas else "")
    r0.font.name = font_name
    r0.font.size = Pt(font_size_pt)
    r0.bold = bold

    for linea in lineas[1:]:
        p_extra = cell.add_paragraph()
        if align is not None:
            p_extra.alignment = align
        r_extra = p_extra.add_run(linea)
        r_extra.font.name = font_name
        r_extra.font.size = Pt(font_size_pt)
        r_extra.bold = bold


def _insertar_parrafo_antes_de(p_referencia, texto="", bold=False, font_size_pt=11.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    """Inserta un párrafo nuevo justo antes de p_referencia con formato Arial."""
    nuevo_p = p_referencia.insert_paragraph_before("")
    nuevo_p.alignment = align
    if texto:
        r = nuevo_p.add_run(texto)
        r.font.name = "Arial"
        r.font.size = Pt(font_size_pt)
        r.bold = bold
    return nuevo_p


def _generar_captura_explorador_nas_automatica(ruta_destino_png: str, ruta_pago_nas: str, total_obligaciones: int):
    """
    Si el contratista aún no ha subido manualmente la captura del explorador de Windows de la NAS,
    genera automáticamente una imagen PNG nítida estilo Explorador de Archivos de Windows 11 mostrando
    la barra de dirección con la ruta NAS según Circular 0009 de 2024 y las carpetas Actividad_1 a Actividad_N.
    """
    ancho, alto = 1100, max(380, 160 + total_obligaciones * 34)
    img = Image.new("RGB", (ancho, alto), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)

    # Barra superior estilo Explorador de Archivos
    draw.rectangle([(0, 0), (ancho, 46)], fill=(243, 243, 243), outline=(218, 218, 218))
    draw.rectangle([(18, 10), (ancho - 18, 36)], fill=(255, 255, 255), outline=(190, 190, 190))
    draw.text((28, 16), f"Red > {ruta_pago_nas}", fill=(30, 30, 30))

    # Encabezado de columnas del explorador
    draw.rectangle([(0, 47), (ancho, 78)], fill=(250, 250, 250), outline=(230, 230, 230))
    draw.text((48, 56), "Nombre", fill=(70, 70, 70))
    draw.text((420, 56), "Fecha de modificación", fill=(70, 70, 70))
    draw.text((670, 56), "Tipo", fill=(70, 70, 70))
    draw.text((880, 56), "Estado Circular 0009/2024", fill=(70, 70, 70))

    for k in range(1, total_obligaciones + 1):
        y_top = 84 + (k - 1) * 30
        if k % 2 == 0:
            draw.rectangle([(10, y_top - 2), (ancho - 10, y_top + 24)], fill=(248, 250, 252))
        # Icono de carpeta amarillo
        draw.rectangle([(22, y_top + 3), (40, y_top + 18)], fill=(248, 203, 70), outline=(214, 163, 36))
        draw.rectangle([(22, y_top + 1), (30, y_top + 4)], fill=(248, 203, 70))
        draw.text((50, y_top + 4), f"Actividad_{k}", fill=(20, 20, 20))
        draw.text((420, y_top + 4), "28/10/2026 05:30 p. m.", fill=(90, 90, 90))
        draw.text((670, y_top + 4), "Carpeta de archivos", fill=(90, 90, 90))
        draw.text((880, y_top + 4), "REPORTE OTI + Anexos/", fill=(22, 101, 52))

    os.makedirs(os.path.dirname(ruta_destino_png), exist_ok=True)
    img.save(ruta_destino_png, "PNG")
    return ruta_destino_png


def generar_documento_informe_principal(informe, detalles_lista: List[Any], ruta_salida_docx: str):
    """
    Genera `1_DOCUMENTOS_CUENTA_DE_COBRO/CTO_<NUM>_<NOMBRE>_INFORME_<XX>.docx`
    clonando la estructura exacta de `CTO_1031_JOHN_VEGA_INFORME_10.docx` (Formato GC-FO-15 Versión 6):
    - Encabezado institucional UAEGRTD con logo y tabla GC-FO-15 intactos.
    - Tabla 0: Condiciones del Contrato (con el control SDT de Número de Pago desenvuelto y actualizado).
    - Tabla 1: Las N obligaciones con su número, cláusula, actividades ejecutadas sanitizadas (Arial 9pt) y ruta NAS exacta (Arial 8pt).
    - Tabla 2: Publicación SECOP II dinámica (informes 1..M).
    - Párrafo de Observaciones Adicionales, bloque de firma del contratista e imagen de firma opcional, y Tabla 3 del Supervisor.
    """
    molde_path = os.path.join(settings.PLANTILLAS_BASE_DIR, "CTO_1031_JOHN_VEGA_INFORME_10.docx")
    doc = Document(molde_path)
    c = informe.contratista

    # 1. Tabla 0: Condiciones del Contrato
    t0 = doc.tables[0]
    _desenvolver_sdt_en_fila(t0.rows[5]._tr)

    tcs_r0 = [_Cell(tc, t0) for tc in t0.rows[0]._tr.tc_lst]
    _set_cell_text_formatted(tcs_r0[1], c.nombres_apellidos, font_size_pt=8.0)
    _set_cell_text_formatted(tcs_r0[3], c.cedula, font_size_pt=8.0)

    tcs_r1 = [_Cell(tc, t0) for tc in t0.rows[1]._tr.tc_lst]
    _set_cell_text_formatted(tcs_r1[1], str(c.numero_contrato), font_size_pt=8.0)
    _set_cell_text_formatted(tcs_r1[3], str(c.anio_contrato), font_size_pt=8.0)

    f_ini = f"{c.fecha_inicio.month}/{c.fecha_inicio.day}/{c.fecha_inicio.year}"
    f_fin = f"{c.fecha_terminacion_actual.month}/{c.fecha_terminacion_actual.day}/{c.fecha_terminacion_actual.year}"
    tcs_r2 = [_Cell(tc, t0) for tc in t0.rows[2]._tr.tc_lst]
    _set_cell_text_formatted(tcs_r2[1], f_ini, font_size_pt=8.0)
    _set_cell_text_formatted(tcs_r2[3], f_fin, font_size_pt=8.0)

    tcs_r4 = [_Cell(tc, t0) for tc in t0.rows[4]._tr.tc_lst]
    _set_cell_text_formatted(
        tcs_r4[0],
        sanitizar_texto_para_reporte(c.objeto_contrato),
        font_size_pt=8.0,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    )

    tcs_r5 = [_Cell(tc, t0) for tc in t0.rows[5]._tr.tc_lst]
    if len(tcs_r5) >= 4:
        _set_cell_text_formatted(tcs_r5[1], str(informe.numero_pago), font_size_pt=8.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text_formatted(tcs_r5[3], c.nombre_supervisor, font_size_pt=8.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    tcs_r6 = [_Cell(tc, t0) for tc in t0.rows[6]._tr.tc_lst]
    _set_cell_text_formatted(tcs_r6[1], c.cargo_supervisor, font_size_pt=8.0)

    tcs_r7 = [_Cell(tc, t0) for tc in t0.rows[7]._tr.tc_lst]
    _set_cell_text_formatted(tcs_r7[1], c.direccion_oficina, font_size_pt=8.0)

    tcs_r8 = [_Cell(tc, t0) for tc in t0.rows[8]._tr.tc_lst]
    _set_cell_text_formatted(tcs_r8[1], c.direccion_territorial, font_size_pt=8.0)

    # 2. Tabla 1: Obligaciones Contractuales (1..N)
    t1 = doc.tables[1]
    fila_molde_tr = deepcopy(t1.rows[1]._tr)
    while len(t1.rows) > 1:
        t1._tbl.remove(t1.rows[-1]._tr)

    for det in detalles_lista:
        nuevo_tr = deepcopy(fila_molde_tr)
        t1._tbl.append(nuevo_tr)
        fila = t1.rows[-1]
        num_k = det.obligacion.numero_obligacion
        clausula_k = sanitizar_texto_para_reporte(det.obligacion.descripcion_clausula)
        actividades_k = sanitizar_texto_para_reporte(det.texto_tabla_principal)
        ruta_nas_k = det.referencia_digital_actividad

        _set_cell_text_formatted(
            fila.cells[0],
            str(num_k),
            font_size_pt=8.0,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )
        _set_cell_text_formatted(
            fila.cells[1],
            clausula_k,
            font_size_pt=8.0,
            align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        )
        _set_cell_text_formatted(
            fila.cells[2],
            actividades_k,
            font_size_pt=9.0,
            align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        )
        _set_cell_text_formatted(
            fila.cells[3],
            ruta_nas_k,
            font_size_pt=8.0,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )

    # 3. Tabla 2: Publicación SECOP II
    if len(doc.tables) >= 3:
        t2 = doc.tables[2]
        fila_secop_molde = deepcopy(t2.rows[1]._tr)
        while len(t2.rows) > 1:
            t2._tbl.remove(t2.rows[-1]._tr)
        total_secop = max(1, informe.informes_secop_publicados)
        for num_inf in range(1, total_secop + 1):
            tr_s = deepcopy(fila_secop_molde)
            t2._tbl.append(tr_s)
            fila_s = t2.rows[-1]
            _set_cell_text_formatted(fila_s.cells[0], str(num_inf), font_size_pt=8.0, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_formatted(fila_s.cells[1], "Sí", font_size_pt=8.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 4. Observaciones adicionales y firma del contratista
    obs_texto = (
        sanitizar_texto_para_reporte(informe.observaciones_adicionales)
        if informe.observaciones_adicionales.strip()
        else "< En este espacio el contratista podrá incluir observaciones relevantes sobre la ejecución contractual, avances significativos, limitaciones o situaciones que impacten el cumplimiento de las obligaciones pactadas.>"
    )
    for idx_p, p in enumerate(doc.paragraphs):
        if "En este espacio el contratista podrá incluir observaciones" in p.text:
            p.clear()
            r = p.add_run(obs_texto)
            r.font.name = "Arial"
            r.font.size = Pt(8.0)
        elif "______________________________________" in p.text and idx_p < 65:
            # Insertar firma digitalizada justo encima de la línea de firma del contratista si existe
            ruta_firma = c.obtener_ruta_firma_fisica()
            if ruta_firma and os.path.exists(ruta_firma):
                p_firma_img = p.insert_paragraph_before("")
                p_firma_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_img = p_firma_img.add_run()
                run_img.add_picture(ruta_firma, width=Inches(1.35))
        elif "John Felipe Vega Ot" in p.text:
            p.clear()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(c.nombres_apellidos)
            r.font.name = "Arial"
            r.font.size = Pt(8.0)
            r.bold = True

    # 5. Tabla 3: Supervisor
    if len(doc.tables) >= 4:
        t3 = doc.tables[3]
        _set_cell_text_formatted(t3.rows[0].cells[0], c.nombre_supervisor, font_size_pt=8.0, bold=True)
        _set_cell_text_formatted(t3.rows[1].cells[0], f"{c.cargo_supervisor}\nSUPERVISOR", font_size_pt=8.0, bold=False)

    os.makedirs(os.path.dirname(ruta_salida_docx), exist_ok=True)
    doc.save(ruta_salida_docx)
    return ruta_salida_docx


def generar_documento_evidencia_pago_nas(informe, total_obligaciones: int, ruta_salida_docx: str):
    """
    Genera `1_DOCUMENTOS_CUENTA_DE_COBRO/Evidencia_Pago_<XX>.docx` siguiendo el molde exacto de
    `plantillas_base/Evidencia_Pago_10.docx` (Circular 0009 de 2024):
    - P[0]: RUTA EVIDENCIAS CUENTAS DE COBRO <MES> - <AÑO>
    - P[1]: Imagen incrustada de la captura del explorador de archivos mostrando las carpetas Actividad_1..N
    - P[2]: RUTA SEGÚN CIRCULAR 0009 DE 2024 (ASEGURAMIENTO ACTIVOS INFORMACIÓN):
    - P[3]: \\\\10.10.36.17\\nivel_central\\oti\\4. sistemas de informacion\\<AÑO>\\01_Evidencias_OPS\\...\\PAGO <XX>
    """
    molde_path = os.path.join(settings.PLANTILLAS_BASE_DIR, "Evidencia_Pago_10.docx")
    doc = Document(molde_path)

    # Párrafo 0: Título con mes y año en mayúsculas
    p0 = doc.paragraphs[0]
    p0.clear()
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run(f"RUTA EVIDENCIAS CUENTAS DE COBRO {informe.nombre_mes_mayus} - {informe.anio}")
    r0.bold = True
    r0.font.name = "Calibri"
    r0.font.size = Pt(11)

    # Párrafo 1: Si el contratista subió una captura personalizada del explorador NAS, incrustarla;
    # de lo contrario, generar automáticamente la vista previa del explorador NAS con la ruta de este contratista y sus Actividad_1..N.
    p1 = doc.paragraphs[1]
    if informe.captura_explorador_nas and hasattr(informe.captura_explorador_nas, "path") and os.path.exists(informe.captura_explorador_nas.path):
        p1.clear()
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.add_run().add_picture(informe.captura_explorador_nas.path, width=Inches(6.14))
    else:
        captura_auto = os.path.join(os.path.dirname(ruta_salida_docx), f"_captura_nas_pago_{informe.numero_pago:02d}.png")
        _generar_captura_explorador_nas_automatica(captura_auto, informe.ruta_pago_nas, total_obligaciones)
        p1.clear()
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.add_run().add_picture(captura_auto, width=Inches(6.14))
        try:
            os.remove(captura_auto)
        except OSError:
            pass

    # Párrafo 2: Subtítulo Circular 0009 de 2024
    p2 = doc.paragraphs[2]
    p2.clear()
    r2 = p2.add_run("EXPEDIENTE DIGITAL DE EVIDENCIAS Y SOPORTES DE PAGO:")
    r2.font.name = "Calibri"
    r2.font.size = Pt(11)

    # Párrafo 3: Ruta NAS raíz del Pago XX
    p3 = doc.paragraphs[3]
    p3.clear()
    ref_digital = informe.enlace_nube_evidencias if informe.enlace_nube_evidencias else f"Expediente Digital de Evidencias / Pago {informe.numero_pago:02d} / Carpetas Actividad_1 a Actividad_{total_obligaciones}"
    r3 = p3.add_run(ref_digital)
    r3.font.name = "Calibri"
    r3.font.size = Pt(11)

    os.makedirs(os.path.dirname(ruta_salida_docx), exist_ok=True)
    doc.save(ruta_salida_docx)
    return ruta_salida_docx


def generar_reporte_individual_actividad_oti(detalle, ruta_salida_docx: str):
    """
    Genera `2_CARPETA_NAS_EVIDENCIAS/Pago <XX>/Actividad_<K>/REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx`
    siguiendo el formato exacto de `plantillas_base/PLANTILLA_REPORTE_INDIVIDUAL_ACTIVIDAD_OTI.docx`:
    - Tabla de Información General (Contratista, Contrato, Periodo reportado, Fecha del reporte).
    - Obligación Contractual K.
    - Descripción detallada de la ejecución de la obligación contractual (párrafos introductorios + viñetas `•`).
    - Resultados Obtenidos, Productos o entregables generados (viñetas `•`).
    - Evidencias y anexos (secciones `• Anexo K.x` y subítems `o archivo: detalle` citando `\\0. Evidencias\\Pago <XX>\\Actividad_<K>\\Anexos\\` y `Obligacion <K>.docx`).
    - Declaración de cumplimiento de la obligación N° K, Firma digitalizada (`firma.jpg`) incrustada, Nombre, C.C. y Fecha.
    """
    molde_path = os.path.join(settings.PLANTILLAS_BASE_DIR, "PLANTILLA_REPORTE_INDIVIDUAL_ACTIVIDAD_OTI.docx")
    doc = Document(molde_path)

    informe = detalle.informe
    c = informe.contratista
    ob = detalle.obligacion
    k = ob.numero_obligacion

    # 1. Tabla 0: Información General (4x2)
    t0 = doc.tables[0]
    _set_cell_text_formatted(t0.rows[0].cells[1], c.nombres_apellidos, font_size_pt=11.0)
    _set_cell_text_formatted(t0.rows[1].cells[1], str(c.numero_contrato), font_size_pt=11.0)
    _set_cell_text_formatted(t0.rows[2].cells[1], informe.nombre_mes, font_size_pt=11.0)
    _set_cell_text_formatted(t0.rows[3].cells[1], informe.fecha_reporte_texto_largo, font_size_pt=11.0)

    # Localizar párrafos ancla en la plantilla original
    p_obligacion = None
    p_desc_header = None
    p_resultados_header = None
    p_evidencias_header = None
    p_declaracion = None
    p_firma_label = None
    p_firma_img = None
    p_nombre_cc = None
    p_fecha_pie = None

    for idx, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if txt.startswith("Obligación Contractual:"):
            p_obligacion = p
        elif txt == "Descripción detallada de la ejecución de la obligación contractual":
            p_desc_header = p
        elif txt == "Resultados Obtenidos, Productos o entregables generados":
            p_resultados_header = p
        elif txt == "Evidencias y anexos":
            p_evidencias_header = p
        elif txt.startswith("La información registrada corresponde a las actividades ejecutadas"):
            p_declaracion = p
        elif txt.startswith("Firma:"):
            p_firma_label = p
            if idx + 1 < len(doc.paragraphs):
                p_firma_img = doc.paragraphs[idx + 1]
        elif "C.C." in txt or "VEGA" in txt.upper():
            p_nombre_cc = p
        elif txt.startswith("Fecha:"):
            p_fecha_pie = p

    # Actualizar párrafo de Obligación Contractual
    if p_obligacion is not None:
        p_obligacion.clear()
        p_obligacion.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_lbl = p_obligacion.add_run("Obligación Contractual: ")
        r_lbl.bold = True
        r_lbl.font.name = "Arial"
        r_lbl.font.size = Pt(11.0)
        r_txt = p_obligacion.add_run(ob.descripcion_clausula)
        r_txt.bold = False
        r_txt.font.name = "Arial"
        r_txt.font.size = Pt(11.0)

    # Eliminar los párrafos de ejemplo del molde entre las secciones dinámicas comparando el nodo XML (_p)
    parrafos_act = list(doc.paragraphs)
    elems_act = [p._p for p in parrafos_act]
    idx_desc = elems_act.index(p_desc_header._p)
    idx_res = elems_act.index(p_resultados_header._p)
    idx_evi = elems_act.index(p_evidencias_header._p)
    idx_dec = elems_act.index(p_declaracion._p)

    for p_borrar in parrafos_act[idx_desc + 1 : idx_res]:
        p_borrar._element.getparent().remove(p_borrar._element)
    for p_borrar in parrafos_act[idx_res + 1 : idx_evi]:
        p_borrar._element.getparent().remove(p_borrar._element)
    for p_borrar in parrafos_act[idx_evi + 1 : idx_dec]:
        p_borrar._element.getparent().remove(p_borrar._element)

    datos_oti = detalle.desarrollo_detallado_oti or {}
    if isinstance(datos_oti, str):
        try:
            datos_oti = json.loads(datos_oti)
        except Exception:
            datos_oti = {
                "parrafos_introductorios": [datos_oti],
                "vinetas_desarrollo": [],
                "resultados_entregables": [],
                "anexos_secciones": [],
            }
    if not isinstance(datos_oti, dict):
        datos_oti = {}

    parrafos_intro = datos_oti.get("parrafos_introductorios", [])
    vinetas_desarrollo = datos_oti.get("vinetas_desarrollo", [])
    resultados_entregables = datos_oti.get("resultados_entregables", [])
    anexos_secciones = datos_oti.get("anexos_secciones", [])

    # Insertar Desarrollo Detallado antes de p_resultados_header
    for intro in parrafos_intro:
        _insertar_parrafo_antes_de(
            p_resultados_header,
            sanitizar_texto_para_reporte(intro),
            bold=False,
            font_size_pt=11.0,
        )

    for vin in vinetas_desarrollo:
        p_v = p_resultados_header.insert_paragraph_before("")
        p_v.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        titulo_v = sanitizar_texto_para_reporte(vin.get("titulo", "")).strip()
        desc_v = sanitizar_texto_para_reporte(vin.get("descripcion", "")).strip()
        if titulo_v:
            r_b = p_v.add_run(f"• {titulo_v}: ")
            r_b.bold = True
            r_b.font.name = "Arial"
            r_b.font.size = Pt(11.0)
        r_d = p_v.add_run(desc_v if titulo_v else f"• {desc_v}")
        r_d.bold = False
        r_d.font.name = "Arial"
        r_d.font.size = Pt(11.0)

    # Insertar Resultados Obtenidos antes de p_evidencias_header
    for res in resultados_entregables:
        p_r = p_evidencias_header.insert_paragraph_before("")
        p_r.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        titulo_r = sanitizar_texto_para_reporte(res.get("titulo", "")).strip()
        desc_r = sanitizar_texto_para_reporte(res.get("descripcion", "")).strip()
        if titulo_r:
            r_b = p_r.add_run(f"• {titulo_r}: ")
            r_b.bold = True
            r_b.font.name = "Arial"
            r_b.font.size = Pt(11.0)
        r_d = p_r.add_run(desc_r if titulo_r else f"• {desc_r}")
        r_d.bold = False
        r_d.font.name = "Arial"
        r_d.font.size = Pt(11.0)

    # Insertar Evidencias y Anexos antes de p_declaracion
    for sec in anexos_secciones:
        p_s = p_declaracion.insert_paragraph_before("")
        p_s.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        titulo_s = sanitizar_texto_para_reporte(sec.get("titulo", f"Anexo {k}.1 — Soportes de Ejecución")).strip()
        desc_s = sanitizar_texto_para_reporte(sec.get("descripcion", "")).strip()
        r_sb = p_s.add_run(f"• {titulo_s}: ")
        r_sb.bold = True
        r_sb.font.name = "Arial"
        r_sb.font.size = Pt(11.0)
        r_sd = p_s.add_run(desc_s)
        r_sd.bold = False
        r_sd.font.name = "Arial"
        r_sd.font.size = Pt(11.0)

        for sub in sec.get("subitems", []):
            p_sub = p_declaracion.insert_paragraph_before("")
            p_sub.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            etiq = sanitizar_texto_para_reporte(sub.get("etiqueta", "")).strip()
            det_sub = sanitizar_texto_para_reporte(sub.get("detalle", "")).strip()
            if etiq:
                r_eb = p_sub.add_run(f"o {etiq}: ")
                r_eb.bold = True
                r_eb.font.name = "Arial"
                r_eb.font.size = Pt(11.0)
            r_ed = p_sub.add_run(det_sub if etiq else f"o {det_sub}")
            r_ed.bold = False
            r_ed.font.name = "Arial"
            r_ed.font.size = Pt(11.0)

    # Actualizar declaración final con el número de obligación K y número de contrato
    if p_declaracion is not None:
        p_declaracion.clear()
        p_declaracion.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_dec = p_declaracion.add_run(
            f"La información registrada corresponde a las actividades ejecutadas en cumplimiento de la obligación N° {k} "
            f"del contrato de prestación de servicios número {c.numero_contrato} y que las evidencias anexas soportan la ejecución reportada."
        )
        r_dec.font.name = "Arial"
        r_dec.font.size = Pt(11.0)

    # Limpiar SIEMPRE la imagen de firma que venía en el molde base (para que jamás aparezca la firma de prueba en otro contratista)
    # e incrustar únicamente la firma propia subida por el contratista actual (si la tiene configurada en su perfil)
    ruta_firma = c.obtener_ruta_firma_fisica()
    if p_firma_img is not None:
        p_firma_img.clear()
        if ruta_firma and os.path.exists(ruta_firma):
            r_img = p_firma_img.add_run()
            r_img.add_picture(ruta_firma, width=Inches(1.45))

    # Actualizar Nombre, C.C. y Fecha al pie
    if p_nombre_cc is not None:
        p_nombre_cc.clear()
        r_nom = p_nombre_cc.add_run(f"{c.nombres_apellidos.upper()}\nC.C. {c.cedula_formateada}")
        r_nom.font.name = "Arial"
        r_nom.font.size = Pt(11.0)

    if p_fecha_pie is not None:
        p_fecha_pie.clear()
        r_flbl = p_fecha_pie.add_run("Fecha: ")
        r_flbl.bold = True
        r_flbl.font.name = "Arial"
        r_flbl.font.size = Pt(11.0)
        r_fval = p_fecha_pie.add_run(informe.fecha_reporte_texto_largo)
        r_fval.bold = False
        r_fval.font.name = "Arial"
        r_fval.font.size = Pt(11.0)

    os.makedirs(os.path.dirname(ruta_salida_docx), exist_ok=True)
    doc.save(ruta_salida_docx)
    return ruta_salida_docx


def _preparar_imagen_optimizada_para_word(ruta_imagen: str) -> io.BytesIO:
    """Comprime y redimensiona en memoria la captura antes de incrustarla en el .docx para que el Word sea ultraligero."""
    with Image.open(ruta_imagen) as img:
        img_rgb = img.convert("RGB")
        if img_rgb.width > 1280:
            ratio = 1280 / float(img_rgb.width)
            nuevo_alto = int(float(img_rgb.height) * ratio)
            img_rgb = img_rgb.resize((1280, nuevo_alto), Image.Resampling.LANCZOS)
        buffer = io.BytesIO()
        img_rgb.save(buffer, format="JPEG", quality=85, optimize=True)
        buffer.seek(0)
        return buffer


def generar_anexo_evidencias_obligacion(detalle, carpeta_anexos_destino: str, ruta_salida_docx: str):
    """
    Genera `2_CARPETA_NAS_EVIDENCIAS/Pago <XX>/Actividad_<K>/Anexos/Obligacion <K>.docx`
    siguiendo el formato exacto de `plantillas_base/PLANTILLA_ANEXO_EVIDENCIAS_OBLIGACION.docx`:
    - P[0]: EVIDENCIAS Y ANEXOS TÉCNICOS — OBLIGACIÓN <K> (<MES> <AÑO>) (Arial Bold, Centrado)
    - P[1]: Contrato N° <NUM> de <AÑO> | Contratista: <Nombre> | <Proyecto/Área> (Arial 10pt, Centrado)
    - Por cada evidencia/bloque numerado (1..M):
      - Título técnico numerado en Arial 11pt Negrita Justificado.
      - Párrafo explicativo formal en Arial 10pt Justificado.
      - Incrustación automática de la imagen (`add_picture(width=Inches(6.1))`) si el archivo es `.png`/`.jpg`,
        o recuadro/referencia técnica formal al archivo físico (Excel, PDF, Word, Script) copiado en esa misma carpeta `Anexos/`.
    """
    molde_path = os.path.join(settings.PLANTILLAS_BASE_DIR, "PLANTILLA_ANEXO_EVIDENCIAS_OBLIGACION.docx")
    doc = Document(molde_path)

    informe = detalle.informe
    c = informe.contratista
    ob = detalle.obligacion
    k = ob.numero_obligacion

    # Conservar únicamente los 2 primeros párrafos de encabezado del molde y limpiar el resto
    while len(doc.paragraphs) > 2:
        p_del = doc.paragraphs[-1]._element
        p_del.getparent().remove(p_del)

    p0 = doc.paragraphs[0]
    p0.clear()
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run(f"EVIDENCIAS Y ANEXOS TÉCNICOS — OBLIGACIÓN {k} ({informe.nombre_mes_mayus} {informe.anio})")
    r0.bold = True
    r0.font.name = "Arial"
    r0.font.size = Pt(11.0)

    p1 = doc.paragraphs[1]
    p1.clear()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(f"Contrato N° {c.numero_contrato} de {c.anio_contrato} | Contratista: {c.nombres_apellidos} | {c.proyecto_area}")
    r1.bold = False
    r1.font.name = "Arial"
    r1.font.size = Pt(10.0)

    # Espacio después del encabezado
    doc.add_paragraph("")

    # Mapear archivos soporte subidos en esta obligación y copiarlos físicamente a `Actividad_<K>/Anexos/`
    os.makedirs(carpeta_anexos_destino, exist_ok=True)
    evidencias_db = list(detalle.archivos_evidencia.all())
    mapa_archivos_copiados: Dict[str, str] = {}

    for ev in evidencias_db:
        if ev.archivo and hasattr(ev.archivo, "path") and os.path.exists(ev.archivo.path):
            nombre_base, ext_base = os.path.splitext(ev.nombre_original)
            # Limpiar caracteres prohibidos en Windows y acotar a 60 caracteres para evitar rebasar MAX_PATH (260)
            nombre_limpio = re.sub(r'[<>:"/\\|?*\x00-\x1f\ufffd]', '_', nombre_base).strip()[:60] + ext_base
            destino_copia = os.path.join(carpeta_anexos_destino, nombre_limpio)
            src_path = ev.archivo.path

            if os.name == "nt" and len(os.path.abspath(destino_copia)) >= 230:
                dst_call = "\\\\?\\" + os.path.abspath(destino_copia)
                src_call = "\\\\?\\" + os.path.abspath(src_path) if not src_path.startswith("\\\\?\\") else src_path
            else:
                dst_call = destino_copia
                src_call = src_path

            if os.path.abspath(src_path) != os.path.abspath(destino_copia):
                shutil.copy2(src_call, dst_call)
            mapa_archivos_copiados[ev.nombre_original] = destino_copia
            mapa_archivos_copiados[nombre_limpio] = destino_copia

    bloques = list(detalle.bloques_evidencias_json or [])
    if not bloques and evidencias_db:
        for idx_ev, ev in enumerate(evidencias_db, start=1):
            bloques.append({
                "numero": idx_ev,
                "titulo": ev.titulo_evidencia or f"Soporte Técnico ({ev.nombre_original})",
                "descripcion": ev.descripcion_tecnica_ia or f"Evidencia documental correspondiente al archivo {ev.nombre_original}.",
                "archivo_asociado": ev.nombre_original,
            })

    archivos_ya_incrustados = set()

    for idx_b, bloque in enumerate(bloques, start=1):
        num_sec = bloque.get("numero", idx_b)
        titulo_sec = str(bloque.get("titulo", f"Evidencia Técnica {num_sec}")).strip()
        # Evitar doble numeración si el título ya empieza por "1."
        if not titulo_sec.startswith(f"{num_sec}."):
            titulo_formateado = f"{num_sec}. {titulo_sec}"
        else:
            titulo_formateado = titulo_sec

        desc_sec = str(bloque.get("descripcion", "")).strip()
        archivo_asoc = str(bloque.get("archivo_asociado", "")).strip()

        # Párrafo de Título Numerado (Arial 11pt Negrita)
        p_tit = doc.add_paragraph()
        p_tit.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_tit = p_tit.add_run(sanitizar_texto_para_reporte(titulo_formateado))
        r_tit.bold = True
        r_tit.font.name = "Arial"
        r_tit.font.size = Pt(11.0)

        # Párrafo de Explicación Técnica (Arial 10pt)
        p_desc = doc.add_paragraph()
        p_desc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_desc = p_desc.add_run(sanitizar_texto_para_reporte(desc_sec))
        r_desc.bold = False
        r_desc.font.name = "Arial"
        r_desc.font.size = Pt(10.0)

        # Si hay un archivo imagen asociado (o si el bloque permite incrustar múltiples capturas como en el molde base)
        archivos_candidatos = []
        if archivo_asoc and archivo_asoc in mapa_archivos_copiados:
            archivos_candidatos.append(archivo_asoc)
        elif idx_b <= len(evidencias_db):
            ev_cand = evidencias_db[idx_b - 1]
            if ev_cand.nombre_original in mapa_archivos_copiados and ev_cand.nombre_original not in archivos_ya_incrustados:
                archivos_candidatos.append(ev_cand.nombre_original)

        for nombre_arch in archivos_candidatos:
            ruta_fisica = mapa_archivos_copiados[nombre_arch]
            archivos_ya_incrustados.add(nombre_arch)
            ext = os.path.splitext(nombre_arch.lower())[1]
            if ext in (".png", ".jpg", ".jpeg", ".bmp", ".webp"):
                try:
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_img = p_img.add_run()
                    r_img.add_picture(_preparar_imagen_optimizada_para_word(ruta_fisica), width=Inches(6.11))
                except Exception:
                    # Si la imagen subida por el usuario estaba corrupta/dañada, no romper el Word: dejar referencia formal
                    p_ref = doc.add_paragraph()
                    p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    r_ref = p_ref.add_run(f"[Soporte Gráfico Adjunto en Carpeta Anexos/: {nombre_arch}]")
                    r_ref.italic = True
                    r_ref.font.name = "Arial"
                    r_ref.font.size = Pt(9.5)
            else:
                p_ref = doc.add_paragraph()
                p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                r_ref = p_ref.add_run(f"[Soporte Digital Adjunto en Carpeta Anexos/: {nombre_arch}]")
                r_ref.italic = True
                r_ref.font.name = "Arial"
                r_ref.font.size = Pt(9.5)

        doc.add_paragraph("")

    # Si quedaron imágenes subidas en la obligación que no fueron referenciadas individualmente en `bloques`,
    # incrustarlas también para garantizar que ningún soporte gráfico quede por fuera de `Obligacion <K>.docx`
    for ev in evidencias_db:
        if ev.nombre_original in mapa_archivos_copiados and ev.nombre_original not in archivos_ya_incrustados:
            ruta_fisica = mapa_archivos_copiados[ev.nombre_original]
            archivos_ya_incrustados.add(ev.nombre_original)
            if ev.es_imagen:
                try:
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(_preparar_imagen_optimizada_para_word(ruta_fisica), width=Inches(6.11))
                except Exception:
                    # Si la imagen subida por el usuario estaba corrupta/dañada, no romper el Word: dejar referencia formal
                    p_ref = doc.add_paragraph()
                    p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    r_ref = p_ref.add_run(f"[Soporte Gráfico Adjunto en Carpeta Anexos/: {ev.nombre_original}]")
                    r_ref.italic = True
                    r_ref.font.name = "Arial"
                    r_ref.font.size = Pt(9.5)
            else:
                p_ref = doc.add_paragraph()
                p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                r_ref = p_ref.add_run(
                    f"[Soporte Digital Adjunto en Carpeta Anexos/: {ev.nombre_original}]"
                )
                r_ref.italic = True
                r_ref.font.name = "Arial"
                r_ref.font.size = Pt(9.5)

    os.makedirs(os.path.dirname(ruta_salida_docx), exist_ok=True)
    doc.save(ruta_salida_docx)
    return ruta_salida_docx


def construir_paquete_completo_pago(informe) -> Dict[str, Any]:
    """
    Construye el árbol completo de carpetas y los 2N+2 documentos Word para un `InformeMensual`,
    empaquetándolos tanto en carpeta local como en un archivo `.zip` listo para descarga directa:

    PAQUETE_COMPLETO_PAGO_<XX>_CTO_<NUM>/
    ├── 1_DOCUMENTOS_CUENTA_DE_COBRO/
    │   ├── CTO_<NUM>_<NOMBRE>_INFORME_<XX>.docx
    │   └── Evidencia_Pago_<XX>.docx
    └── 2_CARPETA_NAS_EVIDENCIAS/
        └── Pago <XX>/
            ├── Actividad_1/
            │   ├── REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx
            │   └── Anexos/
            │       ├── Obligacion 1.docx
            │       └── <archivos soporte originales>
            ...
            └── Actividad_N/
                ├── REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx
                └── Anexos/
                    ├── Obligacion N.docx
                    └── <archivos soporte originales>
    """
    # asegurar_plantillas_dinamicas()

    c = informe.contratista
    num_pago = informe.numero_pago
    num_cto = c.numero_contrato

    detalles = list(
        informe.detalles_obligaciones.select_related("obligacion")
        .prefetch_related("archivos_evidencia")
        .order_by("obligacion__numero_obligacion")
    )
    total_obligaciones = len(detalles)

    nombre_raiz = f"PAQUETE_COMPLETO_PAGO_{num_pago:02d}_CTO_{num_cto}"
    carpeta_raiz = os.path.join(settings.SALIDA_PAQUETES_DIR, nombre_raiz)

    # Limpiar construcción previa si existía para asegurar árbol fresco
    if os.path.exists(carpeta_raiz):
        shutil.rmtree(carpeta_raiz)

    dir_1_cuenta_cobro = os.path.join(carpeta_raiz, "1_DOCUMENTOS_CUENTA_DE_COBRO")
    dir_2_nas_pago = os.path.join(carpeta_raiz, "2_CARPETA_NAS_EVIDENCIAS", f"Pago {num_pago:02d}")
    os.makedirs(dir_1_cuenta_cobro, exist_ok=True)
    os.makedirs(dir_2_nas_pago, exist_ok=True)

    documentos_generados = []

    # a.1) Generar Informe Principal Consolidado: CTO_<NUM>_<NOMBRE>_INFORME_<XX>.docx
    ruta_informe_principal = os.path.join(dir_1_cuenta_cobro, informe.nombre_archivo_informe_principal)
    generar_documento_informe_principal(informe, detalles, ruta_informe_principal)
    documentos_generados.append(ruta_informe_principal)

    # a.2) Generar Documento de Ruta NAS según Circular 0009 de 2024: Evidencia_Pago_<XX>.docx
    ruta_evidencia_pago = os.path.join(dir_1_cuenta_cobro, informe.nombre_archivo_evidencia_pago)
    generar_documento_evidencia_pago_nas(informe, total_obligaciones, ruta_evidencia_pago)
    documentos_generados.append(ruta_evidencia_pago)

    # b) Por cada obligación K (de 1 a N), construir Actividad_<K>/ y Actividad_<K>/Anexos/
    carpetas_actividades = []
    for det in detalles:
        k = det.obligacion.numero_obligacion
        dir_actividad_k = os.path.join(dir_2_nas_pago, f"Actividad_{k}")
        dir_anexos_k = os.path.join(dir_actividad_k, "Anexos")
        os.makedirs(dir_anexos_k, exist_ok=True)
        carpetas_actividades.append(dir_actividad_k)

        # b.1) Actividad_<K>/REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx
        ruta_reporte_oti_k = os.path.join(dir_actividad_k, "REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx")
        generar_reporte_individual_actividad_oti(det, ruta_reporte_oti_k)
        documentos_generados.append(ruta_reporte_oti_k)

        # b.2) Actividad_<K>/Anexos/Obligacion <K>.docx (+ copia de todos los archivos soporte originales)
        ruta_obligacion_k_docx = os.path.join(dir_anexos_k, f"Obligacion {k}.docx")
        generar_anexo_evidencias_obligacion(det, dir_anexos_k, ruta_obligacion_k_docx)
        documentos_generados.append(ruta_obligacion_k_docx)

    # Empaquetar todo el árbol en un archivo .ZIP descargable
    ruta_zip = os.path.join(settings.SALIDA_PAQUETES_DIR, f"{nombre_raiz}.zip")
    with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(carpeta_raiz):
            # Asegurar que incluso carpetas vacías queden registradas en el ZIP
            for d in dirs:
                dir_full = os.path.join(root, d)
                rel_dir = os.path.relpath(dir_full, os.path.dirname(carpeta_raiz)).replace("\\", "/") + "/"
                zf.writestr(rel_dir, "")
            for file in files:
                file_full = os.path.join(root, file)
                rel_file = os.path.relpath(file_full, os.path.dirname(carpeta_raiz)).replace("\\", "/")
                zf.write(file_full, rel_file)

    informe.ruta_carpeta_paquete = carpeta_raiz
    informe.archivo_zip_generado = ruta_zip
    informe.estado = "EMPAQUETADO"
    informe.save()

    return {
        "nombre_raiz": nombre_raiz,
        "carpeta_raiz": carpeta_raiz,
        "ruta_zip": ruta_zip,
        "total_obligaciones": total_obligaciones,
        "total_documentos_word": len(documentos_generados),
        "documentos_generados": documentos_generados,
        "carpetas_actividades": carpetas_actividades,
    }
