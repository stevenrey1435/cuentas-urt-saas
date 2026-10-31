"""
Conector IA Dual con Ahorro Extremo de Tokens (Single-Shot JSON) y Cola Secuencial FIFO.
Integración directa con Ollama Local (HP OMEN 14 RTX 4070) via Cloudflare Tunnel.
"""

import os
import hashlib
import json
import re
import threading
import logging
from typing import Dict, Any, List, Optional
from django.conf import settings
from .file_inspector import sanitizar_texto_para_reporte
from .semantic_matcher import puntuar_obligaciones_semanticamente

logger = logging.getLogger(__name__)

_HERMES_FIFO_LOCK = threading.Lock()


def _normalizar_espacios(texto: str) -> str:
    return re.sub(r"\s+", " ", (texto or "")).strip()


def clasificar_obligacion_por_afinidad(
    texto_item: str,
    detalles_lista: List[Any],
    indice_rotacion: int = 0,
):
    """
    Distribución estricta y balanceada para que todas las obligaciones
    contractuales reciban soporte y ninguna acapare más de 2 archivos.
    """
    if not detalles_lista:
        return None

    mapa_obs = {det.obligacion.numero_obligacion: det for det in detalles_lista}
    txt = _normalizar_espacios(texto_item).lower()

    # 1. Mapeo estricto por nombres y palabras clave inequívocas
    if any(k in txt for k in ("das_hld", "diseño de software", "diseno de software", "entrega documento de diseño")):
        if 1 in mapa_obs:
            return mapa_obs[1]

    if any(k in txt for k in ("seguridad y control", "it-sigaf", "control_de_evidencias")):
        if 3 in mapa_obs:
            return mapa_obs[3]

    if any(k in txt for k in ("cronograma_sigaf", "plant120", "cronograma", "gantt")):
        if 6 in mapa_obs:
            return mapa_obs[6]

    if any(k in txt for k in ("cooperapp", "solicitud de asignación de permisos", "permisos de administrador")):
        if 9 in mapa_obs:
            return mapa_obs[9]

    # 2. Distribución circular equilibrada para capturas e imágenes genéricas
    # Asegura cubrir obligaciones clave: 2 (Mantenimiento), 4 (Soporte), 5 (Socialización),
    # 7 (Git), 8 (Seguimiento), 10 (Reuniones), 11 (Otras)
    conteo_actual = {det.id: det.archivos_evidencia.count() for det in detalles_lista}

    orden_preferente = [7, 8, 5, 2, 4, 10, 11, 7, 8, 5, 2, 4, 10, 11]
    ob_candidata = orden_preferente[indice_rotacion % len(orden_preferente)]

    if ob_candidata in mapa_obs and conteo_actual.get(mapa_obs[ob_candidata].id, 0) < 2:
        return mapa_obs[ob_candidata]

    # Si la preferida ya tiene cupo lleno, asignar a la que tenga menor cantidad de archivos
    detalles_ordenados = sorted(
        detalles_lista,
        key=lambda d: (conteo_actual.get(d.id, 0), d.obligacion.numero_obligacion),
    )
    return detalles_ordenados[0]


def _limpiar_fragmento_significativo(lineas: List[str], nombre_archivo: str, nombre_contratista: str = "") -> str:
    base = os.path.splitext(nombre_archivo)[0]
    variantes_nombre = []
    if nombre_contratista:
        partes = [p for p in re.split(r"\s+", nombre_contratista.strip()) if len(p) > 2]
        if len(partes) >= 2:
            variantes_nombre.append(r"[\s_\-]*" + r"[\s_\-]+".join(partes[:2]) + r".*$")
    variantes_nombre.append(r"[\s_\-]*jes[uú]s[\s_\-]+steven[\s_\-]+rey.*$")

    if re.search(r"(^re[\s_]|^rv[\s_]|^fw[\s_]|outlook|outlo)", base.lower()):
        asunto = re.sub(r"^(re|rv|fw|fwd)[\s_:\-]+", "", base, flags=re.I)
        for pat in variantes_nombre:
            asunto = re.sub(pat, "", asunto, flags=re.I)
        asunto = re.sub(r"[\s_\-]*(outlook|outlo).*$", "", asunto, flags=re.I)
        asunto = re.sub(r"[_\-]+", " ", asunto).strip(" -–—.,;:")
        if len(asunto) > 8:
            return sanitizar_texto_para_reporte(asunto[:90])

    if any(k in base.lower() for k in ("cronograma", "plant120", "plan_trabajo", "gantt")):
        nom_cron = re.sub(r"(\d{1,2})(octubre|noviembre|diciembre|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre)(\d{4})", r" \1 de \2 de \3", base, flags=re.I)
        nom_cron = re.sub(r"[_\-]+", " ", nom_cron)
        nom_cron = re.sub(r"\s+", " ", nom_cron).strip(" -–—.,;:")
        return sanitizar_texto_para_reporte(nom_cron[:90])

    if not re.match(r"^(imagen|image|captura|screenshot|img|paste|foto|whatsapp)", base.lower()):
        base_limpia = base
        for pat in variantes_nombre:
            base_limpia = re.sub(pat, "", base_limpia, flags=re.I)
        base_limpia = re.sub(r"[_\-]+", " ", base_limpia)
        base_limpia = re.sub(r"\s+", " ", base_limpia).strip(" -–—.,;:")
        if len(base_limpia) > 6:
            return sanitizar_texto_para_reporte(base_limpia[:90])

    return ""


def _resumir_contenido_observado_limpio(meta: dict, nombre_archivo: str) -> str:
    base_low = nombre_archivo.lower()
    if "das_hld" in base_low or "arquitectura" in base_low:
        return (
            "Documento de Arquitectura de Software (DAS / HLD) para la herramienta virtual de monitoreo "
            "y seguimiento de órdenes T-120, con definición de alcance funcional, vistas arquitectónicas y componentes"
        )
    if "it-sigaf" in base_low or "seguridad y control" in base_low or "seguridad_y_control" in base_low:
        return (
            "Documento técnico-ejecutivo IT-SIGAF sobre control, inmutabilidad criptográfica (SHA-256), "
            "ciclo de vida y custodia segura de evidencias y adjuntos en cuentas de cobro"
        )
    if "cronograma_sigaf" in base_low:
        return (
            "Matriz de programación y seguimiento de hitos del proyecto SIGAF (Alcance 3), incluyendo "
            "fases de planeación, análisis de requerimientos, control de tiempos y porcentajes de avance"
        )
    if "plant120" in base_low:
        return (
            "Plan de trabajo y programación de actividades del sistema T-120 (Alcance 2), estructurando "
            "entregables metodológicos, levantamiento de requerimientos y seguimiento de ejecución"
        )
    if "cooperapp" in base_low or "permisos" in base_low:
        return (
            "Gestión y confirmación técnica de asignación de permisos de administración y acceso "
            "a funcionalidades de seguimiento de reportes en la plataforma institucional CooperApp"
        )
    if "entrega documento de diseño" in base_low or "diseno de software" in base_low:
        return (
            "Comunicación técnica oficial de entrega y socialización del Documento de Diseño de Software "
            "para el Sistema de Seguimiento T-120 ante el equipo técnico y funcional"
        )

    lineas = list(meta.get("resumen_lineas") or [])
    for clave in ("texto_pdf", "texto_ocr"):
        val = (meta.get(clave) or "").strip()
        if val:
            lineas.append(val)

    fragmentos_limpios = []
    for ln in lineas:
        s = str(ln)
        s = re.sub(r"[‹<]?[\w\.\-]+@[\w\.\-]+\.\w+[›>]?", "", s)
        s = re.sub(r"\.{3,}", " ", s)
        s = re.sub(r"[\ue000-\uf8ff]", "", s)
        if any(r in s.lower() for r in (
            "enviado:", "para:", "cc:", "asunto:", "cordial saludo",
            "nombre de tarea duración", "predecesoras", "página 1 de", "página 2 de"
        )):
            continue
        s = _normalizar_espacios(s).strip(" -•*:.,;|/\\[]")
        if len(s) >= 20:
            fragmentos_limpios.append(s[:160])
        if len(fragmentos_limpios) >= 2:
            break

    if fragmentos_limpios:
        return "; ".join(fragmentos_limpios)[:260]
    return ""


def _inferir_contexto_clausula(
    clausula: str,
    num_ob: int,
    objeto_contrato: str = "",
) -> Dict[str, str]:
    txt = _normalizar_espacios(clausula)
    low = txt.lower()
    ctx_global = f"{low} {_normalizar_espacios(objeto_contrato).lower()}"
    es_invalida = len(txt) < 18 or len(txt.split()) < 3

    if es_invalida:
        return {
            "es_invalida": True,
            "tipo_profesion": "gestión profesional y técnica",
            "tema_corto": f"Ejecución de la Obligación N° {num_ob}",
            "enfoque": "",
            "accion_cierre": "el cumplimiento de los compromisos asignados en el periodo",
        }

    if any(k in ctx_global for k in (
        "software", "sistemas de información", "código fuente", "git",
        "base de datos", "arquitectura", "tecnologías de la información"
    )):
        tipo_profesion = "ingeniería de software y gestión tecnológica"
    else:
        tipo_profesion = "gestión profesional y técnica"

    if any(k in low for k in ("desarrollar e integrar", "arquitectura y las etapas del ciclo de vida", "escalabilidad y mantenibilidad")):
        tema = "Desarrollo, Arquitectura de Componentes e Integración de Sistemas de Información"
        accion = "el diseño modular, estructuración arquitectónica confiable e interoperabilidad de los componentes de software institucionales"
    elif any(k in low for k in ("mantenimiento correctivo", "mejora continua", "actualización de las soluciones")):
        tema = "Mantenimiento Correctivo, Evolutivo y Mejora Continua de Soluciones de Software"
        accion = "la estabilización funcional, remediación oportuna de defectos y actualización evolutiva de los aplicativos de la entidad"
    elif any(k in low for k in ("implementación, parametrización y fortalecimiento", "parametrización", "fortalecimiento técnico")):
        tema = "Parametrización, Implementación y Fortalecimiento Técnico de Sistemas de Información"
        accion = "el fortalecimiento técnico, configuración segura, control de integridad y optimización de los componentes tecnológicos de la OTI"
    elif any(k in low for k in ("asistencia técnica oportuna", "atención ágil de fallos", "continuidad operativa")):
        tema = "Asistencia Técnica Oportuna, Atención Ágil de Fallos y Continuidad Operativa"
        accion = "la atención inmediata de contingencias, aplicación efectiva de soluciones y disponibilidad continua de los sistemas de información"
    elif any(k in low for k in ("socialización y transferencia de conocimiento", "transferencia de conocimiento", "uso y apropiación")):
        tema = "Socialización Técnica, Transferencia de Conocimiento y Acompañamiento Funcional"
        accion = "la transferencia efectiva de conocimiento, apropiación tecnológica y acompañamiento técnico a los equipos y usuarios finales"
    elif any(k in low for k in ("construir cronogramas", "estimar esfuerzos", "actividades planificadas", "compilar informes")):
        tema = "Planeación Técnica, Estimación de Esfuerzo y Seguimiento de Cronogramas de Desarrollo"
        accion = "el control riguroso de hitos, estimación técnica de esfuerzo y consolidación de informes detallados sobre las tareas ejecutadas"
    elif any(k in low for k in ("repositorio git", "ramas asignadas", "documentación estructurada del código", "mezclas")):
        tema = "Control de Versiones en Repositorio Git, Gestión de Ramas y Documentación de Código Fuente"
        accion = "la actualización continua en las ramas asignadas del repositorio Git institucional, ejecución controlada de mezclas (merges) y documentación estructurada del código"
    elif any(k in low for k in ("seguimientos programados", "herramienta de seguimiento", "actualizar las actividades")):
        tema = "Seguimiento de Proyectos Asignados y Actualización en la Herramienta de Control"
        accion = "la participación activa en los seguimientos programados y la actualización continua del estado de las actividades en la herramienta oficial"
    elif any(k in low for k in ("mesa de servicios", "correo electrónico", "requerimientos, solicitudes e incidencias", "comportamientos anómalos")):
        tema = "Gestión Oportuna de Requerimientos, Solicitudes e Incidencias en Mesa de Servicios y Correo"
        accion = "el análisis técnico oportuno y la resolución efectiva de requerimientos, solicitudes de permisos e incidencias reportadas por mesa de servicios o correo institucional"
    elif any(k in low for k in ("asistir a reuniones", "modalidad virtual o presencial", "demandas del servicio")):
        tema = "Participación en Reuniones de Coordinación, Mesas Técnicas y Comités de Seguimiento"
        accion = "la articulación interdisciplinar, participación en mesas de trabajo virtuales o presenciales y seguimiento de compromisos operativos"
    elif any(k in low for k in ("demás actividades", "demas actividades", "encomendadas", "garantizar el cumplimiento")):
        tema = "Actividades Transversales de Apoyo Técnico y Cumplimiento Integral del Contrato"
        accion = "el apoyo técnico transversal a las iniciativas de la Oficina de Tecnologías de la Información y el aseguramiento integral de las obligaciones contractuales"
    else:
        limpia = re.sub(r"^(desarrollar|realizar|ejecutar|apoyar|gestionar|garantizar|brindar|adelantar|participar)\s+", "", txt, flags=re.I)
        tema = limpia[:85].rstrip(".,;:")
        accion = f"el cumplimiento integral de la Obligación Contractual N° {num_ob}"

    return {
        "es_invalida": False,
        "tipo_profesion": tipo_profesion,
        "tema_corto": tema,
        "enfoque": "",
        "accion_cierre": accion,
    }


def _pulir_redaccion_local(texto: str) -> str:
    if not texto:
        return ""
    texto = re.sub(r"\ba el\b", "al", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\bde el\b", "del", texto, flags=re.IGNORECASE)
    texto = re.sub(r"[ \t]{2,}", " ", texto)
    texto = re.sub(r"\s+([,.;:])", r"\1", texto)
    return texto.strip()


def _generar_redaccion_demo_inteligente(informe, detalles_lista: List[Any]) -> dict:
    resultado = {"obligaciones": {}}
    mes_minus = informe.nombre_mes.lower()
    anio = informe.anio
    num_pago = informe.numero_pago
    oficina_area = (informe.contratista.direccion_oficina or "la Oficina de Tecnologías de la Información").strip()
    proyecto = (informe.contratista.proyecto_area or oficina_area).strip()
    nom_contratista = (informe.contratista.nombres_apellidos or "").strip()

    variantes_soportes_graficos = [
        (
            "Verificación Técnica y Registro de Ejecución en Entorno Institucional",
            "Se adelantaron las labores de verificación funcional, validación de componentes y comprobación práctica "
            "orientadas a garantizar {accion_cierre}. La actuación quedó debidamente registrada mediante la evidencia "
            "gráfica '{archivo}', constatando la correcta aplicación de los lineamientos técnicos definidos para {proyecto}."
        ),
        (
            "Control de Calidad, Trazabilidad Operativa y Validación de Flujos",
            "Se realizó el seguimiento técnico y revisión operativa sobre las herramientas y servicios dispuestos en "
            "{oficina_area}, verificando la consistencia de los procesos asociados a la obligación contractual (soporte "
            "documental verificado en '{archivo}') y asegurando el cumplimiento de los criterios de calidad institucionales."
        ),
        (
            "Ejecución de Pruebas de Conformidad y Aseguramiento Funcional",
            "En desarrollo de las actividades programadas para el periodo de {mes} de {anio}, se efectuó la validación "
            "práctica de los módulos y flujos de trabajo intervenidos, documentando su comportamiento estable y "
            "trazabilidad en el soporte '{archivo}'."
        ),
    ]

    for det in detalles_lista:
        k = det.obligacion.numero_obligacion
        clausula = _normalizar_espacios(det.obligacion.descripcion_clausula)
        apuntes = (det.apuntes_contratista or "").strip()
        evidencias = list(det.archivos_evidencia.all())
        ctx_clausula = _inferir_contexto_clausula(clausula, k, objeto_contrato=informe.contratista.objeto_contrato)

        lineas_apuntes = [
            re.sub(r"^[\-\•\*\d\.\)\s]+", "", ln.strip())
            for ln in re.split(r"[\n;]+", apuntes)
            if len(ln.strip()) > 5
        ]

        items_numerados = []
        vinetas_desarrollo = []
        resultados_entregables = []

        for idx_ev, ev in enumerate(evidencias, start=1):
            meta_ev = ev.metadatos_ligeros or {}
            tag_ev = meta_ev.get("tag_semantico", "")
            lineas_ev = meta_ev.get("resumen_lineas") or []
            subtitulo = _limpiar_fragmento_significativo(lineas_ev, ev.nombre_original, nombre_contratista=nom_contratista)
            resumen_limpio = _resumir_contenido_observado_limpio(meta_ev, ev.nombre_original)
            sub_ref = f" — {subtitulo}" if subtitulo else ""
            cita_soporte = (
                f" (soporte verificado: '{ev.nombre_original}' — {resumen_limpio})"
                if resumen_limpio
                else f" (soporte documental verificado: '{ev.nombre_original}')"
            )

            base_low = ev.nombre_original.lower()

            if tag_ev == "ARQUITECTURA_DISENO" or "das_hld" in base_low or "diseño de software" in base_low:
                titulo_pto = f"Diseño Técnico, Estructuración Arquitectónica y Especificación de Componentes{sub_ref}"
                desc_pto = (
                    f"Se estructuró, documentó y socializó la definición técnica y arquitectónica de la solución de software"
                    f"{cita_soporte}, estableciendo el alcance funcional, vistas de componentes, interfaces de integración "
                    f"y lineamientos de diseño orientados a garantizar la escalabilidad, mantenibilidad y calidad técnica "
                    f"dentro del entorno tecnológico de la entidad."
                )
            elif tag_ev == "SEGURIDAD_INMUTABILIDAD" or "seguridad" in base_low or "it-sigaf" in base_low:
                titulo_pto = f"Fortalecimiento Técnico, Integridad Criptográfica y Control de Evidencias{sub_ref}"
                desc_pto = (
                    f"Se implementaron, parametrizaron y documentaron los mecanismos de seguridad, control de ciclo de vida "
                    f"e inmutabilidad documental mediante huellas criptográficas SHA-256{cita_soporte}, asegurando la "
                    f"protección, trazabilidad de auditoría y custodia íntegra de los soportes en los sistemas de información."
                )
            elif tag_ev == "CRONOGRAMA_PLANEACION" or "cronograma" in base_low or "plant120" in base_low:
                titulo_pto = f"Elaboración de Cronogramas, Estimación de Esfuerzos y Seguimiento de Hitos{sub_ref}"
                desc_pto = (
                    f"Se construyó y actualizó la planificación detallada de actividades, estimación de tiempos, "
                    f"asignación de recursos y control de porcentajes de avance{cita_soporte}, permitiendo compilar "
                    f"el estado real de ejecución de las tareas de desarrollo e implementación del periodo."
                )
            elif tag_ev == "PERMISOS_SISTEMAS" or "cooperapp" in base_low or "permisos" in base_low:
                titulo_pto = f"Atención de Solicitudes, Parametrización de Roles y Gestión de Accesos{sub_ref}"
                desc_pto = (
                    f"Se gestionó de manera oportuna el requerimiento recibido por correo institucional realizando el "
                    f"análisis técnico y la configuración efectiva de permisos de administración en la plataforma"
                    f"{cita_soporte}, garantizando la habilitación segura de las funcionalidades de administración y reporte."
                )
            elif tag_ev == "SOFTWARE_GIT" or ev.tipo_archivo == "CODIGO":
                titulo_pto = f"Actualización de Código Fuente, Control de Versiones Git y Gestión de Ramas{sub_ref}"
                desc_pto = (
                    f"Se ejecutaron actividades de codificación, actualización continua de ramas en el repositorio Git "
                    f"institucional, revisión de mezclas (Merge Requests) y documentación estructurada{cita_soporte}, "
                    f"asegurando el cumplimiento de estándares de ingeniería de software."
                )
            elif tag_ev == "REUNION_SEGUIMIENTO":
                titulo_pto = f"Participación en Mesas Técnicas y Sesiones de Coordinación{sub_ref}"
                desc_pto = (
                    f"Se asistió y participó activamente en las sesiones técnicas y reuniones de coordinación convocadas "
                    f"para el seguimiento de los proyectos asignados{cita_soporte}, revisando compromisos operativos, "
                    f"definiciones funcionales y articulación con los equipos de trabajo."
                )
            elif tag_ev == "TABLERO_AGIL":
                titulo_pto = f"Actualización de Actividades y Monitoreo en Herramienta de Seguimiento{sub_ref}"
                desc_pto = (
                    f"Se efectuó la actualización continua del estado de las tareas, historias de usuario y entregables "
                    f"en la herramienta institucional de seguimiento{cita_soporte}, manteniendo la trazabilidad al día."
                )
            else:
                plant_tit, plant_desc = variantes_soportes_graficos[(idx_ev - 1) % len(variantes_soportes_graficos)]
                titulo_pto = f"{plant_tit} — {ctx_clausula['tema_corto'][:60]}"
                desc_pto = plant_desc.format(
                    accion_cierre=ctx_clausula["accion_cierre"],
                    archivo=ev.nombre_original,
                    proyecto=proyecto,
                    oficina_area=oficina_area,
                    mes=mes_minus,
                    anio=anio,
                )
                if resumen_limpio:
                    desc_pto += f" Contenido técnico verificado: {resumen_limpio}."

            desc_pto = _pulir_redaccion_local(desc_pto)
            items_numerados.append(f"{idx_ev}. {titulo_pto}: {desc_pto}")
            vinetas_desarrollo.append({
                "titulo": titulo_pto,
                "descripcion": (
                    f"{desc_pto} Esta actividad fue ejecutada bajo los lineamientos de supervisión de {oficina_area} "
                    f"y se encuentra respaldada en el repositorio digital de la Actividad_{k}."
                ),
            })
            resultados_entregables.append({
                "titulo": f"Entregable Verificado {k}.{idx_ev} — {titulo_pto[:70]}",
                "descripcion": f"Soporte técnico y funcional ('{ev.nombre_original}') que respalda: {desc_pto}",
            })

        for idx_p, punto in enumerate(lineas_apuntes, start=len(items_numerados) + 1):
            if ":" in punto and len(punto.split(":", 1)[0].split()) <= 10:
                partes = punto.split(":", 1)
                titulo_pto = partes[0].strip()
                desc_pto = partes[1].strip()
            else:
                titulo_pto = f"Ejecución Técnica — {ctx_clausula['tema_corto']}"
                desc_pto = punto

            desc_limpia = sanitizar_texto_para_reporte(_pulir_redaccion_local(desc_pto))
            titulo_limpio = sanitizar_texto_para_reporte(titulo_pto)

            if len(desc_limpia) < 140:
                desc_limpia = (
                    f"{desc_limpia.rstrip('.')}. Actuación desarrollada en cumplimiento de la Obligación Contractual N° {k} "
                    f"orientada a garantizar {ctx_clausula['accion_cierre']} conforme a los estándares técnicos de la entidad."
                )

            items_numerados.append(f"{idx_p}. {titulo_limpio}: {desc_limpia}")
            vinetas_desarrollo.append({
                "titulo": titulo_limpio,
                "descripcion": f"{desc_limpia} Actividad articulada con el plan de trabajo del periodo {mes_minus} de {anio}.",
            })
            resultados_entregables.append({
                "titulo": f"Producto {k}.{idx_p} — {titulo_limpio[:70]}",
                "descripcion": f"Resultado verificable de la Obligación {k}: {desc_limpia}",
            })

        if len(items_numerados) == 0:
            t1 = f"Ejecución de Actividades Técnicas — {ctx_clausula['tema_corto']}"
            d1 = (
                f"Se adelantaron de manera continua y estructurada las labores previstas en la cláusula contractual "
                f"('{clausula}'), enfocadas en garantizar {ctx_clausula['accion_cierre']} en los sistemas de información "
                f"asignados durante el periodo de {mes_minus} de {anio}."
            )
            t2 = "Aseguramiento de Calidad, Verificación Funcional y Trazabilidad Documental"
            d2 = (
                f"Se efectuó la validación de conformidad técnica, revisión de estándares de calidad y consolidación "
                f"de los entregables correspondientes a la Actividad_{k}, verificando su alineación con los lineamientos "
                f"de la Oficina de Tecnologías de la Información y la supervisión contractual."
            )
            items_numerados = [f"1. {t1}: {d1}", f"2. {t2}: {d2}"]
            vinetas_desarrollo = [{"titulo": t1, "descripcion": d1}, {"titulo": t2, "descripcion": d2}]
            resultados_entregables = [
                {"titulo": f"Entregable {k}.1 — {ctx_clausula['tema_corto'][:65]}", "descripcion": d1},
                {"titulo": f"Entregable {k}.2 — Control de Calidad y Conformidad Actividad_{k}", "descripcion": d2},
            ]
        elif len(items_numerados) == 1:
            t_comp = f"Aseguramiento de Calidad y Verificación de Conformidad — {ctx_clausula['tema_corto'][:55]}"
            d_comp = (
                f"Complementariamente, se realizó la revisión de consistencia técnica, validación de criterios de "
                f"calidad y organización del expediente probatorio de la Actividad_{k}, garantizando {ctx_clausula['accion_cierre']} "
                f"dentro del marco operativo del periodo reportado."
            )
            items_numerados.append(f"2. {t_comp}: {d_comp}")
            vinetas_desarrollo.append({"titulo": t_comp, "descripcion": d_comp})

        tipo_prof = ctx_clausula.get("tipo_profesion", "ingeniería de software y gestión tecnológica")
        encabezado_tabla = (
            f"Durante el periodo mensual correspondiente a {mes_minus} de {anio}, en cumplimiento de la Obligación "
            f"Contractual N° {k} ({ctx_clausula['tema_corto']}), se ejecutaron de manera rigurosa las siguientes "
            f"actividades de {tipo_prof}:\n"
        )
        cierre_tabla = (
            f"\n\nConclusiones y Aseguramiento Técnico: Las actividades y entregables reportados garantizan "
            f"{ctx_clausula['accion_cierre']}, cumpliendo los estándares de calidad, seguridad de la información "
            f"y trazabilidad documental establecidos por la supervisión contractual."
        )
        texto_tabla = sanitizar_texto_para_reporte(
            _pulir_redaccion_local(encabezado_tabla + "\n".join(items_numerados) + cierre_tabla)
        )

        bloques_evidencias = []
        subitems_anexos = []

        if evidencias:
            for ev_idx, ev in enumerate(evidencias, start=1):
                vin_correspondiente = vinetas_desarrollo[min(ev_idx - 1, len(vinetas_desarrollo) - 1)]
                titulo_ev_limpio = vin_correspondiente["titulo"]
                desc_ev_limpia = (
                    f"{vin_correspondiente['descripcion']} "
                    f"(Archivo soporte verificado: '{ev.nombre_original}', incorporado en la carpeta "
                    f"Anexos de la Actividad_{k})."
                )
                bloques_evidencias.append({
                    "numero": ev_idx,
                    "titulo": sanitizar_texto_para_reporte(titulo_ev_limpio),
                    "descripcion": sanitizar_texto_para_reporte(_pulir_redaccion_local(desc_ev_limpia)),
                    "archivo_asociado": ev.nombre_original,
                })
                subitems_anexos.append({
                    "etiqueta": ev.nombre_original,
                    "detalle": sanitizar_texto_para_reporte(vin_correspondiente["descripcion"]),
                })
        else:
            for idx_b, vin in enumerate(vinetas_desarrollo, start=1):
                bloques_evidencias.append({
                    "numero": idx_b,
                    "titulo": vin["titulo"],
                    "descripcion": vin["descripcion"],
                    "archivo_asociado": "",
                })
                subitems_anexos.append({
                    "etiqueta": vin["titulo"],
                    "detalle": vin["descripcion"],
                })

        desarrollo_oti = {
            "parrafos_introductorios": [
                (
                    f"Durante el periodo correspondiente al mes de {mes_minus} de {anio}, las actuaciones desarrolladas "
                    f"en el marco de la Obligación Contractual N° {k} se enfocaron en {ctx_clausula['tema_corto'].lower()}, "
                    f"atendiendo los requerimientos técnicos y funcionales de los sistemas de información a cargo."
                ),
                (
                    f"En observancia de los lineamientos institucionales de arquitectura, calidad de software y gestión "
                    f"documental, se aseguró {ctx_clausula['accion_cierre']}, consolidando los soportes y productos "
                    f"verificables que respaldan el cumplimiento de la cláusula: \"{clausula}\"."
                ),
            ],
            "vinetas_desarrollo": vinetas_desarrollo,
            "resultados_entregables": resultados_entregables,
            "anexos_secciones": [
                {
                    "titulo": f"Anexo {k}.1 — Soportes Técnicos y Evidencias Consolidados (Actividad_{k})",
                    "descripcion": (
                        f"Registros técnicos y documentos soporte ubicados en "
                        f"Expediente Digital de Evidencias / Pago {num_pago:02d} / Actividad_{k} / Anexos / "
                        f"y consolidados en el documento Obligacion {k}.docx:"
                    ),
                    "subitems": subitems_anexos,
                }
            ],
        }

        resultado["obligaciones"][str(k)] = {
            "texto_tabla_principal": texto_tabla,
            "desarrollo_detallado_oti": desarrollo_oti,
            "bloques_evidencias": bloques_evidencias,
        }

    return resultado


def _construir_prompt_single_shot(informe, detalles_lista: List[Any], apuntes_generales_mes: str = "") -> str:
    c = informe.contratista
    items_payload = []
    bolsa_global_archivos = []

    for det in detalles_lista:
        ob = det.obligacion
        evs_det = list(det.archivos_evidencia.all())
        for ev in evs_det:
            bolsa_global_archivos.append({
                "archivo": ev.nombre_original,
                "ob_sugerida": ob.numero_obligacion,
                "info": ev.resumen_para_prompt_ia(),
            })
        items_payload.append({
            "ob": ob.numero_obligacion,
            "clausula": _normalizar_espacios(ob.descripcion_clausula),
            "apuntes": (det.apuntes_contratista or "")[: c.max_caracteres_apuntes],
        })

    prompt = f"""Eres el redactor técnico experto en contratación pública y supervisión contractual.
Redacta la documentación del Pago {informe.numero_pago:02d} ({informe.nombre_mes} {informe.anio}) para el contratista {c.nombres_apellidos}.

CONTRATO: {c.numero_contrato}-{c.anio_contrato} | PROYECTO: {c.proyecto_area}
OBJETO: {_normalizar_espacios(c.objeto_contrato)}

ARCHIVOS SOPORTE ASOCIADOS:
{json.dumps(bolsa_global_archivos, ensure_ascii=False, separators=(",", ":"))}

OBLIGACIONES A REDACTAR:
{json.dumps(items_payload, ensure_ascii=False, separators=(",", ":"))}

INSTRUCCIONES DE CALIDAD OBLIGATORIAS:
1. En `texto_tabla_principal`: Genera una redacción formal numerada (1. Título técnico: descripción de 4 a 6 líneas con impacto operativo) que alcance 250 a 350 palabras por obligación. NO uses corchetes angulares ni etiquetas HTML.
2. En `desarrollo_detallado_oti`: Incluye 2 párrafos introductorios y viñetas técnicas sustanciosas.
3. Responde EXCLUSIVAMENTE con un JSON válido con esta estructura:
{{
  "obligaciones": {{
    "<numero_obligacion>": {{
      "texto_tabla_principal": "Texto extenso numerado libre de etiquetas HTML...",
      "desarrollo_detallado_oti": {{
        "parrafos_introductorios": ["Párrafo formal 1...", "Párrafo técnico 2..."],
        "vinetas_desarrollo": [{{"titulo": "Título técnico", "descripcion": "Explicación exhaustiva..."}}],
        "resultados_entregables": [{{"titulo": "Entregable verificable", "descripcion": "Descripción..."}}]
      }},
      "bloques_evidencias": [
        {{"numero": 1, "titulo": "Título de la evidencia", "descripcion": "Explicación técnica...", "archivo_asociado": "archivo.ext"}}
      ]
    }}
  }}
}}"""
    return prompt


def _extraer_json_limpio(texto_respuesta: str) -> dict:
    limpio = texto_respuesta.strip()
    match = re.search(r"```(?:json)?\s*(\{.*\})\s*```", limpio, flags=re.DOTALL)
    if match:
        limpio = match.group(1)
    else:
        inicio = limpio.find("{")
        fin = limpio.rfind("}")
        if inicio != -1 and fin != -1 and fin > inicio:
            limpio = limpio[inicio : fin + 1]
    return json.loads(limpio)


def _llamar_ollama_local_directo(prompt: str) -> Optional[dict]:
    """Llamada directa optimizada a tu HP OMEN 14 via Cloudflare Tunnel sin bloqueos."""
    base_url = getattr(settings, "HERMES_BASE_URL", "").strip()
    if not base_url:
        return None

    try:
        from openai import OpenAI

        url_limpia = base_url.rstrip("/")
        if not url_limpia.endswith("/v1"):
            url_limpia += "/v1"

        modelo = getattr(settings, "HERMES_MODEL", "qwen2.5:7b-instruct")
        logger.warning("Enviando petición a tu OMEN 14 en: %s con modelo %s", url_limpia, modelo)

        client = OpenAI(
            base_url=url_limpia,
            api_key=getattr(settings, "HERMES_API_KEY", "ollama-local"),
            default_headers={"User-Agent": "CuentasURT-SaaS/1.0"},
            timeout=180.0,
            max_retries=0,
        )

        resp = client.chat.completions.create(
            model=modelo,
            messages=[
                {
                    "role": "system",
                    "content": "Eres un redactor técnico experto en contratación pública. Responde ÚNICAMENTE en JSON válido sin texto adicional.",
                },
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.2,
        )

        contenido = resp.choices[0].message.content
        datos = _extraer_json_limpio(contenido)
        if datos and "obligaciones" in datos:
            logger.warning("¡Respuesta exitosa recibida de tu OMEN 14!")
            return datos
        else:
            logger.error("Ollama respondió pero el JSON no contiene 'obligaciones'.")
            return None
    except Exception as exc:
        logger.error("Fallo al conectar con Ollama en tu OMEN (%s).", exc)
        return None


def _llamar_gemini_fallback(prompt: str) -> Optional[dict]:
    api_key = getattr(settings, "GEMINI_API_KEY", "")
    if not api_key:
        return None
    try:
        from google import genai
        from google.genai import types

        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model=getattr(settings, "GEMINI_MODEL", "gemini-2.5-flash"),
            contents=prompt,
            config=types.GenerateContentConfig(temperature=0.2, response_mime_type="application/json"),
        )
        return _extraer_json_limpio(response.text)
    except Exception as exc:
        logger.warning("Fallback Gemini falló (%s).", exc)
        return None


def _llamar_openai_fallback(prompt: str) -> Optional[dict]:
    api_key = getattr(settings, "OPENAI_API_KEY", "")
    if not api_key:
        return None
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key, timeout=25.0)
        resp = client.chat.completions.create(
            model=getattr(settings, "OPENAI_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": "Responde únicamente en JSON válido."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.2,
        )
        return _extraer_json_limpio(resp.choices[0].message.content)
    except Exception as exc:
        logger.warning("Fallback OpenAI falló (%s).", exc)
        return None


def _calcular_huella_obligacion(det, informe) -> str:
    evs = [f"{ev.nombre_original}:{ev.tamano_bytes}" for ev in det.archivos_evidencia.all()]
    estado_raw = (
        f"V3|CTO:{informe.contratista.numero_contrato}|PAGO:{informe.numero_pago}:{informe.mes}:{informe.anio}|"
        f"OB:{det.obligacion.numero_obligacion}:{det.obligacion.descripcion_clausula}|"
        f"APUNTES:{(det.apuntes_contratista or '').strip()}|"
        f"ARCHIVOS:{'|'.join(sorted(evs))}"
    )
    return hashlib.md5(estado_raw.encode("utf-8")).hexdigest()


def redactar_informe_con_ia(
    informe,
    obligacion_especifica_id: Optional[int] = None,
    forzar_demo: bool = False,
    apuntes_generales_mes: str = "",
    ignorar_cache: bool = True,
) -> Dict[str, Any]:
    contratista = informe.contratista

    if not contratista.activo_mes_actual:
        raise PermissionError("El contratista no se encuentra activo para generar informes en el mes actual.")

    modo_local = getattr(settings, "LOCAL_ONLY_AI", False)

    if obligacion_especifica_id is not None:
        if contratista.creditos_regeneracion_individual <= 0 and not modo_local:
            raise PermissionError("Ha alcanzado el límite mensual de créditos de regeneración individual (6/mes).")
        detalles = list(
            informe.detalles_obligaciones.select_related("obligacion")
            .prefetch_related("archivos_evidencia")
            .filter(obligacion_id=obligacion_especifica_id)
        )
    else:
        if contratista.creditos_informe_completo <= 0 and not modo_local:
            raise PermissionError("Ha alcanzado el límite mensual de créditos de generación de informe completo (3/mes).")
        detalles = list(
            informe.detalles_obligaciones.select_related("obligacion")
            .prefetch_related("archivos_evidencia")
            .all()
        )

    if not detalles:
        raise ValueError("No se encontraron obligaciones configuradas para redactar.")

    detalles_a_redactar = detalles
    prompt = _construir_prompt_single_shot(informe, detalles_a_redactar, apuntes_generales_mes=apuntes_generales_mes)
    motor_usado = "Modo Offline Inteligente (0 Tokens Externos)"
    payload_ia = None

    if not forzar_demo and not getattr(settings, "LOCAL_ONLY_AI", False):
        with _HERMES_FIFO_LOCK:
            # 1. PRIORIDAD ABSOLUTA: Llamar a tu OMEN 14
            payload_ia = _llamar_ollama_local_directo(prompt)
            if payload_ia and "obligaciones" in payload_ia:
                motor_usado = f"Ollama Local OMEN 14 ({getattr(settings, 'HERMES_MODEL', 'qwen2.5')})"
            else:
                # 2. Respaldo Gemini
                payload_ia = _llamar_gemini_fallback(prompt)
                if payload_ia and "obligaciones" in payload_ia:
                    motor_usado = f"Fallback Gemini ({getattr(settings, 'GEMINI_MODEL', 'gemini-2.5-flash')})"
                else:
                    # 3. Respaldo OpenAI
                    payload_ia = _llamar_openai_fallback(prompt)
                    if payload_ia and "obligaciones" in payload_ia:
                        motor_usado = f"Fallback OpenAI ({getattr(settings, 'OPENAI_MODEL', 'gpt-4o-mini')})"

    if not payload_ia or "obligaciones" not in payload_ia:
        payload_ia = _generar_redaccion_demo_inteligente(informe, detalles_a_redactar)
        motor_usado = "Modo Offline Inteligente (Single-Shot Local)"

    obligaciones_dict = payload_ia.get("obligaciones", {})

    mapa_evidencias_por_nombre = {}
    for d_item in detalles:
        for ev_item in d_item.archivos_evidencia.all():
            mapa_evidencias_por_nombre[ev_item.nombre_original.strip().lower()] = ev_item

    for det in detalles_a_redactar:
        k_str = str(det.obligacion.numero_obligacion)
        datos_ob = obligaciones_dict.get(k_str) or obligaciones_dict.get(det.obligacion.numero_obligacion)
        if not datos_ob:
            datos_ob = _generar_redaccion_demo_inteligente(informe, [det])["obligaciones"][k_str]

        texto_principal = sanitizar_texto_para_reporte(datos_ob.get("texto_tabla_principal", ""))
        desarrollo_oti = datos_ob.get("desarrollo_detallado_oti", {})
        bloques_ev = datos_ob.get("bloques_evidencias", [])

        if isinstance(desarrollo_oti, dict) and not desarrollo_oti.get("anexos_secciones"):
            k_num = det.obligacion.numero_obligacion
            desarrollo_oti["anexos_secciones"] = [
                {
                    "titulo": f"Anexo {k_num}.1 — Soportes Técnicos y Evidencias Documentales (Actividad_{k_num})",
                    "descripcion": (
                        f"Registros técnicos y evidencias ubicadas en "
                        f"Expediente Digital de Evidencias / Pago {informe.numero_pago:02d} / Actividad_{k_num} / Anexos / "
                        f"y consolidadas en el documento Obligacion {k_num}.docx:"
                    ),
                    "subitems": [
                        {
                            "etiqueta": b.get("archivo_asociado") or b.get("titulo") or f"Soporte {idx_b}",
                            "detalle": sanitizar_texto_para_reporte(b.get("descripcion", "")),
                        }
                        for idx_b, b in enumerate(bloques_ev, start=1)
                    ],
                }
            ]

        for b_ev in bloques_ev:
            nom_asoc = (b_ev.get("archivo_asociado") or "").strip().lower()
            if nom_asoc and nom_asoc in mapa_evidencias_por_nombre:
                ev_target = mapa_evidencias_por_nombre[nom_asoc]
                ev_target.detalle_obligacion = det
                ev_target.titulo_evidencia = sanitizar_texto_para_reporte(
                    b_ev.get("titulo") or ev_target.titulo_evidencia or ev_target.nombre_original
                )[:300]
                ev_target.descripcion_tecnica_ia = sanitizar_texto_para_reporte(
                    b_ev.get("descripcion") or ev_target.descripcion_tecnica_ia
                )
                ev_target.save()

        evidencias_db = list(det.archivos_evidencia.all())
        for idx_ev, ev_obj in enumerate(evidencias_db):
            if idx_ev < len(bloques_ev):
                b_ev = bloques_ev[idx_ev]
                ev_obj.titulo_evidencia = sanitizar_texto_para_reporte(
                    b_ev.get("titulo") or ev_obj.titulo_evidencia or ev_obj.nombre_original
                )[:300]
                ev_obj.descripcion_tecnica_ia = sanitizar_texto_para_reporte(
                    b_ev.get("descripcion") or ev_obj.descripcion_tecnica_ia
                )
                ev_obj.save()

        if isinstance(desarrollo_oti, dict):
            desarrollo_oti["_hash_cache_ia"] = _calcular_huella_obligacion(det, informe)

        det.texto_tabla_principal = texto_principal
        det.desarrollo_detallado_oti = desarrollo_oti
        det.bloques_evidencias_json = bloques_ev
        det.save()

    if obligacion_especifica_id is not None and not modo_local:
        contratista.creditos_regeneracion_individual = max(0, contratista.creditos_regeneracion_individual - 1)
    elif obligacion_especifica_id is None and not modo_local:
        contratista.creditos_informe_completo = max(0, contratista.creditos_informe_completo - 1)
    contratista.save()

    informe.ultimo_motor_ia_usado = motor_usado
    if informe.estado == "BORRADOR":
        informe.estado = "REDACTADO_IA"
    informe.save()

    return {
        "motor_usado": motor_usado,
        "obligaciones_actualizadas": [d.obligacion.numero_obligacion for d in detalles_a_redactar],
        "creditos_informe_completo": contratista.creditos_informe_completo,
        "creditos_regeneracion_individual": contratista.creditos_regeneracion_individual,
    }