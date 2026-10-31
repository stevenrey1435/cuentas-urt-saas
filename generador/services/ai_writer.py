"""
Conector IA Dual con Ahorro Extremo de Tokens (Single-Shot JSON) y Cola Secuencial FIFO.

Arquitectura de Redacción Coherente en UNA SOLA LLAMADA:
- Primaria: Hermes Agent (`openai-codex` OAuth / endpoint local compatible con OpenAI) en cola secuencial FIFO (1 informe a la vez).
- Respaldo 1 (Fallback Automático): Google Gemini API (`gemini-2.5-flash` via SDK `google-genai`) si `GEMINI_API_KEY` está configurada.
- Respaldo 2 (Fallback Automático): OpenAI API (`gpt-4o-mini` via SDK `openai`) si `OPENAI_API_KEY` está configurada.
- Respaldo 3 (Modo Offline Inteligente): Motor determinista estructurado de alta precisión técnica para contratación pública.
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

# Cola secuencial (Semáforo FIFO = 1 informe a la vez) para proteger la cuenta OAuth de Hermes Agent
_HERMES_FIFO_LOCK = threading.Lock()


FEW_SHOT_STYLE_REFERENCE = """
ESTÁNDAR DE EXTENSIÓN, PROFUNDIDAD Y RIGOR TÉCNICO EN CONTRATACIÓN PÚBLICA:
CRITERIO DE CALIDAD: La supervisión técnica y las auditorías de contratación pública exigen redacciones EXTENSAS, EXHAUSTIVAS, PROFUNDAS Y DENSAS. Está TERMINANTEMENTE PROHIBIDO generar resúmenes cortos, viñetas de una sola línea, textos superficiales o mensajes de error de lectura. Cada texto debe transmitir solvencia y rigor profesional.

1. En `texto_tabla_principal` (para la columna 'Actividades Desarrolladas' del formato oficial GC-FO-15):
   - Estructura obligatoria (extensión objetivo: 250 a 450 palabras por obligación):
     * Párrafo introductorio formal: Contextualización del periodo reportado, vinculación al proyecto o área, marco institucional y directrices de supervisión.
     * Numerales técnicos detallados ('1. Título Técnico: ...', '2. Título Técnico: ...'): Cada numeral debe ser un bloque denso de 4 a 7 líneas que detalle componentes intervenidos, metodologías aplicadas, patrones de trabajo, validaciones o actas de sesión técnica.
     * Párrafo de impacto o cierre técnico: Certificación de la estabilidad operativa, trazabilidad transaccional, calidad del trabajo y aseguramiento de directrices institucionales y contractuales.
   - PROHIBIDO incluir etiquetas como `<para>`, `</para>`, `<br>`, correos electrónicos personales o corchetes angulares `<` `>` sin escapar.

2. En `desarrollo_detallado_oti` (para `REPORTE DE OBLIGACIONES CONTRACTUALES.docx` de cada `Actividad_K`):
   - `parrafos_introductorios`: Mínimo 2 párrafos formales extensos y bien articulados.
   - `vinetas_desarrollo`: Mínimo 2 a 4 viñetas técnicas sustanciosas (`{"titulo": "Título Técnico", "descripcion": "Explicación exhaustiva de 5 a 8 líneas..."}`).
   - `resultados_entregables`: 2 a 4 productos o entregables verificables.

3. En `bloques_evidencias` (para `Actividad_<K>/Anexos/Obligacion <K>.docx`):
   - Lista numerada de evidencias donde cada bloque debe tener una descripción técnica profunda de 4 a 6 líneas explicando qué demuestra la captura o documento y su impacto comprobable en el proyecto.
"""


def _normalizar_espacios(texto: str) -> str:
    """Colapsa saltos de línea y espacios múltiples para que los match de frases no fallen por saltos de línea."""
    return re.sub(r"\s+", " ", (texto or "")).strip()


def clasificar_obligacion_por_afinidad(
    texto_item: str,
    detalles_lista: List[Any],
    indice_rotacion: int = 0,
):
    """
    Analiza el nombre del archivo, su ruta dentro del .zip y sus metadatos ligeros y determina
    en cuál de las N obligaciones del contratista encaja mejor según el texto de cada cláusula.
    Garantiza además que los soportes genéricos (capturas PNG) cubran primero las obligaciones vacías.
    """
    if not detalles_lista:
        return None

    txt = _normalizar_espacios(texto_item).lower()

    # 1. Mención explícita en el nombre o carpeta ("obligacion K", "obK", "act_K", "actividad K")
    m_exp = re.search(r"(?:obligaci[oó]n|ob|actividad|act)[\s_\-]*(\d{1,2})\b", txt)
    if m_exp:
        num_buscado = int(m_exp.group(1))
        for det in detalles_lista:
            if det.obligacion.numero_obligacion == num_buscado:
                return det

    puntajes_semanticos = puntuar_obligaciones_semanticamente(texto_item, detalles_lista)

    # 2. Detección de familias de archivos específicos
    es_reunion = any(t in txt for t in (
        "reunion_seguimiento", "asistencia", "participantes", "resumen de la reunión",
        "resumen de la reunion", "unirse a la reunión", "teams", "convocatoria", "mesa técnica"
    ))
    es_desarrollo_git = any(t in txt for t in (
        "software_git", "merge request", "hotfix", "commit ", "commits", "gitlab", "branch", "feat(", "fix("
    ))
    es_tablero_agil = any(t in txt for t in (
        "tablero_agil", "kanban", "in progress", "sprint", "jira", "azure boards", "planner"
    ))
    es_cronograma = any(t in txt for t in (
        "cronograma_planeacion", "cronograma", "gantt", "plan de trabajo", "plant120", "alcance2", "alcance"
    ))
    es_arquitectura = any(t in txt for t in (
        "arquitectura_diseno", "das_hld", "das/", "hld", "arquitectura de software",
        "diseño de alto nivel", "diseño de software", "diseno de software"
    ))
    es_seguridad = any(t in txt for t in (
        "seguridad_inmutabilidad", "sha-256", "control de evidencias", "inmutabilidad",
        "seguridad_datos", "it-sigaf", "seguridad y control"
    ))
    es_permisos = any(t in txt for t in (
        "permisos_sistemas", "permisos de administrador", "cooperapp", "solicitud de asignación de permisos",
        "asignación de permisos", "mesa de servicios"
    ))

    # Mapeo directo por tipo de documento inequívoco
    for det in detalles_lista:
        clausula_norm = _normalizar_espacios(det.obligacion.descripcion_clausula).lower()
        if es_arquitectura and any(k in clausula_norm for k in ("arquitectura", "desarrollar e integrar")):
            return det
        if es_cronograma and any(k in clausula_norm for k in ("cronograma", "estimar esfuerzos", "actividades planificadas")):
            return det
        if es_permisos and any(k in clausula_norm for k in ("mesa de servicios", "correo electrónico", "requerimientos", "solicitudes")):
            return det
        if es_seguridad and any(k in clausula_norm for k in ("parametrización", "fortalecimiento técnico", "implementación")):
            return det
        if es_desarrollo_git and any(k in clausula_norm for k in ("repositorio git", "ramas asignadas", "código fuente")):
            return det
        if es_tablero_agil and any(k in clausula_norm for k in ("herramienta de seguimiento", "seguimientos programados")):
            return det
        if es_reunion and any(k in clausula_norm for k in ("asistir a reuniones", "modalidad virtual o presencial")):
            return det

    # 3. Puntuación léxica + semántica para documentos con texto
    stopwords_genericas = {
        "para", "como", "este", "esta", "documento", "archivo", "soporte", "evidencia",
        "sistema", "entrega", "informe", "gestion", "actividad", "captura", "pantalla", "imagen"
    }
    palabras_item = set(re.findall(r"[a-záéíóúñ]{4,}", txt)) - stopwords_genericas

    mejor_det = None
    mejor_puntaje = 0.0

    for det in detalles_lista:
        clausula_low = _normalizar_espacios(det.obligacion.descripcion_clausula).lower()
        palabras_clausula = set(re.findall(r"[a-záéíóúñ]{4,}", clausula_low)) - stopwords_genericas
        comunes = palabras_item.intersection(palabras_clausula)
        puntaje = len(comunes) * 4.0

        similitud = (puntajes_semanticos or {}).get(det.obligacion.numero_obligacion, 0.0)
        if similitud > 0.18:
            puntaje += similitud * 10.0

        if puntaje > 0:
            puntaje_ajustado = puntaje - (det.archivos_evidencia.count() * 1.5)
            if puntaje_ajustado > mejor_puntaje:
                mejor_puntaje = puntaje_ajustado
                mejor_det = det

    if mejor_det is not None and mejor_puntaje > 2.0:
        return mejor_det

    # 4. Distribución equitativa inteligente para capturas genéricas (`Captura de pantalla...`, `imagen (N).png`)
    # Prioriza obligaciones que aún tienen 0 evidencias para que ninguna obligación quede huérfana en el PDF
    detalles_ordenados = sorted(
        detalles_lista,
        key=lambda d: (d.archivos_evidencia.count(), d.obligacion.numero_obligacion),
    )
    return detalles_ordenados[0]


def _limpiar_fragmento_significativo(lineas: List[str], nombre_archivo: str, nombre_contratista: str = "") -> str:
    """
    Extrae un subtítulo limpio y profesional a partir del nombre del archivo o sus primeras líneas,
    eliminando nombres propios del contratista, etiquetas de Outlook ('RE_', 'Outlo'), puntos suspensivos
    de tablas de contenido y metadatos crudos.
    """
    base = os.path.splitext(nombre_archivo)[0]

    # Limpiar dinámicamente el nombre del contratista y variantes sin tilde
    variantes_nombre = []
    if nombre_contratista:
        partes = [p for p in re.split(r"\s+", nombre_contratista.strip()) if len(p) > 2]
        if len(partes) >= 2:
            variantes_nombre.append(r"[\s_\-]*" + r"[\s_\-]+".join(partes[:2]) + r".*$")
    variantes_nombre.append(r"[\s_\-]*jes[uú]s[\s_\-]+steven[\s_\-]+rey.*$")

    # 1. Correos exportados de Outlook
    if re.search(r"(^re[\s_]|^rv[\s_]|^fw[\s_]|outlook|outlo)", base.lower()):
        asunto = re.sub(r"^(re|rv|fw|fwd)[\s_:\-]+", "", base, flags=re.I)
        for pat in variantes_nombre:
            asunto = re.sub(pat, "", asunto, flags=re.I)
        asunto = re.sub(r"[\s_\-]*(outlook|outlo).*$", "", asunto, flags=re.I)
        asunto = re.sub(r"[_\-]+", " ", asunto).strip(" -–—.,;:")
        if len(asunto) > 8:
            return sanitizar_texto_para_reporte(asunto[:90])

    # 2. Cronogramas y Planes de Trabajo
    if any(k in base.lower() for k in ("cronograma", "plant120", "plan_trabajo", "gantt")):
        nom_cron = re.sub(r"(\d{1,2})(octubre|noviembre|diciembre|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre)(\d{4})", r" \1 de \2 de \3", base, flags=re.I)
        nom_cron = re.sub(r"[_\-]+", " ", nom_cron)
        nom_cron = re.sub(r"\s+", " ", nom_cron).strip(" -–—.,;:")
        return sanitizar_texto_para_reporte(nom_cron[:90])

    # 3. Documentos técnicos con nombre descriptivo (ej. DAS_HLD_T120_URT, IT-SIGAF_Seguridad...)
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
    """
    Sintetiza de forma limpia y ejecutiva el contenido del archivo SIN volcar correos electrónicos
    personales (Ley 1581 de 2012), ni puntos de tabla de contenido ('.....'), ni filas crudas de MS Project.
    """
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

    # Limpieza general para cualquier otro PDF/Word/Excel futuro
    lineas = list(meta.get("resumen_lineas") or [])
    for clave in ("texto_pdf", "texto_ocr"):
        val = (meta.get(clave) or "").strip()
        if val:
            lineas.append(val)

    fragmentos_limpios = []
    for ln in lineas:
        s = str(ln)
        # Eliminar correos electrónicos <...> y direcciones @
        s = re.sub(r"[‹<]?[\w\.\-]+@[\w\.\-]+\.\w+[›>]?", "", s)
        # Eliminar secuencias de puntos de índice (.....)
        s = re.sub(r"\.{3,}", " ", s)
        # Eliminar caracteres privados de fuentes de íconos (Unicode Private Use Area)
        s = re.sub(r"[\ue000-\uf8ff]", "", s)
        # Eliminar cabeceras crudas de correo o MS Project
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
    """
    Analiza la cláusula contractual (normalizando primero los saltos de línea) para determinar con
    exactitud milimétrica el título temático, tipo de profesión y conclusión técnica.
    """
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

    # 1. Inferencia de Tipo de Profesión
    if any(k in ctx_global for k in (
        "software", "sistemas de información", "código fuente", "git",
        "base de datos", "arquitectura", "tecnologías de la información"
    )):
        tipo_profesion = "ingeniería de software y gestión tecnológica"
    elif any(k in low for k in ("demanda", "tutela", "juzgado", "tribunal", "sentencia", "sustanciación")):
        tipo_profesion = "gestión jurídica y procesal"
    elif any(k in low for k in ("predio", "catastro", "catastral", "georreferenc", "topográf", "lindero")):
        tipo_profesion = "gestión técnico-predial y catastral"
    elif any(k in low for k in ("víctima", "comunidad", "psicosocial", "enfoque diferencial")):
        tipo_profesion = "intervención social y comunitaria"
    else:
        tipo_profesion = "gestión profesional y técnica"

    # 2. Clasificación exacta de la Cláusula (usando texto sin saltos de línea)
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
    # Ramas multidisciplinares para otros contratos (solo si NO es contrato de software)
    elif any(k in low for k in ("defensa judicial", "tutela", "juzgado", "tribunal", "demanda", "procesal")):
        tema = "Defensa Judicial, Sustanciación Procesal y Atención de Acciones Constitucionales"
        accion = "la oportuna defensa judicial, sustanciación técnica de memoriales y atención estricta de términos procesales"
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
    """Aplica correcciones gramaticales conservadoras."""
    if not texto:
        return ""
    texto = re.sub(r"\ba el\b", "al", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\bde el\b", "del", texto, flags=re.IGNORECASE)
    texto = re.sub(r"[ \t]{2,}", " ", texto)
    texto = re.sub(r"\s+([,.;:])", r"\1", texto)
    return texto.strip()


def _generar_redaccion_demo_inteligente(informe, detalles_lista: List[Any]) -> dict:
    """
    Motor Redactor Experto Local 0-Tokens:
    Produce redacciones extensas, formales y 100% alineadas con cada cláusula contractual y sus
    soportes físicos, sin mensajes robóticos de fallo de lectura ni volcados de correo crudos.
    """
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
        ctx_clausula = _inferir_contexto_clausula(
            clausula,
            k,
            objeto_contrato=informe.contratista.objeto_contrato,
        )

        lineas_apuntes = [
            re.sub(r"^[\-\•\*\d\.\)\s]+", "", ln.strip())
            for ln in re.split(r"[\n;]+", apuntes)
            if len(ln.strip()) > 5
        ]

        items_numerados = []
        vinetas_desarrollo = []
        resultados_entregables = []

        # 1. Redactar numerales técnicos ricos por cada evidencia asociada
        for idx_ev, ev in enumerate(evidencias, start=1):
            meta_ev = ev.metadatos_ligeros or {}
            tag_ev = meta_ev.get("tag_semantico", "")
            lineas_ev = meta_ev.get("resumen_lineas") or []
            hojas_ev = meta_ev.get("hojas") or []
            cols_ev = meta_ev.get("encabezados_fila_1") or {}

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
            elif ev.tipo_archivo == "EXCEL" and (hojas_ev or cols_ev):
                titulo_pto = f"Estructuración, Procesamiento y Validación Técnica de Datos ({ev.nombre_original})"
                desc_pto = (
                    f"Se consolidó y verificó la estructura de datos contenida en el archivo '{ev.nombre_original}' "
                    f"(hojas: {', '.join(hojas_ev[:4])}){cita_soporte}, validando la integridad de los registros."
                )
            else:
                # Soportes gráficos / capturas de pantalla: Redacción técnica limpia y contextualizada a la obligación K
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

        # 2. Incorporar apuntes del contratista si existen
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

        # 3. Asegurar al menos 2 numerales densos y profundos en todas las obligaciones
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

        # 4. Construir bloques_evidencias para `Obligacion K.docx` y subitems para `REPORTE OTI.docx`
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
    """Construye un prompt Single-Shot JSON ultra-optimizado en tokens."""
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

    bloque_apuntes_globales = ""
    if apuntes_generales_mes and apuntes_generales_mes.strip():
        bloque_apuntes_globales = (
            f"\nAPUNTES GLOBALES DEL MES (Distribúyelos en la obligación más afín):\n"
            f"{apuntes_generales_mes.strip()[:3000]}\n"
        )

    prompt = f"""Eres el redactor técnico experto en informes de supervisión y cuentas de cobro para contratistas del Estado.
Redacta en UN SOLO OBJETO JSON ESTRUCTURADO la documentación del Pago {informe.numero_pago:02d} ({informe.nombre_mes} {informe.anio}).

DATOS DEL CONTRATO:
Contratista: {c.nombres_apellidos} | Contrato: {c.numero_contrato}-{c.anio_contrato} | Proyecto(s): {c.proyecto_area}
Objeto: {_normalizar_espacios(c.objeto_contrato)}
{bloque_apuntes_globales}
ARCHIVOS SOPORTE SUBIDOS (Asócialos en `bloques_evidencias` de su obligación más afín usando el nombre exacto en `archivo_asociado`):
{json.dumps(bolsa_global_archivos, ensure_ascii=False, separators=(",", ":"))}

{FEW_SHOT_STYLE_REFERENCE}

OBLIGACIONES A REDACTAR:
{json.dumps(items_payload, ensure_ascii=False, separators=(",", ":"))}

RESPONDE EXCLUSIVAMENTE CON UN JSON VÁLIDO CON ESTA ESTRUCTURA EXACTA:
{{
  "obligaciones": {{
    "<numero_obligacion>": {{
      "texto_tabla_principal": "Texto extenso y numerado (1. Título: ..., 2. Título: ...) de 250-400 palabras libre de etiquetas XML/HTML",
      "desarrollo_detallado_oti": {{
        "parrafos_introductorios": ["Párrafo 1 formal...", "Párrafo 2 técnico..."],
        "vinetas_desarrollo": [
          {{"titulo": "Título técnico", "descripcion": "Explicación rigurosa de 5-7 líneas..."}}
        ],
        "resultados_entregables": [
          {{"titulo": "Entregable verificable", "descripcion": "Descripción del producto obtenido."}}
        ]
      }},
      "bloques_evidencias": [
        {{
          "numero": 1,
          "titulo": "Título técnico formal de la evidencia",
          "descripcion": "Explicación técnica de 4-5 líneas de lo que demuestra este soporte.",
          "archivo_asociado": "nombre_exacto_archivo.ext"
        }}
      ]
    }}
  }}
}}"""
    return prompt


def _extraer_json_limpio(texto_respuesta: str) -> dict:
    """Extrae y parsea el objeto JSON devuelto por el LLM."""
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


def _llamar_hermes_agent_fifo(prompt: str) -> Optional[dict]:
    """Conecta con Hermes Agent respetando la cola FIFO."""
    import base64

    hermes_auth_path = os.path.expandvars(r"%LOCALAPPDATA%\hermes\auth.json")
    hermes_model = getattr(settings, "HERMES_MODEL", "gpt-5.4")
    hermes_reasoning = getattr(settings, "HERMES_REASONING", "low")

    with _HERMES_FIFO_LOCK:
        if os.path.exists(hermes_auth_path):
            try:
                from openai import OpenAI

                with open(hermes_auth_path, "r", encoding="utf-8") as f_auth:
                    auth_data = json.load(f_auth)

                codex_tokens = (
                    auth_data.get("providers", {})
                    .get("openai-codex", {})
                    .get("tokens", {})
                )
                access_token = codex_tokens.get("access_token", "").strip()
                if access_token:
                    headers = {
                        "User-Agent": "HermesAgent/0.11.0",
                        "originator": "hermes-agent",
                    }
                    try:
                        parts = access_token.split(".")
                        if len(parts) >= 2:
                            payload_b64 = parts[1] + "=" * (-len(parts[1]) % 4)
                            jwt_auth = json.loads(base64.urlsafe_b64decode(payload_b64)).get(
                                "https://api.openai.com/auth", {}
                            )
                            acct_id = jwt_auth.get("chatgpt_account_id")
                            if acct_id:
                                headers["ChatGPT-Account-ID"] = acct_id
                    except Exception:
                        pass

                    client = OpenAI(
                        api_key=access_token,
                        base_url="https://chatgpt.com/backend-api/codex",
                        default_headers=headers,
                        timeout=180.0,
                        max_retries=0,
                    )
                    stream = client.responses.create(
                        model=hermes_model,
                        instructions=(
                            "Eres el redactor técnico experto en supervisión y cuentas de cobro para contratistas. "
                            "Responde EXCLUSIVAMENTE con un objeto JSON válido."
                        ),
                        input=[{"role": "user", "content": [{"type": "input_text", "text": prompt}]}],
                        store=False,
                        reasoning={"effort": hermes_reasoning},
                        stream=True,
                    )
                    partes_texto = []
                    for ev in stream:
                        if getattr(ev, "type", "") == "response.output_text.delta":
                            partes_texto.append(getattr(ev, "delta", ""))
                    texto_completo = "".join(partes_texto).strip()
                    if texto_completo and "{" in texto_completo:
                        return _extraer_json_limpio(texto_completo)
            except Exception as exc:
                logger.info("Hermes OAuth Codex directo no retornó JSON (%s).", exc)

        # B. Modo Servidor HTTP Compatible con OpenAI / Ollama Local
        base_url = getattr(settings, "HERMES_BASE_URL", "")
        if base_url:
            try:
                from openai import OpenAI

                client = OpenAI(
                    base_url=base_url,
                    api_key=getattr(settings, "HERMES_API_KEY", "ollama-local"),
                    timeout=90.0,
                    max_retries=1,
                )
                resp = client.chat.completions.create(
                    model=hermes_model,
                    messages=[
                        {
                            "role": "system",
                            "content": (
                                "Eres un redactor técnico experto en contratación pública y cuentas de cobro. "
                                "Responde ÚNICAMENTE con un objeto JSON válido según la estructura requerida, sin texto adicional."
                            ),
                        },
                        {"role": "user", "content": prompt},
                    ],
                    response_format={"type": "json_object"},
                    temperature=0.2,
                )
                contenido = resp.choices[0].message.content
                return _extraer_json_limpio(contenido)
            except Exception as exc:
                logger.warning("Llamada a LLM Local (Ollama/Hermes) falló (%s). Activando fallback...", exc)

    return None


def _llamar_gemini_fallback(prompt: str) -> Optional[dict]:
    """Fallback automático #1: Google Gemini API."""
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
            config=types.GenerateContentConfig(
                temperature=0.2,
                response_mime_type="application/json",
            ),
        )
        return _extraer_json_limpio(response.text)
    except Exception as exc:
        logger.warning("Fallback Gemini API falló (%s).", exc)
        return None


def _llamar_openai_fallback(prompt: str) -> Optional[dict]:
    """Fallback automático #2: OpenAI API directa."""
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
        logger.warning("Fallback OpenAI API falló (%s).", exc)
        return None


def _calcular_huella_obligacion(det, informe) -> str:
    """Calcula un hash MD5 único con el estado exacto de la obligación K (incluyendo versión de motor v2)."""
    evs = [f"{ev.nombre_original}:{ev.tamano_bytes}" for ev in det.archivos_evidencia.all()]
    estado_raw = (
        f"V2|CTO:{informe.contratista.numero_contrato}|PAGO:{informe.numero_pago}:{informe.mes}:{informe.anio}|"
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
) -> Dict[str, Any]:
    """Punto de entrada principal de `services/ai_writer.py`."""
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

    huellas_actuales = {det.id: _calcular_huella_obligacion(det, informe) for det in detalles}
    detalles_a_redactar = []

    if obligacion_especifica_id is not None or bool(apuntes_generales_mes and apuntes_generales_mes.strip()):
        detalles_a_redactar = detalles
    else:
        for det in detalles:
            datos_previos = det.desarrollo_detallado_oti if isinstance(det.desarrollo_detallado_oti, dict) else {}
            huella_previa = datos_previos.get("_hash_cache_ia", "")
            tiene_texto_valido = bool(
                det.texto_tabla_principal
                and len(det.texto_tabla_principal.strip()) > 200
                and "no pudo ser leído" not in det.texto_tabla_principal
                and datos_previos.get("vinetas_desarrollo")
            )
            if not tiene_texto_valido or huella_previa != huellas_actuales[det.id]:
                detalles_a_redactar.append(det)

    if not detalles_a_redactar:
        return {
            "motor_usado": f"Caché Instantáneo 0-Tokens (Sin cambios detectados — {informe.ultimo_motor_ia_usado or 'IA'})",
            "obligaciones_actualizadas": [],
            "creditos_informe_completo": contratista.creditos_informe_completo,
            "creditos_regeneracion_individual": contratista.creditos_regeneracion_individual,
        }

    prompt = _construir_prompt_single_shot(informe, detalles_a_redactar, apuntes_generales_mes=apuntes_generales_mes)
    motor_usado = "Modo Offline Inteligente (0 Tokens Externos)"
    payload_ia = None

    if not forzar_demo and not getattr(settings, "LOCAL_ONLY_AI", False):
        payload_ia = _llamar_hermes_agent_fifo(prompt)
        if payload_ia and "obligaciones" in payload_ia:
            motor_usado = f"Hermes Agent FIFO ({getattr(settings, 'HERMES_MODEL', 'openai-codex')})"
        else:
            payload_ia = _llamar_gemini_fallback(prompt)
            if payload_ia and "obligaciones" in payload_ia:
                motor_usado = f"Fallback API Key Gemini ({getattr(settings, 'GEMINI_MODEL', 'gemini-2.5-flash')})"
            else:
                payload_ia = _llamar_openai_fallback(prompt)
                if payload_ia and "obligaciones" in payload_ia:
                    motor_usado = f"Fallback API Key OpenAI ({getattr(settings, 'OPENAI_MODEL', 'gpt-4o-mini')})"

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