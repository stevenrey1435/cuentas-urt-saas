"""
Conector IA Dual con Ahorro Extremo de Tokens (Single-Shot JSON) y Cola Secuencial FIFO.

Arquitectura de Redacción Coherente en UNA SOLA LLAMADA:
- Primaria: Hermes Agent (`openai-codex` OAuth / endpoint local compatible con OpenAI) en cola secuencial FIFO (1 informe a la vez).
- Respaldo 1 (Fallback Automático): Google Gemini API (`gemini-2.5-flash` via SDK `google-genai`) si `GEMINI_API_KEY` está configurada.
- Respaldo 2 (Fallback Automático): OpenAI API (`gpt-4o-mini` via SDK `openai`) si `OPENAI_API_KEY` está configurada.
- Respaldo 3 (Modo Offline Inteligente): Motor determinista estructurado con Few-Shot real de los Pagos 9 y 10 para poder probar y generar el paquete completo de 20 documentos incluso antes de encender Hermes o configurar llaves externas.

Garantías de Ingeniería:
1. Jamás envía archivos crudos pesados al modelo de IA (solo nombres, tipo y encabezados de Fila 1 si es Excel).
2. En 1 sola petición estructurada produce:
   - `texto_tabla_principal`: Resumen numerado sanitizado para el Informe Principal (`CTO_<NUM>_INFORME_<XX>.docx`) y plataformas de reporte.
   - `desarrollo_detallado_oti`: Introducción, viñetas técnicas, productos/entregables y listado de anexos para `Actividad_<K>/REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx`.
   - `bloques_evidencias`: Títulos numerados y explicaciones técnicas para `Actividad_<K>/Anexos/Obligacion <K>.docx`.
"""

import os
import hashlib
import json
import re
import threading
import logging
from typing import Dict, Any, List, Optional
from django.conf import settings
from .file_inspector import sanitizar_texto_para_reporte, sanitizar_texto_para_reporte
from .semantic_matcher import puntuar_obligaciones_semanticamente

logger = logging.getLogger(__name__)

# Cola secuencial (Semáforo FIFO = 1 informe a la vez) para proteger la cuenta OAuth de Hermes Agent
_HERMES_FIFO_LOCK = threading.Lock()


FEW_SHOT_STYLE_REFERENCE = """
ESTÁNDAR DE EXTENSIÓN, PROFUNDIDAD Y RIGOR TÉCNICO EN CONTRATACIÓN PÚBLICA:
CRITERIO DE CALIDAD: La supervisión técnica y las auditorías de contratación pública exigen redacciones EXTENSAS, EXHAUSTIVAS, PROFUNDAS Y DENSAS. Está TERMINANTEMENTE PROHIBIDO generar resúmenes cortos, viñetas de una sola línea o textos superficiales. Cada texto debe transmitir solvencia y rigor profesional.

1. En `texto_tabla_principal` (para la columna 'Actividades Desarrolladas' del formato oficial):
   - Estructura obligatoria (extensión objetivo: 250 a 450 palabras por obligación):
     * Párrafo introductorio formal: Contextualización del periodo reportado, vinculación al proyecto o área, marco institucional y directrices de supervisión.
     * Numerales técnicos detallados ('1. Título Técnico: ...', '2. Título Técnico: ...', '3. Título Técnico: ...'): Cada numeral debe ser un bloque denso de 4 a 7 líneas que detalle componentes intervenidos, metodologías aplicadas, patrones de trabajo, validaciones o actas de sesión técnica.
     * Párrafo de impacto o cierre técnico: Certificación de la estabilidad operativa, trazabilidad transaccional, calidad del trabajo y aseguramiento de directrices institucionales y contractuales.
   - PROHIBIDO incluir etiquetas como `<para>`, `</para>`, `<br>` o corchetes angulares `<` `>` sin escapar, para compatibilidad 100% con procesadores de texto y plataformas de reporte.

2. En `desarrollo_detallado_oti` (para `REPORTE DE OBLIGACIONES CONTRACTUALES.docx` de cada `Actividad_K`):
   - `parrafos_introductorios`: Mínimo 2 o 3 párrafos formales extensos y bien articulados que desarrollen el contexto técnico, la necesidad institucional y los lineamientos de la entidad contratante.
   - `vinetas_desarrollo`: Mínimo 3 a 5 viñetas técnicas sustanciosas (`{"titulo": "Título Técnico en Negrita", "descripcion": "Explicación exhaustiva de 5 a 8 líneas detallando la implementación, metodología de desarrollo, lógica de negocio aplicada y pruebas ejecutadas."}`).
   - `resultados_entregables`: 2 a 4 productos o entregables verificables bajo la sección 'Resultados Obtenidos, Productos o entregables generados', describiendo qué se produjo y su beneficio técnico.
   - `anexos_secciones`: Desglose pormenorizado de las categorías de anexos citando la ruta NAS oficial (`\\0. Evidencias\\Pago <XX>\\Actividad_<K>\\Anexos\\`) y el documento consolidado `Obligacion <K>.docx`.

3. En `bloques_evidencias` (para `Actividad_<K>/Anexos/Obligacion <K>.docx`):
   - Lista numerada de evidencias donde cada bloque debe tener una descripción técnica profunda de 4 a 6 líneas explicando qué demuestra la captura o documento, su fecha, contexto procesal o funcional y su impacto comprobable en el proyecto.
"""


def clasificar_obligacion_por_afinidad(
    texto_item: str,
    detalles_lista: List[Any],
    indice_rotacion: int = 0,
):
    """
    Analiza el nombre del archivo, su ruta dentro del .zip, sus metadatos ligeros (encabezados Excel,
    párrafos de Word, funciones .py/.sql) o una línea de apunte/commit y determina en cuál de las
    N obligaciones del contratista encaja mejor según el texto de cada cláusula.
    """
    if not detalles_lista:
        return None

    txt = (texto_item or "").lower()

    # 1. ¿El nombre del archivo, carpeta del ZIP o apunte menciona explícitamente "obligacion K", "obK", "act_K", "actividad K"?
    m_exp = re.search(r"(?:obligaci[oó]n|ob|actividad|act)[\s_\-]*(\d{1,2})\b", txt)
    if m_exp:
        num_buscado = int(m_exp.group(1))
        for det in detalles_lista:
            if det.obligacion.numero_obligacion == num_buscado:
                return det

    # La similitud local aporta contexto semántico para nombres genéricos,
    # documentos multidisciplinares y obligaciones que no comparten palabras exactas.
    # Las reglas explícitas de abajo siguen teniendo prioridad en casos inequívocos.
    puntajes_semanticos = puntuar_obligaciones_semanticamente(
        texto_item,
        detalles_lista,
    )

    # 2. Detección de tags semánticos extraídos por OCR o PDF
    es_reunion = any(t in txt for t in (
        "tag_semantico': 'reunion_seguimiento'", "reunion_seguimiento", "asistencia", "participantes",
        "resumen de la reunión", "resumen de la reunion", "unirse a la reunión", "salas para sesión",
        "teams", "convocatoria de reunión"
    ))
    es_desarrollo_git = any(t in txt for t in (
        "tag_semantico': 'software_git'", "software_git", "merge request", "hotfix", "commit ", "commits", "gitlab", "branch"
    ))
    es_tablero_agil = any(t in txt for t in (
        "tag_semantico': 'tablero_agil'", "tablero_agil", "kanban", "closed", "in progress", "sprint"
    ))
    es_cronograma = any(t in txt for t in (
        "tag_semantico': 'cronograma_planeacion'", "cronograma_planeacion", "cronograma", "gantt", "plan de trabajo", "plant120"
    ))
    es_arquitectura = any(t in txt for t in (
        "tag_semantico': 'arquitectura_diseno'", "arquitectura_diseno", "das_hld", "das/", "hld", "arquitectura de software", "diseño de alto nivel"
    ))
    es_seguridad = any(t in txt for t in (
        "tag_semantico': 'seguridad_inmutabilidad'", "seguridad_inmutabilidad", "sha-256", "control de evidencias", "inmutabilidad", "seguridad_datos"
    ))
    es_permisos = any(t in txt for t in (
        "tag_semantico': 'permisos_sistemas'", "permisos_sistemas", "permisos de administrador", "cooperapp", "solicitud de permisos"
    ))

    # Regla estricta: Las reuniones y asistencias NUNCA van en desarrollo, mantenimiento o arquitectura
    candidatos = list(detalles_lista)
    if es_reunion:
        candidatos_reunion = [
            d for d in detalles_lista
            if any(k in (d.obligacion.descripcion_clausula or "").lower() for k in (
                "asistir", "reunion", "reunión", "comité", "comite", "seguimiento", "mesa", "socializacion", "socialización"
            ))
        ]
        if candidatos_reunion:
            candidatos = candidatos_reunion

    # Mapa de familias semánticas -> palabras clave de cláusulas contractuales
    familias_semanticas = [
        (
            ("cronograma", "plan", "plant120", "alcance", "estimacion", "esfuerzo", "planeacion", "calendario", "hito"),
            ("cronograma", "estimar", "esfuerzo", "planificad", "compilar informes"),
        ),
        (
            ("das", "hld", "diseño", "diseno", "arquitectura", "ciclo de vida", "componente", "integracion", "modelo"),
            ("arquitectura", "ciclo de vida", "desarrollar e integrar", "escalabilidad", "mantenibilidad"),
        ),
        (
            ("seguridad", "control de evidencias", "seguridad_datos", "parametrizacion", "fortalecimiento", "configuracion", "blindaje", "owasp", "sha-256", "inmutabilidad"),
            ("parametrización", "fortalecimiento", "implementación", "calidad", "seguridad"),
        ),
        (
            ("solicitud", "permisos", "administrador", "cooperapp", "outlook", "correo", "mesa de servicios", "glpi", "ticket", "requerimiento", "incidencia"),
            ("mesa de servicios", "correo electrónico", "solicitudes", "incidencias", "requerimientos", "gestionar de manera oportuna"),
        ),
        (
            ("git", "commit", "rama", "branch", "merge", "mezcla", "pr", "pull request", "repositorio", ".py", ".js", ".ts", "def ", "class "),
            ("repositorio", "git", "ramas", "código fuente", "mezclas", "documentación estructurada del código"),
        ),
        (
            ("hotfix", "fix", "bug", "error", "fallo", "correctivo", "mantenimiento", "mejora continua", "actualizacion", "terminacion anticipada"),
            ("mantenimiento correctivo", "mejora continua", "atención ágil de fallos", "continuidad operativa"),
        ),
        (
            ("socializacion", "transferencia", "capacitacion", "acompañamiento", "apropiacion", "usuario", "manual", "guia", "orfeo", "expedientes digitales"),
            ("socialización", "transferencia de conocimiento", "acompañamiento", "uso y apropiación", "usuarios finales"),
        ),
        (
            ("seguimiento", "tablero", "herramienta de seguimiento", "avance", "sprint", "jira", "azure", "kanban"),
            ("seguimientos programados", "herramienta de seguimiento", "actualizar las actividades"),
        ),
        (
            ("reunion", "mesa", "comite", "equipo", "sesion", "teams", "meet", "virtual", "presencial", "acta", "modulaciones"),
            ("asistir a reuniones", "modalidad virtual o presencial", "reuni", "comité"),
        ),
        (
            ("incidente", "regimen tributario", "tributario", "retencion", "rol consulta", "soporte", "mesa técnica"),
            ("soporte", "incidentes", "mesa de servicios", "solicitudes", "mantenimiento"),
        ),
        (
            ("bd", "sql", "query", "consulta", "tabla", "migracion", "excel", "xlsx", "csv", "datos", "reporte", "estadistica"),
            ("base de datos", "datos", "información", "reporte", "consulta", "estructur", "indicador"),
        ),
        # Familias Semánticas Multidisciplinares Universales (Jurídica, Catastral, Social, Administrativa)
        (
            ("demanda", "fallo", "tutela", "juzgado", "tribunal", "auto", "recurso", "apelacion", "concepto", "notificacion", "expediente", "sentencia", "acto administrativo", "resolucion"),
            ("jurídic", "legal", "judicial", "fallo", "sentencia", "concepto", "defensa", "tutela", "procesal", "normativ", "sustanciación"),
        ),
        (
            ("predio", "catastro", "plano", "georreferenciacion", "ortofoto", "coordenadas", "lindero", "poligono", "shx", "shp", "sig", "cartografia", "topograf", "geodatabase", "inspeccion ocular"),
            ("predio", "catastral", "topográfic", "cartograf", "georreferenc", "geográfica", "inspección", "lindero", "espacial"),
        ),
        (
            ("taller", "comunidad", "victima", "beneficiario", "encuesta", "acta de socializacion", "etnico", "asamblea", "caracterizacion", "atencion al ciudadano", "psicosocial"),
            ("social", "comunitari", "víctima", "atención", "socialización", "participación", "beneficiari", "étnic", "acompañamiento"),
        ),
        (
            ("factura", "cuenta de cobro", "pago", "presupuesto", "cdp", "crp", "pac", "archivo", "tabla de retencion", "trd", "radicado", "orfeo", "memorando", "inventario documental"),
            ("financier", "presupuest", "contabl", "archivo", "documental", "administrativ", "gestión documental", "radicación", "trd"),
        ),
    ]

    mejor_det = None
    mejor_puntaje = 0

    palabras_item = set(re.findall(r"[a-záéíóúñ]{4,}", txt))
    # Ignorar palabras demasiado genéricas para evitar falsos positivos
    stopwords_genericas = {"para", "como", "este", "esta", "documento", "archivo", "soporte", "evidencia", "sistema", "entrega", "informe", "gestion", "actividad"}
    palabras_item = palabras_item - stopwords_genericas

    # Priorizar obligaciones que tengan cláusulas reales (>= 20 caracteres)
    candidatos_validos = [
        d for d in candidatos
        if len((d.obligacion.descripcion_clausula or "").strip()) >= 20
    ]
    pool_candidatos = candidatos_validos if candidatos_validos else candidatos

    for det in pool_candidatos:
        clausula_low = (det.obligacion.descripcion_clausula or "").lower()
        palabras_clausula = set(re.findall(r"[a-záéíóúñ]{4,}", clausula_low)) - stopwords_genericas

        # Coincidencia directa de palabras significativas entre el archivo/apunte y la cláusula
        comunes = palabras_item.intersection(palabras_clausula)
        puntaje = len(comunes) * 3

        # Coincidencia por familias semánticas OTI
        for triggers_item, triggers_clausula in familias_semanticas:
            hits_item = sum(1 for t in triggers_item if t in txt)
            hits_clausula = sum(1 for tc in triggers_clausula if tc in clausula_low)
            if hits_item > 0 and hits_clausula > 0:
                puntaje += (hits_item * 5) + (hits_clausula * 4)

        # Reglas contextuales de alta prioridad por tipo de activo
        if es_arquitectura and any(k in clausula_low for k in ("arquitectura", "desarrollar e integrar")):
            puntaje += 30
        if es_desarrollo_git:
            if "hotfix" in txt or "terminacion anticipada" in txt:
                if any(k in clausula_low for k in ("mantenimiento", "correctivo", "fallos")):
                    puntaje += 35
            elif any(k in clausula_low for k in ("desarrollar", "integrar", "componentes")):
                puntaje += 30
        if es_tablero_agil and any(k in clausula_low for k in ("herramienta de seguimiento", "actualizar las actividades")):
            puntaje += 30
        if es_cronograma and any(k in clausula_low for k in ("cronograma", "estimar el esfuerzo", "informes")):
            puntaje += 30
        if es_permisos and any(k in clausula_low for k in ("mesa de servicios", "correo", "solicitudes")):
            puntaje += 30
        if es_seguridad and any(k in clausula_low for k in ("seguridad", "calidad", "parametrización")):
            puntaje += 30
        if es_reunion:
            if "modulaciones" in txt and any(k in clausula_low for k in ("reuniones", "comité", "presencial o virtual")):
                puntaje += 35
            elif "orfeo" in txt and any(k in clausula_low for k in ("socialización", "transferencia", "usuarios")):
                puntaje += 35
            elif any(k in clausula_low for k in ("seguimientos programados", "actualizar las actividades")):
                puntaje += 25

        # Penalizar ligeramente obligaciones que ya tienen muchos archivos para balancear cuando hay empate
        similitud = (puntajes_semanticos or {}).get(det.obligacion.numero_obligacion, 0.0)
        # La similitud se usa como señal adicional, sin desplazar coincidencias
        # explícitas de etiquetas, rutas o familias documentales.
        puntaje_ajustado = puntaje + max(0.0, similitud) * 12 - (det.archivos_evidencia.count() * 0.5)

        if puntaje > 0 and puntaje_ajustado > mejor_puntaje:
            mejor_puntaje = puntaje_ajustado
            mejor_det = det

    if mejor_det is not None and mejor_puntaje > 0:
        return mejor_det

    # Si es un archivo genérico, balancear únicamente dentro de los candidatos compatibles
    detalles_ordenados = sorted(
        pool_candidatos,
        key=lambda d: (d.archivos_evidencia.count(), d.obligacion.numero_obligacion),
    )
    return detalles_ordenados[indice_rotacion % len(detalles_ordenados)]


def _construir_prompt_single_shot(informe, detalles_lista: List[Any], apuntes_generales_mes: str = "") -> str:
    """
    Construye un prompt Single-Shot JSON ultra-optimizado en tokens:
    - Sin espacios ni indentación superflua en los payloads JSON.
    - Sin duplicar listas de archivos ni rutas NAS.
    - No le pide al LLM repetir `anexos_secciones` (se construye en Python a partir de `bloques_evidencias` con 0 tokens).
    """
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
            "clausula": ob.descripcion_clausula,
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
Objeto: {c.objeto_contrato}
{bloque_apuntes_globales}
ARCHIVOS SOPORTE SUBIDOS (Asócialos en `bloques_evidencias` de su obligación más afín usando el nombre exacto en `archivo_asociado`):
{json.dumps(bolsa_global_archivos, ensure_ascii=False, separators=(",", ":"))}

{FEW_SHOT_STYLE_REFERENCE}

OBLIGACIONES A REDACTAR:
{json.dumps(items_payload, ensure_ascii=False, separators=(",", ":"))}

VALIDACIÓN DE COHERENCIA OBLIGACIÓN vs EVIDENCIA:
- Si el texto de `clausula` es inválido/prueba (ej. 'primer', 'segundo', 'test') o NO guarda ninguna relación lógica ni técnica con los archivos soporte o apuntes, NO inventes actividades falsas: inicia `texto_tabla_principal` con "[ADVERTENCIA DE COHERENCIA: La cláusula registrada ('<clausula>') no especifica un alcance técnico compatible con las evidencias adjuntas. Por favor actualice el texto literal de su obligación contractual.]" seguido de la descripción objetiva de los archivos subidos.

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
    """Extrae y parsea el objeto JSON devuelto por el LLM incluso si viene envuelto en ```json ... ```."""
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
    """
    Conecta directamente en Python con el transporte oficial `openai-codex` de Hermes Agent
    leyendo la sesión OAuth local de `C:\\Users\\ACER\\AppData\\Local\\hermes\\auth.json` (sin abrir
    subprocesos lentos de terminal ni activar el bucle de herramientas de CLI).
    Respeta estrictamente la cola secuencial FIFO (`_HERMES_FIFO_LOCK` = 1 informe a la vez)
    y usa `gpt-6.1-sol` con `reasoning={"effort": "low"}` para responder en pocos segundos.
    """
    import base64

    hermes_auth_path = os.path.expandvars(r"%LOCALAPPDATA%\hermes\auth.json")
    hermes_model = getattr(settings, "HERMES_MODEL", "gpt-6.1-sol")
    hermes_reasoning = getattr(settings, "HERMES_REASONING", "low")

    with _HERMES_FIFO_LOCK:
        # A. Transporte Directo Python -> Hermes OAuth Codex (`https://chatgpt.com/backend-api/codex`)
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
                            residency = jwt_auth.get("chatgpt_data_residency") or jwt_auth.get(
                                "chatgpt_compute_residency"
                            )
                            if residency:
                                headers["x-openai-internal-codex-residency"] = residency.strip()
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
                            "Responde EXCLUSIVAMENTE con un objeto JSON válido, sin usar herramientas ni bloques adicionales."
                        ),
                        input=[
                            {
                                "role": "user",
                                "content": [{"type": "input_text", "text": prompt}],
                            }
                        ],
                        store=False,
                        reasoning={"effort": hermes_reasoning},
                        stream=True,
                    )
                    partes_texto = []
                    for ev in stream:
                        ev_type = getattr(ev, "type", "")
                        if ev_type == "response.output_text.delta":
                            partes_texto.append(getattr(ev, "delta", ""))
                    texto_completo = "".join(partes_texto).strip()
                    if texto_completo and "{" in texto_completo:
                        return _extraer_json_limpio(texto_completo)
            except Exception as exc:
                logger.info("Hermes OAuth Codex directo no retornó JSON (%s). Probando siguiente opción...", exc)

        # B. Modo Servidor HTTP Compatible con OpenAI (si se configuró HERMES_BASE_URL)
        base_url = getattr(settings, "HERMES_BASE_URL", "")
        if base_url:
            try:
                from openai import OpenAI

                client = OpenAI(
                    base_url=base_url,
                    api_key=getattr(settings, "HERMES_API_KEY", "hermes-local"),
                    timeout=15.0,
                    max_retries=0,
                )
                resp = client.chat.completions.create(
                    model=hermes_model,
                    messages=[
                        {
                            "role": "system",
                            "content": "Eres un redactor técnico experto en contratación pública y cuentas de cobro. Responde únicamente en JSON válido.",
                        },
                        {"role": "user", "content": prompt},
                    ],
                    temperature=0.2,
                )
                contenido = resp.choices[0].message.content
                return _extraer_json_limpio(contenido)
            except Exception as exc:
                logger.info("Hermes HTTP local no disponible (%s). Activando Fallback...", exc)

    return None


def _llamar_gemini_fallback(prompt: str) -> Optional[dict]:
    """Fallback automático #1: Google Gemini API (`google-genai` SDK)."""
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
        logger.warning("Fallback Gemini API falló (%s). Probando siguiente respaldo...", exc)
        return None


def _llamar_openai_fallback(prompt: str) -> Optional[dict]:
    """Fallback automático #2: OpenAI API directa (`gpt-4o-mini`)."""
    api_key = getattr(settings, "OPENAI_API_KEY", "")
    if not api_key:
        return None
    try:
        from openai import OpenAI

        client = OpenAI(api_key=api_key, timeout=25.0)
        resp = client.chat.completions.create(
            model=getattr(settings, "OPENAI_MODEL", "gpt-4o-mini"),
            messages=[
                {
                    "role": "system",
                    "content": "Eres un redactor técnico experto en contratación pública y cuentas de cobro. Responde únicamente en JSON válido.",
                },
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.2,
        )
        return _extraer_json_limpio(resp.choices[0].message.content)
    except Exception as exc:
        logger.warning("Fallback OpenAI API falló (%s). Activando Motor Demo/Offline Inteligente...", exc)
        return None


def _limpiar_fragmento_significativo(lineas: List[str], nombre_archivo: str, nombre_contratista: str = "") -> str:
    """
    Extrae el título o concepto técnico limpio de un archivo o captura:
    - Si es un correo PDF de Outlook ('RE_ ...', '- Outlook.pdf'), extrae el Asunto Técnico limpio del nombre.
    - Limpia dinámicamente el nombre del contratista si aparece en el nombre del archivo.
    - Ignora saludos ('Cordial saludo'), nombres de remitentes con sílabas rotas ('Cl audi a...'),
      encabezados de tablas de MS Project ('Nombre de tarea Duración Comienzo Fin Predecesoras') y numeración de índices ('1. PROPÓSITO').
    - Balancea paréntesis para jamás dejar '(DAS' sin cerrar.
    """
    base = os.path.splitext(nombre_archivo)[0]

    # Armar patrón dinámico con los nombres/apellidos del contratista actual para limpiarlos de los nombres de archivo
    patron_persona = ""
    if nombre_contratista:
        partes_nom = [re.escape(p) for p in nombre_contratista.lower().split() if len(p) > 2]
        if len(partes_nom) >= 2:
            patron_persona = r"[\s_\-]*" + r"[\s_\-]+".join(partes_nom[:3]) + r".*$"

    # 1. Si el archivo es un correo exportado de Outlook (ej. 'RE_ Entrega documento... - Outlook.pdf')
    # el mejor resumen técnico es el ASUNTO limpio del correo, no el 'Cordial saludo' ni el remitente.
    if re.search(r"(^re[\s_]|^rv[\s_]|^fw[\s_]|outlook|outlo)", base.lower()):
        asunto = re.sub(r"^(re|rv|fw|fwd)[\s_:\-]+", "", base, flags=re.I)
        if patron_persona:
            asunto = re.sub(patron_persona, "", asunto, flags=re.I)
        asunto = re.sub(r"[\s_\-]*(outlook|outlo).*$", "", asunto, flags=re.I)
        asunto = re.sub(r"[_\-]+", " ", asunto).strip(" -–—.,;:")
        if len(asunto) > 8:
            return sanitizar_texto_para_reporte(asunto[:95])

    # 2. Si es un archivo de Cronograma/Plan (ej. 'Cronograma_General27Octubre2026.pdf' o 'PlanT120_Alcance2...'),
    # usar el nombre limpio del cronograma en vez de la fila de columnas 'Nombre de tarea Duración Comienzo...'
    if any(k in base.lower() for k in ("cronograma", "plant120", "plan_trabajo", "gantt")):
        nom_cron = re.sub(r"[_\-]+", " ", base).strip()
        return sanitizar_texto_para_reporte(nom_cron[:90])

    ruido_subcadenas = (
        "cordial saludo", "cor di al", "remito adjunto", "r emi t o", "nombre de tarea",
        "predecesoras", "duración comienzo", "propósito y alcance", "proposito y alcance",
        "cl audi a", "para:", "de:", "asunto:", "enviado el:", "unidad administrativa",
        "código: gc-fo", "versión:", "página ",
    )

    for ln in (lineas or []):
        limpia = sanitizar_texto_para_reporte(str(ln))
        # Quitar numeración inicial tipo '1. ' o '1.1 '
        limpia = re.sub(r"^\d+(\.\d+)*[\.\)\-\s]+", "", limpia).strip()
        limpia = re.sub(r"\s+", " ", limpia).strip(" -•*:.,;|/\\[]")

        # Detectar texto corrupto de PDF con sílabas separadas (ej. 'Cl audi a P eña' -> muchas palabras de 1-2 letras)
        palabras_todas = limpia.split()
        palabras_cortas = [p for p in palabras_todas if len(p) <= 2]
        if len(palabras_todas) >= 5 and (len(palabras_cortas) / len(palabras_todas)) > 0.45:
            continue

        if len(limpia) < 14 or len(palabras_todas) < 3:
            continue
        if any(r in limpia.lower() for r in ruido_subcadenas):
            continue
        if re.match(r"^(lun|mar|mié|mie|jue|vie|sáb|sab|dom)?\s*\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}", limpia.lower()):
            continue

        recorte = limpia[:95]
        # Si quedó un paréntesis abierto sin cerrar, cerrarlo o limpiarlo
        if "(" in recorte and ")" not in recorte:
            recorte = recorte + ")"
        return sanitizar_texto_para_reporte(recorte)

    # 3. Fallback al nombre del archivo si no es un nombre genérico ('imagen (3).png', 'Captura de pantalla...')
    if re.match(r"^(imagen|image|captura|screenshot|img|paste|foto|whatsapp)[\s_\-\(\d\)]*", base.lower()):
        return ""
    base_limpia = base
    if patron_persona:
        base_limpia = re.sub(patron_persona, "", base_limpia, flags=re.I)
    base_limpia = re.sub(r"[_\-]+", " ", base_limpia).strip(" -–—.,;:")
    return sanitizar_texto_para_reporte(base_limpia[:90]) if len(base_limpia) > 6 else ""


def _contraer_al(texto: str) -> str:
    """Corrige la contracción castellana en frases ensambladas por plantillas."""
    return re.sub(r"\ba el\b", "al", texto, flags=re.IGNORECASE)


def _inferir_contexto_clausula(
    clausula: str,
    num_ob: int,
    objeto_contrato: str = "",
) -> Dict[str, str]:
    """
    Analiza la cláusula contractual del contratista para determinar el área disciplinar
    (Jurídica, Catastral/SIG, Social, Administrativa/Financiera o Tecnológica), adaptando el título
    temático, la profesión y la conclusión al alcance real de la obligación.
    """
    txt = (clausula or "").strip()
    low = txt.lower()
    ctx_global = f"{low} {(objeto_contrato or '').lower()}"
    es_invalida = len(txt) < 18 or len(txt.split()) < 3

    if es_invalida:
        return {
            "es_invalida": True,
            "tipo_profesion": "gestión profesional y técnica",
            "tema_corto": f"Ejecución de la Obligación N° {num_ob}",
            "enfoque": (
                f"[NOTA DE AUDITORÍA: El texto registrado para la Obligación {num_ob} ('{txt}') es un marcador de prueba. "
                f"Se recomienda registrar la cláusula literal del contrato para garantizar la trazabilidad semántica.] "
            ),
            "accion_cierre": "el cumplimiento de los compromisos asignados en el periodo",
        }

    # Inferencia de Tipo de Profesión / Disciplina
    if any(k in ctx_global for k in ("software", "sistemas de información", "código fuente", "git", "base de datos", "arquitectura", "desarrollo")):
        tipo_profesion = "gestión tecnológica y de sistemas"
    elif any(k in ctx_global for k in ("demanda", "fallo", "tutela", "juzgado", "tribunal", "auto", "recurso", "apelación", "apelacion", "jurídic", "juridic", "legal", "judicial", "sentencia", "concepto", "defensa", "procesal", "normativ", "sustanciación", "sustanciacion")):
        tipo_profesion = "gestión jurídica y legal"
    elif any(k in low for k in ("predio", "catastro", "catastral", "plano", "georreferenc", "ortofoto", "coordenadas", "lindero", "polígono", "poligono", "cartograf", "topográf", "geodatabase", "inspección ocular", "inspeccion ocular", "espacial")):
        tipo_profesion = "gestión técnico-predial y catastral"
    elif any(k in low for k in ("taller", "comunidad", "víctima", "victima", "beneficiario", "socialización", "socializacion", "étnico", "etnico", "asamblea", "caracterización", "caracterizacion", "atención al ciudadano", "psicosocial", "comunitari")):
        tipo_profesion = "intervención social y comunitaria"
    elif any(k in low for k in ("factura", "cuenta de cobro", "presupuest", "contabl", "financier", "cdp", "crp", "archivo", "tabla de retención", "tabla de retencion", "trd", "radicado", "orfeo", "gestión documental", "gestion documental")):
        tipo_profesion = "gestión administrativa y financiera"
    elif any(k in low for k in ("software", "desarrollo", "sistema", "tecnolog", "arquitectura", "base de datos", "código", "codigo", "git", "plataforma", "incidente", "redes", "servidor", "aplicativo")):
        tipo_profesion = "gestión tecnológica y de sistemas"
    else:
        tipo_profesion = "gestión profesional y técnica"

    # Clasificación Temática y Acción de Cierre. Las obligaciones transversales
    # tienen prioridad para no contaminarse con el área del contrato.
    if any(k in low for k in ("asistir a reuniones", "participar en reuniones", "asistir a las reuniones", "reuniones en modalidad")):
        tema = "Participación en Reuniones de Coordinación, Mesas Técnicas y Comités de Seguimiento"
        accion = "la articulación interdisciplinar, asistencia a sesiones de trabajo y seguimiento de compromisos operativos"
    elif any(k in low for k in ("seguimientos programados", "herramienta de seguimiento", "actualizar las actividades")):
        tema = "Seguimiento de Actividades y Actualización en la Herramienta Institucional de Control"
        accion = "el reporte continuo de tareas, trazabilidad de avances y actualización de las herramientas de seguimiento"
    elif any(k in low for k in ("demás actividades", "demas actividades", "encomendadas", "necesarias para garantizar el cumplimiento")):
        tema = "Actividades Transversales de Apoyo a la Gestión y Cumplimiento Contractual"
        accion = "el soporte transversal a las iniciativas del área y el aseguramiento integral de las metas del contrato"
    elif any(k in low for k in ("cronograma", "estimar esfuerzos", "actividades planificadas", "compilar informes", "plan de acción")):
        tema = "Planeación Técnica, Estimación de Esfuerzo y Seguimiento de Cronogramas de Trabajo"
        accion = "el control de hitos, monitoreo de capacidad técnica y consolidación de informes de avance"

    # 1. Clasificación específica del área
    elif any(k in low for k in ("demanda", "fallo", "tutela", "juzgado", "tribunal", "auto", "recurso", "apelación", "apelacion", "jurídic", "juridic", "legal", "judicial", "sentencia", "concepto", "defensa", "procesal", "normativ", "sustanciación", "sustanciacion")):
        # Rama Jurídica / Legal
        tema = "Defensa Judicial, Sustanciación Procesal y Atención de Acciones Constitucionales"
        accion = "la oportuna defensa judicial, sustanciación técnica de memoriales y atención estricta de términos procesales"
    elif any(k in low for k in ("defensa judicial", "tutela", "juzgado", "tribunal", "demanda", "audiencia", "procesal")):
        tema = "Defensa Judicial, Sustanciación Procesal y Atención de Acciones Constitucionales"
        accion = "la oportuna defensa judicial, sustanciación técnica de memoriales y atención estricta de términos procesales"
    elif any(k in low for k in ("concepto jurídico", "concepto legal", "proyectar actos", "resolución", "acto administrativo", "estudio jurídico")):
        tema = "Emisión de Conceptos Jurídicos, Control de Legalidad y Proyección de Actos Administrativos"
        accion = "el análisis normativo riguroso, sustanciación jurídica y control de legalidad en las decisiones institucionales"
    
    # 2. Rama Catastral / Topográfica / SIG
    elif any(k in low for k in ("topograf", "levantamiento", "lindero", "coordenadas", "inspección ocular", "plano")):
        tema = "Levantamiento Topográfico, Validación de Linderos e Inspección Predial"
        accion = "la determinación física exacta de linderos, verificación espacial y consistencia de la información predial"
    elif any(k in low for k in ("georreferenc", "cartograf", "geodatabase", "sig", "análisis espacial", "ortofoto")):
        tema = "Análisis Geográfico, Estructuración de Geodatabases y Cartografía Digital"
        accion = "la consolidación de capas espaciales, consistencia geográfica y confiabilidad de los datos catastrales"

    # 3. Rama Social / Comunitaria / Víctimas
    elif any(k in low for k in ("caracterización", "caracterizacion", "víctima", "victima", "beneficiario", "psicosocial", "enfoque diferencial")):
        tema = "Acompañamiento Social, Caracterización de Beneficiarios y Atención Diferencial"
        accion = "la atención integral a beneficiarios, garantía de enfoque diferencial y levantamiento participativo de información"
    elif any(k in low for k in ("taller", "asamblea", "comunidad", "socialización comunitaria", "concertación", "concertacion")):
        tema = "Socialización Comunitaria, Realización de Talleres y Articulación con Comunidades"
        accion = "la participación ciudadana efectiva, concertación en territorio y apropiación de los procesos comunitarios"

    # 4. Rama Administrativa / Financiera / Archivo
    elif any(k in low for k in ("presupuest", "cdp", "crp", "financier", "contabl", "cuenta de cobro", "pago")):
        tema = "Gestión Presupuestal, Trámite Financiero y Seguimiento Contable"
        accion = "la correcta ejecución presupuestal, revisión documental de pagos y aseguramiento financiero de los procesos"
    elif any(k in low for k in ("trd", "archivo", "gestión documental", "gestion documental", "expediente", "foliación", "custodia")):
        tema = "Gestión Documental, Organización de Expedientes y Aplicación de Tablas de Retención"
        accion = "la debida custodia, foliación, organización técnica y conservación del acervo documental"

    # 5. Rama Tecnológica / Sistemas
    elif any(k in low for k in ("mantenimiento correctivo", "mejora continua", "actualización de las soluciones")):
        tema = "Mantenimiento Correctivo, Evolutivo y Mejora Continua de Soluciones de Software"
        accion = "la estabilización funcional, remediación de defectos y mejora continua de los aplicativos institucionales"
    elif any(k in low for k in ("parametrización", "fortalecimiento técnico", "colaborar en la implementación")):
        tema = "Parametrización, Implementación y Fortalecimiento Técnico de Sistemas de Información"
        accion = "el fortalecimiento técnico, configuración segura y optimización de los componentes tecnológicos de la dependencia supervisora"
    elif any(k in low for k in ("asistencia técnica", "atención ágil de fallos", "continuidad operativa")):
        tema = "Asistencia Técnica Oportuna, Atención de Fallos y Continuidad Operativa"
        accion = "la atención ágil de incidentes, aplicación de soluciones efectivas y continuidad operativa de las plataformas"
    elif any(k in low for k in ("socialización", "transferencia de conocimiento", "acompañamiento a equipos", "uso y apropiación")):
        tema = "Socialización Técnica, Transferencia de Conocimiento y Acompañamiento a Usuarios"
        accion = "la transferencia efectiva de conocimiento, apropiación tecnológica y acompañamiento funcional a las áreas"
    elif any(k in low for k in ("cronograma", "estimar esfuerzos", "actividades planificadas", "compilar informes")):
        tema = "Planeación Técnica, Estimación de Esfuerzo y Seguimiento de Cronogramas de Trabajo"
        accion = "el control de hitos, monitoreo de capacidad técnica y consolidación de informes de avance"
    elif any(k in low for k in ("repositorio git", "ramas asignadas", "documentación estructurada del código", "mezclas")):
        tema = "Control de Versiones en Repositorio, Gestión de Ramas y Documentación de Código Fuente"
        accion = "la integración continua mediante Merge Requests, versionamiento seguro y documentación estructurada del código"
    elif any(k in low for k in ("seguimientos programados", "herramienta de seguimiento", "actualizar las actividades")):
        tema = "Seguimiento de Proyectos y Actualización de Actividades en Herramienta de Control"
        accion = "la gestión oportuna de tareas, trazabilidad de entregables y actualización del tablero de seguimiento"
    elif any(k in low for k in ("mesa de servicios", "correo electrónico", "requerimientos, solicitudes e incidencias")):
        tema = "Gestión de Requerimientos, Solicitudes e Incidencias Operativas"
        accion = "el diagnóstico técnico, atención oportuna de solicitudes de servicio y resolución efectiva de anomalías"
    elif any(k in low for k in ("asistir a reuniones", "modalidad virtual o presencial", "demandas del servicio")):
        tema = "Participación en Mesas Técnicas, Sesiones de Coordinación y Comités de Seguimiento"
        accion = "la articulación interdisciplinar, participación en sesiones técnicas y seguimiento de compromisos operativos"
    elif any(k in low for k in ("inteligencia artificial", "aprendizaje automático", "lenguaje natural", "análisis predictivo")):
        tema = "Desarrollo y Optimización de Componentes Inteligentes y Procesamiento de Datos"
        accion = "el entrenamiento, ajuste e integración de modelos analíticos y automatización de procesos"
    elif any(k in low for k in ("validar", "historias de usuario", "criterios de aceptación", "pruebas de calidad")):
        tema = "Validación Funcional, Aseguramiento de Calidad (QA) y Verificación de Criterios de Aceptación"
        accion = "la verificación rigurosa de entregables, pruebas integrales y conformidad funcional"
    elif any(k in low for k in ("scripts de ajuste", "base de datos", "despliegue en los entornos")):
        tema = "Estructuración de Scripts de Base de Datos, Configuración de Entornos y Despliegue Técnico"
        accion = "la preparación de artefactos de persistencia, parametrización de entornos y verificación controlada"
    elif any(k in low for k in ("desarrollar e integrar", "arquitectura", "ciclo de vida", "escalabilidad y mantenibilidad")):
        tema = "Estructuración de Soluciones, Arquitectura de Componentes e Integración de Sistemas"
        accion = "el diseño modular, estructuración confiable e interoperabilidad de los componentes misionales"
    elif any(k in low for k in ("demás actividades", "encomendadas", "cumplimiento de las obligaciones")):
        tema = "Actividades Transversales de Apoyo a la Gestión y Cumplimiento Contractual"
        accion = "el soporte transversal a las iniciativas del área y el aseguramiento integral de las metas del contrato"
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

def _generar_redaccion_demo_inteligente(informe, detalles_lista: List[Any]) -> dict:
    """
    Motor Redactor Experto Local 0-Tokens (100% Gratuito y Autónomo):
    Analiza el texto real extraído por RapidOCR / PyPDF / python-docx / openpyxl, lo cruza con
    el alcance semántico de cada cláusula contractual y redacta los 20 documentos con variedad
    léxica, precisión técnica y formato de auditoría sin repetir párrafos genéricos.
    """
    resultado = {"obligaciones": {}}
    mes_minus = informe.nombre_mes.lower()
    anio = informe.anio
    num_pago = informe.numero_pago
    oficina_area = (informe.contratista.direccion_oficina or "la dependencia supervisora").strip()
    proyecto = (informe.contratista.proyecto_area or oficina_area).strip()
    nom_contratista = (informe.contratista.nombres_apellidos or "").strip()

    # Banco de plantillas léxicas variadas para evitar repetición entre numerales de imágenes/documentos
    variantes_verbos_soporte = [
        (
            "Ejecución y Validación Técnica de Componentes",
            "Se adelantaron las labores especializadas y verificación funcional orientadas a {accion_cierre}, "
            "dejando registro comprobable de la operación y estabilidad del componente intervenido{contexto_extra}."
        ),
        (
            "Aseguramiento Operativo y Verificación Funcional",
            "Se realizó el control técnico y comprobación práctica sobre el entorno de trabajo de {proyecto}, "
            "constatando el correcto procesamiento de reglas operativas y la trazabilidad de la información{contexto_extra}."
        ),
        (
            "Implementación y Seguimiento de Entregables",
            "Se desarrollaron las actuaciones previstas para el periodo de {mes} de {anio}, articulando los ajustes "
            "operativos, revisión de flujos y consolidación de evidencias de cumplimiento{contexto_extra}."
        ),
        (
            "Consolidación de Activos y Control de Calidad",
            "Se efectuó la revisión detallada y estructuración de los elementos asociados a la obligación contractual, "
            "garantizando su alineación con los estándares de calidad y lineamientos de {oficina_area}{contexto_extra}."
        ),
        (
            "Verificación de Procesos y Flujos de Información",
            "Se ejecutaron pruebas de validación sobre los procesos y herramientas de {proyecto}, "
            "confirmando la integridad de la información y el cumplimiento de los criterios institucionales establecidos{contexto_extra}."
        ),
    ]

    for det in detalles_lista:
        k = det.obligacion.numero_obligacion
        clausula = (det.obligacion.descripcion_clausula or "").strip()
        apuntes = (det.apuntes_contratista or "").strip()
        evidencias = list(det.archivos_evidencia.all())
        ctx_clausula = _inferir_contexto_clausula(
            clausula,
            k,
            objeto_contrato=informe.contratista.objeto_contrato,
        )

        # Extraer apuntes reales del contratista (sin meter frases rellenas repetidas si ya hay evidencias)
        lineas_apuntes = [
            re.sub(r"^[\-\•\*\d\.\)\s]+", "", ln.strip())
            for ln in re.split(r"[\n;]+", apuntes)
            if len(ln.strip()) > 5
        ]

        items_numerados = []
        vinetas_desarrollo = []
        resultados_entregables = []

        # 1. Redactar numerales ricos a partir de cada evidencia y su contenido interno (OCR / PDF / Word / Excel)
        for idx_ev, ev in enumerate(evidencias, start=1):
            meta_ev = ev.metadatos_ligeros or {}
            tag_ev = meta_ev.get("tag_semantico", "")
            lineas_ev = meta_ev.get("resumen_lineas") or []
            hojas_ev = meta_ev.get("hojas") or []
            cols_ev = meta_ev.get("encabezados_fila_1") or {}

            extracto_limpio = _limpiar_fragmento_significativo(lineas_ev, ev.nombre_original, nombre_contratista=nom_contratista)
            sub_ref = f" — {extracto_limpio}" if extracto_limpio else ""
            ctx_extra = f" (conforme a la evidencia '{extracto_limpio}' registrada en {ev.nombre_original})" if extracto_limpio else f" (soportado en el registro '{ev.nombre_original}')"

            if tag_ev == "SOFTWARE_GIT" or ev.tipo_archivo == "CODIGO":
                titulo_pto = f"Desarrollo de Software, Versionamiento y Control de Cambios{sub_ref}"
                desc_pto = (
                    f"Se ejecutaron actividades de codificación, estructuración de lógica e integración continua "
                    f"en el repositorio de la entidad{ctx_extra}. Se verificó la correcta aplicación de estándares de calidad, "
                    f"revisión de ramas de trabajo y consistencia técnica dentro de {proyecto}."
                )
            elif tag_ev == "ARQUITECTURA_DISENO":
                titulo_pto = f"Diseño Técnico, Estructuración y Especificación de Soluciones{sub_ref}"
                desc_pto = (
                    f"Se estructuró y documentó la definición técnica y arquitectónica de la solución{ctx_extra}, estableciendo los diagramas "
                    f"funcionales, interfaces, flujos de datos y lineamientos técnicos orientados a garantizar "
                    f"la escalabilidad, mantenibilidad y consistencia operativa en {oficina_area}."
                )
            elif tag_ev == "SEGURIDAD_INMUTABILIDAD":
                titulo_pto = f"Aseguramiento de Seguridad, Inmutabilidad y Control de Evidencias{sub_ref}"
                desc_pto = (
                    f"Se implementaron y verificaron los mecanismos de control, integridad documental (SHA-256) "
                    f"y trazabilidad de pistas de auditoría{ctx_extra}, asegurando la protección e inmutabilidad de los registros "
                    f"funcionales conforme a los lineamientos institucionales y de gestión de la entidad."
                )
            elif tag_ev == "REUNION_SEGUIMIENTO":
                titulo_pto = f"Participación en Mesas Técnicas y Sesiones de Seguimiento{sub_ref}"
                desc_pto = (
                    f"Se participó activamente en las sesiones de trabajo y mesas de coordinación interdisciplinar convocadas en el marco "
                    f"de {proyecto}{ctx_extra}, revisando el avance de compromisos, concertando definiciones "
                    f"y articulando acciones conjuntas con el equipo de {oficina_area}."
                )
            elif tag_ev == "TABLERO_AGIL":
                titulo_pto = f"Gestión Ágil de Tareas y Monitoreo de Avance en Tablero de Control{sub_ref}"
                desc_pto = (
                    f"Se actualizó y gestionó el flujo de actividades dentro de la herramienta de seguimiento{ctx_extra}, "
                    f"documentando el progreso de los compromisos, el cierre de tareas operativas y la priorización "
                    f"de entregables del periodo {mes_minus} de {anio}."
                )
            elif tag_ev == "CRONOGRAMA_PLANEACION":
                titulo_pto = f"Planeación Operativa, Estimación de Esfuerzo y Control de Cronograma{sub_ref}"
                desc_pto = (
                    f"Se formuló y actualizó la planificación detallada de actividades, hitos de entrega y distribución de cargas de trabajo"
                    f"{ctx_extra}, efectuando el seguimiento riguroso a los porcentajes de cumplimiento de {proyecto}."
                )
            elif tag_ev == "PERMISOS_SISTEMAS":
                titulo_pto = f"Gestión de Accesos, Roles y Configuración de Plataformas Institucionales{sub_ref}"
                desc_pto = (
                    f"Se adelantó la gestión y configuración de privilegios operativos sobre los entornos y herramientas "
                    f"misionales{ctx_extra}, garantizando el control de acceso seguro y la disponibilidad de los servicios requeridos."
                )
            elif tag_ev == "SOPORTE_INCIDENTES":
                titulo_pto = f"Diagnóstico Oportuno, Atención de Requerimientos y Soporte Especializado{sub_ref}"
                desc_pto = (
                    f"Se brindó atención oportuna a los requerimientos e incidentes reportados en la operación{ctx_extra}, "
                    f"aplicando los ajustes correctivos necesarios y comprobando el restablecimiento satisfactorio del servicio."
                )
            elif ev.tipo_archivo == "EXCEL" and (hojas_ev or cols_ev):
                muestra_cols = []
                for h_nom, c_list in list(cols_ev.items())[:2]:
                    if c_list:
                        muestra_cols.append(f"{h_nom} ({', '.join(c_list[:5])})")
                detalle_excel = f" con estructura de datos [{'; '.join(muestra_cols)}]" if muestra_cols else ""
                titulo_pto = f"Procesamiento, Estructuración y Validación de Datos ({ev.nombre_original})"
                desc_pto = (
                    f"Se consolidó y verificó la información contenida en el libro '{ev.nombre_original}' (hojas: {', '.join(hojas_ev[:4])})"
                    f"{detalle_excel}, asegurando la consistencia de los registros, cruces de validación e insumo analítico para {proyecto}."
                )
            else:
                # Rotar entre las 5 plantillas léxicas variadas usando el texto real extraído por OCR/PDF/Word
                plantilla_tit, plantilla_desc = variantes_verbos_soporte[(idx_ev - 1) % len(variantes_verbos_soporte)]
                titulo_pto = f"{plantilla_tit}{sub_ref}"
                desc_pto = plantilla_desc.format(
                    accion_cierre=ctx_clausula["accion_cierre"],
                    proyecto=proyecto,
                    oficina_area=oficina_area,
                    mes=mes_minus,
                    anio=anio,
                    contexto_extra=ctx_extra,
                )

            desc_pto = _contraer_al(desc_pto)
            items_numerados.append(f"{idx_ev}. {titulo_pto}: {desc_pto}")
            vinetas_desarrollo.append({
                "titulo": titulo_pto,
                "descripcion": (
                    f"{desc_pto} Esta actuación se desarrolló aplicando los estándares técnicos y operativos de {oficina_area} "
                    f"y quedó debidamente respaldada en el expediente digital de la Actividad_{k}."
                ),
            })
            resultados_entregables.append({
                "titulo": f"Producto Verificado {k}.{idx_ev} — {titulo_pto[:75]}",
                "descripcion": f"Entregable y soporte ('{ev.nombre_original}') que acredita: {desc_pto}",
            })

        # 2. Incorporar los apuntes escritos por el contratista (SIN duplicar el título antes y después de ':')
        for idx_p, punto in enumerate(lineas_apuntes, start=len(items_numerados) + 1):
            if ":" in punto and len(punto.split(":", 1)[0].split()) <= 10:
                partes = punto.split(":", 1)
                titulo_pto = partes[0].strip()
                desc_pto = partes[1].strip()
            else:
                titulo_pto = f"Gestión y Desarrollo — {ctx_clausula['tema_corto']}"
                desc_pto = punto

            desc_limpia = sanitizar_texto_para_reporte(_contraer_al(desc_pto))
            titulo_limpio = sanitizar_texto_para_reporte(titulo_pto)

            # Si el apunte era corto, expandirlo con redacción formal
            if len(desc_limpia) < 140:
                desc_limpia = (
                    f"{desc_limpia.rstrip('.')}. Actividad ejecutada en cumplimiento de la Obligación Contractual N° {k} "
                    f"dentro de {proyecto}, verificando el cumplimiento de los criterios de calidad, "
                    f"funcionalidad y trazabilidad definidos por la supervisión."
                )

            items_numerados.append(f"{idx_p}. {titulo_limpio}: {desc_limpia}")
            vinetas_desarrollo.append({
                "titulo": titulo_limpio,
                "descripcion": (
                    f"{desc_limpia} El desarrollo de esta tarea se articuló con los lineamientos de {oficina_area} "
                    f"para el periodo de {mes_minus} de {anio}."
                ),
            })
            resultados_entregables.append({
                "titulo": f"Entregable {k}.{idx_p} — {titulo_limpio[:75]}",
                "descripcion": f"Resultado verificable en el marco de la Obligación {k}: {desc_limpia}",
            })

        # 3. Si la obligación no tenía ni archivos ni apuntes aún, generar 2 numerales sólidos basados en su cláusula real
        if not items_numerados:
            t1 = f"Desarrollo de Actividades — {ctx_clausula['tema_corto']}"
            d1 = (
                f"Se adelantaron las actividades correspondientes a la cláusula '{clausula[:140]}', "
                f"orientadas a {ctx_clausula['accion_cierre']} dentro de {proyecto} durante el mes de {mes_minus} de {anio}."
            )
            t2 = "Verificación de Calidad, Trazabilidad y Consolidación Documental"
            d2 = (
                f"Se realizó la revisión de conformidad técnica y documental, organización de entregables y aseguramiento probatorio "
                f"de la Actividad_{k} bajo los estándares de supervisión de la entidad contratante."
            )
            items_numerados = [f"1. {t1}: {d1}", f"2. {t2}: {d2}"]
            vinetas_desarrollo = [
                {"titulo": t1, "descripcion": d1},
                {"titulo": t2, "descripcion": d2},
            ]
            resultados_entregables = [
                {"titulo": f"Entregable {k}.1 — {ctx_clausula['tema_corto']}", "descripcion": d1},
                {"titulo": f"Entregable {k}.2 — Aseguramiento Documental Actividad_{k}", "descripcion": d2},
            ]

        tipo_prof = ctx_clausula.get("tipo_profesion", "gestión profesional y técnica")
        encabezado_tabla = (
            f"{ctx_clausula['enfoque']}"
            f"Durante el periodo mensual correspondiente a {mes_minus} de {anio}, en cumplimiento de la Obligación "
            f"Contractual N° {k} ({ctx_clausula['tema_corto']}) asociada a {proyecto}, se ejecutaron de manera "
            f"rigurosa las siguientes actividades de {tipo_prof}:\n"
        )
        cierre_tabla = (
            f"\n\nConclusiones y Aseguramiento Técnico: Las actividades y entregables reportados garantizan "
            f"{ctx_clausula['accion_cierre']}, cumpliendo los estándares de calidad e inmutabilidad "
            f"de soportes y productos establecidos en los lineamientos de supervisión y gestión documental."
        )
        texto_tabla = sanitizar_texto_para_reporte(
            _contraer_al(encabezado_tabla + "\n".join(items_numerados) + cierre_tabla)
        )

        # 4. Construir bloques_evidencias para `Obligacion K.docx` y subitems para `REPORTE OTI.docx`
        bloques_evidencias = []
        subitems_anexos = []

        if evidencias:
            for ev_idx, ev in enumerate(evidencias, start=1):
                vin_correspondiente = vinetas_desarrollo[ev_idx - 1]
                titulo_ev_limpio = vin_correspondiente["titulo"]
                desc_ev_limpia = (
                    f"{vin_correspondiente['descripcion']} "
                    f"(Archivo soporte físico verificado: '{ev.nombre_original}', tipo {ev.tipo_archivo}, "
                    f"incorporado en la carpeta Anexos de la Actividad_{k})."
                )
                bloques_evidencias.append({
                    "numero": ev_idx,
                    "titulo": sanitizar_texto_para_reporte(titulo_ev_limpio),
                    "descripcion": sanitizar_texto_para_reporte(desc_ev_limpia),
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
                    f"{ctx_clausula['enfoque']}"
                    f"Durante el periodo correspondiente al mes de {mes_minus} de {anio}, las actuaciones desarrolladas en el marco "
                    f"de la Obligación Contractual N° {k} se centraron en {ctx_clausula['tema_corto'].lower()}, atendiendo los "
                    f"requerimientos técnicos del proyecto {proyecto} y los lineamientos de la supervisión contractual."
                ),
                (
                    f"En observancia de los lineamientos sobre aseguramiento documental y directrices institucionales "
                    f"aplicables, se garantizó {ctx_clausula['accion_cierre']}, "
                    f"consolidando los soportes y evidencias verificables que se detallan a continuación en relación con la cláusula: \"{clausula}\"."
                ),
            ],
            "vinetas_desarrollo": vinetas_desarrollo,
            "resultados_entregables": resultados_entregables,
            "anexos_secciones": [
                {
                    "titulo": f"Anexo {k}.1 — Soportes Técnicos y Evidencias Consolidados (Actividad_{k})",
                    "descripcion": (
                        f"Registros gráficos y documentos soporte ubicados en "
                        f"Expediente Digital de Evidencias / Pago {num_pago:02d} / Actividad_{k} / Anexos / y consolidados en el documento "
                        f"Obligacion {k}.docx:"
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

def _calcular_huella_obligacion(det, informe) -> str:
    """
    Calcula un hash MD5 único con el estado exacto de la obligación K:
    (cláusula + apuntes del contratista + nombres y metadatos de sus archivos soporte).
    Si el contratista no cambió ni una letra ni subió archivos nuevos en esta obligación,
    el hash será idéntico al de la última redacción.
    """
    evs = [
        f"{ev.nombre_original}:{ev.tamano_bytes}"
        for ev in det.archivos_evidencia.all()
    ]
    estado_raw = (
        f"CTO:{informe.contratista.numero_contrato}|PAGO:{informe.numero_pago}:{informe.mes}:{informe.anio}|"
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
    """
    Punto de entrada principal de `services/ai_writer.py`.
    - Si `obligacion_especifica_id` es None: clasifica las evidencias y redacta en UNA SOLA LLAMADA (Single-Shot JSON)
      todas las N obligaciones del informe y descuenta 1 crédito de `creditos_informe_completo`.
    - Si `obligacion_especifica_id` está presente: redacta únicamente esa obligación K y descuenta 1 crédito
      de `creditos_regeneracion_individual`.
    - Aplica sanitización transversal contra errores `paraparser` de ReportLab (`GC-FO-15`).
    """
    contratista = informe.contratista

    if not contratista.activo_mes_actual:
        raise PermissionError("El contratista no se encuentra activo para generar informes en el mes actual.")

    if obligacion_especifica_id is not None:
        if contratista.creditos_regeneracion_individual <= 0:
            raise PermissionError(
                "Ha alcanzado el límite mensual de créditos de regeneración individual (6/mes)."
            )
        detalles = list(
            informe.detalles_obligaciones.select_related("obligacion")
            .prefetch_related("archivos_evidencia")
            .filter(obligacion_id=obligacion_especifica_id)
        )
    else:
        if contratista.creditos_informe_completo <= 0:
            raise PermissionError(
                "Ha alcanzado el límite mensual de créditos de generación de informe completo (3/mes)."
            )
        detalles = list(
            informe.detalles_obligaciones.select_related("obligacion")
            .prefetch_related("archivos_evidencia")
            .all()
        )

    if not detalles:
        raise ValueError("No se encontraron obligaciones configuradas para redactar.")

    # =========================================================================
    # CACHÉ INTELIGENTE 0-TOKENS POR OBLIGACIÓN (Evita re-redactar lo que no cambió)
    # =========================================================================
    huellas_actuales = {det.id: _calcular_huella_obligacion(det, informe) for det in detalles}
    detalles_a_redactar = []

    if obligacion_especifica_id is not None or bool(apuntes_generales_mes and apuntes_generales_mes.strip()):
        # Si el usuario pidió explícitamente regenerar 1 obligación o mandó apuntes globales sin clasificar, redactarlas
        detalles_a_redactar = detalles
    else:
        # En "Redactar Todo": filtrar solo las obligaciones que tuvieron cambios o que están vacías
        for det in detalles:
            datos_previos = det.desarrollo_detallado_oti if isinstance(det.desarrollo_detallado_oti, dict) else {}
            huella_previa = datos_previos.get("_hash_cache_ia", "")
            tiene_texto_valido = bool(
                det.texto_tabla_principal
                and len(det.texto_tabla_principal.strip()) > 200
                and datos_previos.get("vinetas_desarrollo")
            )
            if not tiene_texto_valido or huella_previa != huellas_actuales[det.id]:
                detalles_a_redactar.append(det)

    # Si NINGUNA obligación cambió desde la última vez, retornar en 0.01s gastando 0 tokens y 0 créditos
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
        # 1. Intentar Hermes Agent en cola FIFO
        payload_ia = _llamar_hermes_agent_fifo(prompt)
        if payload_ia and "obligaciones" in payload_ia:
            motor_usado = f"Hermes Agent FIFO ({getattr(settings, 'HERMES_MODEL', 'openai-codex')})"
        else:
            # 2. Fallback automático a Gemini 2.5 Flash
            payload_ia = _llamar_gemini_fallback(prompt)
            if payload_ia and "obligaciones" in payload_ia:
                motor_usado = f"Fallback API Key Gemini ({getattr(settings, 'GEMINI_MODEL', 'gemini-2.5-flash')})"
            else:
                # 3. Fallback automático a OpenAI API
                payload_ia = _llamar_openai_fallback(prompt)
                if payload_ia and "obligaciones" in payload_ia:
                    motor_usado = f"Fallback API Key OpenAI ({getattr(settings, 'OPENAI_MODEL', 'gpt-4o-mini')})"

    if not payload_ia or "obligaciones" not in payload_ia:
        payload_ia = _generar_redaccion_demo_inteligente(informe, detalles_a_redactar)
        motor_usado = "Modo Offline Inteligente (Single-Shot Local)"

    obligaciones_dict = payload_ia.get("obligaciones", {})

    # Mapa de todas las evidencias del informe por nombre de archivo para reacomodarlas donde el Agente IA decida
    mapa_evidencias_por_nombre = {}
    for d_item in detalles:
        for ev_item in d_item.archivos_evidencia.all():
            mapa_evidencias_por_nombre[ev_item.nombre_original.strip().lower()] = ev_item

    # Guardar los resultados coherentes en cada DetalleEjecucionObligacion y reacomodar sus ArchivoEvidencia
    for det in detalles_a_redactar:
        k_str = str(det.obligacion.numero_obligacion)
        datos_ob = obligaciones_dict.get(k_str) or obligaciones_dict.get(det.obligacion.numero_obligacion)
        if not datos_ob:
            datos_ob = _generar_redaccion_demo_inteligente(informe, [det])["obligaciones"][k_str]

        texto_principal = sanitizar_texto_para_reporte(datos_ob.get("texto_tabla_principal", ""))
        desarrollo_oti = datos_ob.get("desarrollo_detallado_oti", {})
        bloques_ev = datos_ob.get("bloques_evidencias", [])

        # AHORRO DE ~1.800 TOKENS DE SALIDA: Construir `anexos_secciones` en Python (0 Tokens) a partir de `bloques_evidencias`
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

        # 1. Primero mover y actualizar los ArchivoEvidencia que el Agente IA asoció a esta obligación K
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

        # 2. Actualizar títulos y descripciones técnicas en los objetos ArchivoEvidencia que quedaron en esta obligación
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

        # 3. AHORA SÍ: Con los archivos y títulos ya definitivos en BD, guardar la huella exacta en el caché
        if isinstance(desarrollo_oti, dict):
            desarrollo_oti["_hash_cache_ia"] = _calcular_huella_obligacion(det, informe)

        det.texto_tabla_principal = texto_principal
        det.desarrollo_detallado_oti = desarrollo_oti
        det.bloques_evidencias_json = bloques_ev
        det.save()

    # Descontar crédito correspondiente:
    # - Si usó el botón pequeño de una sola obligación: descuenta de regeneración individual
    # - Si usó el botón principal de redactar el informe: descuenta 1 Pase de Cuenta de Cobro
    if obligacion_especifica_id is not None:
        contratista.creditos_regeneracion_individual = max(0, contratista.creditos_regeneracion_individual - 1)
    else:
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
