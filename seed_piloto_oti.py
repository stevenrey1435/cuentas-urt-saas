"""
Script Semilla Oficial (`seed_piloto_oti.py`) para OTI-PayAgent.
Carga el perfil real del Contratista Piloto:
- John Felipe Vega Otalora (C.C. 1006552477)
- Contrato 1031 de 2026 (Inicio: 2026-01-13, Terminación: 2026-11-30)
- Supervisora: DIANA MARCELA RINCON LOPEZ (PROFESIONAL ESPECIALIZADO 2028 GRADO 17)
- Firma digitalizada: plantillas_base/firma.jpg
- Sus 9 obligaciones contractuales reales
- Pago 10 (Octubre 2026) precargado con sus apuntes reales, textos técnicos de referencia y evidencias reales
  (incluyendo las capturas extraídas de PLANTILLA_ANEXO_EVIDENCIAS_OBLIGACION.docx y un Excel de muestra con cabeceras).
"""

import os
import sys
import shutil
from datetime import date
from PIL import Image, ImageDraw
import openpyxl

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "oti_cuentas.settings")
import django
django.setup()

from django.conf import settings
from django.contrib.auth.models import User
from django.core.files import File
from docx import Document

from generador.models import (
    PerfilContratista,
    ObligacionContractual,
    InformeMensual,
    DetalleEjecucionObligacion,
    ArchivoEvidencia,
)
from generador.services.template_converter import asegurar_plantillas_dinamicas
from generador.services.file_inspector import extraer_metadatos_ligeros_archivo, sanitizar_texto_para_sigaf_reportlab


OBLIGACIONES_REALES_1031 = [
    (
        1,
        "Desarrollar, actualizar, optimizar e implementar soluciones o componentes basados en inteligencia artificial, aplicando modelos de aprendizaje automático, procesamiento de lenguaje natural y análisis predictivo, conforme a los requerimientos funcionales del proyecto y los estándares tecnológicos establecidos.",
    ),
    (
        2,
        "Validar que los componentes de desarrollo entregados se ajusten a los requerimientos, historias de usuario y criterios de aceptación establecidos durante la fase de análisis del proyecto.",
    ),
    (
        3,
        "Proporcionar junto con la solución desarrollada, los scripts de ajuste de base de datos, configuración y demás documentos técnicos requeridos para llevar a cabo las pruebas de calidad del software implementado o actualizado.",
    ),
    (
        4,
        "Formular cronogramas, estimar esfuerzos, implementar actividades planificadas y compilar informes detallados sobre las tareas de desarrollo ejecutadas.",
    ),
    (
        5,
        "Gestionar la actualización continua y la documentación estructurada del código fuente en las ramas asignadas del repositorio GIT institucional.",
    ),
    (
        6,
        "Asistir a los seguimientos programados de los proyectos asignados y actualizar las actividades en la herramienta de seguimiento.",
    ),
    (
        7,
        "Gestionar y dar respuesta a los requerimientos e incidencias asignadas a través de la herramienta de mesa de servicios y/o correo electrónico reportando los seguimientos y soluciones implementadas.",
    ),
    (
        8,
        "Asistir a reuniones en modalidad virtual o presencial, según las necesidades operativas y/o demandas del servicio.",
    ),
    (
        9,
        "Las demás actividades que le sean encomendadas, necesarias para garantizar el cumplimiento de las obligaciones contractuales.",
    ),
]


DATOS_PAGO_10_POR_OBLIGACION = {
    1: {
        "apuntes": (
            "[MODO_SEED_EXACTO]\n"
            "- Sanitización y Escape de Caracteres XML/HTML en el Motor PDF ReportLab (apps/informes/views_pdf.py - Commit d77f67b6 / MR !134): Solución de raíz a ValueError paraparser syntax error en formato GC-FO-15 aplicando html.escape en procesar_texto_largo, val_obj_con y ruta_acceso.\n"
            "- Corrección de Lógica de Visualización de Terminación Anticipada en Formato GF-FO-05 (apps/informes/views_pdf.py - Commit 12e7bd6c): Refactorización del bloque # 5. TERMINACIÓN ANTICIPADA en InformeDePagos eliminando restricción mes_term <= mes_inf e incorporando fallback a contrato.fecha_terminacion_anticipada.\n"
            "- Integración de Persistencia Atómica en MongoDB GridFS, Vigencia Dinámica y Controles de Supervisión (Commits cd3540c4 / MR !129, 3c558134 y 5e853db3): Verificación post-guardado en GridFS, Fecha de terminación actual en GC-FO-15 y campo requiere_correccion."
        ),
        "texto_tabla_principal": (
            "Durante el periodo de octubre de 2026, se ejecutó el desarrollo, optimización y blindaje de seguridad de los componentes del Sistema de Gestión Administrativa y Financiera (SIGAF - Alcance 3), destacándose las siguientes intervenciones en el código fuente:\n"
            "1. Sanitización y Escape de Caracteres XML/HTML en el Motor PDF ReportLab (apps/informes/views_pdf.py - Commit d77f67b6 / MR !134): Se diagnosticó y solucionó de raíz la excepción 'ValueError: paraparser: syntax error' que generaba errores HTTP 500 al expedir el formato GC-FO-15 cuando las descripciones de actividades o rutas contenían secuencias como 'para' o caracteres especiales. Se implementó la función estándar html.escape de manera transversal sobre procesar_texto_largo, el objeto contractual (val_obj_con) y la ruta de acceso de evidencias (ruta_acceso).\n"
            "2. Corrección de Lógica de Visualización de Terminación Anticipada en el Formato GF-FO-05 (apps/informes/views_pdf.py - Commit 12e7bd6c): Se refactorizó el bloque '# 5. TERMINACIÓN ANTICIPADA' en la vista InformeDePagos, eliminando la restricción condicional por mes en trámite (mes_term ‹= mes_inf) que impedía reflejar la marcación '¿TERMINACIÓN ANTICIPADA?: SÍ' cuando el pago tramitado correspondía a un periodo previo a la fecha de efectividad de la terminación (05 de octubre de 2026), e incorporando un mecanismo de respaldo (fallback) hacia contrato.fecha_terminacion_anticipada.\n"
            "3. Integración de Persistencia Atómica en MongoDB GridFS, Vigencia Dinámica y Controles de Supervisión (Commits cd3540c4 / MR !129, 3c558134 y 5e853db3): Se consolidó en la rama develop la verificación post-guardado en GridFS, el cálculo dinámico de la 'Fecha de terminación actual' en el formato GC-FO-15 y la sincronización con el control de edición selectiva de obligaciones (requiere_correccion) para informes devueltos con apoyo a la supervisión."
        ),
        "desarrollo_detallado_oti": {
            "parrafos_introductorios": [
                "Durante el periodo correspondiente al mes de octubre de 2026, las actividades de ingeniería de software se centraron en la optimización, blindaje de seguridad y estabilización del motor de generación de reportes oficiales PDF (ReportLab), así como en la integración, despliegue a pruebas y aseguramiento de calidad de los controles de cambio del Alcance 3 del Sistema de Gestión Administrativa y Financiera (SIGAF), garantizando la continuidad operativa de los trámites de pago de la entidad.",
                "En el componente de sanitización y desarrollo seguro del motor de reportes PDF (Formato GC-FO-15 - Informe de Ejecución y Supervisión de Actividades), se diagnosticó y solucionó de raíz una vulnerabilidad crítica de interpretación XML/HTML en la librería ReportLab:",
            ],
            "vinetas_desarrollo": [
                {
                    "titulo": "Sanitización y Escape de Caracteres XML/HTML en ReportLab (apps/informes/views_pdf.py - Commit d77f67b6 / MR !134)",
                    "descripcion": "Se identificó que el motor interno de ReportLab (paraparser) interpretaba secuencias de texto ingresadas por los contratistas como etiquetas de marcado XML (por ejemplo, el uso de secuencias 'para' reportado en el Contrato 147 o la inclusión de símbolos especiales en descripciones de actividades y rutas de red), provocando excepciones fatales 'ValueError: paraparser: syntax error: parse ended with 1 unclosed tags para' y errores HTTP 500 al guardar o remitir cuentas de cobro. Para erradicar este fallo bajo estándares de desarrollo seguro, se implementó la función estándar html.escape de forma transversal sobre la función procesar_texto_largo, así como sobre la cadena del objeto contractual (val_obj_con) y la ruta de almacenamiento de evidencias (ruta_acceso), preservando intactas las etiquetas internas de formato controladas por el sistema.",
                },
                {
                    "titulo": "Corrección de Lógica de Visualización de Terminación Anticipada en Certificación de Pagos GF-FO-05 (apps/informes/views_pdf.py - Commit 12e7bd6c)",
                    "descripcion": "A partir del análisis técnico de los contratos con modificaciones contractuales concurrentes (como Reducción y Terminación Anticipada en el Contrato 1464 con fecha 05 de octubre de 2026), se detectó que la vista InformeDePagos condicionaba la marcación de '¿TERMINACIÓN ANTICIPADA?: SÍ' a que el mes de la fecha de terminación fuera menor o igual al mes del informe en trámite. Se refactorizó el bloque '# 5. TERMINACIÓN ANTICIPADA' eliminando la restricción por mes en trámite e incorporando un mecanismo de respaldo (fallback) directo hacia el atributo contrato.fecha_terminacion_anticipada cuando no exista registro individual en ModificacionContractual.",
                },
                {
                    "titulo": "Integración de Paquetes de Estabilización GridFS, Vigencia Contractual y Controles de Supervisión (Commits cd3540c4, 381dbfd6, 3c558134 y 5e853db3)",
                    "descripcion": "Se ejecutó la integración en la rama develop (Merge Requests !129 y !134) y la sincronización bidireccional con la rama funcional feature/alcance-3-controles-cambio durante octubre de 2026, consolidando la persistencia atómica de archivos adjuntos en MongoDB GridFS, el cálculo dinámico de la 'Fecha de terminación actual' en el formato GC-FO-15 y la compatibilidad con el control de habilitación selectiva de edición de obligaciones (requiere_correccion) para contratistas con apoyos a la supervisión asignados.",
                },
            ],
            "resultados_entregables": [
                {
                    "titulo": "Motor de Sanitización XML/HTML para Generación Segura de Reportes PDF (Formato GC-FO-15)",
                    "descripcion": "Componente desplegado que neutraliza mediante html.escape cualquier carácter especial o pseudo-etiqueta ingresada por los usuarios en descripciones de obligaciones, objeto contractual o rutas UNC, eliminando los errores HTTP 500 del analizador paraparser de ReportLab.",
                },
                {
                    "titulo": "Módulo Optimizado de Certificación de Pagos con Detección Integral de Terminación Anticipada (Formato GF-FO-05)",
                    "descripcion": "Actualización de la clase InformeDePagos en apps/informes/views_pdf.py que garantiza la visualización precisa de la marcación '¿TERMINACIÓN ANTICIPADA?: SÍ' y su respectiva fecha de efectividad sin depender del mes del pago en trámite.",
                },
                {
                    "titulo": "Ramas de Integración y Control de Cambios Sincronizadas (develop y feature/alcance-3-controles-cambio)",
                    "descripcion": "Código fuente unificado, libre de conflictos de fusión y desplegado exitosamente en el ambiente de pruebas de SIGAF durante los ciclos de liberación del 06 y 19 de octubre de 2026.",
                },
            ],
            "anexos_secciones": [
                {
                    "titulo": "Anexo 1.1 — Componentes de Código Fuente Desarrollados y Optimizados",
                    "descripcion": "Archivos de código fuente intervenidos y versionados en el repositorio institucional de SIGAF en GitLab:",
                    "subitems": [
                        {
                            "etiqueta": "apps/informes/views_pdf.py",
                            "detalle": "Implementación de sanitización HTML/XML (from html import escape) en procesar_texto_largo, val_obj_con y ruta_acceso para el formato GC-FO-15; y refactorización del bloque de Terminación Anticipada en InformeDePagos para el formato GF-FO-05.",
                        },
                        {
                            "etiqueta": "apps/informes/views.py y core/storage_backends.py",
                            "detalle": "Integración y despliegue en rama develop de la verificación atómica de almacenamiento en MongoDB GridFS y manejo resiliente de descargas.",
                        },
                        {
                            "etiqueta": "apps/informes/migrations/0012_informeobligacion_requiere_correccion.py",
                            "detalle": "Sincronización de esquemas de control de corrección de obligaciones con apoyo a la supervisión.",
                        },
                    ],
                },
                {
                    "titulo": "Anexo 1.2 — Soportes de Pruebas y Evidencias Gráficas",
                    "descripcion": "Registros gráficos exclusivos ubicados en la carpeta \\0. Evidencias\\Pago 10\\Actividad_1\\Anexos\\ y consolidados en el documento Obligacion 1.docx:",
                    "subitems": [
                        {
                            "etiqueta": "Prueba_Captura_Etiquetas_XML.png",
                            "detalle": "Evidencia de ingreso de etiquetas críticas en el formulario de obligaciones para prueba de estrés del generador PDF.",
                        },
                        {
                            "etiqueta": "Sanitizacion_Caracteres_XML_GC_FO_15.png",
                            "detalle": "Evidencia del PDF oficial GC-FO-15 generado exitosamente renderizando la etiqueta como texto plano sin error de compilación.",
                        },
                        {
                            "etiqueta": "Terminacion_Anticipada_GF_FO_05.png",
                            "detalle": "Evidencia del formato GF-FO-05 (Certificación para Pago de Contratistas) reflejando correctamente '¿TERMINACIÓN ANTICIPADA?: SÍ' y su fecha.",
                        },
                    ],
                },
                {
                    "titulo": "Anexo 1.3 — Trazabilidad de Control de Versiones (Commits de Octubre de 2026)",
                    "descripcion": "Registro de commits auditables en el repositorio institucional GitLab (https://git.urt.gov.co/oti/sigaf):",
                    "subitems": [
                        {
                            "etiqueta": "cd3540c4 (05/10/2026)",
                            "detalle": "Merge branch 'feature/alcance-3-controles-cambio' into 'develop' (MR !129 - Persistencia atómica GridFS y fecha de terminación actual en GC-FO-15).",
                        },
                        {
                            "etiqueta": "d77f67b6 (09/10/2026)",
                            "detalle": "fix(pdf): sanitizar caracteres XML/HTML con escape en ReportLab para prevenir error paraparser.",
                        },
                        {
                            "etiqueta": "381dbfd6 (09/10/2026)",
                            "detalle": "Merge branch 'feature/alcance-3-controles-cambio' into 'develop' (MR !134 - Sanitización XML/HTML en ReportLab).",
                        },
                        {
                            "etiqueta": "3c558134 y 5e853db3 (19/10/2026)",
                            "detalle": "Sincronización de 'origin/develop' hacia 'feature/alcance-3-controles-cambio'.",
                        },
                        {
                            "etiqueta": "12e7bd6c (19/10/2026)",
                            "detalle": "fix(informes): mostrar terminacion anticipada en certificacion de pago independientemente del mes en tramite.",
                        },
                    ],
                },
            ],
        },
        "bloques_evidencias": [
            {
                "numero": 1,
                "titulo": "Sanitización y Escape de Caracteres XML/HTML en el Motor PDF ReportLab (Formato GC-FO-15 - Commit d77f67b6 / MR !134)",
                "descripcion": "Se implementó en apps/informes/views_pdf.py la sanitización transversal mediante html.escape sobre la función procesar_texto_largo, el objeto contractual (val_obj_con) y la ruta de acceso de evidencias (ruta_acceso). Con ello se erradicó el error HTTP 500 ('paraparser: syntax error: parse ended with 1 unclosed tags para') cuando los contratistas ingresan secuencias especiales en sus informes de actividades.",
                "archivo_asociado": "Prueba_Captura_Etiquetas_XML.png",
            },
            {
                "numero": 2,
                "titulo": "Verificación de Renderizado Seguro en el Formato Oficial GC-FO-15 (Commit 381dbfd6 / MR !134)",
                "descripcion": "Comprobación del documento PDF oficial GC-FO-15 expedido por ReportLab renderizando las secuencias XML/HTML como texto plano íntegro sin interrumpir el flujo de radicación de cuentas de cobro en SIGAF.",
                "archivo_asociado": "Sanitizacion_Caracteres_XML_GC_FO_15.png",
            },
            {
                "numero": 3,
                "titulo": "Corrección de Visualización de Terminación Anticipada en Certificación de Pagos (Formato GF-FO-05 - Commit 12e7bd6c)",
                "descripcion": "Se optimizó el bloque '# 5. TERMINACIÓN ANTICIPADA' dentro de la vista InformeDePagos (apps/informes/views_pdf.py), eliminando la restricción condicional por mes del pago en trámite e incorporando un respaldo (fallback) hacia contrato.fecha_terminacion_anticipada. De esta forma, el formato oficial GF-FO-05 marca invariablemente 'SÍ' e imprime la fecha de efectividad siempre que el contrato posea una terminación anticipada activa.",
                "archivo_asociado": "Terminacion_Anticipada_GF_FO_05.png",
            },
        ],
    },
    2: {
        "apuntes": (
            "[MODO_SEED_EXACTO]\n"
            "- Validación de Resiliencia y Escape XML en el Formato GC-FO-15 (Caso CTO 147): Pruebas de estrés ingresando etiquetas XML y caracteres especiales, verificando respuesta HTTP 200.\n"
            "- Validación de Reglas de Negocio en Certificación de Pagos GF-FO-05 (Caso CTO 1464): Verificación de coexistencia de Reducción y Terminación Anticipada con fecha 05 de octubre de 2026.\n"
            "- Acompañamiento y Verificación en los Ciclos 1 y 2 de Pruebas Internas QA (Alcance 3): Validación de criterios de aceptación en entregas del 06 y 19 de octubre de 2026."
        ),
        "texto_tabla_principal": (
            "Se realizaron las pruebas funcionales, de regresión y de validación de criterios de aceptación sobre los componentes entregados en el mes de octubre de 2026:\n"
            "1. Validación de Resiliencia y Escape XML en el Formato GC-FO-15 (Caso CTO 147): Se ejecutaron pruebas de estrés ingresando etiquetas 'para' y caracteres especiales en los campos de ejecución de obligaciones, comprobando que el documento PDF GC-FO-15 se genera exitosamente (HTTP 200) y renderiza las cadenas como texto plano íntegro.\n"
            "2. Validación de Reglas de Negocio en Certificación de Pagos GF-FO-05 (Caso CTO 1464): Se verificó en ambiente local y de pruebas la coexistencia de modificaciones de Reducción y Terminación Anticipada con fecha 05 de octubre de 2026, confirmando que el formato GF-FO-05 marque 'SÍ' en '¿TERMINACIÓN ANTICIPADA?' e imprima la fecha exacta registrada.\n"
            "3. Acompañamiento y Verificación en los Ciclos 1 y 2 de Pruebas Internas QA (Alcance 3): Se validó junto con el equipo funcional el cumplimiento de los criterios de aceptación en las entregas desplegadas el 06 y el 19 de octubre de 2026."
        ),
    },
    3: {
        "apuntes": (
            "[MODO_SEED_EXACTO]\n"
            "- Primer Despliegue a Pruebas (05 y 06 de octubre de 2026 - MR !129 / Commit cd3540c4): Entrega de paquete técnico e instrucción de python manage.py migrate sobre PostgreSQL y MongoDB GridFS.\n"
            "- Sincronización de Migraciones Relacionales (19 de octubre de 2026): Integración de la migración 0012_informeobligacion_requiere_correccion.py.\n"
            "- Segundo Despliegue a Pruebas (19 de octubre de 2026 - Commits 381dbfd6 y 12e7bd6c): Suministro de notas técnicas de liberación para actualización del motor PDF (GC-FO-15 y GF-FO-05)."
        ),
        "texto_tabla_principal": (
            "Se entregaron oportunamente las instrucciones de migración de base de datos, configuraciones y notas técnicas requeridas para los dos (2) despliegues realizados en el ambiente de pruebas (QA) de SIGAF durante octubre de 2026:\n"
            "1. Primer Despliegue a Pruebas (05 y 06 de octubre de 2026 - MR !129 / Commit cd3540c4): Entrega de paquete técnico e instrucción explícita de ejecución del comando 'python manage.py migrate' sobre PostgreSQL para consolidar las estructuras del modelo Informe y la capa de almacenamiento en MongoDB GridFS.\n"
            "2. Sincronización de Migraciones Relacionales (19 de octubre de 2026): Integración y verificación técnica de la migración 0012_informeobligacion_requiere_correccion.py en la rama de trabajo.\n"
            "3. Segundo Despliegue a Pruebas (19 de octubre de 2026 - Commits 381dbfd6 y 12e7bd6c): Suministro de notas técnicas de liberación para la actualización del motor PDF (formatos GC-FO-15 y GF-FO-05), precisando la exención del comando migrate y confirmando la disponibilidad del entorno para validación QA."
        ),
    },
    4: {
        "apuntes": (
            "[MODO_SEED_EXACTO]\n"
            "- Actualización y control del cronograma oficial del proyecto SIGAF - Alcance 3 (Cronograma_SIGAF13Octubre2026.pdf) con corte al 13 de octubre de 2026 registrando avance global del 65%.\n"
            "- Estimación de esfuerzos para resolución de incidencias críticas de producción y elaboración del informe técnico mensual INFORME DESARROLLADORES DEASARROLLO.pdf."
        ),
        "texto_tabla_principal": (
            "Durante el mes de octubre de 2026, se participó en la actualización y control del cronograma oficial del proyecto SIGAF - Alcance 3 ('Cronograma_SIGAF13Octubre2026.pdf', con corte al 13 de octubre de 2026), registrando un avance global del 65% (Planeación 100%, Desarrollo 94%, Despliegues a Pruebas 1 y 2, Pruebas Internas QA Ciclos 1 y 2 al 100%, Corrección de Incidencias al 100%, Pruebas de Aceptación Funcional UAT, Uso y Apropiación, y Preparación de Salida a Producción).\n"
            "Asimismo, se realizó la estimación de esfuerzos para la resolución de las incidencias críticas de producción y se elaboró el informe técnico mensual 'INFORME DESARROLLADORES DEASARROLLO.pdf' detallando las actividades, procedimientos y resultados obtenidos en octubre de 2026."
        ),
    },
    5: {
        "apuntes": (
            "[MODO_SEED_EXACTO]\n"
            "- Commit cd3540c4 (05/10/2026): Merge branch feature/alcance-3-controles-cambio into develop (MR !129).\n"
            "- Commit d77f67b6 (09/10/2026): fix(pdf): sanitizar caracteres XML/HTML con escape en ReportLab.\n"
            "- Commit 381dbfd6 (09/10/2026): Merge branch feature/alcance-3-controles-cambio into develop (MR !134).\n"
            "- Commits 3c558134 y 5e853db3 (19/10/2026): Sincronización de origin/develop hacia feature/alcance-3-controles-cambio.\n"
            "- Commit 12e7bd6c (19/10/2026): fix(informes): mostrar terminacion anticipada en certificacion de pago."
        ),
        "texto_tabla_principal": (
            "Se mantuvo rigurosamente actualizado y documentado el repositorio oficial del proyecto SIGAF en la plataforma institucional GitLab (https://git.urt.gov.co/oti/sigaf.git) sobre las ramas 'feature/alcance-3-controles-cambio' y 'develop' bajo el estándar de Conventional Commits durante octubre de 2026:\n"
            "- Commit cd3540c4 (05/10/2026): Merge branch 'feature/alcance-3-controles-cambio' into 'develop' (Merge Request !129).\n"
            "- Commit d77f67b6 (09/10/2026): fix(pdf): sanitizar caracteres XML/HTML con escape en ReportLab para prevenir error paraparser.\n"
            "- Commit 381dbfd6 (09/10/2026): Merge branch 'feature/alcance-3-controles-cambio' into 'develop' (Merge Request !134).\n"
            "- Commits 3c558134 y 5e853db3 (19/10/2026): Sincronización de 'origin/develop' hacia 'feature/alcance-3-controles-cambio'.\n"
            "- Commit 12e7bd6c (19/10/2026): fix(informes): mostrar terminacion anticipada en certificacion de pago independientemente del mes en tramite."
        ),
    },
    6: {
        "apuntes": (
            "[MODO_SEED_EXACTO]\n"
            "- Jueves 1 de octubre de 2026 (9:00 a.m. - 9:12 a.m.): Seguimiento al cierre de desarrollo del Alcance 3 y apertura del Ciclo 1 de QA.\n"
            "- Jueves 8 de octubre de 2026 (9:00 a.m. - 9:24 a.m.): Revisión del primer despliegue a pruebas (06/10/2026) y avance de validación QA.\n"
            "- Martes 13 de octubre de 2026 (9:00 a.m. - 9:34 a.m.): Actualización del cronograma maestro del proyecto y apertura del Ciclo 2 de QA.\n"
            "- Martes 20 de octubre de 2026 (8:59 a.m. - 9:04 a.m.): Balance del segundo despliegue a pruebas (19/10/2026), cierre de incidencias GLPI y preparación de UAT."
        ),
        "texto_tabla_principal": (
            "Se asistió puntualmente al 100% de las sesiones programadas de seguimiento técnico del proyecto SIGAF convocadas durante el mes de octubre de 2026 a través de Microsoft Teams, presentando el avance de los desarrollos, despliegues y validaciones QA:\n"
            "1. Jueves 1 de octubre de 2026 (9:00 a.m. - 9:12 a.m.): Seguimiento al cierre de desarrollo del Alcance 3 y apertura del Ciclo 1 de QA.\n"
            "2. Jueves 8 de octubre de 2026 (9:00 a.m. - 9:24 a.m.): Revisión del primer despliegue a pruebas (06/10/2026) y avance de validación QA.\n"
            "3. Martes 13 de octubre de 2026 (9:00 a.m. - 9:34 a.m.): Actualización del cronograma maestro del proyecto y apertura del Ciclo 2 de QA.\n"
            "4. Martes 20 de octubre de 2026 (8:59 a.m. - 9:04 a.m.): Balance del segundo despliegue a pruebas (19/10/2026), cierre de incidencias GLPI y preparación de UAT / salida a producción.\n"
            "Adicionalmente, se mantuvieron actualizadas las herramientas de seguimiento y control de tickets de la OTI."
        ),
    },
    7: {
        "apuntes": (
            "[MODO_SEED_EXACTO]\n"
            "- Ticket GLPI #0169062 (06 al 09 de octubre de 2026) — ERROR ENVIO CUENTA SIGAF .- GENERACIÓN PDF INFORME ACTIVIDADES (Contrato 147): Diagnóstico y solución al error paraparser mediante sanitización HTML/XML en ReportLab.\n"
            "- Ticket GLPI #0169455 (15 al 19 de octubre de 2026) — SIGAF ERROR GENERACIÓN FORMATO GF-FO-05 CTO 1464: Diagnóstico y ajuste en InformeDePagos para visualizar TERMINACIÓN ANTICIPADA: SÍ (05 de octubre de 2026).\n"
            "- Gestión por Correo Electrónico Institucional (05–06 y 19 de octubre de 2026): Solicitud y coordinación de los dos (2) despliegues de la rama develop al ambiente de pruebas."
        ),
        "texto_tabla_principal": (
            "Se atendieron, solucionaron en código y respondieron dentro de los tiempos establecidos los requerimientos e incidencias asignados mediante la plataforma GLPI y correo electrónico institucional en octubre de 2026:\n"
            "1. Ticket GLPI #0169062 (06 al 09 de octubre de 2026) — 'ERROR ENVIO CUENTA SIGAF .- GENERACIÓN PDF INFORME ACTIVIDADES' (Contrato 147): Diagnóstico de causa raíz y solución al error 'paraparser' ocasionado por la etiqueta 'para' mediante sanitización HTML/XML en ReportLab, con respuesta y validación concedida el 09/10/2026.\n"
            "2. Ticket GLPI #0169455 (15 al 19 de octubre de 2026) — 'SIGAF ERROR GENERACIÓN FORMATO GF-FO-05 CTO 1464': Diagnóstico y ajuste en InformeDePagos para visualizar la marcación '¿TERMINACIÓN ANTICIPADA?: SÍ' y su fecha (05 de octubre de 2026) sin restricción por el mes del pago en trámite, con respuesta y validación concedida el 19/10/2026.\n"
            "3. Gestión por Correo Electrónico Institucional (05–06 y 19 de octubre de 2026): Solicitud y coordinación de los dos (2) despliegues de la rama develop al ambiente de pruebas de SIGAF."
        ),
    },
    8: {
        "apuntes": (
            "[MODO_SEED_EXACTO]\n"
            "- Mesa Técnica Extraordinaria Revisión Incidente Régimen Tributario SIGAF (Jueves 1 de octubre de 2026, 15:17 - 16:17): Reunión con supervisora Ing. Diana Marcela Rincón López y enlace financiero Julieth Alejandra Ladino Leguizamón.\n"
            "- Sesiones Virtuales de Articulación Técnica y Seguimiento SIGAF (1, 8, 13 y 20 de octubre de 2026): Mesas de coordinación con QA y liderazgo del proyecto."
        ),
        "texto_tabla_principal": (
            "Se asistió a las reuniones virtuales y mesas técnicas extraordinarias convocadas por la supervisión del contrato y los enlaces funcionales de acuerdo con las necesidades del servicio durante octubre de 2026:\n"
            "1. Mesa Técnica Extraordinaria 'Revisión Incidente Régimen Tributario SIGAF' (Jueves 1 de octubre de 2026, 15:17 - 16:17): Reunión con participación de la supervisora del contrato (Ing. Diana Marcela Rincón López), el enlace funcional financiero (Julieth Alejandra Ladino Leguizamón) y el equipo técnico OTI para definir el tratamiento del incidente de régimen tributario.\n"
            "2. Sesiones Virtuales de Articulación Técnica y Seguimiento SIGAF (1, 8, 13 y 20 de octubre de 2026): Mesas de coordinación con el equipo de QA y liderazgo del proyecto para la revisión de despliegues, pruebas de aceptación y alistamiento de salida a producción."
        ),
    },
    9: {
        "apuntes": (
            "[MODO_SEED_EXACTO]\n"
            "- Análisis Forense No Intrusivo y Reproducción Segura de Casos Críticos de Producción: Simulación en entorno local de flujos complejos de cuentas de cobro (casos CTO 147 y CTO 1464).\n"
            "- Acompañamiento Técnico en Ciclos 1 y 2 de QA, Pruebas de Aceptación (UAT) y Preparación de Salida a Producción: Soporte permanente al equipo funcional y de calidad.\n"
            "- Organización y Consolidación de Evidencias Contractuales (Circular 0009 de 2024): Estructuración integral del expediente digital de evidencias del Pago 10 en la NAS de la OTI."
        ),
        "texto_tabla_principal": (
            "En atención a las directrices de la supervisión del contrato para garantizar el cumplimiento de los objetivos del proyecto SIGAF durante octubre de 2026, se ejecutaron las siguientes actividades complementarias:\n"
            "1. Análisis Forense No Intrusivo y Reproducción Segura de Casos Críticos de Producción: Simulación controlada en entorno local de flujos complejos de cuentas de cobro (devolución desde rol de par financiero, edición de borradores por el contratista y coexistencia de modificaciones de reducción y terminación anticipada en los casos CTO 147 y CTO 1464) para diagnosticar y validar soluciones sin afectar la integridad de los datos institucionales.\n"
            "2. Acompañamiento Técnico en Ciclos 1 y 2 de QA, Pruebas de Aceptación (UAT) y Preparación de Salida a Producción: Soporte permanente al equipo funcional y de calidad en la validación del Alcance 3 y preparación de insumos para la Mesa de Gestión de Cambios.\n"
            "3. Organización y Consolidación de Evidencias Contractuales (Circular 0009 de 2024): Estructuración integral del expediente digital de evidencias y anexos correspondientes al Pago 10 en el repositorio de red de la OTI."
        ),
    },
}


def _extraer_capturas_reales_de_plantilla_anexo(carpeta_destino: str) -> list:
    """
    Extrae las 3 capturas PNG reales que vienen dentro de `plantillas_base/PLANTILLA_ANEXO_EVIDENCIAS_OBLIGACION.docx`
    y las guarda con sus nombres oficiales del Pago 10:
    - Prueba_Captura_Etiquetas_XML.png
    - Sanitizacion_Caracteres_XML_GC_FO_15.png
    - Terminacion_Anticipada_GF_FO_05.png
    """
    os.makedirs(carpeta_destino, exist_ok=True)
    doc = Document(os.path.join(settings.PLANTILLAS_BASE_DIR, "PLANTILLA_ANEXO_EVIDENCIAS_OBLIGACION.docx"))
    nombres_oficiales = [
        ("rId4", "Prueba_Captura_Etiquetas_XML.png"),
        ("rId5", "Sanitizacion_Caracteres_XML_GC_FO_15.png"),
        ("rId6", "Terminacion_Anticipada_GF_FO_05.png"),
    ]
    rutas_extraidas = []
    for rel_id, nombre_png in nombres_oficiales:
        if rel_id in doc.part.rels:
            blob = doc.part.rels[rel_id].target_part.blob
            ruta_out = os.path.join(carpeta_destino, nombre_png)
            with open(ruta_out, "wb") as f:
                f.write(blob)
            rutas_extraidas.append((nombre_png, ruta_out))
    return rutas_extraidas


def _generar_evidencia_grafica_para_obligacion(ruta_png: str, k: int, titulo: str, subtitulo: str):
    """Genera una captura de evidencia limpia para las obligaciones 2..9 en el seed de prueba."""
    os.makedirs(os.path.dirname(ruta_png), exist_ok=True)
    img = Image.new("RGB", (1050, 420), color=(248, 250, 252))
    draw = ImageDraw.Draw(img)
    draw.rectangle([(0, 0), (1050, 56)], fill=(15, 23, 42))
    draw.text((24, 18), f"SIGAF - OTI URT | Evidencia Técnica Verificada - Obligación {k} (Pago 10 - Octubre 2026)", fill=(255, 255, 255))
    draw.rectangle([(24, 80), (1026, 392)], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
    draw.text((48, 108), f"Contrato N° 1031 de 2026 | Contratista: John Felipe Vega Otalora", fill=(71, 85, 105))
    draw.text((48, 145), f" Evidencia: {titulo}", fill=(15, 23, 42))
    # Dividir subtítulo en líneas de 95 caracteres
    palabras = subtitulo.split()
    lineas = []
    actual = ""
    for p in palabras:
        if len(actual) + len(p) + 1 <= 95:
            actual = f"{actual} {p}".strip()
        else:
            lineas.append(actual)
            actual = p
    if actual:
        lineas.append(actual)
    for idx_l, ln in enumerate(lineas[:6]):
        draw.text((48, 190 + idx_l * 26), ln, fill=(51, 65, 85))
    draw.rectangle([(48, 345), (420, 375)], fill=(220, 252, 231), outline=(134, 239, 172))
    draw.text((62, 353), "ESTADO: VERIFICADO Y ALMACENADO EN NAS", fill=(22, 101, 52))
    img.save(ruta_png, "PNG")
    return ruta_png


def _crear_excel_muestra_matriz_qa(ruta_xlsx: str):
    """
    Crea un archivo Excel de soporte en la Obligación 2 (`Matriz_Validacion_QA_Alcance3_Octubre2026.xlsx`)
    para demostrar la extracción 0-Tokens de nombres de hojas y fila 1 de columnas con `openpyxl(read_only=True)`.
    """
    os.makedirs(os.path.dirname(ruta_xlsx), exist_ok=True)
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Casos_Prueba_Alcance3"
    ws1.append(["ID_Caso", "Modulo_SIGAF", "Contrato_Prueba", "Formato_PDF", "Resultado_Esperado", "Estado_QA", "Fecha_Ejecucion"])
    ws1.append(["CP-101", "Motor PDF ReportLab", "CTO-147", "GC-FO-15", "Escape HTML/XML sin error paraparser", "APROBADO", "2026-10-09"])
    ws1.append(["CP-102", "Certificación Pagos", "CTO-1464", "GF-FO-05", "Marcación Terminación Anticipada SÍ", "APROBADO", "2026-10-19"])

    ws2 = wb.create_sheet("Ciclos_Despliegue_QA")
    ws2.append(["Ciclo_QA", "Fecha_Despliegue", "Merge_Request", "Commit_Hash", "Responsable_Tecnico", "Estado_Liberacion"])
    ws2.append(["Ciclo 1", "2026-10-06", "MR !129", "cd3540c4", "John Felipe Vega Otalora", "COMPLETADO"])
    ws2.append(["Ciclo 2", "2026-10-19", "MR !134", "381dbfd6 / 12e7bd6c", "John Felipe Vega Otalora", "COMPLETADO"])
    wb.save(ruta_xlsx)
    wb.close()
    return ruta_xlsx


def ejecutar_seed():
    print("=" * 80)
    print("INICIANDO SEED PILOTO OTI-URT (Contrato 1031 de 2026 - John Felipe Vega Otalora)")
    print("=" * 80)

    # 1. Asegurar plantillas dinámicas en plantillas_dinamicas/
    rutas_plantillas = asegurar_plantillas_dinamicas()
    print(f"[OK] Plantillas dinamicas generadas: {list(rutas_plantillas.keys())}")

    # 2. Crear usuario y PerfilContratista
    user, _ = User.objects.get_or_create(
        username="john.vega",
        defaults={
            "first_name": "John Felipe",
            "last_name": "Vega Otalora",
            "email": "john.vega@restituciondetierras.gov.co",
            "is_staff": True,
            "is_superuser": True,
        },
    )
    user.set_password("oti2026*")
    user.save()

    perfil, _ = PerfilContratista.objects.update_or_create(
        cedula="1006552477",
        defaults={
            "user": user,
            "nombres_apellidos": "John Felipe Vega Otalora",
            "apellidos_nombres_nas": "VEGA_OTALORA_JOHN_FELIPE",
            "nombre_corto_archivo": "JOHN_VEGA",
            "numero_contrato": "1031",
            "anio_contrato": 2026,
            "fecha_inicio": date(2026, 1, 13),
            "fecha_terminacion_actual": date(2026, 11, 30),
            "objeto_contrato": (
                "Prestar sus servicios de apoyo a la gestión con plena autonomía técnica y administrativa "
                "a la Oficina de Tecnologías de la Información, en el desarrollo, actualización, optimización "
                "e implementación de soluciones o componentes basados en inteligencia artificial."
            ),
            "nombre_supervisor": "DIANA MARCELA RINCON LOPEZ",
            "cargo_supervisor": "PROFESIONAL ESPECIALIZADO 2028 GRADO 17",
            "direccion_oficina": "Oficina Tecnologías de la Información",
            "direccion_territorial": "Nivel Central",
            "subcarpeta_rol_nas": "Desarrolladores",
            "proyecto_area": "Sistema SIGAF - Alcance 3",
            "servidor_nas_prefix": r"\\10.10.36.17\nivel_central\oti\4. sistemas de informacion",
            "activo_mes_actual": True,
            "creditos_informe_completo": 3,
            "creditos_regeneracion_individual": 6,
            "max_caracteres_apuntes": 2500,
        },
    )

    # Asociar firma.jpg si existe en plantillas_base/
    firma_origen = os.path.join(settings.PLANTILLAS_BASE_DIR, "firma.jpg")
    if os.path.exists(firma_origen) and not perfil.firma_imagen:
        with open(firma_origen, "rb") as f_img:
            perfil.firma_imagen.save("firma.jpg", File(f_img), save=True)

    print(f"[OK] PerfilContratista cargado: {perfil} | Ruta NAS: {perfil.ruta_base_nas}")

    # 3. Cargar las 9 Obligaciones Contractuales reales
    obligaciones_objs = {}
    for num_ob, texto_clausula in OBLIGACIONES_REALES_1031:
        ob_obj, _ = ObligacionContractual.objects.update_or_create(
            contratista=perfil,
            numero_obligacion=num_ob,
            defaults={"descripcion_clausula": texto_clausula},
        )
        obligaciones_objs[num_ob] = ob_obj
    print(f"[OK] {len(obligaciones_objs)} Obligaciones Contractuales cargadas.")

    # 4. Crear o actualizar el InformeMensual del Pago 10 (Octubre 2026)
    informe_10, _ = InformeMensual.objects.update_or_create(
        contratista=perfil,
        numero_pago=10,
        anio=2026,
        defaults={
            "mes": 10,
            "fecha_reporte": date(2026, 10, 28),
            "informes_secop_publicados": 5,
            "estado": "REDACTADO_IA",
            "ultimo_motor_ia_usado": "Seed Oficial Pago 10 (Octubre 2026 - OTI URT)",
        },
    )

    # Carpeta temporal para preparar archivos semilla de evidencias
    seed_files_dir = os.path.join(settings.MEDIA_ROOT, "seed_assets_pago_10")
    os.makedirs(seed_files_dir, exist_ok=True)
    capturas_ob1 = _extraer_capturas_reales_de_plantilla_anexo(seed_files_dir)

    # Crear los 9 DetalleEjecucionObligacion y sus archivos soporte reales
    from generador.services.ai_writer import _generar_redaccion_demo_inteligente

    for k in range(1, 10):
        ob_obj = obligaciones_objs[k]
        info_k = DATOS_PAGO_10_POR_OBLIGACION[k]
        det, _ = DetalleEjecucionObligacion.objects.update_or_create(
            informe=informe_10,
            obligacion=ob_obj,
            defaults={
                "apuntes_contratista": info_k["apuntes"],
                "texto_tabla_principal": sanitizar_texto_para_sigaf_reportlab(info_k["texto_tabla_principal"]),
            },
        )

        # Limpiar evidencias previas para evitar duplicados al re-ejecutar el seed
        det.archivos_evidencia.all().delete()

        if k == 1:
            # Cargar las 3 capturas reales extraídas de PLANTILLA_ANEXO_EVIDENCIAS_OBLIGACION.docx
            bloques_ob1 = info_k["bloques_evidencias"]
            for idx_img, (nombre_png, ruta_png) in enumerate(capturas_ob1):
                bloque_meta = bloques_ob1[idx_img] if idx_img < len(bloques_ob1) else {}
                with open(ruta_png, "rb") as f_bin:
                    ev = ArchivoEvidencia(
                        detalle_obligacion=det,
                        orden=idx_img + 1,
                        nombre_original=nombre_png,
                        tipo_archivo="IMAGEN",
                        tamano_bytes=os.path.getsize(ruta_png),
                        titulo_evidencia=bloque_meta.get("titulo", nombre_png),
                        descripcion_tecnica_ia=bloque_meta.get("descripcion", ""),
                        metadatos_ligeros={"nombre_archivo": nombre_png, "extension": ".png"},
                    )
                    ev.archivo.save(nombre_png, File(f_bin), save=True)
            det.desarrollo_detallado_oti = info_k["desarrollo_detallado_oti"]
            det.bloques_evidencias_json = info_k["bloques_evidencias"]
            det.save()
        else:
            # Crear captura gráfica de soporte para la obligación K
            nombre_cap = f"Evidencia_Tecnica_Obligacion_{k}_Pago10.png"
            ruta_cap = os.path.join(seed_files_dir, nombre_cap)
            _generar_evidencia_grafica_para_obligacion(
                ruta_cap,
                k=k,
                titulo=f"Soporte de Ejecución Obligación {k} - Octubre 2026",
                subtitulo=info_k["texto_tabla_principal"].replace("\n", " "),
            )
            with open(ruta_cap, "rb") as f_bin:
                ev_img = ArchivoEvidencia(
                    detalle_obligacion=det,
                    orden=1,
                    nombre_original=nombre_cap,
                    tipo_archivo="IMAGEN",
                    tamano_bytes=os.path.getsize(ruta_cap),
                    titulo_evidencia=f"Soporte Gráfico de Ejecución y Verificación — Obligación {k} (Octubre 2026)",
                    descripcion_tecnica_ia=sanitizar_texto_para_sigaf_reportlab(info_k["texto_tabla_principal"]),
                    metadatos_ligeros={"nombre_archivo": nombre_cap, "extension": ".png"},
                )
                ev_img.archivo.save(nombre_cap, File(f_bin), save=True)

            # En la Obligación 2 adjuntar además un archivo Excel estructurado para probar el filtro 0-Tokens de cabeceras
            if k == 2:
                nombre_xlsx = "Matriz_Validacion_QA_Alcance3_Octubre2026.xlsx"
                ruta_xlsx = os.path.join(seed_files_dir, nombre_xlsx)
                _crear_excel_muestra_matriz_qa(ruta_xlsx)
                meta_xlsx = extraer_metadatos_ligeros_archivo(ruta_xlsx, nombre_xlsx)
                with open(ruta_xlsx, "rb") as f_x:
                    ev_x = ArchivoEvidencia(
                        detalle_obligacion=det,
                        orden=2,
                        nombre_original=nombre_xlsx,
                        tipo_archivo="EXCEL",
                        tamano_bytes=os.path.getsize(ruta_xlsx),
                        titulo_evidencia="Matriz de Validación y Casos de Prueba QA — SIGAF Alcance 3 (Octubre 2026)",
                        descripcion_tecnica_ia=(
                            "Libro de control de calidad en formato Excel ubicado en la carpeta Anexos de la Actividad_2, "
                            "con las hojas Casos_Prueba_Alcance3 y Ciclos_Despliegue_QA que documentan la verificación de "
                            "los casos CTO-147 (GC-FO-15) y CTO-1464 (GF-FO-05)."
                        ),
                        metadatos_ligeros=meta_xlsx,
                    )
                    ev_x.archivo.save(nombre_xlsx, File(f_x), save=True)

            # Generar estructura detallada OTI y bloques de evidencias para las obligaciones 2..9
            generado_k = _generar_redaccion_demo_inteligente(informe_10, [det])["obligaciones"][str(k)]
            det.desarrollo_detallado_oti = generado_k["desarrollo_detallado_oti"]
            det.bloques_evidencias_json = generado_k["bloques_evidencias"]
            det.texto_tabla_principal = sanitizar_texto_para_sigaf_reportlab(info_k["texto_tabla_principal"])
            det.save()

    print(f"[OK] Pago 10 precargado con sus 9 obligaciones, textos detallados y soportes en BD.")
    print("=" * 80)


if __name__ == "__main__":
    ejecutar_seed()
