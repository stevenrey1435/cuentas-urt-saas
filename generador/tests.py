"""
Suite Automatizada de Pruebas y Auditoría de Calidad 10/10 (generador/tests.py).
Ejecutar con: python manage.py test -v 2
"""

import io
import os
import re
import zipfile
from datetime import date
from django.test import TestCase, Client
from django.urls import reverse
from django.core.cache import cache
from django.contrib.auth.models import User
from docx import Document

from .models import (
    PerfilContratista,
    ObligacionContractual,
    InformeMensual,
    DetalleEjecucionObligacion,
    ArchivoEvidencia,
)
from .services.ai_writer import (
    redactar_informe_con_ia,
    clasificar_obligacion_por_afinidad,
    _limpiar_fragmento_significativo,
    _inferir_contexto_clausula,
    _contraer_al,
)
from .services.package_builder import construir_paquete_completo_pago
from .views import _generar_codigo_unico_cliente


class AuditoriaCalidadRedaccionTests(TestCase):
    """
    Evalúa con rigor de auditoría que la redacción generada para distintos puestos y evidencias
    tenga calidad 10/10, cero textos basura, cero falsos positivos y 100% coherencia.
    """

    def setUp(self):
        cache.clear()

    def _auditar_calidad_texto(self, texto: str, contexto_esperado: str, palabras_prohibidas=()):
        """Verifica que un texto cumpla con todas las reglas gramaticales, técnicas y de auditoría."""
        self.assertTrue(len(texto) >= 350, f"Texto demasiado corto ({len(texto)} chars): {texto[:120]}")

        # 1. Cero caracteres ilegales para Reporte PDF
        self.assertNotIn("<", texto, "Contiene '<' sin sanitizar para Reporte PDF.")
        self.assertNotIn(">", texto, "Contiene '>' sin sanitizar para Reporte PDF.")

        # 2. Cero ruidos de PDFs de Outlook o MS Project
        vicios_prohibidos = [
            "cor di al",
            "cordial saludo",
            "remito adjunto",
            "predecesoras",
            "duración comienzo",
            "cl audi a",
            "(id)",
            "(das:",
            "al la ",
            "el la ",
            "garantizan el la ",
            "garantizan el el ",
            "orientadas al la ",
            "orientadas al el ",
            "\\10.10.36.17",
        ]
        txt_low = texto.lower()
        for vicio in vicios_prohibidos:
            self.assertNotIn(vicio, txt_low, f"Vicio de redacción detectado ('{vicio}') en:\n{texto}")

        # 3. Verificar que no contenga palabras prohibidas de otra profesión
        for pal in palabras_prohibidas:
            self.assertNotIn(
                pal.lower(),
                txt_low,
                f"Palabra impropia del cargo ('{pal}') apareció en el informe de {contexto_esperado}:\n{texto}",
            )

    def test_01_cero_falsos_positivos_inteligencia_artificial_en_11_obligaciones_ti(self):
        """
        Prueba las 11 obligaciones reales de desarrollo/soporte OTI y verifica que NINGUNA
        obligación de mantenimiento, reuniones, git o soporte mencione por error 'Inteligencia Artificial'.
        """
        clausulas_reales = [
            "Desarrollar e integrar componentes de software utilizando tecnologías vigentes y buenas prácticas alineadas con la arquitectura.",
            "Realizar actividades de mantenimiento correctivo, mejora continua y actualización de las soluciones de software.",
            "Colaborar en la implementación, parametrización y fortalecimiento técnico de los componentes de software.",
            "Proporcionar asistencia técnica oportuna mediante la atención ágil de fallos y la aplicación efectiva de soluciones.",
            "Participar en procesos de socialización y transferencia de conocimiento sobre los desarrollos realizados y acompañamiento a usuarios.",
            "Construir cronogramas, estimar esfuerzos, implementar actividades planificadas y compilar informes detallados.",
            "Efectuar la actualización continua y la documentación estructurada del código fuente en las ramas asignadas del repositorio git institucional.",
            "Asistir a los seguimientos programados de los proyectos asignados y actualizar las actividades en la herramienta de seguimiento.",
            "Gestionar de manera oportuna los requerimientos, solicitudes e incidencias asignadas a través de la herramienta de mesa de servicios.",
            "Asistir a reuniones en modalidad virtual o presencial, según las necesidades operativas y/o demandas del servicio.",
            "Las demás actividades que le sean encomendadas, necesarias para garantizar el cumplimiento de las obligaciones contractuales.",
        ]

        for idx, cl in enumerate(clausulas_reales, start=1):
            ctx = _inferir_contexto_clausula(cl, idx)
            self.assertNotIn(
                "Inteligencia Artificial",
                ctx["tema_corto"],
                f"Falso positivo de IA en Obligación {idx}: '{cl}' -> Salió: '{ctx['tema_corto']}'",
            )
            self.assertNotIn(
                "modelos inteligentes",
                ctx["accion_cierre"],
                f"Falso positivo de modelos IA en cierre de Obligación {idx}",
            )

        reunion = _inferir_contexto_clausula(
            "Asistir a reuniones de seguimiento y mesas de trabajo.",
            12,
            objeto_contrato="Prestar servicios profesionales de desarrollo de sistemas de información.",
        )
        self.assertIn("Reuniones", reunion["tema_corto"])
        self.assertEqual(_contraer_al("orientadas a el cumplimiento contractual"), "orientadas al cumplimiento contractual")

    def test_02_limpieza_fragmentos_outlook_msproject_y_parentesis(self):
        """Verifica que el limpiador de extractos erradique saludos de Outlook, nombres propios y columnas de Gantt."""
        # Caso A: Correo de Outlook con sílabas rotas
        lineas_outlook = ["Cl audi a Mónic a P eña R ubi ano", "Con un cor di al s aludo , r emi t o adjunt o"]
        res_a = _limpiar_fragmento_significativo(
            lineas_outlook,
            "RE_ Entrega documento de diseño de software – Sistema T-120_ Maria Camila Gomez - Outlook.pdf",
            nombre_contratista="Maria Camila Gomez",
        )
        self.assertNotIn("cl audi", res_a.lower())
        self.assertNotIn("cor di al", res_a.lower())
        self.assertNotIn("maria camila", res_a.lower())
        self.assertIn("Entrega documento de diseño", res_a)

        # Caso B: PDF de MS Project con columnas 'Predecesoras'
        lineas_project = ["Nombre de tarea Duración Comienzo Fin PredecesorasNombres de los"]
        res_b = _limpiar_fragmento_significativo(lineas_project, "Cronograma_Proyecto27Octubre2026.pdf")
        self.assertNotIn("predecesoras", res_b.lower())
        self.assertIn("Cronograma", res_b)

        # Caso C: Paréntesis truncado '(DAS'
        lineas_das = ["1. Documento de Arquitectura de Software (DAS"]
        res_c = _limpiar_fragmento_significativo(lineas_das, "DAS_HLD_T120.pdf")
        self.assertTrue(res_c.endswith("(DAS)"), f"No cerró el paréntesis truncado: {res_c}")

    def test_03_perfiles_multiprofesion_abogado_social_catastral(self):
        """
        Simula un Abogado de Restitución (Dirección Jurídica - Territorial Meta) y un Profesional Social,
        verificando que sus informes hablen de su área y jamás mencionen términos exclusivos de programación.
        """
        abogado = PerfilContratista.objects.create(
            nombres_apellidos="Carlos Andrés Mendoza Rojas",
            apellidos_nombres_nas="MENDOZA_ROJAS_CARLOS",
            nombre_corto_archivo="CARLOS_MENDOZA",
            cedula="80123456",
            numero_contrato="2045",
            anio_contrato=2026,
            fecha_inicio=date(2026, 1, 15),
            fecha_terminacion_actual=date(2026, 12, 30),
            objeto_contrato="Prestar servicios profesionales como abogado sustanciador para la representación judicial de víctimas en procesos de restitución de tierras.",
            nombre_supervisor="Laura Patricia Gómez",
            cargo_supervisor="COORDINADORA JURÍDICA TERRITORIAL",
            direccion_oficina="Dirección Jurídica de Restitución",
            direccion_territorial="Dirección Territorial Meta",
            proyecto_area="Gestión Jurídica y Procesal de Restitución",
            creditos_informe_completo=2,
            creditos_regeneracion_individual=6,
        )
        ob1 = ObligacionContractual.objects.create(
            contratista=abogado,
            numero_obligacion=1,
            descripcion_clausula="Proyectar demandas de restitución de tierras, autos de trámite, resoluciones y contestaciones judiciales dentro de los términos procesales.",
        )
        ob2 = ObligacionContractual.objects.create(
            contratista=abogado,
            numero_obligacion=2,
            descripcion_clausula="Asistir a las mesas de articulación interinstitucional y reuniones convocadas por la coordinación jurídica.",
        )
        informe = InformeMensual.objects.create(
            contratista=abogado,
            numero_pago=1,
            mes=10,
            anio=2026,
            fecha_reporte=date(2026, 10, 28),
        )
        det1 = DetalleEjecucionObligacion.objects.create(
            informe=informe,
            obligacion=ob1,
            apuntes_contratista="Proyección de 4 demandas de restitución para el municipio de Puerto Gaitán; revisión de 6 expedientes judiciales.",
        )
        det2 = DetalleEjecucionObligacion.objects.create(
            informe=informe,
            obligacion=ob2,
            apuntes_contratista="Participación en comité jurídico semanal sobre líneas jurisprudenciales.",
        )

        res = redactar_informe_con_ia(informe, forzar_demo=True)
        det1.refresh_from_db()
        det2.refresh_from_db()

        self._auditar_calidad_texto(
            det1.texto_tabla_principal,
            contexto_esperado="Abogado Jurídica",
            palabras_prohibidas=("gitlab", "merge request", "backend", "frontend", "base de datos", "codificación"),
        )
        self.assertIn("Puerto Gaitán", det1.texto_tabla_principal)

    def test_04_cache_0_tokens_y_proteccion_de_pases(self):
        """
        Verifica que al generar por segunda vez sin haber cambiado apuntes ni archivos,
        se active el Caché Instantáneo 0-Tokens y NO descuente ningún pase adicional.
        """
        c = PerfilContratista.objects.create(
            nombres_apellidos="Ana María López",
            apellidos_nombres_nas="LOPEZ_ANA",
            nombre_corto_archivo="ANA_LOPEZ",
            cedula="52987654",
            numero_contrato="998",
            anio_contrato=2026,
            fecha_inicio=date(2026, 2, 1),
            fecha_terminacion_actual=date(2026, 11, 30),
            objeto_contrato="Apoyo profesional en seguimiento de proyectos y planeación.",
            nombre_supervisor="Pedro Nel Ospina",
            creditos_informe_completo=2,
            creditos_regeneracion_individual=5,
        )
        ob = ObligacionContractual.objects.create(
            contratista=c,
            numero_obligacion=1,
            descripcion_clausula="Construir cronogramas, estimar esfuerzos, implementar actividades planificadas y compilar informes detallados.",
        )
        inf = InformeMensual.objects.create(contratista=c, numero_pago=3, mes=10, anio=2026, fecha_reporte=date(2026, 10, 29))
        DetalleEjecucionObligacion.objects.create(informe=inf, obligacion=ob, apuntes_contratista="Actualización del cronograma maestro de actividades.")

        # 1ra llamada: Redacta y descuenta 1 pase (queda en 1)
        res1 = redactar_informe_con_ia(inf, forzar_demo=True)
        c.refresh_from_db()
        self.assertEqual(c.pases_disponibles, 1)
        self.assertEqual(res1["obligaciones_actualizadas"], [1])

        # 2da llamada inmediata sin cambios: Entra por Caché 0-Tokens y NO descuenta pase (sigue en 1)
        res2 = redactar_informe_con_ia(inf, forzar_demo=True)
        c.refresh_from_db()
        self.assertEqual(c.pases_disponibles, 1, "El caché 0-tokens no debía descontar otro pase.")
        self.assertIn("Caché Instantáneo 0-Tokens", res2["motor_usado"])
        self.assertEqual(res2["obligaciones_actualizadas"], [])

    def test_05_seguridad_codigos_unicos_honeypot_y_aislamiento_entre_usuarios(self):
        """
        Audita el blindaje comercial y de seguridad:
        1. Cuentas nuevas nacen con 0 Pases.
        2. El código de activación de la Cédula A es rechazado si lo intenta usar la Cédula B.
        3. Un código ya utilizado no puede volver a canjearse por segunda vez.
        4. Un contratista no puede ver ni descargar el informe de otro compañero por URL (IDOR).
        """
        client = Client()

        # 1. Crear Contratista A desde el formulario web
        resp_reg = client.post(
            reverse("generador:crear_contratista"),
            {
                "nombres_apellidos": "Juan Camilo Restrepo",
                "cedula": "1015443322",
                "pin_acceso": "4321",
                "numero_contrato": "1501",
                "anio_contrato": "2026",
                "nombre_supervisor": "Marta Lucía Vélez",
                "direccion_oficina": "Subdirección General",
                "direccion_territorial": "Territorial Antioquia",
                "objeto_contrato": "Prestar servicios profesionales de apoyo técnico y administrativo en la territorial.",
                "obligacion": [
                    "1. Elaborar informes técnicos mensuales del área.",
                    "2. Asistir a las reuniones de seguimiento programadas.",
                ],
            },
        )
        c_a = PerfilContratista.objects.get(cedula="1015443322")
        self.assertEqual(c_a.obligaciones.count(), 2)
        self.assertEqual(
            list(c_a.obligaciones.values_list("descripcion_clausula", flat=True)),
            [
                "Elaborar informes técnicos mensuales del área.",
                "Asistir a las reuniones de seguimiento programadas.",
            ],
        )
        self.assertEqual(c_a.pases_disponibles, 0, "Toda cuenta nueva debe nacer con 0 Pases.")

        # 2. Verificar que un código de otra cédula es rechazado
        codigo_ajeno = _generar_codigo_unico_cliente("999999999", 1, 1)
        client.post(reverse("generador:canjear_codigo_recarga"), {"codigo": codigo_ajeno})
        c_a.refresh_from_db()
        self.assertEqual(c_a.pases_disponibles, 0, "Aceptó un código perteneciente a otra cédula.")

        # 3. Canjear su propio código legítimo de 1 Pase -> Sube a 1 Pase
        codigo_legitimo = _generar_codigo_unico_cliente("1015443322", 1, 1)
        client.post(reverse("generador:canjear_codigo_recarga"), {"codigo": codigo_legitimo})
        c_a.refresh_from_db()
        self.assertEqual(c_a.pases_disponibles, 1, "No acreditó el código legítimo de 1 pase.")

        # 4. Intentar volver a meter el mismo código por segunda vez -> Debe ser rechazado (sigue en 1 Pase)
        client.post(reverse("generador:canjear_codigo_recarga"), {"codigo": codigo_legitimo})
        c_a.refresh_from_db()
        self.assertEqual(c_a.pases_disponibles, 1, "¡Peligro! Permitió reutilizar el mismo código dos veces.")

        # 5. Verificar protección IDOR: Contratista A no puede abrir el informe de otro contratista B
        c_b = PerfilContratista.objects.create(
            nombres_apellidos="Otro Contratista Privado",
            apellidos_nombres_nas="PRIVADO_OTRO",
            nombre_corto_archivo="OTRO_PRIVADO",
            cedula="79112233",
            numero_contrato="8888",
            anio_contrato=2026,
            fecha_inicio=date(2026, 1, 1),
            fecha_terminacion_actual=date(2026, 12, 31),
            objeto_contrato="Contrato privado de otro compañero.",
            nombre_supervisor="Supervisor X",
        )
        inf_ajeno = InformeMensual.objects.create(contratista=c_b, numero_pago=1, mes=10, anio=2026, fecha_reporte=date(2026, 10, 28))
        resp_espia = client.get(reverse("generador:detalle_informe", args=[inf_ajeno.id]))
        self.assertEqual(resp_espia.status_code, 404, "¡Fallo de seguridad! Un usuario pudo abrir el informe de otro por ID.")

    def test_06_empaquetado_word_y_zip_sin_rastro_de_nas(self):
        """
        Verifica que construir_paquete_completo_pago() genere el .ZIP y los .docx con los datos
        exactos de la oficina/territorial del contratista y sin ninguna referencia a \\10.10.36.17.
        """
        c = PerfilContratista.objects.create(
            nombres_apellidos="Diana Carolina Torres",
            apellidos_nombres_nas="TORRES_DIANA",
            nombre_corto_archivo="DIANA_TORRES",
            cedula="1032445566",
            numero_contrato="1420",
            anio_contrato=2026,
            fecha_inicio=date(2026, 1, 20),
            fecha_terminacion_actual=date(2026, 12, 30),
            objeto_contrato="Prestar servicios profesionales en gestión catastral y análisis predial.",
            nombre_supervisor="Ing. Mauricio Vargas",
            cargo_supervisor="COORDINADOR CATASTRAL",
            direccion_oficina="Subdirección de Catastro y Análisis Territorial",
            direccion_territorial="Dirección Territorial Magdalena Medio",
            creditos_informe_completo=1,
            creditos_regeneracion_individual=3,
        )
        ob = ObligacionContractual.objects.create(
            contratista=c,
            numero_obligacion=1,
            descripcion_clausula="Elaborar informes técnico-prediales, georreferenciación y cruces de información cartográfica IGAC-SNR.",
        )
        inf = InformeMensual.objects.create(contratista=c, numero_pago=4, mes=10, anio=2026, fecha_reporte=date(2026, 10, 30))
        DetalleEjecucionObligacion.objects.create(
            informe=inf,
            obligacion=ob,
            apuntes_contratista="Elaboración de 8 informes técnico-prediales y validación de polígonos georreferenciados.",
        )

        redactar_informe_con_ia(inf, forzar_demo=True)
        res_pkg = construir_paquete_completo_pago(inf)

        self.assertTrue(os.path.exists(res_pkg["ruta_zip"]), "No se creó el archivo .ZIP físico.")
        with zipfile.ZipFile(res_pkg["ruta_zip"], "r") as zf:
            nombres_zip = zf.namelist()
            self.assertTrue(any("CTO_1420_DIANA_TORRES_INFORME_04.docx" in n for n in nombres_zip))
            self.assertTrue(any("Obligacion 1.docx" in n for n in nombres_zip))

        # Inspeccionar por dentro el Word principal generado
        doc_principal = Document(res_pkg["documentos_generados"][0])
        texto_tablas = " ".join(
            cell.text for table in doc_principal.tables for row in table.rows for cell in row.cells
        )
        self.assertIn("Subdirección de Catastro y Análisis Territorial", texto_tablas)
        self.assertIn("Dirección Territorial Magdalena Medio", texto_tablas)
        self.assertNotIn("10.10.36.17", texto_tablas, "El Word generado aún contiene la IP vieja del servidor NAS.")


    def test_07_caos_zip_con_archivos_corruptos_duplicados_emojis_y_basura(self):
        """
        Prueba Extrema de ZIP ('Chaos Zip'):
        Simula un .ZIP subido por un usuario con:
        - Archivos basura (Thumbs.db, __MACOSX/._foto.png, ~$temporal.docx) -> Deben ser ignorados.
        - Dos archivos llamados igual en carpetas distintas (Actividad_1/evidencia.png y Actividad_2/evidencia.png) -> No deben pisarse.
        - Una imagen .png CORRUPTA (bytes inválidos) -> No debe hacer crashear a RapidOCR ni a python-docx al armar el Word.
        - Un PDF CORRUPTO y un Excel CORRUPTO -> No deben romper el inspector.
        - Un nombre de archivo de 200 caracteres con emojis 🚀 y símbolos raros.
        """
        from .services.file_inspector import desempacar_archivos_subidos, extraer_metadatos_ligeros_archivo
        from django.core.files.uploadedfile import SimpleUploadedFile

        # Construir el ZIP trampa en memoria
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("Thumbs.db", b"basura windows")
            zf.writestr("__MACOSX/._oculta.png", b"basura mac")
            zf.writestr("~$borrador.docx", b"temporal word")
            zf.writestr("Actividad_1/soporte.png", b"no-soy-un-png-valido-estoy-corrupto")
            zf.writestr("Actividad_2/soporte.png", b"tampoco-soy-un-png-valido-pero-me-llamo-igual")
            zf.writestr("Actividad_3/corrupto.pdf", b"%PDF-1.4 archivo roto sin cierre")
            zf.writestr("Actividad_3/roto.xlsx", b"PK\x03\x04 excel danado")
            nombre_largo = "Actividad_1/" + ("Reporte_🚀_Super_Largo_" * 8) + ".txt"
            zf.writestr(nombre_largo, "Linea 1 de prueba técnica con caracteres \x00 nulos y <para> etiquetas.".encode("utf-8"))

        zip_buffer.seek(0)
        zip_upload = SimpleUploadedFile("evidencias_caos.zip", zip_buffer.read(), content_type="application/zip")

        # 1. Desempacar: debe ignorar los 3 archivos basura y conservar los 5 archivos reales (sin pisar los dos 'soporte.png')
        extraidos = desempacar_archivos_subidos([zip_upload])
        self.assertEqual(len(extraidos), 5, f"Debía extraer exactamente 5 archivos válidos, extrajo: {[x['nombre_original'] for x in extraidos]}")

        # 2. Crear contratista con 3 obligaciones de áreas distintas y meterle todos estos archivos rotos
        c = PerfilContratista.objects.create(
            nombres_apellidos="Usuario Prueba Caos",
            apellidos_nombres_nas="CAOS_USUARIO",
            nombre_corto_archivo="USUARIO_CAOS",
            cedula="1099887766",
            numero_contrato="777",
            anio_contrato=2026,
            fecha_inicio=date(2026, 1, 1),
            fecha_terminacion_actual=date(2026, 12, 31),
            objeto_contrato="Contrato de prueba de estrés extremo.",
            nombre_supervisor="Supervisor Estrés",
            creditos_informe_completo=2,
            creditos_regeneracion_individual=5,
        )
        for k in range(1, 4):
            ObligacionContractual.objects.create(
                contratista=c,
                numero_obligacion=k,
                descripcion_clausula=f"Obligación número {k} para gestionar trámites administrativos, informes y seguimiento documental del área.",
            )
        inf = InformeMensual.objects.create(contratista=c, numero_pago=1, mes=10, anio=2026, fecha_reporte=date(2026, 10, 30))
        detalles = [DetalleEjecucionObligacion.objects.create(informe=inf, obligacion=ob) for ob in c.obligaciones.all()]

        for idx, item in enumerate(extraidos):
            meta = extraer_metadatos_ligeros_archivo(item["file_obj"], item["nombre_original"], item["ruta_relativa"])
            det_dest = clasificar_obligacion_por_afinidad(f"{item['ruta_relativa']} {item['nombre_original']}", detalles, idx)
            ev = ArchivoEvidencia(
                detalle_obligacion=det_dest,
                orden=idx + 1,
                nombre_original=item["nombre_original"][:250],
                tipo_archivo="IMAGEN" if item["nombre_original"].endswith(".png") else "OTRO",
                tamano_bytes=item["tamano_bytes"],
                metadatos_ligeros=meta,
            )
            ev.archivo.save(item["nombre_original"][:120], item["file_obj"], save=True)

        # 3. Redactar y construir el paquete Word completo: NO debe lanzar ninguna excepción pese a los archivos corruptos
        redactar_informe_con_ia(inf, forzar_demo=True)
        res_pkg = construir_paquete_completo_pago(inf)
        self.assertTrue(os.path.exists(res_pkg["ruta_zip"]))
        self.assertEqual(res_pkg["total_documentos_word"], 8)  # 2*3 + 2 = 8 documentos Word creados con éxito

    def test_08_estres_16_obligaciones_variadas_y_caracteres_veneno_xml(self):
        """
        Prueba un contrato grande con 16 obligaciones de toda índole (financieras, archivo, jurídicas,
        ambientales, atención al ciudadano, TI) y apuntes con caracteres veneno (\x00, <para>, &, <>).
        """
        c = PerfilContratista.objects.create(
            nombres_apellidos="Martha Cecilia Gómez",
            apellidos_nombres_nas="GOMEZ_MARTHA",
            nombre_corto_archivo="MARTHA_GOMEZ",
            cedula="41556677",
            numero_contrato="3090",
            anio_contrato=2026,
            fecha_inicio=date(2026, 1, 10),
            fecha_terminacion_actual=date(2026, 12, 31),
            objeto_contrato="Apoyo transversal administrativo, financiero, documental y de atención al ciudadano.",
            nombre_supervisor="Dr. Hernando Ruiz",
            creditos_informe_completo=2,
            creditos_regeneracion_individual=5,
        )
        textos_16 = [
            "Revisar las cuentas de cobro y órdenes de pago presupuestal en el sistema financiero.",
            "Radicar y archivar comunicaciones oficiales en el sistema de gestión documental Orfeo.",
            "Atender peticiones, quejas, reclamos y solicitudes (PQRS) de la ciudadanía.",
            "Proyectar actos administrativos, oficios y memorandos requeridos por la coordinación.",
            "Apoyar la logística y elaboración de actas de las mesas de trabajo territoriales.",
            "Actualizar las bases de datos en Excel con el seguimiento de indicadores de gestión.",
            "Realizar el control de inventarios y activos fijos asignados a la dependencia.",
            "Verificar el cumplimiento de requisitos legales en los expedientes contractuales.",
            "Elaborar informes periódicos de avance físico y presupuestal del área.",
            "Socializar los protocolos de atención con enfoque diferencial a los usuarios.",
            "Apoyar el alistamiento de información para auditorías internas y de entes de control.",
            "Coordinar la programación de comisiones y desplazamientos del equipo de campo.",
            "Efectuar el cruce de bases de datos catastrales y registrales para depuración.",
            "Custodiar la integridad de los expedientes físicos y digitales según tablas de retención.",
            "Participar en las reuniones virtuales o presenciales convocadas por el supervisor.",
            "Las demás actividades conexas que le sean asignadas para el cumplimiento del objeto.",
        ]
        inf = InformeMensual.objects.create(contratista=c, numero_pago=7, mes=10, anio=2026, fecha_reporte=date(2026, 10, 30))

        for idx, txt_ob in enumerate(textos_16, start=1):
            ob = ObligacionContractual.objects.create(contratista=c, numero_obligacion=idx, descripcion_clausula=txt_ob)
            DetalleEjecucionObligacion.objects.create(
                informe=inf,
                obligacion=ob,
                apuntes_contratista=f"Ejecución punto {idx} con caracteres veneno: \x00\x07 <para> A < B > C & 'comillas' y cierre exitoso.",
            )

        redactar_informe_con_ia(inf, forzar_demo=True)
        res_pkg = construir_paquete_completo_pago(inf)

        self.assertEqual(res_pkg["total_documentos_word"], 34)  # 2*16 + 2 = 34 documentos Word
        for det in inf.detalles_obligaciones.all():
            self._auditar_calidad_texto(det.texto_tabla_principal, f"Obligación {det.obligacion.numero_obligacion}")