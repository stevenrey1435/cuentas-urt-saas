"""
Modelos de Base de Datos para Cuentas al Día (Generador Automatizado de Cuentas de Cobro para Contratistas).
"""

import os
from django.db import models
from django.contrib.auth.models import User
from django.conf import settings


class PerfilContratista(models.Model):
    """
    Datos contractuales fijos del contratista, rutas documentales
    y control de créditos mensuales.
    """
    user = models.OneToOneField(
        User,
        on_delete=models.CASCADE,
        related_name="perfil_contratista",
        null=True,
        blank=True,
        verbose_name="Usuario del Sistema",
    )
    nombres_apellidos = models.CharField(
        max_length=200,
        verbose_name="Nombres y Apellidos del Contratista",
    )
    apellidos_nombres_nas = models.CharField(
        max_length=200,
        help_text="Formato APELLIDOS_NOMBRES para la carpeta en el servidor NAS",
        verbose_name="Apellidos_Nombres (Formato NAS)",
    )
    nombre_corto_archivo = models.CharField(
        max_length=100,
        help_text="Ej: NOMBRE_APELLIDO para nombrar CTO_<NUM>_<NOMBRE>_INFORME_<XX>.docx",
        verbose_name="Nombre Corto para Archivo",
    )
    cedula = models.CharField(
        max_length=30,
        unique=True,
        verbose_name="Número Documento de Identificación (C.C.)",
    )
    numero_contrato = models.CharField(
        max_length=30,
        verbose_name="Número de Contrato",
    )
    anio_contrato = models.PositiveIntegerField(
        default=2026,
        verbose_name="Año del Contrato",
    )
    fecha_inicio = models.DateField(
        verbose_name="Fecha de Inicio",
    )
    fecha_terminacion_actual = models.DateField(
        verbose_name="Fecha de Terminación Actual",
    )
    objeto_contrato = models.TextField(
        verbose_name="Objeto del Contrato",
    )
    nombre_supervisor = models.CharField(
        max_length=200,
        verbose_name="Nombre(s) Supervisor(es)",
    )
    cargo_supervisor = models.CharField(
        max_length=200,
        default="PROFESIONAL ESPECIALIZADO",
        verbose_name="Cargo Supervisor(es)",
    )
    direccion_oficina = models.CharField(
        max_length=200,
        default="Oficina de Gestión Contractual",
        verbose_name="Dirección Técnica, Oficina o Grupo",
    )
    direccion_territorial = models.CharField(
        max_length=120,
        default="Nivel Central",
        verbose_name="Dirección Territorial",
    )
    subcarpeta_rol_nas = models.CharField(
        max_length=100,
        default="Contratistas",
        help_text="Rol o subcarpeta documental",
        verbose_name="Subcarpeta de Rol",
    )
    proyecto_area = models.CharField(
        max_length=200,
        default="Gestión y Supervisión Contractual",
        help_text="Subtítulo que aparece en los compendios",
        verbose_name="Proyecto o Área Asignada",
    )
    servidor_nas_prefix = models.CharField(
        max_length=255,
        default=r"\\10.10.36.17\nivel_central\oti\4. sistemas de informacion",
        verbose_name="Prefijo Servidor NAS",
    )
    firma_imagen = models.FileField(
        upload_to="firmas/",
        null=True,
        blank=True,
        verbose_name="Firma Digitalizada (firma.jpg / .png)",
    )

    # Control de acceso y tope mensual de créditos IA
    activo_mes_actual = models.BooleanField(
        default=True,
        verbose_name="Activo para Generar en el Mes Actual",
    )
    creditos_informe_completo = models.PositiveSmallIntegerField(
        default=0,
        verbose_name="Pases Disponibles (Cuentas de Cobro)",
    )
    creditos_regeneracion_individual = models.PositiveSmallIntegerField(
        default=0,
        verbose_name="Ajustes Individuales Restantes",
    )
    max_caracteres_apuntes = models.PositiveIntegerField(
        default=2500,
        verbose_name="Máximo de Caracteres en Apuntes por Obligación",
    )
    creado_en = models.DateTimeField(auto_now_add=True)
    actualizado_en = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Perfil de Contratista OTI"
        verbose_name_plural = "Perfiles de Contratistas OTI"

    def __str__(self):
        return f"CTO {self.numero_contrato}-{self.anio_contrato} | {self.nombres_apellidos} (CC {self.cedula})"

    @property
    def cedula_formateada(self) -> str:
        """Retorna la cédula con puntos de mil (ej. 1.006.552.477)."""
        limpia = "".join(ch for ch in str(self.cedula) if ch.isdigit())
        if not limpia:
            return str(self.cedula)
        return f"{int(limpia):,}".replace(",", ".")

    @property
    def pases_disponibles(self) -> int:
        """Pases disponibles para generar cuentas de cobro completas."""
        return self.creditos_informe_completo

    def autocalcular_campos_internos(self):
        """
        Calcula de forma transparente y sin molestar al usuario:
        - apellidos_nombres_nas: Formato APELLIDOS_NOMBRES
        - nombre_corto_archivo: Formato NOMBRE_APELLIDO
        - subcarpeta_rol_nas: Rol predeterminado
        """
        import unicodedata, re
        norm = unicodedata.normalize('NFKD', self.nombres_apellidos or '').encode('ASCII', 'ignore').decode('utf-8')
        palabras = [re.sub(r'[^A-Za-z0-9]', '', p).upper() for p in norm.split() if p.strip()]
        if not palabras:
            palabras = ["CONTRATISTA", "INFORME"]
        
        if len(palabras) == 1:
            self.nombre_corto_archivo = palabras[0]
            self.apellidos_nombres_nas = palabras[0]
        elif len(palabras) == 2:
            self.nombre_corto_archivo = f"{palabras[0]}_{palabras[1]}"
            self.apellidos_nombres_nas = f"{palabras[1]}_{palabras[0]}"
        elif len(palabras) == 3:
            self.nombre_corto_archivo = f"{palabras[0]}_{palabras[1]}"
            self.apellidos_nombres_nas = f"{palabras[1]}_{palabras[2]}_{palabras[0]}"
        else:
            nombres = palabras[:-2]
            apellidos = palabras[-2:]
            self.nombre_corto_archivo = f"{nombres[0]}_{apellidos[0]}"
            self.apellidos_nombres_nas = f"{'_'.join(apellidos)}_{'_'.join(nombres)}"

        if not self.subcarpeta_rol_nas:
            self.subcarpeta_rol_nas = "Profesionales"
        if not self.servidor_nas_prefix:
            self.servidor_nas_prefix = r"\\10.10.36.17\nivel_central\oti\4. sistemas de informacion"

    @property
    def ruta_base_nas(self) -> str:
        """
        Retorna la ruta base NAS exacta según la Circular 0009 de 2024:
        \\\\10.10.36.17\\nivel_central\\oti\\4. sistemas de informacion\\<AÑO>\\01_Evidencias_OPS\\<SUBCARPETA_ROL>\\<NUM_CTO>_<AÑO>_CC_<CEDULA>_<APELLIDOS_NOMBRES>
        """
        prefix = self.servidor_nas_prefix.rstrip("\\")
        carpeta_contratista = f"{self.numero_contrato}_{self.anio_contrato}_CC_{self.cedula}_{self.apellidos_nombres_nas}"
        return f"{prefix}\\{self.anio_contrato}\\01_Evidencias_OPS\\{self.subcarpeta_rol_nas}\\{carpeta_contratista}"

    def obtener_ruta_pago_nas(self, numero_pago: int, enlace_personalizado: str = "") -> str:
        """Retorna el enlace de Drive/OneDrive si el usuario lo puso, o una ruta de expediente limpia y genérica."""
        if enlace_personalizado and enlace_personalizado.strip():
            return enlace_personalizado.strip()
        carpeta_cto = f"CTO_{self.numero_contrato}_{self.anio_contrato}_{self.apellidos_nombres_nas}"
        return f"Expediente Digital de Evidencias / {carpeta_cto} / Pago {int(numero_pago):02d}"

    def obtener_ruta_actividad_nas(self, numero_pago: int, numero_obligacion: int, enlace_personalizado: str = "") -> str:
        """Ruta limpia para la columna 4 de la tabla del Informe GC-FO-15 (sin IPs ni servidores NAS internos)."""
        base = self.obtener_ruta_pago_nas(numero_pago, enlace_personalizado)
        return f"{base} / Actividad_{int(numero_obligacion)}"

    def obtener_ruta_firma_fisica(self) -> str:
        """
        Retorna la ruta en disco de la firma propia subida por este contratista en su perfil (`firma_imagen`).
        JAMÁS usa la firma de otro contratista ni un fallback global: si el contratista no ha subido su firma,
        retorna cadena vacía ("") para dejar el espacio de firma en blanco.
        """
        if self.firma_imagen and hasattr(self.firma_imagen, "path") and os.path.exists(self.firma_imagen.path):
            return self.firma_imagen.path
        return ""


class ObligacionContractual(models.Model):
    """
    Clausulado específico de obligaciones contractuales (1..N) asociadas a un PerfilContratista.
    """
    contratista = models.ForeignKey(
        PerfilContratista,
        on_delete=models.CASCADE,
        related_name="obligaciones",
        verbose_name="Contratista",
    )
    numero_obligacion = models.PositiveSmallIntegerField(
        verbose_name="Número de Obligación (1..N)",
    )
    descripcion_clausula = models.TextField(
        verbose_name="Texto Literal de la Obligación Contractual",
    )

    class Meta:
        ordering = ["numero_obligacion"]
        unique_together = [("contratista", "numero_obligacion")]
        verbose_name = "Obligación Contractual"
        verbose_name_plural = "Obligaciones Contractuales"

    def __str__(self):
        return f"Obligación {self.numero_obligacion} (CTO {self.contratista.numero_contrato}): {self.descripcion_clausula[:65]}..."


class InformeMensual(models.Model):
    """
    Representa una cuenta de cobro / Pago Mensual (ej. Pago 9, Pago 10) de un contratista.
    """
    MESES_CHOICES = [
        (1, "Enero"),
        (2, "Febrero"),
        (3, "Marzo"),
        (4, "Abril"),
        (5, "Mayo"),
        (6, "Junio"),
        (7, "Julio"),
        (8, "Agosto"),
        (9, "Septiembre"),
        (10, "Octubre"),
        (11, "Noviembre"),
        (12, "Diciembre"),
    ]

    ESTADOS_CHOICES = [
        ("BORRADOR", "Borrador en edición"),
        ("REDACTADO_IA", "Textos redactados por IA"),
        ("EMPAQUETADO", "Árbol de carpetas y ZIP generados"),
    ]

    contratista = models.ForeignKey(
        PerfilContratista,
        on_delete=models.CASCADE,
        related_name="informes_mensuales",
        verbose_name="Contratista",
    )
    numero_pago = models.PositiveSmallIntegerField(
        default=10,
        verbose_name="Número de Pago",
    )
    mes = models.PositiveSmallIntegerField(
        choices=MESES_CHOICES,
        default=10,
        verbose_name="Mes Reportado",
    )
    anio = models.PositiveIntegerField(
        default=2026,
        verbose_name="Año Reportado",
    )
    fecha_reporte = models.DateField(
        verbose_name="Fecha de Emisión del Reporte",
    )
    observaciones_adicionales = models.TextField(
        blank=True,
        default="",
        verbose_name="Sección 4: Observaciones o Información Adicional (Opcional)",
    )
    enlace_nube_evidencias = models.CharField(
        max_length=500,
        blank=True,
        default="",
        verbose_name="Enlace de Carpeta Digital en la Nube (Opcional - OneDrive / Drive)",
    )
    captura_explorador_nas = models.FileField(
        upload_to="evidencias_nas/",
        null=True,
        blank=True,
        verbose_name="Captura del Explorador de Archivos NAS (Para Evidencia_Pago_XX.docx)",
    )
    informes_secop_publicados = models.PositiveSmallIntegerField(
        default=5,
        help_text="Cantidad de informes anteriores relacionados como 'Sí' en la Tabla 2 de SECOP II",
        verbose_name="Informes Publicados en SECOP II",
    )
    estado = models.CharField(
        max_length=30,
        choices=ESTADOS_CHOICES,
        default="BORRADOR",
        verbose_name="Estado del Informe",
    )
    ultimo_motor_ia_usado = models.CharField(
        max_length=120,
        blank=True,
        default="",
        verbose_name="Último Motor IA Utilizado",
    )
    ruta_carpeta_paquete = models.CharField(
        max_length=500,
        blank=True,
        default="",
        verbose_name="Ruta Local del Paquete Generado",
    )
    archivo_zip_generado = models.CharField(
        max_length=500,
        blank=True,
        default="",
        verbose_name="Ruta del Archivo .ZIP Descargable",
    )
    creado_en = models.DateTimeField(auto_now_add=True)
    actualizado_en = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-anio", "-numero_pago"]
        unique_together = [("contratista", "numero_pago", "anio")]
        verbose_name = "Informe Mensual de Pago"
        verbose_name_plural = "Informes Mensuales de Pago"

    def __str__(self):
        return f"Pago {self.numero_pago:02d} ({self.nombre_mes} {self.anio}) - {self.contratista.nombres_apellidos}"

    @property
    def nombre_mes(self) -> str:
        return dict(self.MESES_CHOICES).get(self.mes, "Octubre")

    @property
    def nombre_mes_mayus(self) -> str:
        return self.nombre_mes.upper()

    @property
    def fecha_reporte_texto_largo(self) -> str:
        """Ejemplo: '28 de octubre de 2026'."""
        mes_minus = self.nombre_mes.lower()
        return f"{self.fecha_reporte.day} de {mes_minus} de {self.fecha_reporte.year}"

    @property
    def ruta_pago_nas(self) -> str:
        enlace = (getattr(self, "enlace_nube_evidencias", "") or "").strip()
        if not enlace:
            pref = (self.contratista.servidor_nas_prefix or "").strip()
            if pref.startswith("http"):
                enlace = pref
        return self.contratista.obtener_ruta_pago_nas(self.numero_pago, enlace)

    @property
    def nombre_archivo_informe_principal(self) -> str:
        return f"CTO_{self.contratista.numero_contrato}_{self.contratista.nombre_corto_archivo}_INFORME_{self.numero_pago:02d}.docx"

    @property
    def nombre_archivo_evidencia_pago(self) -> str:
        return f"Evidencia_Pago_{self.numero_pago:02d}.docx"


class DetalleEjecucionObligacion(models.Model):
    """
    Almacena tanto los apuntes cortos del contratista para una Obligación K en un Pago X,
    como los tres bloques coherentes redactados por la IA:
    1) texto_tabla_principal (para CTO_<NUM>_INFORME_<XX>.docx y botón Copiar Actividad)
    2) desarrollo_detallado_oti (JSON o texto con introducción, viñetas de desarrollo, productos y anexos para REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx)
    3) bloques_evidencias_json (títulos y descripciones técnicas numeradas para Anexos/Obligacion <K>.docx)
    """
    informe = models.ForeignKey(
        InformeMensual,
        on_delete=models.CASCADE,
        related_name="detalles_obligaciones",
        verbose_name="Informe Mensual",
    )
    obligacion = models.ForeignKey(
        ObligacionContractual,
        on_delete=models.CASCADE,
        related_name="ejecuciones_mensuales",
        verbose_name="Obligación Contractual",
    )
    apuntes_contratista = models.TextField(
        blank=True,
        default="",
        help_text="Apuntes cortos de las actividades realizadas en el mes para esta obligación",
        verbose_name="Apuntes Cortos del Contratista",
    )
    # Salida 1: Celda de Actividades Desarrolladas en CTO_<NUM>_<NOMBRE>_INFORME_<XX>.docx
    texto_tabla_principal = models.TextField(
        blank=True,
        default="",
        verbose_name="Resumen Numerado (Tabla Informe Principal)",
    )
    # Salida 2: Cuerpo estructurado para REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx
    desarrollo_detallado_oti = models.JSONField(
        default=dict,
        blank=True,
        verbose_name="Estructura Detallada para Reporte Individual OTI",
    )
    # Salida 3: Bloques de evidencias para Anexos/Obligacion <K>.docx
    bloques_evidencias_json = models.JSONField(
        default=list,
        blank=True,
        verbose_name="Bloques Explicativos para Documento Obligacion K.docx",
    )
    actualizado_en = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["obligacion__numero_obligacion"]
        unique_together = [("informe", "obligacion")]
        verbose_name = "Detalle de Ejecución por Obligación"
        verbose_name_plural = "Detalles de Ejecución por Obligación"

    def __str__(self):
        return f"Pago {self.informe.numero_pago:02d} - Obligación {self.obligacion.numero_obligacion}"

    @property
    def referencia_digital_actividad(self) -> str:
        """Referencia limpia y universal para el informe de supervisión."""
        enlace = (getattr(self.informe, "enlace_nube_evidencias", "") or "").strip()
        if not enlace:
            pref = (self.informe.contratista.servidor_nas_prefix or "").strip()
            if pref.startswith("http"):
                enlace = pref
        if enlace:
            return f"{enlace} - Actividad_{self.obligacion.numero_obligacion}"
        return f"Expediente Digital de Evidencias / Pago {self.informe.numero_pago:02d} / Actividad_{self.obligacion.numero_obligacion}/Anexos"

    @property
    def ruta_nas_actividad(self) -> str:
        return self.referencia_digital_actividad


def ruta_subida_evidencia(instance, filename: str) -> str:
    """Organiza los archivos subidos por contrato, año y pago."""
    det = instance.detalle_obligacion
    inf = det.informe
    c = inf.contratista
    return f"evidencias_soporte/cto_{c.numero_contrato}_{c.anio_contrato}/pago_{inf.numero_pago:02d}/ob_{det.obligacion.numero_obligacion}/{filename}"


class ArchivoEvidencia(models.Model):
    """
    Representa un soporte documental o gráfico (captura de Teams, GitLab, PR, SQL, Excel, PDF, etc.)
    subido por el contratista y asociado a una obligación K.
    """
    TIPOS_ARCHIVO = [
        ("IMAGEN", "Captura de Pantalla / Imagen (.png, .jpg)"),
        ("EXCEL", "Hoja de Cálculo (.xlsx, .csv)"),
        ("PDF", "Documento PDF (.pdf)"),
        ("WORD", "Documento Word (.docx)"),
        ("CODIGO", "Script / Código Fuente (.py, .sql, .js, .sh)"),
        ("TEXTO", "Texto Plano / Logs (.txt, .log)"),
        ("OTRO", "Otro Tipo de Soporte"),
    ]

    detalle_obligacion = models.ForeignKey(
        DetalleEjecucionObligacion,
        on_delete=models.CASCADE,
        related_name="archivos_evidencia",
        verbose_name="Obligación Asociada",
    )
    orden = models.PositiveSmallIntegerField(
        default=1,
        verbose_name="Orden en el Documento de Anexos",
    )
    nombre_original = models.CharField(
        max_length=255,
        verbose_name="Nombre Original del Archivo",
    )
    archivo = models.FileField(
        upload_to=ruta_subida_evidencia,
        null=True,
        blank=True,
        verbose_name="Archivo Físico",
    )
    tipo_archivo = models.CharField(
        max_length=20,
        choices=TIPOS_ARCHIVO,
        default="IMAGEN",
        verbose_name="Tipo de Soporte",
    )
    tamano_bytes = models.BigIntegerField(
        default=0,
        verbose_name="Tamaño en Bytes",
    )
    titulo_evidencia = models.CharField(
        max_length=300,
        blank=True,
        default="",
        verbose_name="Título Técnico de la Evidencia",
    )
    descripcion_tecnica_ia = models.TextField(
        blank=True,
        default="",
        verbose_name="Descripción Técnica Formal (Para Obligacion K.docx)",
    )
    # Metadatos ultraligeros extraídos sin tocar filas de datos (ej. nombres de hojas y fila 1 de Excel)
    metadatos_ligeros = models.JSONField(
        default=dict,
        blank=True,
        help_text="Si es Excel gigante: {'hojas': [...], 'encabezados_fila_1': {...}}. Jamás guarda filas crudas.",
        verbose_name="Metadatos Ultraligeros (0 Tokens Crudos)",
    )
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["orden", "id"]
        verbose_name = "Archivo de Evidencia / Soporte"
        verbose_name_plural = "Archivos de Evidencia / Soportes"

    def __str__(self):
        return f"[Ob.{self.detalle_obligacion.obligacion.numero_obligacion}] {self.nombre_original} ({self.tipo_archivo})"

    @property
    def es_imagen(self) -> bool:
        ext = os.path.splitext(self.nombre_original.lower())[1]
        return self.tipo_archivo == "IMAGEN" or ext in (".png", ".jpg", ".jpeg", ".bmp", ".webp")

    def resumen_para_prompt_ia(self) -> str:
        """
        Representación hiper-compacta de alta densidad semántica (Máx ~65 tokens por archivo).
        Elimina ruido de interfaz de Windows/Teams, membretes repetitivos y exceso de columnas Excel.
        """
        partes = [f"{self.nombre_original} ({self.tipo_archivo})"]
        meta = self.metadatos_ligeros or {}

        tag = meta.get("tag_semantico")
        if tag:
            partes.append(f"Tag:{tag}")

        hojas = meta.get("hojas") or []
        cols = meta.get("encabezados_fila_1") or {}
        if hojas:
            # Solo primeras 3 hojas y primeras 6 columnas de las 2 hojas principales
            cols_compactas = {h: c[:6] for h, c in list(cols.items())[:2] if c}
            partes.append(f"Hojas:{hojas[:3]} Cols:{cols_compactas}")

        lineas = meta.get("resumen_lineas") or []
        if lineas:
            # Filtrar líneas basura de membretes o interfaz de Windows/Teams
            ruido = ("unidad administrativa", "restitución de tierras", "código:", "versión:", "página", "buscar", "minimizar", "http")
            lineas_utiles = [
                ln[:80] for ln in lineas
                if len(ln.strip()) > 4 and not any(r in ln.lower() for r in ruido)
            ][:4]
            if lineas_utiles:
                partes.append("Extracto:" + " | ".join(lineas_utiles))

        if self.titulo_evidencia and not self.titulo_evidencia.startswith("Soporte Técnico Obligación"):
            partes.append(f"Nota:{self.titulo_evidencia[:90]}")

        return " ".join(partes)[:280]


class CodigoRecarga(models.Model):
    """
    Cupones y códigos de activación prepagados (Nequi, Daviplata, Bre-B o convenios).
    Permite a los contratistas recargar pases de cuentas de cobro al instante.
    """
    codigo = models.CharField(
        max_length=50,
        unique=True,
        verbose_name="Código de Activación (PIN de Recarga)",
    )
    pases = models.PositiveSmallIntegerField(
        default=1,
        verbose_name="Pases de Cuentas de Cobro que Otorga",
    )
    canjeado = models.BooleanField(
        default=False,
        verbose_name="¿Ya fue canjeado?",
    )
    canjeado_por = models.ForeignKey(
        PerfilContratista,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="codigos_canjeados",
        verbose_name="Canjeado por",
    )
    fecha_canje = models.DateTimeField(
        null=True,
        blank=True,
        verbose_name="Fecha de Canje",
    )
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-creado_en"]
        verbose_name = "Código de Recarga / Cupón"
        verbose_name_plural = "Códigos de Recarga / Cupones"

    def __str__(self):
        estado = "CANJEADO" if self.canjeado else "DISPONIBLE"
        return f"{self.codigo} (+{self.pases} pases) - [{estado}]"
