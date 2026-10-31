"""
Vistas de la plataforma web OTI-PayAgent (Sin chat libre).
Incluye:
- Dashboard del contratista con contador de créditos mensuales (`creditos_informe_completo` y `creditos_regeneracion_individual`),
  datos del Contrato 1031, ruta NAS (Circular 0009 de 2024) y listado de pagos mensuales.
- Creación / Edición de Pago Mensual con tarjetas interactivas por cada obligación (1..N).
- Subida de evidencias con filtro inteligente 0-Tokens (extracción instantánea de hojas y fila 1 para Excels pesados).
- Regeneración individual de una sola obligación (`Regenerar solo esta obligación`).
- Generación Single-Shot JSON de todo el informe y construcción automática del árbol completo de carpetas + los 2N+2 documentos `.docx` + `.zip`.
- Descarga directa del `.zip` o de cualquiera de los 2N+2 documentos `.docx` generados.
"""

import os
import re
from datetime import date
from django.shortcuts import render, get_object_or_404, redirect
from django.http import JsonResponse, FileResponse, Http404
from django.conf import settings
from django.contrib import messages
from django.views.decorators.http import require_POST
from django.core.files import File

from .models import (
    PerfilContratista,
    ObligacionContractual,
    InformeMensual,
    DetalleEjecucionObligacion,
    ArchivoEvidencia,
    CodigoRecarga,
)
from .services.file_inspector import (
    clasificar_tipo_archivo,
    desempacar_archivos_subidos,
    extraer_metadatos_ligeros_archivo,
    sanitizar_texto_para_reporte,
)
from .services.ai_writer import redactar_informe_con_ia, clasificar_obligacion_por_afinidad
from .services.package_builder import construir_paquete_completo_pago
import hashlib
from django.core.cache import cache


def _obtener_ip_cliente(request) -> str:
    """Obtiene la IP real del usuario incluso detrás de proxies (Railway / GCP / Cloudflare)."""
    x_forwarded = request.META.get("HTTP_X_FORWARDED_FOR")
    if x_forwarded:
        return x_forwarded.split(",")[0].strip()
    return request.META.get("REMOTE_ADDR", "0.0.0.0")


def _verificar_rate_limit(clave: str, max_intentos: int, ventana_segundos: int) -> bool:
    """
    Retorna True si está bloqueado por exceder el límite de intentos en la ventana de tiempo.
    Usa la caché en memoria de Django (0 costo de base de datos).
    """
    intentos = cache.get(clave, 0)
    if intentos >= max_intentos:
        return True
    cache.set(clave, intentos + 1, timeout=ventana_segundos)
    return False


def _generar_codigo_unico_cliente(cedula: str, cantidad_pases: int, consecutivo_uso: int, prefijo: str = "CD") -> str:
    """
    Genera el código único e irrepetible atado a la cédula del contratista,
    a la cantidad de pases (1, 3, 6, 12) y a la llave secreta del servidor.
    Formato: <PREFIJO>-<PASES>P-<HASH6> (Ej: CD-1P-8F3A9C o URT-1P-8F3A9C)
    """
    from django.conf import settings

    semilla = f"{settings.SECRET_KEY}|CC:{cedula}|PASES:{cantidad_pases}|USO:{consecutivo_uso}"
    firma = hashlib.sha256(semilla.encode("utf-8")).hexdigest()[:6].upper()
    return f"{prefijo}-{cantidad_pases}P-{firma}"


def _obtener_contratista_activo(request) -> PerfilContratista:
    """
    Obtiene únicamente el perfil del contratista que ha iniciado sesión en la sesión actual
    (`contratista_activo_id`) o el asociado al usuario autenticado.
    Jamás muestra datos de otros contratistas por defecto.
    """
    cid = request.session.get("contratista_activo_id")
    if cid:
        perfil = PerfilContratista.objects.filter(id=cid).first()
        if perfil:
            return perfil
    if request.user.is_authenticated and hasattr(request.user, "perfil_contratista"):
        return request.user.perfil_contratista
    return None


def dashboard(request):
    """
    Si el contratista no ha iniciado sesión con su cédula, muestra la pantalla limpia de
    Ingreso / Registro Inicial de Contrato (sin datos precargados de ningún tercero).
    Si ya inició sesión, muestra exclusivamente SU propio contrato, obligaciones y pagos.
    """
    contratista = _obtener_contratista_activo(request)

    if not contratista:
        return render(
            request,
            "generador/sin_perfil.html",
            {"anio_actual": date.today().year, "hoy": date.today().isoformat()},
        )

    informes = contratista.informes_mensuales.all()
    obligaciones = contratista.obligaciones.all()

    context = {
        "contratista": contratista,
        "informes": informes,
        "obligaciones": obligaciones,
        "total_obligaciones": obligaciones.count(),
        "total_docs_por_pago": obligaciones.count() * 2 + 2,
        "meses_choices": InformeMensual.MESES_CHOICES,
        "hoy": date.today().isoformat(),
    }
    return render(request, "generador/dashboard.html", context)


@require_POST
def crear_informe_mensual(request):
    """Crea un nuevo Pago Mensual e inicializa automáticamente sus N tarjetas de DetalleEjecucionObligacion."""
    contratista = _obtener_contratista_activo(request)
    if not contratista:
        messages.warning(request, "Por favor inicia sesión con tu cédula y PIN para acceder a tus informes.")
        return redirect("generador:dashboard")

    numero_pago = int(request.POST.get("numero_pago", 11))
    mes = int(request.POST.get("mes", 11))
    anio = int(request.POST.get("anio", 2026))
    fecha_reporte_str = request.POST.get("fecha_reporte") or date.today().isoformat()
    informes_secop = int(request.POST.get("informes_secop_publicados", max(1, numero_pago - 2)))

    informe, created = InformeMensual.objects.get_or_create(
        contratista=contratista,
        numero_pago=numero_pago,
        anio=anio,
        defaults={
            "mes": mes,
            "fecha_reporte": date.fromisoformat(fecha_reporte_str),
            "informes_secop_publicados": informes_secop,
            "estado": "BORRADOR",
        },
    )

    for ob in contratista.obligaciones.all():
        DetalleEjecucionObligacion.objects.get_or_create(
            informe=informe,
            obligacion=ob,
            defaults={
                "apuntes_contratista": "",
                "texto_tabla_principal": "",
            },
        )

    if created:
        messages.success(
            request,
            f"Pago {informe.numero_pago:02d} ({informe.nombre_mes} {informe.anio}) creado con {contratista.obligaciones.count()} obligaciones listas.",
        )
    else:
        messages.info(
            request,
            f"El Pago {informe.numero_pago:02d} ({informe.anio}) ya existía. Abriendo editor del pago.",
        )
    return redirect("generador:detalle_informe", informe_id=informe.id)


def _desempacar_zips_pendientes_en_informe(informe: InformeMensual, detalles_lista: list) -> list:
    """
    Si en la base de datos quedó registrado un archivo `.zip` crudo como `ArchivoEvidencia`
    (por ejemplo subido antes o desde otra vista), lo abre automáticamente, ignora `Thumbs.db`
    y archivos basura del sistema, clasifica cada archivo interno en la obligación K más afín
    y elimina el registro contenedor `.zip` para que solo queden las evidencias reales.
    """
    resumen = []
    if not detalles_lista:
        return resumen

    zips_pendientes = []
    nombres_existentes = set()
    for det in detalles_lista:
        for ev in list(det.archivos_evidencia.all()):
            if ev.nombre_original.lower().endswith(".zip") and ev.archivo and os.path.exists(ev.archivo.path):
                zips_pendientes.append(ev)
            else:
                nombres_existentes.add(ev.nombre_original.strip().lower())

    for ev_zip in zips_pendientes:
        try:
            with open(ev_zip.archivo.path, "rb") as f_raw:
                items_extraidos = desempacar_archivos_subidos([f_raw])

            for idx_item, item in enumerate(items_extraidos):
                f_obj = item["file_obj"]
                nombre_orig = item["nombre_original"]
                if nombre_orig.strip().lower() in nombres_existentes:
                    continue
                nombres_existentes.add(nombre_orig.strip().lower())

                ruta_rel = item["ruta_relativa"]
                tamano = item["tamano_bytes"]
                tipo = clasificar_tipo_archivo(nombre_orig)
                metadatos = extraer_metadatos_ligeros_archivo(f_obj, nombre_orig, ruta_relativa=ruta_rel)
                if hasattr(f_obj, "seek"):
                    f_obj.seek(0)

                pista_semantica = f"{ruta_rel} {nombre_orig} {metadatos}"
                det_destino = clasificar_obligacion_por_afinidad(
                    pista_semantica, detalles_lista, indice_rotacion=idx_item
                ) or detalles_lista[0]

                orden_actual = det_destino.archivos_evidencia.count() + 1
                nuevo_ev = ArchivoEvidencia(
                    detalle_obligacion=det_destino,
                    orden=orden_actual,
                    nombre_original=nombre_orig,
                    tipo_archivo=tipo,
                    tamano_bytes=tamano,
                    titulo_evidencia=f"Soporte Técnico Obligación {det_destino.obligacion.numero_obligacion} ({nombre_orig})",
                    descripcion_tecnica_ia="",
                    metadatos_ligeros=metadatos,
                )
                nuevo_ev.archivo.save(nombre_orig, f_obj, save=True)
                resumen.append(f"'{nombre_orig}' → Obligación {det_destino.obligacion.numero_obligacion}")

            ev_zip.delete()
        except Exception as exc:
            pass

    return resumen


def detalle_informe(request, informe_id: int):
    """
    Vista guiada del Pago Mensual (Sin chat libre):
    - Tarjetas por cada obligación (1..N) con apuntes cortos (con contador de caracteres),
      subida de capturas/soportes (con filtro 0-Tokens para Excels pesados y desempaquetado automático de .ZIP),
      vista previa del texto redactado para la tabla principal y el reporte individual OTI,
      botón 'Regenerar solo esta obligación',
      botones 'Copiar Actividad',
      y botón maestro para generar y descargar el paquete `.zip` con los 2N+2 documentos `.docx`.
    """
    contratista_sesion = _obtener_contratista_activo(request)
    if not contratista_sesion:
        messages.warning(request, "Por favor inicia sesión con tu cédula y PIN para acceder a tus informes.")
        return redirect("generador:dashboard")
    informe = get_object_or_404(
        InformeMensual.objects.select_related("contratista"),
        id=informe_id,
        contratista=contratista_sesion,  # <-- Impide que un contratista vea o descargue el informe de otro por URL
    )
    if (
        informe.contratista.pases_disponibles <= 0
        and informe.estado == "BORRADOR"
        and not getattr(settings, "LOCAL_ONLY_AI", False)
    ):
        messages.warning(
            request,
            "⚡ Este informe está listo para editar. Para generar tus documentos y paquete .ZIP, primero activa un Pase de Cobro en '⚡ Recargar Pases'.",
        )
    contratista = informe.contratista

    # Asegurar que todas las obligaciones del contratista tengan su registro en este informe
    for ob in contratista.obligaciones.all():
        DetalleEjecucionObligacion.objects.get_or_create(informe=informe, obligacion=ob)

    detalles_inicial = list(
        informe.detalles_obligaciones.select_related("obligacion")
        .prefetch_related("archivos_evidencia")
        .order_by("obligacion__numero_obligacion")
    )
    # Si había un archivo .zip crudo pendiente, desempaquetarlo automáticamente en sus obligaciones
    desempacados_auto = _desempacar_zips_pendientes_en_informe(informe, detalles_inicial)
    if desempacados_auto:
        messages.info(
            request,
            f"Se desempaquetaron y acomodaron automáticamente {len(desempacados_auto)} archivos de tu .ZIP (ignorando Thumbs.db y archivos del sistema).",
        )

    detalles = (
        informe.detalles_obligaciones.select_related("obligacion")
        .prefetch_related("archivos_evidencia")
        .order_by("obligacion__numero_obligacion")
    )

    zip_disponible = bool(informe.archivo_zip_generado and os.path.exists(informe.archivo_zip_generado))
    total_obs = detalles.count()
    request.session["contratista_activo_id"] = contratista.id

    context = {
        "informe": informe,
        "contratista": contratista,
        "detalles": detalles,
        "total_obligaciones": total_obs,
        "total_documentos_esperados": total_obs * 2 + 2,
        "zip_disponible": zip_disponible,
        "meses_choices": InformeMensual.MESES_CHOICES,
    }
    return render(request, "generador/detalle_informe.html", context)


@require_POST
def guardar_datos_generales_informe(request, informe_id: int):
    """Actualiza fecha de reporte, mes, informes SECOP II, observaciones adicionales y captura del explorador NAS."""
    contratista_sesion = _obtener_contratista_activo(request)
    if not contratista_sesion:
        messages.warning(request, "Por favor inicia sesión con tu cédula y PIN para acceder a tus informes.")
        return redirect("generador:dashboard")
    informe = get_object_or_404(
        InformeMensual.objects.select_related("contratista"),
        id=informe_id,
        contratista=contratista_sesion,
    )

    # Capa de Defensa 1: Si no tiene pases disponibles y el informe nunca se ha generado
    if informe.contratista.pases_disponibles <= 0 and informe.estado == "BORRADOR":
        messages.warning(
            request,
            "⚡ No tienes Pases activos para generar esta cuenta. Haz clic en '⚡ Recargar Pases' para solicitar tu código de activación por WhatsApp.",
        )
        return redirect("generador:detalle_informe", informe_id=informe.id)
    mes = request.POST.get("mes")
    if mes:
        informe.mes = int(mes)
    fecha_str = request.POST.get("fecha_reporte")
    if fecha_str:
        informe.fecha_reporte = date.fromisoformat(fecha_str)
    secop = request.POST.get("informes_secop_publicados")
    if secop is not None and str(secop).isdigit():
        informe.informes_secop_publicados = int(secop)

    informe.observaciones_adicionales = sanitizar_texto_para_reporte(
        request.POST.get("observaciones_adicionales", "")
    )
    enlace_nube = (request.POST.get("enlace_nube_evidencias") or "").strip()
    if enlace_nube:
        informe.enlace_nube_evidencias = enlace_nube
        informe.contratista.servidor_nas_prefix = enlace_nube
        informe.contratista.save()

    if "captura_explorador_nas" in request.FILES:
        informe.captura_explorador_nas = request.FILES["captura_explorador_nas"]

    informe.save()
    messages.success(request, f"Datos del Pago {informe.numero_pago:02d} actualizados correctamente.")
    return redirect("generador:detalle_informe", informe_id=informe.id)


@require_POST
def guardar_obligacion(request, detalle_id: int):
    """
    Guarda los apuntes cortos del contratista (respetando el límite de caracteres),
    permite editar manualmente el texto de la tabla principal si lo desea, y procesa
    los archivos soporte subidos (incluyendo desempaquetado automático de .ZIP) aplicando el Filtro 0-Tokens.
    """
    contratista_sesion = _obtener_contratista_activo(request)
    if not contratista_sesion:
        messages.warning(request, "Por favor inicia sesión con tu cédula y PIN para acceder a tus informes.")
        return redirect("generador:dashboard")

    detalle = get_object_or_404(
        DetalleEjecucionObligacion.objects.select_related("informe__contratista", "obligacion"),
        id=detalle_id,
        informe__contratista=contratista_sesion,
    )
    contratista = detalle.informe.contratista
    max_chars = contratista.max_caracteres_apuntes

    apuntes = request.POST.get("apuntes_contratista", "").strip()
    if len(apuntes) > max_chars:
        apuntes = apuntes[:max_chars]
    detalle.apuntes_contratista = apuntes

    if "texto_tabla_principal" in request.POST:
        detalle.texto_tabla_principal = sanitizar_texto_para_reporte(
            request.POST.get("texto_tabla_principal", "")
        )

    detalle.save()

    # Actualizar notas/contexto de evidencias existentes si el usuario las editó
    for ev_existente in detalle.archivos_evidencia.all():
        key_nota = f"nota_evidencia_{ev_existente.id}"
        if key_nota in request.POST:
            nueva_nota = request.POST.get(key_nota, "").strip()
            if nueva_nota != ev_existente.titulo_evidencia:
                ev_existente.titulo_evidencia = nueva_nota[:300]
                ev_existente.save()

    # Procesar archivos soporte subidos en la obligación K (si viene un .zip, extrae todos sus archivos válidos)
    archivos_raw = request.FILES.getlist("archivos_soporte")
    items_desempacados = desempacar_archivos_subidos(archivos_raw)
    titulo_sugerido = request.POST.get("titulo_evidencia_nuevo", "").strip()
    
    evs_existentes_informe = {
        ev.nombre_original.strip().lower(): ev
        for ev in ArchivoEvidencia.objects.filter(detalle_obligacion__informe=detalle.informe)
    }

    for item in items_desempacados:
        f_obj = item["file_obj"]
        nombre_orig = item["nombre_original"]
        nombre_key = nombre_orig.strip().lower()
        ruta_rel = item["ruta_relativa"]
        tamano = item["tamano_bytes"]
        tipo = clasificar_tipo_archivo(nombre_orig)
        metadatos = extraer_metadatos_ligeros_archivo(f_obj, nombre_orig, ruta_relativa=ruta_rel)
        if hasattr(f_obj, "seek"):
            f_obj.seek(0)

        if nombre_key in evs_existentes_informe:
            ev_existente = evs_existentes_informe[nombre_key]
            ev_existente.detalle_obligacion = detalle
            ev_existente.metadatos_ligeros = metadatos
            if titulo_sugerido:
                ev_existente.titulo_evidencia = titulo_sugerido
            ev_existente.save()
        else:
            orden_actual = detalle.archivos_evidencia.count() + 1
            ev = ArchivoEvidencia(
                detalle_obligacion=detalle,
                orden=orden_actual,
                nombre_original=nombre_orig,
                tipo_archivo=tipo,
                tamano_bytes=tamano,
                titulo_evidencia=titulo_sugerido or f"Soporte Técnico Obligación {detalle.obligacion.numero_obligacion} ({nombre_orig})",
                descripcion_tecnica_ia="",
                metadatos_ligeros=metadatos,
            )
            ev.archivo.save(nombre_orig, f_obj, save=True)
            evs_existentes_informe[nombre_key] = ev

    if request.headers.get("x-requested-with") == "XMLHttpRequest":
        archivos_data = [
            {
                "id": ev_item.id,
                "nombre_original": ev_item.nombre_original,
                "tipo_archivo": ev_item.tipo_archivo,
                "tamano_bytes": ev_item.tamano_bytes,
            }
            for ev_item in detalle.archivos_evidencia.all().order_by("orden", "id")
        ]
        return JsonResponse({
            "ok": True,
            "mensaje": f"Obligación {detalle.obligacion.numero_obligacion} guardada ({len(items_desempacados)} archivos nuevos)." if items_desempacados else f"Obligación {detalle.obligacion.numero_obligacion} guardada.",
            "texto_tabla_principal": detalle.texto_tabla_principal,
            "detalle_id": detalle.id,
            "total_soportes": len(archivos_data),
            "archivos": archivos_data,
        })

    messages.success(
        request,
        f"Obligación {detalle.obligacion.numero_obligacion} actualizada correctamente ({len(items_desempacados)} archivos procesados).",
    )
    return redirect(f"/informes/{detalle.informe_id}/#obligacion-{detalle.obligacion.numero_obligacion}")


@require_POST
def eliminar_evidencia(request, evidencia_id: int):
    """Elimina un archivo soporte de una obligación."""
    contratista_sesion = _obtener_contratista_activo(request)
    if not contratista_sesion:
        if request.headers.get("x-requested-with") == "XMLHttpRequest":
            return JsonResponse({"ok": False, "error": "Sesión expirada"}, status=401)
        messages.warning(request, "Por favor inicia sesión con tu cédula y PIN para acceder a tus informes.")
        return redirect("generador:dashboard")

    ev = get_object_or_404(
        ArchivoEvidencia.objects.select_related("detalle_obligacion__obligacion", "detalle_obligacion__informe"),
        id=evidencia_id,
        detalle_obligacion__informe__contratista=contratista_sesion,
    )
    detalle = ev.detalle_obligacion
    informe_id = detalle.informe_id
    num_ob = detalle.obligacion.numero_obligacion
    nombre = ev.nombre_original
    try:
        if ev.archivo and os.path.exists(ev.archivo.path):
            os.remove(ev.archivo.path)
    except Exception:
        pass
    ev.delete()
    total_restantes = detalle.archivos_evidencia.count()

    if request.headers.get("x-requested-with") == "XMLHttpRequest":
        return JsonResponse({
            "ok": True,
            "mensaje": f"Soporte '{nombre}' eliminado de la Obligación {num_ob}.",
            "evidencia_id": evidencia_id,
            "detalle_id": detalle.id,
            "total_soportes": total_restantes,
        })

    messages.info(request, f"Soporte '{nombre}' eliminado de la Obligación {num_ob}.")
    return redirect(f"/informes/{informe_id}/#obligacion-{num_ob}")


@require_POST
def regenerar_obligacion_individual(request, detalle_id: int):
    """
    Botón 'Regenerar solo esta obligación':
    Guarda los apuntes y notas de commits/capturas de la obligación K, consume 1 crédito de `creditos_regeneracion_individual`
    y ejecuta 1 llamada pequeña a `services/ai_writer.py` para redactar exclusivamente esa obligación.
    """
    contratista_sesion = _obtener_contratista_activo(request)
    if not contratista_sesion:
        messages.warning(request, "Por favor inicia sesión con tu cédula y PIN para acceder a tus informes.")
        return redirect("generador:dashboard")

    detalle = get_object_or_404(
        DetalleEjecucionObligacion.objects.select_related("informe__contratista", "obligacion"),
        id=detalle_id,
        informe__contratista=contratista_sesion,
    )
    informe = detalle.informe
    contratista = informe.contratista

    if "apuntes_contratista" in request.POST:
        detalle.apuntes_contratista = request.POST.get("apuntes_contratista", "")[: contratista.max_caracteres_apuntes]
        detalle.save()

    for ev_existente in detalle.archivos_evidencia.all():
        key_nota = f"nota_evidencia_{ev_existente.id}"
        if key_nota in request.POST:
            nueva_nota = request.POST.get(key_nota, "").strip()
            if nueva_nota != ev_existente.titulo_evidencia:
                ev_existente.titulo_evidencia = nueva_nota[:300]
                ev_existente.save()

    try:
        res_ia = redactar_informe_con_ia(informe, obligacion_especifica_id=detalle.obligacion_id)
        construir_paquete_completo_pago(informe)
        msg = (
            f"Obligación {detalle.obligacion.numero_obligacion} regenerada con éxito ({res_ia['motor_usado']}) "
            f"y documentos Word actualizados. Créditos individuales restantes: {res_ia['creditos_regeneracion_individual']}."
        )
        if request.headers.get("x-requested-with") == "XMLHttpRequest":
            detalle.refresh_from_db()
            return JsonResponse({
                "ok": True,
                "mensaje": msg,
                "texto_tabla_principal": detalle.texto_tabla_principal,
                "creditos_regeneracion_individual": res_ia["creditos_regeneracion_individual"],
            })
        messages.success(request, msg)
    except Exception as exc:
        if request.headers.get("x-requested-with") == "XMLHttpRequest":
            return JsonResponse({"ok": False, "error": str(exc)}, status=400)
        messages.error(request, f"No se pudo regenerar la obligación: {exc}")

    return redirect(f"/informes/{informe.id}/#obligacion-{detalle.obligacion.numero_obligacion}")


@require_POST
def generar_paquete_completo_view(request, informe_id: int):
    """
    Recibe opcionalmente la carga masiva de archivos (`archivos_masivos`, incluyendo cualquier `.zip` con
    imágenes, PDFs, Word, Excels, `.py`, `.sql`, etc.) y apuntes/commits generales del mes (`apuntes_generales_mes`),
    desempaqueta automáticamente los `.zip` ignorando `Thumbs.db` y archivos basura, acomoda cada archivo
    y apunte en la obligación K (1..N) más afín según la cláusula contractual, y si
    `modo_accion == 'redactar_ia_y_empaquetar'`, ejecuta la llamada Single-Shot al Agente IA para redactar las N
    obligaciones y construir el árbol de carpetas + los 2N+2 documentos `.docx` + `.zip`.
    """
    contratista_sesion = _obtener_contratista_activo(request)
    if not contratista_sesion:
        messages.warning(request, "Por favor inicia sesión con tu cédula y PIN para acceder a tus informes.")
        return redirect("generador:dashboard")

    informe = get_object_or_404(
        InformeMensual.objects.select_related("contratista"),
        id=informe_id,
        contratista=contratista_sesion,
    )

    # Capa de Defensa 1: Si no tiene pases disponibles y el informe nunca se ha generado
    if informe.contratista.pases_disponibles <= 0 and informe.estado == "BORRADOR":
        messages.warning(
            request,
            "⚡ No tienes Pases activos para generar esta cuenta. Haz clic en '⚡ Recargar Pases' para solicitar tu código de activación por WhatsApp.",
        )
        return redirect("generador:detalle_informe", informe_id=informe.id)
    modo = request.POST.get("modo_accion", "ensamblar_y_redactar")
    apuntes_generales = (request.POST.get("apuntes_generales_mes") or "").strip()
    nota_masiva = (request.POST.get("nota_evidencias_masivas") or "").strip()
    archivos_masivos_raw = request.FILES.getlist("archivos_masivos")

    detalles_lista = list(
        informe.detalles_obligaciones.select_related("obligacion")
        .prefetch_related("archivos_evidencia")
        .order_by("obligacion__numero_obligacion")
    )

    try:
        resumen_acomodo = []

        # La carga de embeddings puede consumir cientos de MB y descargar el modelo
        # durante la primera petición. En Render se mantiene desactivada salvo que
        # LOCAL_SEMANTIC_ENABLED=True se configure explícitamente.
        if not getattr(settings, "LOCAL_SEMANTIC_ENABLED", False):
            from .services import semantic_matcher
            semantic_matcher._obtener_modelo.cache_clear()

        # 0. Si había algún .zip crudo guardado previamente en la BD, desempaquetarlo primero
        resumen_acomodo.extend(_desempacar_zips_pendientes_en_informe(informe, detalles_lista))

        # 1. Si el contratista pegó apuntes/commits generales del mes, distribuir cada línea en su obligación más afín
        if apuntes_generales and detalles_lista:
            lineas_apuntes = [ln.strip() for ln in apuntes_generales.splitlines() if ln.strip()]
            for idx_ln, linea in enumerate(lineas_apuntes):
                det_destino = clasificar_obligacion_por_afinidad(linea, detalles_lista, indice_rotacion=idx_ln)
                if det_destino:
                    actual = (det_destino.apuntes_contratista or "").strip()
                    if linea not in actual:
                        nuevo_txt = f"{actual}\n{linea}".strip() if actual else linea
                        det_destino.apuntes_contratista = nuevo_txt[: informe.contratista.max_caracteres_apuntes]
                        det_destino.save()

        # 2. Si el contratista subió evidencias en bloque o un archivo .ZIP, desempaquetar, filtrar Thumbs.db y clasificar
        items_desempacados = desempacar_archivos_subidos(archivos_masivos_raw)
        
        # Mapa de evidencias existentes en este informe para deduplicación estricta (cero duplicados)
        mapa_existentes = {
            ev.nombre_original.strip().lower(): ev
            for det in detalles_lista
            for ev in det.archivos_evidencia.all()
        }

        for idx_f, item in enumerate(items_desempacados):
            f_obj = item["file_obj"]
            nombre_orig = item["nombre_original"]
            nombre_key = nombre_orig.strip().lower()
            ruta_rel = item["ruta_relativa"]
            tamano = item["tamano_bytes"]
            tipo = clasificar_tipo_archivo(nombre_orig)
            metadatos = extraer_metadatos_ligeros_archivo(f_obj, nombre_orig, ruta_relativa=ruta_rel)
            if hasattr(f_obj, "seek"):
                f_obj.seek(0)

            pista_semantica = f"{ruta_rel} {nombre_orig} {nota_masiva} {metadatos}"
            det_destino = clasificar_obligacion_por_afinidad(pista_semantica, detalles_lista, indice_rotacion=idx_f)
            if not det_destino and detalles_lista:
                det_destino = detalles_lista[0]

            if det_destino:
                if nombre_key in mapa_existentes:
                    # El archivo ya existe: actualizar su obligación destino y metadatos sin duplicar
                    ev_existente = mapa_existentes[nombre_key]
                    if ev_existente.detalle_obligacion_id != det_destino.id:
                        ev_existente.detalle_obligacion = det_destino
                    ev_existente.metadatos_ligeros = metadatos
                    if nota_masiva:
                        ev_existente.titulo_evidencia = nota_masiva
                    ev_existente.save()
                    resumen_acomodo.append(f"'{nombre_orig}' (actualizado) → Obligación {det_destino.obligacion.numero_obligacion}")
                else:
                    orden_actual = det_destino.archivos_evidencia.count() + 1
                    ev = ArchivoEvidencia(
                        detalle_obligacion=det_destino,
                        orden=orden_actual,
                        nombre_original=nombre_orig,
                        tipo_archivo=tipo,
                        tamano_bytes=tamano,
                        titulo_evidencia=nota_masiva or f"Soporte Técnico Obligación {det_destino.obligacion.numero_obligacion} ({nombre_orig})",
                        descripcion_tecnica_ia="",
                        metadatos_ligeros=metadatos,
                    )
                    ev.archivo.save(nombre_orig, f_obj, save=True)
                    mapa_existentes[nombre_key] = ev
                    resumen_acomodo.append(f"'{nombre_orig}' → Obligación {det_destino.obligacion.numero_obligacion}")

        # Validar que las obligaciones no sean textos de prueba vacíos ('primer', 'segundo', etc.)
        obligaciones_invalidas = [
            det.obligacion.numero_obligacion
            for det in detalles_lista
            if len((det.obligacion.descripcion_clausula or "").strip()) < 20
        ]
        if obligaciones_invalidas and modo == "redactar_ia_y_empaquetar":
            messages.warning(
                request,
                f"⚠️ Atención: Las obligaciones {obligaciones_invalidas} tienen textos demasiado cortos o de prueba (ej. 'primer', 'segundo') que no corresponden a cláusulas reales. Actualízalas en tu perfil para que la IA valide la afinidad real con tus evidencias.",
            )

        if modo == "solo_acomodar_masivo":
            resumen_corto = ", ".join(resumen_acomodo[:8])
            if len(resumen_acomodo) > 8:
                resumen_corto += f" y {len(resumen_acomodo) - 8} más"
            detalle_msg = f"{len(resumen_acomodo)} archivos clasificados ({resumen_corto})" if resumen_acomodo else "Apuntes distribuidos en las obligaciones."
            messages.success(
                request,
                f"¡Evidencias desempaquetadas y acomodadas automáticamente en tus obligaciones! {detalle_msg}. "
                f"Revisa abajo cómo quedaron o haz clic en 'Redactar Todo con IA + Generar ZIP'.",
            )
        elif modo == "redactar_ia_y_empaquetar":
            res_ia = redactar_informe_con_ia(
                informe,
                obligacion_especifica_id=None,
                apuntes_generales_mes="",  # <-- Ya fueron clasificados y guardados en cada obligación arriba (Ahorra ~1.000 tokens)
            )
            res_pkg = construir_paquete_completo_pago(informe)
            extra_acomodo = f" ({len(resumen_acomodo)} archivos extraídos y clasificados automáticamente)." if resumen_acomodo else ""
            messages.success(
                request,
                f"¡Paquete completo generado! El Agente acomodó las evidencias y redactó las {res_pkg['total_obligaciones']} obligaciones en 1 sola pasada ({res_ia['motor_usado']}), "
                f"creando los {res_pkg['total_documentos_word']} documentos Word y las carpetas Actividad_1..{res_pkg['total_obligaciones']}.{extra_acomodo} "
                f"Créditos restantes: {res_ia['creditos_informe_completo']}.",
            )
        else:
            # Solo ensamblar el árbol de carpetas y los 2N+2 documentos Word con los textos actuales (0 créditos IA)
            res_pkg = construir_paquete_completo_pago(informe)
            messages.success(
                request,
                f"¡Árbol de carpetas y {res_pkg['total_documentos_word']} documentos Word ensamblados y comprimidos en .ZIP exitosamente (0 créditos consumidos)!",
            )
    except Exception as exc:
        messages.error(request, f"Error al construir el paquete del Pago {informe.numero_pago:02d}: {exc}")

    return redirect("generador:detalle_informe", informe_id=informe.id)


def descargar_zip_paquete(request, informe_id: int):
    """Descarga el archivo `.zip` con todo el árbol `1_DOCUMENTOS_CUENTA_DE_COBRO` y `2_CARPETA_NAS_EVIDENCIAS`."""
    contratista_sesion = _obtener_contratista_activo(request)
    if not contratista_sesion:
        messages.warning(request, "Por favor inicia sesión con tu cédula y PIN para acceder a tus informes.")
        return redirect("generador:dashboard")
    informe = get_object_or_404(
        InformeMensual.objects.select_related("contratista"),
        id=informe_id,
        contratista=contratista_sesion,  # <-- Impide que un contratista vea o descargue el informe de otro por URL
    )
    if informe.contratista.pases_disponibles <= 0 and informe.estado == "BORRADOR":
        messages.warning(
            request,
            "⚡ Para descargar tus documentos y paquete .ZIP, primero activa un Pase de Cobro en '⚡ Recargar Pases'.",
        )
        return redirect("generador:detalle_informe", informe_id=informe.id)
    if not informe.archivo_zip_generado or not os.path.exists(informe.archivo_zip_generado):
        construir_paquete_completo_pago(informe)

    if not os.path.exists(informe.archivo_zip_generado):
        raise Http404("El archivo ZIP no pudo ser localizado.")

    nombre_descarga = os.path.basename(informe.archivo_zip_generado)
    return FileResponse(
        open(informe.archivo_zip_generado, "rb"),
        as_attachment=True,
        filename=nombre_descarga,
    )


def descargar_documento_individual(request, informe_id: int, tipo_doc: str, numero_obligacion: int = 0):
    """
    Permite descargar individualmente cualquiera de los 2N+2 documentos Word generados:
    - tipo_doc == 'informe_principal' -> CTO_<NUM>_<NOMBRE>_INFORME_<XX>.docx
    - tipo_doc == 'evidencia_pago' -> Evidencia_Pago_<XX>.docx
    - tipo_doc == 'reporte_oti' -> Pago <XX>/Actividad_<K>/REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx
    - tipo_doc == 'anexo_obligacion' -> Pago <XX>/Actividad_<K>/Anexos/Obligacion <K>.docx
    """
    contratista_sesion = _obtener_contratista_activo(request)
    if not contratista_sesion:
        messages.warning(request, "Por favor inicia sesión con tu cédula y PIN para acceder a tus informes.")
        return redirect("generador:dashboard")
    informe = get_object_or_404(
        InformeMensual.objects.select_related("contratista"),
        id=informe_id,
        contratista=contratista_sesion,  # <-- Impide que un contratista vea o descargue el informe de otro por URL
    )
    if informe.contratista.pases_disponibles <= 0 and informe.estado == "BORRADOR":
        messages.warning(
            request,
            "⚡ Para descargar tus documentos y paquete .ZIP, primero activa un Pase de Cobro en '⚡ Recargar Pases'.",
        )
        return redirect("generador:detalle_informe", informe_id=informe.id)
    if not informe.ruta_carpeta_paquete or not os.path.exists(informe.ruta_carpeta_paquete):
        construir_paquete_completo_pago(informe)

    raiz = informe.ruta_carpeta_paquete
    num_pago = informe.numero_pago

    if tipo_doc == "informe_principal":
        ruta = os.path.join(raiz, "1_DOCUMENTOS_CUENTA_DE_COBRO", informe.nombre_archivo_informe_principal)
        nombre_out = informe.nombre_archivo_informe_principal
    elif tipo_doc == "evidencia_pago":
        ruta = os.path.join(raiz, "1_DOCUMENTOS_CUENTA_DE_COBRO", informe.nombre_archivo_evidencia_pago)
        nombre_out = informe.nombre_archivo_evidencia_pago
    elif tipo_doc == "reporte_oti":
        ruta = os.path.join(
            raiz,
            "2_CARPETA_NAS_EVIDENCIAS",
            f"Pago {num_pago:02d}",
            f"Actividad_{numero_obligacion}",
            "REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx",
        )
        nombre_out = f"Actividad_{numero_obligacion}_REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx"
    elif tipo_doc == "anexo_obligacion":
        ruta = os.path.join(
            raiz,
            "2_CARPETA_NAS_EVIDENCIAS",
            f"Pago {num_pago:02d}",
            f"Actividad_{numero_obligacion}",
            "Anexos",
            f"Obligacion {numero_obligacion}.docx",
        )
        nombre_out = f"Obligacion {numero_obligacion}.docx"
    else:
        raise Http404("Tipo de documento no válido.")

    if not os.path.exists(ruta):
        construir_paquete_completo_pago(informe)

    if not os.path.exists(ruta):
        raise Http404("Archivo no encontrado en el paquete.")

    return FileResponse(open(ruta, "rb"), as_attachment=True, filename=nombre_out)


@require_POST
def recargar_creditos_demo(request):
    """Permite restablecer a 3 y 6 los créditos mensuales en el entorno local de pruebas."""
    contratista = _obtener_contratista_activo(request)
    if contratista:
        contratista.creditos_informe_completo = 3
        contratista.creditos_regeneracion_individual = 6
        contratista.activo_mes_actual = True
        contratista.save()
        messages.success(request, "Créditos mensuales restablecidos (3 Informes Completos / 6 Regeneraciones Individuales).")
    return redirect(request.META.get("HTTP_REFERER", "generador:dashboard"))


@require_POST
def actualizar_firma_contratista(request):
    """
    Permite que cada contratista suba, actualice o elimine su propia firma digitalizada (`firma_imagen`)
    en su perfil personal, garantizando que ningún contratista use la firma de otro.
    """
    contratista = _obtener_contratista_activo(request)
    if not contratista:
        return redirect("generador:dashboard")

    if request.POST.get("accion_firma") == "eliminar":
        if contratista.firma_imagen:
            contratista.firma_imagen.delete(save=True)
        messages.info(
            request,
            "Tu firma digitalizada fue removida del perfil. Los reportes se generarán con el espacio de firma en blanco.",
        )
    elif "firma_imagen" in request.FILES:
        archivo_firma = request.FILES["firma_imagen"]
        contratista.firma_imagen = archivo_firma
        contratista.save()
        messages.success(
            request,
            f"Firma personal actualizada correctamente para {contratista.nombres_apellidos}.",
        )
    return redirect(request.META.get("HTTP_REFERER", "generador:dashboard"))


@require_POST
def cambiar_contratista(request):
    """
    Maneja el inicio de sesión privado del contratista mediante su número de Cédula + PIN.
    Protegido contra fuerza bruta: Máx 5 intentos fallidos cada 15 min.
    """
    import re
    from django.contrib.auth.models import User

    if request.POST.get("accion") == "salir":
        request.session.pop("contratista_activo_id", None)
        messages.info(request, "Sesión cerrada correctamente.")
        return redirect("generador:dashboard")

    ip = _obtener_ip_cliente(request)
    cedula_ingresada = re.sub(r"\D", "", request.POST.get("cedula") or "")
    pin_ingresado = (request.POST.get("pin_acceso") or "").strip()

    # Bloquear si esta IP o esta Cédula lleva más de 5 intentos fallidos en 15 minutos (900 seg)
    clave_bloqueo_login = f"login_fail_{ip}_{cedula_ingresada}"
    if cache.get(clave_bloqueo_login, 0) >= 5:
        messages.error(
            request,
            "🔒 Por seguridad, el acceso fue pausado temporalmente tras varios intentos fallidos. Intenta de nuevo en 15 minutos.",
        )
        return redirect("generador:dashboard")

    if not cedula_ingresada:
        messages.error(request, "Debes ingresar tu número de cédula.")
        return redirect("generador:dashboard")

    perfil = PerfilContratista.objects.select_related("user").filter(cedula=cedula_ingresada).first()
    if not perfil:
        _verificar_rate_limit(clave_bloqueo_login, 5, 900)
        messages.error(
            request,
            f"No se encontró un contrato registrado con la cédula {cedula_ingresada}. Regístralo por única vez en el formulario de al lado.",
        )
        return redirect("generador:dashboard")

    # Asegurar que el perfil tenga un User asociado para validar su PIN encriptado
    if not perfil.user:
        user_auto = User.objects.create_user(username=f"cto_{perfil.numero_contrato}_{perfil.cedula[-4:]}")
        user_auto.set_password(perfil.cedula[-4:])
        user_auto.save()
        perfil.user = user_auto
        perfil.save()

    # Si el usuario fue creado por seed sin contraseña usable, su PIN inicial son los últimos 4 dígitos de su cédula
    if not perfil.user.has_usable_password():
        if pin_ingresado == perfil.cedula[-4:]:
            perfil.user.set_password(pin_ingresado)
            perfil.user.save()
        else:
            _verificar_rate_limit(clave_bloqueo_login, 5, 900)
            messages.error(request, "PIN incorrecto. (Para cuentas iniciales de prueba usa los últimos 4 dígitos de tu cédula).")
            return redirect("generador:dashboard")
    elif not perfil.user.check_password(pin_ingresado):
        _verificar_rate_limit(clave_bloqueo_login, 5, 900)
        messages.error(request, "Cédula o PIN de acceso incorrectos.")
        return redirect("generador:dashboard")

    # Acceso exitoso: limpiar intentos fallidos
    cache.delete(clave_bloqueo_login)
    request.session["contratista_activo_id"] = perfil.id
    messages.success(
        request,
        f"Bienvenido(a), {perfil.nombres_apellidos} (Contrato N° {perfil.numero_contrato} de {perfil.anio_contrato}).",
    )
    return redirect("generador:dashboard")


@require_POST
def canjear_codigo_recarga(request):
    """
    Valida y canjea un código de activación único para la cédula del contratista activo.
    - Imposible de adivinar (protegido por Rate Limit + SHA256 con SECRET_KEY).
    - Imposible de compartir con otro compañero (atado a la cédula).
    - Imposible de usar 2 veces (el consecutivo_uso sube automáticamente tras cada canje).
    """
    contratista = _obtener_contratista_activo(request)
    if not contratista:
        return redirect("generador:dashboard")

    ip = _obtener_ip_cliente(request)
    clave_intentos = f"canje_fail_{ip}_{contratista.cedula}"
    if cache.get(clave_intentos, 0) >= 5:
        messages.error(request, "🔒 Demasiados intentos inválidos. Espera 15 minutos o contáctanos por WhatsApp.")
        return redirect(request.META.get("HTTP_REFERER", "generador:dashboard"))

    codigo_ingresado = (request.POST.get("codigo") or "").strip().upper()
    if not codigo_ingresado:
        messages.error(request, "Por favor ingresa tu código de activación.")
        return redirect(request.META.get("HTTP_REFERER", "generador:dashboard"))

    usos_previos = max(0, contratista.max_caracteres_apuntes - 2500)
    siguiente_uso = usos_previos + 1

    pases_detectados = 0
    for cant in (1, 2, 3, 6, 12):
        for pref in ("CD", "URT", "MES"):
            if codigo_ingresado == _generar_codigo_unico_cliente(contratista.cedula, cant, siguiente_uso, pref):
                pases_detectados = cant
                break
        if pases_detectados > 0:
            break

    # Si no es un código generado por algoritmo para este cliente, verificar catálogo de bienvenida / cupones
    if pases_detectados <= 0:
        catalogo_fijo = {
            "PASES-DEMO": 2, "PASES-URT-DEMO": 2,
            "MES-1": 1, "URT-MES-1": 1,
            "TRIMESTRE-3": 3, "URT-TRIMESTRE-3": 3,
            "SEMESTRE-6": 6, "URT-SEMESTRE-6": 6,
            "ANUAL-12": 12, "URT-ANUAL-12": 12,
        }
        if codigo_ingresado in catalogo_fijo:
            pases_detectados = catalogo_fijo[codigo_ingresado]
        else:
            cupon_bd = CodigoRecarga.objects.filter(codigo__iexact=codigo_ingresado, canjeado=False).first()
            if cupon_bd:
                from django.utils import timezone
                pases_detectados = cupon_bd.pases
                cupon_bd.canjeado = True
                cupon_bd.canjeado_por = contratista
                cupon_bd.fecha_canje = timezone.now()
                cupon_bd.save()

    if pases_detectados <= 0:
        _verificar_rate_limit(clave_intentos, 5, 900)
        messages.error(
            request,
            f"El código '{codigo_ingresado}' no es válido para la cédula {contratista.cedula} o ya fue utilizado anteriormente.",
        )
        return redirect(request.META.get("HTTP_REFERER", "generador:dashboard"))

    # Código válido: acreditar pases y avanzar el contador de uso
    cache.delete(clave_intentos)
    contratista.creditos_informe_completo += pases_detectados
    contratista.creditos_regeneracion_individual += (pases_detectados * 5)
    contratista.max_caracteres_apuntes = 2500 + siguiente_uso
    contratista.activo_mes_actual = True
    contratista.save()

    messages.success(
        request,
        f"🎉 ¡Código activado con éxito! Se agregaron +{pases_detectados} {'Pase' if pases_detectados == 1 else 'Pases'} a tu cuenta. "
        f"Total disponible: {contratista.pases_disponibles}.",
    )
    return redirect(request.META.get("HTTP_REFERER", "generador:dashboard"))


@require_POST
def crear_contratista(request):
    """
    Crea un nuevo contratista desde la interfaz web con 4 Capas de Defensa:
    1. Trampa Honeypot invisible contra bots.
    2. Rate limit estricto por IP (máximo 2 registros cada 24 horas).
    3. Validación de Cédula Colombiana real (6 a 10 dígitos numéricos).
    4. Cuentas nuevas nacen con 0 Pases de IA (100% Blindado contra consumo indebido).
    """
    import re
    from django.contrib.auth.models import User

    ip = _obtener_ip_cliente(request)

    # 1. Trampa Honeypot Anti-Bots (si el campo oculto viene lleno, es un bot automático)
    if request.POST.get("website_url_hp"):
        messages.error(request, "Solicitud rechazada por filtro de seguridad.")
        return redirect("generador:dashboard")

    # 2. Límite estricto por IP: Máximo 2 registros de contrato por IP cada 24 horas (86400 seg)
    clave_registro_ip = f"reg_ip_{ip}"
    if cache.get(clave_registro_ip, 0) >= 2:
        messages.error(
            request,
            "🛡️ Ya se registraron cuentas recientemente desde tu conexión. Si necesitas registrar un contrato adicional, contáctanos por WhatsApp.",
        )
        return redirect("generador:dashboard")

    nombres_apellidos = (request.POST.get("nombres_apellidos") or "").strip()
    cedula = re.sub(r"\D", "", request.POST.get("cedula") or "")
    pin_acceso = (request.POST.get("pin_acceso") or (cedula[-4:] if len(cedula) >= 4 else "1234")).strip()

    # 3. Validar formato real de Cédula Colombiana (6 a 10 dígitos, sin números repetidos falsos tipo 1111111)
    if not (6 <= len(cedula) <= 10) or len(set(cedula)) <= 2:
        messages.error(request, "Por favor ingresa un número de Cédula de Ciudadanía válido (entre 6 y 10 dígitos reales).")
        return redirect("generador:dashboard")

    if PerfilContratista.objects.filter(cedula=cedula).exists():
        messages.error(request, f"Ya existe una cuenta registrada con la cédula {cedula}. Ingresa en la pestaña 'Ingresar a mi cuenta'.")
        return redirect("generador:dashboard")

    numero_contrato = (request.POST.get("numero_contrato") or "").strip()
    anio_contrato = int(request.POST.get("anio_contrato") or date.today().year)
    fecha_inicio_str = request.POST.get("fecha_inicio") or f"{anio_contrato}-01-19"
    fecha_term_str = request.POST.get("fecha_terminacion_actual") or f"{anio_contrato}-12-30"
    objeto_contrato = (request.POST.get("objeto_contrato") or "").strip()
    nombre_supervisor = (request.POST.get("nombre_supervisor") or "").strip()
    cargo_supervisor = (request.POST.get("cargo_supervisor") or "Supervisor(a) del Contrato").strip()
    direccion_oficina = (request.POST.get("direccion_oficina") or "Oficina de Gestión Contractual").strip()
    direccion_territorial = (request.POST.get("direccion_territorial") or "Nivel Central").strip()
    proyecto_area = (request.POST.get("proyecto_area") or direccion_oficina).strip()
    obligaciones_enviadas = request.POST.getlist("obligacion")
    if not obligaciones_enviadas:
        # Compatibilidad con formularios antiguos que todavía enviaban un textarea.
        obligaciones_enviadas = [request.POST.get("obligaciones_texto") or ""]
    obligaciones = [obligacion.strip() for obligacion in obligaciones_enviadas if obligacion.strip()]

    if not nombres_apellidos or not cedula or not numero_contrato or not obligaciones:
        messages.error(request, "Tu nombre completo, cédula, número de contrato y al menos una obligación son obligatorios.")
        return redirect("generador:dashboard")

    username_base = f"cto_{numero_contrato}_{cedula[-4:]}"
    username = username_base
    idx = 1
    while User.objects.filter(username=username).exists():
        idx += 1
        username = f"{username_base}_{idx}"

    partes_nom = nombres_apellidos.title().split()
    first_name = " ".join(partes_nom[:2]) if len(partes_nom) >= 2 else nombres_apellidos.title()
    last_name = " ".join(partes_nom[2:]) if len(partes_nom) > 2 else ""

    user = User.objects.create_user(
        username=username,
        first_name=first_name,
        last_name=last_name,
        password=pin_acceso,
    )

    perfil = PerfilContratista(
        user=user,
        nombres_apellidos=nombres_apellidos.title(),
        cedula=cedula,
        numero_contrato=numero_contrato,
        anio_contrato=anio_contrato,
        fecha_inicio=date.fromisoformat(fecha_inicio_str),
        fecha_terminacion_actual=date.fromisoformat(fecha_term_str),
        objeto_contrato=objeto_contrato,
        nombre_supervisor=nombre_supervisor,
        cargo_supervisor=cargo_supervisor,
        direccion_oficina=direccion_oficina,
        direccion_territorial=direccion_territorial,
        proyecto_area=proyecto_area,
        activo_mes_actual=True,
        creditos_informe_completo=0,          # <-- NACE CON 0 PASES (100% Seguro)
        creditos_regeneracion_individual=0,   # <-- NACE CON 0 AJUSTES
    )
    perfil.autocalcular_campos_internos()
    perfil.save()

    if "firma_imagen" in request.FILES:
        perfil.firma_imagen = request.FILES["firma_imagen"]
        perfil.save()

    # Procesar cada obligación enviada en su propio campo.
    for num_ob, obligacion in enumerate(obligaciones, start=1):
        desc_limpia = re.sub(r"^(?:obligaci[oó]n\s*\d+[:.\-\s]*|\d+\s*[.):\-]\s*)", "", obligacion, flags=re.IGNORECASE).strip()
        if not desc_limpia:
            desc_limpia = obligacion
        ObligacionContractual.objects.create(
            contratista=perfil,
            numero_obligacion=num_ob,
            descripcion_clausula=desc_limpia,
        )

    # Registrar conteo por IP en caché
    _verificar_rate_limit(clave_registro_ip, 2, 86400)
    request.session["contratista_activo_id"] = perfil.id
    messages.info(
        request,
        f"¡Cuenta creada para {perfil.nombres_apellidos}! Para generar tu primera cuenta de cobro, haz clic en '⚡ Recargar Pases' y solicita tu código de activación por WhatsApp.",
    )
    return redirect("generador:dashboard")


@require_POST
def editar_contratista(request):
    """Permite editar los datos básicos del contrato y supervisor del contratista activo."""
    import re

    contratista = _obtener_contratista_activo(request)
    if not contratista:
        return redirect("generador:dashboard")

    contratista.nombres_apellidos = (request.POST.get("nombres_apellidos") or contratista.nombres_apellidos).strip()
    cedula_nueva = re.sub(r"\D", "", request.POST.get("cedula") or contratista.cedula)
    contratista.cedula = cedula_nueva

    contratista.numero_contrato = (request.POST.get("numero_contrato") or contratista.numero_contrato).strip()
    contratista.anio_contrato = int(request.POST.get("anio_contrato") or contratista.anio_contrato)
    if request.POST.get("fecha_inicio"):
        contratista.fecha_inicio = date.fromisoformat(request.POST.get("fecha_inicio"))
    if request.POST.get("fecha_terminacion_actual"):
        contratista.fecha_terminacion_actual = date.fromisoformat(request.POST.get("fecha_terminacion_actual"))
    contratista.objeto_contrato = (request.POST.get("objeto_contrato") or contratista.objeto_contrato).strip()
    contratista.nombre_supervisor = (request.POST.get("nombre_supervisor") or contratista.nombre_supervisor).strip()
    contratista.cargo_supervisor = (request.POST.get("cargo_supervisor") or contratista.cargo_supervisor).strip()
    contratista.direccion_oficina = (request.POST.get("direccion_oficina") or contratista.direccion_oficina).strip()
    contratista.direccion_territorial = (request.POST.get("direccion_territorial") or contratista.direccion_territorial).strip()
    contratista.proyecto_area = (request.POST.get("proyecto_area") or contratista.proyecto_area).strip()
    contratista.autocalcular_campos_internos()
    contratista.save()

    messages.success(
        request,
        f"Datos de tu Contrato N° {contratista.numero_contrato} actualizados correctamente.",
    )
    return redirect("generador:dashboard")


@require_POST
def gestionar_obligaciones_contratista(request):
    """
    Permite agregar, editar o eliminar obligaciones contractuales del contratista activo.
    Soporta peticiones AJAX para guardar sin recargar la pantalla ni perder lo digitado.
    """
    contratista = _obtener_contratista_activo(request)
    if not contratista:
        return redirect("generador:dashboard")

    is_ajax = (
        request.headers.get("x-requested-with") == "XMLHttpRequest"
        or "application/json" in request.headers.get("accept", "")
    )
    accion = request.POST.get("accion")
    mensaje = ""
    nueva_info = None

    if accion == "agregar":
        texto = (request.POST.get("descripcion_clausula") or "").strip()
        if texto:
            lineas = [l.strip() for l in texto.splitlines() if l.strip()]
            agregadas = 0
            for linea in lineas:
                desc_limpia = re.sub(
                    r"^(?:obligaci[oó]n\s*\d+[:.\-\s]*|\d+\s*[.):\-]\s*)",
                    "",
                    linea,
                    flags=re.IGNORECASE,
                ).strip() or linea
                siguiente_num = contratista.obligaciones.count() + 1
                nueva_ob = ObligacionContractual.objects.create(
                    contratista=contratista,
                    numero_obligacion=siguiente_num,
                    descripcion_clausula=desc_limpia,
                )
                for inf in contratista.informes_mensuales.all():
                    DetalleEjecucionObligacion.objects.get_or_create(informe=inf, obligacion=nueva_ob)
                agregadas += 1
                nueva_info = {
                    "id": nueva_ob.id,
                    "numero": nueva_ob.numero_obligacion,
                    "texto": nueva_ob.descripcion_clausula,
                }
            if agregadas == 1:
                mensaje = f"Obligación N° {nueva_info['numero']} agregada al contrato."
            else:
                mensaje = f"Se agregaron {agregadas} obligaciones al contrato."
            messages.success(request, mensaje)

    elif accion == "guardar_todas":
        # Guarda todas las obligaciones presentes en el formulario a la vez
        guardadas = 0
        for ob in contratista.obligaciones.all():
            campo = f"ob_{ob.id}"
            if campo in request.POST:
                nuevo_texto = request.POST.get(campo, "").strip()
                if nuevo_texto and nuevo_texto != ob.descripcion_clausula:
                    ob.descripcion_clausula = nuevo_texto
                    ob.save()
                    guardadas += 1
        mensaje = f"Se guardaron los cambios en {guardadas} obligaciones." if guardadas else "Todas las obligaciones están al día."
        messages.success(request, mensaje)

    elif accion == "editar":
        ob_id = request.POST.get("obligacion_id")
        texto = (request.POST.get("descripcion_clausula") or "").strip()
        if ob_id and texto:
            ob = get_object_or_404(ObligacionContractual, id=int(ob_id), contratista=contratista)
            ob.descripcion_clausula = texto
            ob.save()
            mensaje = f"Obligación N° {ob.numero_obligacion} actualizada."
            messages.success(request, mensaje)

    elif accion == "eliminar":
        ob_id = request.POST.get("obligacion_id")
        if ob_id:
            ob = get_object_or_404(ObligacionContractual, id=int(ob_id), contratista=contratista)
            num_borrado = ob.numero_obligacion
            ob.delete()
            for idx, ob_restante in enumerate(contratista.obligaciones.order_by("numero_obligacion"), start=1):
                if ob_restante.numero_obligacion != idx:
                    ob_restante.numero_obligacion = idx
                    ob_restante.save()
            mensaje = f"Obligación N° {num_borrado} eliminada y numeración reajustada."
            messages.info(request, mensaje)

    if is_ajax:
        return JsonResponse({
            "ok": True,
            "mensaje": mensaje,
            "nueva_obligacion": nueva_info,
            "total_obligaciones": contratista.obligaciones.count(),
        })

    return redirect("/?seccion=obligaciones")



from django.contrib.admin.views.decorators import staff_member_required


@staff_member_required
def diagnostico_ia_admin_view(request):
    """Permite al Superadmin probar en vivo la conexión entre Render y Ollama en su PC."""
    import time

    base_url = getattr(settings, "HERMES_BASE_URL", "").strip()
    model_name = getattr(settings, "HERMES_MODEL", "qwen2.5:7b-instruct")
    api_key = getattr(settings, "HERMES_API_KEY", "ollama-local")
    gemini_key = getattr(settings, "GEMINI_API_KEY", "")
    openai_key = getattr(settings, "OPENAI_API_KEY", "")
    local_only = getattr(settings, "LOCAL_ONLY_AI", False)

    resultado_ping = None
    if request.method == "POST" or request.GET.get("probar") == "1":
        inicio = time.time()
        url_endpoint = base_url.rstrip("/")
        if not url_endpoint.endswith("/v1"):
            url_endpoint = f"{url_endpoint}/v1"

        if not base_url:
            resultado_ping = {
                "ok": False,
                "error": "HERMES_BASE_URL no está configurada en las variables de entorno de este servidor.",
                "tiempo_segundos": 0,
            }
        else:
            try:
                from openai import OpenAI

                client = OpenAI(
                    base_url=url_endpoint,
                    api_key=api_key or "ollama-local",
                    default_headers={
                        "ngrok-skip-browser-warning": "true",
                        "User-Agent": "CuentasURT/1.0",
                    },
                    timeout=30.0,
                    max_retries=0,
                )
                resp = client.chat.completions.create(
                    model=model_name,
                    messages=[
                        {
                            "role": "system",
                            "content": "Eres un asistente técnico de diagnóstico. Responde ÚNICAMENTE con un objeto JSON.",
                        },
                        {
                            "role": "user",
                            "content": "Devuelve un objeto JSON con status='ok' y mensaje='Conexión exitosa desde Render hacia tu PC'.",
                        },
                    ],
                    response_format={"type": "json_object"},
                    temperature=0.1,
                )
                tiempo = round(time.time() - inicio, 2)
                contenido = resp.choices[0].message.content
                resultado_ping = {
                    "ok": True,
                    "tiempo_segundos": tiempo,
                    "contenido": contenido,
                    "endpoint_probado": url_endpoint,
                    "modelo_probado": model_name,
                }
            except Exception as exc:
                tiempo = round(time.time() - inicio, 2)
                resultado_ping = {
                    "ok": False,
                    "error": str(exc),
                    "tiempo_segundos": tiempo,
                    "endpoint_probado": url_endpoint,
                    "modelo_probado": model_name,
                }

    ctx = {
        "title": "Diagnóstico de Conexión IA (Render ↔ PC Local)",
        "base_url": base_url,
        "model_name": model_name,
        "api_key_masked": f"{api_key[:4]}***" if len(api_key) > 4 else ("Configurada" if api_key else "Sin configurar"),
        "gemini_configurada": bool(gemini_key),
        "openai_configurada": bool(openai_key),
        "local_only": local_only,
        "resultado_ping": resultado_ping,
    }
    return render(request, "admin/diagnostico_ia.html", ctx)
