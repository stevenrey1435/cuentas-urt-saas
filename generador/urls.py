from django.urls import path
from . import views

app_name = "generador"

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("informes/crear/", views.crear_informe_mensual, name="crear_informe_mensual"),
    path("informes/<int:informe_id>/", views.detalle_informe, name="detalle_informe"),
    path(
        "informes/<int:informe_id>/guardar-general/",
        views.guardar_datos_generales_informe,
        name="guardar_datos_generales_informe",
    ),
    path(
        "informes/<int:informe_id>/generar-paquete/",
        views.generar_paquete_completo_view,
        name="generar_paquete_completo",
    ),
    path(
        "informes/<int:informe_id>/descargar-zip/",
        views.descargar_zip_paquete,
        name="descargar_zip_paquete",
    ),
    path(
        "informes/<int:informe_id>/descargar-doc/<str:tipo_doc>/<int:numero_obligacion>/",
        views.descargar_documento_individual,
        name="descargar_documento_individual",
    ),
    path(
        "obligaciones/<int:detalle_id>/guardar/",
        views.guardar_obligacion,
        name="guardar_obligacion",
    ),
    path(
        "obligaciones/<int:detalle_id>/regenerar/",
        views.regenerar_obligacion_individual,
        name="regenerar_obligacion_individual",
    ),
    path(
        "evidencias/<int:evidencia_id>/eliminar/",
        views.eliminar_evidencia,
        name="eliminar_evidencia",
    ),
    path(
        "evidencias/<int:evidencia_id>/guardar-observacion/",
        views.guardar_observacion_evidencia,
        name="guardar_observacion_evidencia",
    ),
    path(
        "creditos/recargar-demo/",
        views.recargar_creditos_demo,
        name="recargar_creditos_demo",
    ),
    path(
        "pases/canjear/",
        views.canjear_codigo_recarga,
        name="canjear_codigo_recarga",
    ),
    path(
        "perfil/firma/",
        views.actualizar_firma_contratista,
        name="actualizar_firma_contratista",
    ),
    path(
        "contratistas/cambiar/",
        views.cambiar_contratista,
        name="cambiar_contratista",
    ),
    path(
        "contratistas/crear/",
        views.crear_contratista,
        name="crear_contratista",
    ),
    path(
        "contratistas/editar/",
        views.editar_contratista,
        name="editar_contratista",
    ),
    path(
        "contratistas/obligaciones/",
        views.gestionar_obligaciones_contratista,
        name="gestionar_obligaciones_contratista",
    ),
]
