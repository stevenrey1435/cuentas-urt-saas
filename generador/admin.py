from django.contrib import admin
from .models import (
    PerfilContratista,
    ObligacionContractual,
    InformeMensual,
    DetalleEjecucionObligacion,
    ArchivoEvidencia,
    CodigoRecarga,
)
from .views import _generar_codigo_unico_cliente


class ObligacionInline(admin.TabularInline):
    model = ObligacionContractual
    extra = 0


@admin.register(PerfilContratista)
class PerfilContratistaAdmin(admin.ModelAdmin):
    list_display = (
        "nombres_apellidos",
        "cedula",
        "numero_contrato",
        "creditos_informe_completo",
        "codigo_1_pase_prueba",
        "codigo_3_pases",
        "codigo_6_pases",
        "activo_mes_actual",
    )
    list_editable = ("creditos_informe_completo",)
    list_filter = ("activo_mes_actual", "anio_contrato")
    search_fields = ("nombres_apellidos", "cedula", "numero_contrato")
    inlines = [ObligacionInline]

    def _sig_uso(self, obj):
        return max(0, obj.max_caracteres_apuntes - 2500) + 1

    @admin.display(description="🎁 Código 1 Pase (Prueba/Mes)")
    def codigo_1_pase_prueba(self, obj):
        return _generar_codigo_unico_cliente(obj.cedula, 1, self._sig_uso(obj))

    @admin.display(description="⚡ Código 3 Pases")
    def codigo_3_pases(self, obj):
        return _generar_codigo_unico_cliente(obj.cedula, 3, self._sig_uso(obj))

    @admin.display(description="💎 Código 6 Pases")
    def codigo_6_pases(self, obj):
        return _generar_codigo_unico_cliente(obj.cedula, 6, self._sig_uso(obj))


class DetalleEjecucionInline(admin.TabularInline):
    model = DetalleEjecucionObligacion
    extra = 0


@admin.register(InformeMensual)
class InformeMensualAdmin(admin.ModelAdmin):
    list_display = (
        "numero_pago",
        "mes",
        "anio",
        "contratista",
        "estado",
        "ultimo_motor_ia_usado",
        "actualizado_en",
    )
    list_filter = ("estado", "anio", "mes")
    inlines = [DetalleEjecucionInline]


@admin.register(DetalleEjecucionObligacion)
class DetalleEjecucionObligacionAdmin(admin.ModelAdmin):
    list_display = ("informe", "obligacion", "actualizado_en")
    list_filter = ("informe",)


@admin.register(ArchivoEvidencia)
class ArchivoEvidenciaAdmin(admin.ModelAdmin):
    list_display = ("nombre_original", "tipo_archivo", "detalle_obligacion", "creado_en")
    list_filter = ("tipo_archivo",)


@admin.register(CodigoRecarga)
class CodigoRecargaAdmin(admin.ModelAdmin):
    list_display = ("codigo", "pases", "canjeado", "canjeado_por", "creado_en", "fecha_canje")
    list_filter = ("canjeado", "pases")
    search_fields = ("codigo",)
