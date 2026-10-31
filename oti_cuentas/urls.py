from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static
from generador import views

urlpatterns = [
    path("admin/diagnostico-ia/", views.diagnostico_ia_admin_view, name="admin_diagnostico_ia"),
    path("admin/", admin.site.urls),
    path("", include("generador.urls")),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
