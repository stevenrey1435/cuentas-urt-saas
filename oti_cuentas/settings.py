"""
Django settings for oti_cuentas project (OTI-PayAgent).
Plataforma automatizada de generación de Cuentas de Cobro y Árbol de Evidencias NAS para Contratistas Contratistas.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")

SECRET_KEY = os.environ.get(
    "DJANGO_SECRET_KEY",
    "django-insecure-oti-payagent-app-2026-clave-local-segura-998877"
)

DEBUG = os.environ.get("DJANGO_DEBUG", "True").lower() in ("true", "1", "yes")

ALLOWED_HOSTS = ["*"]

# Soporte para orígenes HTTPS en Railway / GCP
CSRF_TRUSTED_ORIGINS = [
    origin.strip()
    for origin in os.environ.get("CSRF_TRUSTED_ORIGINS", "http://localhost:8000,http://127.0.0.1:8000").split(",")
    if origin.strip()
]

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "generador",
]

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "whitenoise.middleware.WhiteNoiseMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

ROOT_URLCONF = "oti_cuentas.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [BASE_DIR / "templates"],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.debug",
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]

WSGI_APPLICATION = "oti_cuentas.wsgi.application"

# Si existe variable DATA_DIR (en Docker/Cloud), guarda ahí la BD y los archivos subidos para que nunca se borren
DATA_DIR = Path(os.environ.get("DATA_DIR", BASE_DIR))

DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.sqlite3",
        "NAME": DATA_DIR / "db.sqlite3",
    }
}

AUTH_PASSWORD_VALIDATORS = [
    {"NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator"},
    {"NAME": "django.contrib.auth.password_validation.MinimumLengthValidator"},
]

LANGUAGE_CODE = "es-co"
TIME_ZONE = "America/Bogota"
USE_I18N = True
USE_TZ = True

STATIC_URL = "static/"
STATICFILES_DIRS = [BASE_DIR / "static"] if (BASE_DIR / "static").exists() else []
STATIC_ROOT = BASE_DIR / "staticfiles"

MEDIA_URL = "media/"
MEDIA_ROOT = DATA_DIR / "media"

# Directorios específicos de OTI-PayAgent
PLANTILLAS_BASE_DIR = BASE_DIR / "plantillas_base"
PLANTILLAS_DINAMICAS_DIR = BASE_DIR / "plantillas_dinamicas"
SALIDA_PAQUETES_DIR = MEDIA_ROOT / "paquetes_generados"

# Límites de negocio y ahorro de tokens
MAX_CARACTERES_APUNTES_OBLIGACION = 2500
CREDITOS_MENSUALES_INFORME_COMPLETO = 3
CREDITOS_MENSUALES_REGENERACION_INDIVIDUAL = 6

# Configuración del Conector IA Dual (Hermes Agent Primary -> API Key Fallback -> Smart Demo Offline)
HERMES_PROVIDER = os.environ.get("HERMES_PROVIDER", "openai-codex")
HERMES_MODEL = os.environ.get("HERMES_MODEL", "gpt-6.1-sol")
HERMES_REASONING = os.environ.get("HERMES_REASONING", "low")
HERMES_BASE_URL = os.environ.get("HERMES_BASE_URL", "")
HERMES_API_KEY = os.environ.get("HERMES_API_KEY", "hermes-oauth-codex")

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")

OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

# Permitir subida de archivos de soporte grandes sin cargar en memoria para IA
DATA_UPLOAD_MAX_MEMORY_SIZE = 52428800  # 50 MB
FILE_UPLOAD_MAX_MEMORY_SIZE = 10485760  # 10 MB

# Configuración detrás del Proxy HTTPS de Render / Cloudflare
SECURE_PROXY_SSL_HEADER = ("HTTP_X_FORWARDED_PROTO", "https")
USE_X_FORWARDED_HOST = True
SECURE_SSL_REDIRECT = False


import dj_database_url

# Si existe DATABASE_URL (en Render), usa PostgreSQL.
# Si estás en tu PC local sin esa variable, sigue usando tu db.sqlite3 de siempre.
DATABASES = {
    "default": dj_database_url.config(
        default=f"sqlite:///{BASE_DIR / 'db.sqlite3'}",
        conn_max_age=600,
    )
}