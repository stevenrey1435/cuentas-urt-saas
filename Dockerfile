FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DATA_DIR=/app/data

WORKDIR /app

# Librerías del sistema para OpenCV Headless, Pillow y fuentes Arial/Liberation para los Word
RUN apt-get update && apt-get install -y --no-install-recommends \
    libglib2.0-0 \
    libsm6 \
    libxext6 \
    libxrender1 \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /app/data/media/paquetes_generados /app/staticfiles && \
    python manage.py collectstatic --noinput

EXPOSE 8000

# Al iniciar en producción:
# 1. Aplica migraciones en el disco persistente (/app/data/db.sqlite3)
# 2. Crea automáticamente el superusuario de /admin/ si definiste DJANGO_SUPERUSER_PASSWORD
# 3. Levanta Gunicorn con timeout de 180s para paquetes grandes
CMD ["sh", "-c", "python manage.py migrate --noinput && (python manage.py createsuperuser --noinput || true) && gunicorn oti_cuentas.wsgi:application --bind 0.0.0.0:${PORT:-8000} --workers 1 --threads 4 --timeout 180"]