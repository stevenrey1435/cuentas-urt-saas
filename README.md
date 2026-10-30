# OTI-PayAgent 🏛️⚡
> **Generador Automatizado de Cuentas de Cobro, Redacción Técnica con IA y Estructuración de Árbol NAS**  
> *Oficina de Tecnologías de la Información (OTI) — Unidad de Restitución de Tierras (URT)*  
> *Conforme a la Circular 0009 de 2024 (Aseguramiento de Activos de Información) y formatos oficiales GC-FO-15 / GF-FO-05.*

---

## 📌 Visión General

**OTI-PayAgent** es una plataforma web desarrollada en **Django 5.2** que resuelve de forma integral y automatizada el flujo mensual de legalización de cuentas de cobro para contratistas de la OTI - URT.

El sistema transforma apuntes crudos de desarrollo (`git log`, tickets GLPI, actas de Teams) y archivos de evidencia dispersos (`.png`, `.jpg`, `.pdf`, `.xlsx`, `.docx`) en un **paquete mensual completo y listo para firma y radicación en SIGAF y SECOP II**, ensamblando:
1. **Informe Mensual Principal GC-FO-15 v6:** Con redacción técnica densa y profunda (250 a 450 palabras por obligación), sanitizado al 100% contra caídas del motor ReportLab.
2. **Certificación de Ruta NAS (`Evidencia_Pago_XX.docx`):** Conforme a la Circular 0009 de 2024.
3. **Árbol de Carpetas NAS (`2_CARPETA_NAS_EVIDENCIAS/Pago XX/`):**
   - Subcarpetas `Actividad_1` a `Actividad_N` para cada obligación contractual.
   - `REPORTE DE OBLIGACIONES CONTRACTUALES OTI.docx` por cada actividad (con descripción detallada, productos/entregables verificables y desglose de anexos).
   - Subcarpeta `Anexos/` con el compendio `Obligacion K.docx` (imágenes incrustadas con títulos técnicos y párrafos de contexto procesal).

---

## 🚀 Capacidades y Arquitectura de Ingeniería

### 1. 🔍 Motor OCR Local "0-Tokens" (`rapidocr-onnxruntime`)
* **Extracción local en CPU:** Analiza capturas de pantalla, Merge Requests, reportes de Teams y consolas en milisegundos sin consumir tokens externos ni exponer datos confidenciales a la nube.
* **Etiquetado Semántico Automático:** Identifica familias de contenido (`SOFTWARE_GIT`, `REUNION_SEGUIMIENTO`, `TABLERO_AGIL`, `CRONOGRAMA_PLANEACION`, `ARQUITECTURA_DISENO`, `SOPORTE_INCIDENTES`).

### 2. 🧠 Clasificador Semántico Inteligente (Anti-Contaminación)
* **Acomodación Automática:** Los contratistas pueden subir un archivo comprimido `.zip` o seleccionar 20 evidencias juntas sin necesidad de clasificarlas a mano; el sistema analiza las cláusulas contractuales y acomoda cada evidencia en su respectiva actividad.
* **Aislamiento Estricto de Reuniones:** Bloquea que evidencias de reuniones o asistencias se mezclen erróneamente en obligaciones de desarrollo de software o arquitectura.

### 3. 🤖 Conector IA Dual con Cola FIFO (Single-Shot JSON)
* **Hermes Agent (`gpt-6.1-sol`):** Conexión directa Python mediante transporte OAuth Codex (`https://chatgpt.com/backend-api/codex`) protegido con semáforo FIFO (`1 informe a la vez`).
* **Fallback Automático en Cascada:**
  1. Google Gemini API (`gemini-2.5-flash`).
  2. OpenAI API (`gpt-4o-mini`).
  3. **Modo Demo/Offline Inteligente OTI-URT:** Motor local estructurado que garantiza operatividad 100% incluso sin conexión externa.

### 4. 🛡️ Blindaje ReportLab SIGAF (`GC-FO-15` & `GF-FO-05`)
* **Sanitización Transversal de XML/HTML:** Neutraliza errores fatales del analizador `paraparser` de ReportLab (`ValueError: paraparser: syntax error: parse ended with 1 unclosed tags para`) generados por caracteres `< > &` o etiquetas accidentales.

### 5. ⏱️ Experiencia de Usuario y Métricas en Tiempo Real
* **Modal de Progreso con Cronómetro:** Indicador visual paso a paso durante el desempaquetado, OCR, clasificación, redacción IA y ensamble de archivos `.docx`.
* **Contador de Palabras Dinámico:** Supervisa en vivo que cada obligación cumpla con los estándares de extensión y rigor de auditoría.

---

## 🛠️ Estructura del Proyecto

```text
Proyecto_Agente_Cuentas_OTI/
├── generador/                  # Aplicación principal de Django
│   ├── models.py               # Modelos de PerfilContratista, Obligacion, Informe, ArchivoEvidencia
│   ├── views.py                # Controladores del flujo de carga, redacción y empaquetado
│   ├── services/
│   │   ├── ai_writer.py        # Conector IA (Hermes Agent, Gemini, OpenAI, Motor Offline)
│   │   ├── file_inspector.py   # OCR local (RapidOCR), lectura de PDFs, Word y metadatos Excel
│   │   ├── package_builder.py  # Ensamblador de documentos Word (.docx) y árbol de carpetas
│   │   └── template_converter.py # Gestor de plantillas base oficiales URT
│   └── templates/generador/    # Vistas HTML con TailwindCSS
├── oti_cuentas/                # Configuración global del proyecto Django (settings, urls)
├── plantillas_base/            # Moldes oficiales (.docx) suministrados por la OTI
├── seed_piloto_oti.py          # Script de precarga de contratos y obligaciones piloto
├── requirements.txt            # Dependencias del proyecto
└── manage.py
```

---

## ⚙️ Instalación y Puesta en Marcha

### Prerrequisitos
* Python 3.11 o superior.
* Git.

### 1. Clonar el repositorio y crear entorno virtual
```bash
git clone <URL_DEL_REPOSITORIO>
cd Proyecto_Agente_Cuentas_OTI

# Crear y activar entorno virtual
python -m venv .venv

# En Windows PowerShell:
.\.venv\Scripts\Activate.ps1

# En Linux / macOS:
source .venv/bin/activate
```

### 2. Instalar dependencias
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

### 3. Configuración de Variables de Entorno (Opcional)
Copia el archivo `.env.example` como `.env`:
```bash
cp .env.example .env
```
Configura tus llaves de API si deseas respaldos externos:
```env
GEMINI_API_KEY=tu_gemini_api_key
OPENAI_API_KEY=tu_openai_api_key
```
*(Nota: Si usas Hermes Agent localmente, el sistema detecta de forma automática la sesión OAuth en `%LOCALAPPDATA%\hermes\auth.json`).*

### 4. Ejecutar Migraciones y Precargar Datos
```bash
python manage.py migrate
python seed_piloto_oti.py
```

### 5. Iniciar el Servidor de Desarrollo
```bash
python manage.py runserver 127.0.0.1:8000
```
Abre tu navegador en: [http://127.0.0.1:8000/](http://127.0.0.1:8000/)

---

## 🔒 Buenas Prácticas de Git (Evitar Conflictos entre GitLab y GitHub)

Si tu equipo ya tiene configuradas credenciales globales para el **GitLab institucional de la URT** (`git.urt.gov.co`), puedes subir este proyecto a **GitHub** de forma completamente aislada sin alterar tu configuración global:

```bash
# 1. Inicializar git en la raíz de este proyecto (si no está inicializado)
git init

# 2. Configurar usuario y correo LOCALES solo para este repositorio
git config user.name "TuNombreEnGitHub"
git config user.email "tu_correo@github.com"

# 3. Agregar los archivos respetando el .gitignore
git add .
git commit -m "feat: inicializacion de OTI-PayAgent con motor OCR local y conector Hermes Agent"

# 4. Vincular el repositorio remoto de GitHub
git branch -M main
git remote add origin https://github.com/TU_USUARIO_GITHUB/NOMBRE_DEL_REPO.git

# 5. Enviar a GitHub
git push -u origin main
```

---

## 📄 Licencia y Marco Institucional
Desarrollado para la **Oficina de Tecnologías de la Información (OTI)** de la **Unidad de Restitución de Tierras (URT)** bajo los lineamientos de Gobierno Digital y la Circular 0009 de 2024.
