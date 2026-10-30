"""
Servicios del motor OTI-PayAgent:
- template_converter.py: Convierte las plantillas originales de plantillas_base/ en plantillas dinámicas docxtpl / python-docx.
- file_inspector.py: Filtro inteligente de archivos pesados (0 tokens para Excels gigantes / PDFs) y sanitizador ReportLab GC-FO-15.
- ai_writer.py: Conector IA Dual (Hermes Agent OAuth FIFO -> API Key Fallback -> Modo Demo/Offline Inteligente) con llamada única Single-Shot JSON.
- package_builder.py: Motor constructor del árbol completo de carpetas y los 2N+2 documentos Word (.docx) + empaquetado .zip.
"""
