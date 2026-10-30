"""Clasificación semántica local de evidencias y obligaciones.

El modelo se carga de forma perezosa para no ralentizar el arranque de Django.
Si fastembed o el modelo no están disponibles, el llamador debe conservar su
clasificador heurístico y continuar sin ocultar el motivo en los logs.
"""

import logging
from functools import lru_cache
from typing import Any, Optional

import numpy as np
from django.conf import settings

logger = logging.getLogger(__name__)


@lru_cache(maxsize=1)
def _obtener_modelo():
    if not getattr(settings, "LOCAL_SEMANTIC_ENABLED", True):
        return None
    try:
        from fastembed import TextEmbedding

        kwargs = {"model_name": getattr(settings, "LOCAL_SEMANTIC_MODEL", "intfloat/multilingual-e5-small")}
        cache_dir = getattr(settings, "LOCAL_SEMANTIC_CACHE_DIR", "")
        if cache_dir:
            kwargs["cache_dir"] = cache_dir
        return TextEmbedding(**kwargs)
    except Exception as exc:
        logger.warning("Clasificador semántico local no disponible: %s", exc)
        return None


def _normalizar_vector(vector: Any) -> np.ndarray:
    resultado = np.asarray(vector, dtype=np.float32)
    norma = np.linalg.norm(resultado)
    return resultado / norma if norma else resultado


@lru_cache(maxsize=128)
def _vectorizar_obligaciones(model_name: str, obligaciones: tuple[str, ...]) -> tuple[np.ndarray, ...]:
    modelo = _obtener_modelo()
    if modelo is None:
        return ()
    usa_prefijos_e5 = model_name.lower().startswith("intfloat/multilingual-e5")
    textos = [
        f"passage: {clausula}" if usa_prefijos_e5 else clausula
        for clausula in obligaciones
    ]
    return tuple(_normalizar_vector(vector) for vector in modelo.embed(textos))


def puntuar_obligaciones_semanticamente(
    texto_evidencia: str,
    detalles: list,
) -> Optional[dict[int, float]]:
    """Devuelve similitudes coseno por obligación o None si no puede cargar el modelo."""
    texto = (texto_evidencia or "").strip()
    if not texto or not detalles:
        return None

    modelo = _obtener_modelo()
    if modelo is None:
        return None

    obligaciones = [
        (det.obligacion.numero_obligacion, (det.obligacion.descripcion_clausula or "").strip())
        for det in detalles
    ]
    obligaciones_validas = [(numero, clausula) for numero, clausula in obligaciones if clausula]
    if not obligaciones_validas:
        return None

    try:
        model_name = getattr(
            settings,
            "LOCAL_SEMANTIC_MODEL",
            "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2",
        )
        usa_prefijos_e5 = model_name.lower().startswith("intfloat/multilingual-e5")
        evidencia_texto = f"query: {texto}" if usa_prefijos_e5 else texto
        evidencia = _normalizar_vector(next(modelo.embed([evidencia_texto])))
        vectores_obligaciones = _vectorizar_obligaciones(
            model_name,
            tuple(clausula for _, clausula in obligaciones_validas),
        )
        return {
            numero: float(np.dot(evidencia, vector))
            for (numero, _), vector in zip(obligaciones_validas, vectores_obligaciones)
        }
    except Exception as exc:
        logger.warning("No se pudo calcular afinidad semántica local: %s", exc)
        return None
