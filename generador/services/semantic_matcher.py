"""Clasificación semántica local de evidencias y obligaciones."""

import logging
import re
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
    res = np.asarray(vector, dtype=np.float32)
    norm = np.linalg.norm(res)
    return res / norm if norm else res


@lru_cache(maxsize=64)
def _vectorizar_obligaciones(model_name: str, obligaciones: tuple[str, ...]) -> tuple[np.ndarray, ...]:
    modelo = _obtener_modelo()
    if modelo is None:
        return ()
    usa_prefijos = model_name.lower().startswith("intfloat/multilingual-e5")
    textos = [f"passage: {c}" if usa_prefijos else c for c in obligaciones]
    return tuple(_normalizar_vector(v) for v in modelo.embed(textos))


def puntuar_obligaciones_semanticamente(
    texto_evidencia: str,
    detalles: list,
) -> Optional[dict[int, float]]:
    texto = (texto_evidencia or "").strip()
    if not texto or not detalles:
        return None

    # Si es solo un nombre de captura genérico (ej. "imagen (2).png"), NO usar embeddings para evitar falsos positivos
    nombre_limpio = os.path.basename(texto).lower()
    if re.match(r"^(imagen|image|captura|screenshot|img|paste|imgpsh)[\s_\-\(\d\)]*\.(png|jpg|jpeg)$", nombre_limpio):
        return None

    modelo = _obtener_modelo()
    if modelo is None:
        return None

    obligaciones_validas = [
        (d.obligacion.numero_obligacion, (d.obligacion.descripcion_clausula or "").strip())
        for d in detalles
        if (d.obligacion.descripcion_clausula or "").strip()
    ]
    if not obligaciones_validas:
        return None

    try:
        model_name = getattr(settings, "LOCAL_SEMANTIC_MODEL", "intfloat/multilingual-e5-small")
        usa_prefijos = model_name.lower().startswith("intfloat/multilingual-e5")
        query_text = f"query: {texto}" if usa_prefijos else texto
        evidencia_vec = _normalizar_vector(next(modelo.embed([query_text])))
        vectores_ob = _vectorizar_obligaciones(model_name, tuple(c for _, c in obligaciones_validas))
        return {
            num: float(np.dot(evidencia_vec, vec))
            for (num, _), vec in zip(obligaciones_validas, vectores_ob)
        }
    except Exception as exc:
        logger.warning("No se pudo calcular afinidad semántica local: %s", exc)
        return None