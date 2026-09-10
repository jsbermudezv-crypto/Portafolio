"""Utilidades generales para procesamiento de datos."""

import logging
import re
import unicodedata

logger = logging.getLogger(__name__)


def clean_text(text: str) -> str:
    """
    Limpia y normaliza texto.
    
    Args:
        text: Texto a limpiar
        
    Returns:
        Texto limpio
    """
    # Remover acentos
    text = (
        unicodedata.normalize("NFD", text)
        .encode("ascii", "ignore")
        .decode("utf-8")
    )
    
    # Convertir a minúsculas
    text = text.lower()
    
    # Remover caracteres especiales
    text = re.sub(r"[^a-z0-9\s]", "", text)
    
    # Remover espacios múltiples
    text = " ".join(text.split())
    
    return text


def normalize_names(names: list) -> list:
    """
    Normaliza una lista de nombres.
    
    Args:
        names: Lista de nombres
        
    Returns:
        Lista de nombres normalizados
    """
    return [clean_text(name) for name in names]


def is_valid_email(email: str) -> bool:
    """Valida formato de email."""
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email) is not None
