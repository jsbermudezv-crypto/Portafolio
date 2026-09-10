"""Tests para utilidades compartidas."""

import pytest

from src.utils.data_utils import clean_text, is_valid_email, normalize_names


class TestCleanText:
    """Tests para limpieza de texto."""
    
    def test_acentos(self):
        """Debe remover acentos."""
        assert clean_text("café") == "cafe"
        assert clean_text("Mañana") == "manana"
    
    def test_minusculas(self):
        """Debe convertir a minúsculas."""
        assert clean_text("HELLO") == "hello"
        assert clean_text("HeLLo") == "hello"
    
    def test_caracteres_especiales(self):
        """Debe remover caracteres especiales."""
        assert clean_text("hello-world!") == "hello world"
        assert clean_text("test@123") == "test 123"
    
    def test_espacios_multiples(self):
        """Debe remover espacios múltiples."""
        assert clean_text("hello   world") == "hello world"
        assert clean_text("  hello  world  ") == "hello world"


class TestEmailValidation:
    """Tests para validación de emails."""
    
    def test_email_valido(self):
        """Debe aceptar emails válidos."""
        assert is_valid_email("test@example.com")
        assert is_valid_email("user.name+tag@example.co.uk")
    
    def test_email_invalido(self):
        """Debe rechazar emails inválidos."""
        assert not is_valid_email("notanemail")
        assert not is_valid_email("@example.com")
        assert not is_valid_email("test@")


class TestNormalizeNames:
    """Tests para normalización de nombres."""
    
    def test_lista_nombres(self):
        """Debe normalizar una lista de nombres."""
        names = ["José", "María-José", "JUAN PABLO"]
        result = normalize_names(names)
        assert result == ["jose", "maria jose", "juan pablo"]
