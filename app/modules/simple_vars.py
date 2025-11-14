"""
Módulo para manejar variables simples.
"""
from typing import Dict, Any


def get_default_values(cfg_simple: dict) -> dict:
    """
    Obtiene valores por defecto para variables simples.

    Args:
        cfg_simple: Configuración de variables_simples.yaml

    Returns:
        Diccionario con valores por defecto
    """
    defaults = {}

    for var in cfg_simple.get("simple_variables", []):
        var_id = var["id"]
        var_type = var.get("type", "text")

        # Valores por defecto según el tipo
        if var_type == "text":
            defaults[var_id] = ""
        elif var_type == "long_text":
            defaults[var_id] = ""
        elif var_type == "number":
            defaults[var_id] = 0.0
        elif var_type == "percent":
            defaults[var_id] = 0.0
        elif var_type == "integer":
            defaults[var_id] = 0
        elif var_type == "email":
            defaults[var_id] = ""
        else:
            defaults[var_id] = ""

    return defaults


def validate_simple_vars(cfg_simple: dict, inputs: dict) -> list:
    """
    Valida las variables simples ingresadas.

    Args:
        cfg_simple: Configuración de variables simples
        inputs: Valores ingresados

    Returns:
        Lista de errores (vacía si todo es válido)
    """
    errors = []

    for var in cfg_simple.get("simple_variables", []):
        var_id = var["id"]
        label = var["label"]
        var_type = var.get("type", "text")
        marker = var.get("marker")
        optional = var.get("optional", False)

        value = inputs.get(var_id)

        # Si tiene marcador y NO es opcional, es requerido
        if marker is not None and not optional:
            if value is None or (isinstance(value, str) and not value.strip()):
                errors.append(f"El campo '{label}' es requerido")
                continue

        # Validar tipos
        if value is not None and value != "":
            if var_type == "email":
                if "@" not in str(value):
                    errors.append(f"El campo '{label}' debe ser un email válido")
            elif var_type in ["number", "percent"]:
                try:
                    float(value)
                except (ValueError, TypeError):
                    errors.append(f"El campo '{label}' debe ser un número")
            elif var_type == "integer":
                try:
                    int(value)
                except (ValueError, TypeError):
                    errors.append(f"El campo '{label}' debe ser un número entero")

    return errors
