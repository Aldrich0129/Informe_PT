"""
Módulo de utilidades para construir el contexto completo
que se pasa al motor de generación de Word.
"""
from typing import Dict, List, Any, Tuple
import json
from datetime import datetime


def build_simple_context(cfg_simple: dict, simple_inputs: dict) -> dict:
    """
    Construye el contexto de variables simples.

    Args:
        cfg_simple: Configuración de variables_simples.yaml
        simple_inputs: Diccionario con {id_variable: valor} ingresados por el usuario

    Returns:
        Diccionario con {marker: valor} para reemplazar en la plantilla
    """
    context = {}

    for var in cfg_simple.get("simple_variables", []):
        var_id = var["id"]
        marker = var.get("marker")  # puede ser None (ej: ejercicio_anterior)

        value = simple_inputs.get(var_id)

        # Si no hay marcador, lo omitimos en plantilla,
        # pero el valor estará disponible para tablas
        if marker and value is not None:
            # Formatear según el tipo
            var_type = var.get("type", "text")

            if var_type == "percent" and isinstance(value, (int, float)):
                # Convertir a porcentaje (ej: 0.35 -> "35%")
                context[marker] = f"{value * 100:.2f}%"
            elif var_type == "number" and isinstance(value, (int, float)):
                # Formatear números con separadores de miles
                context[marker] = f"{value:,.2f}"
            else:
                # Texto, long_text, email, etc.
                context[marker] = str(value)

    return context


def build_conditions_context(cfg_cond: dict, condition_inputs: dict) -> Tuple[dict, list]:
    """
    Construye el contexto de condiciones.

    Args:
        cfg_cond: Configuración de variables_condicionales.yaml
        condition_inputs: Diccionario con {id_condicion: "Sí"/"No"}

    Returns:
        Tupla con:
        - context_markers: {marker: ""} para limpiar si es "No"
        - docs_to_insert: lista de {marker, file} para insertar bloques Word
    """
    context_markers = {}
    docs_to_insert = []

    for cond in cfg_cond.get("conditions", []):
        cond_id = cond["id"]
        marker = cond["marker"]
        word_file = cond["word_file"]

        answer = condition_inputs.get(cond_id, "No")

        if answer == "Sí":
            # Marcar para insertar el bloque de Word
            docs_to_insert.append({
                "marker": marker,
                "file": word_file
            })
            # Dejar el marcador vacío por ahora (se reemplazará con el bloque)
            context_markers[marker] = ""
        else:
            # Limpieza: borrar el marcador
            context_markers[marker] = ""

    return context_markers, docs_to_insert


def build_operations_list_context(cfg_simple: dict, table_inputs: dict) -> dict:
    """
    Construye el contexto para las operaciones vinculadas (texto).

    Args:
        cfg_simple: Configuración que contiene la sección 'operations'
        table_inputs: Datos de tablas (incluye operaciones)

    Returns:
        Diccionario con marcadores de texto de operaciones
    """
    context = {}

    operations = cfg_simple.get("operations", {}).get("items", [])
    operaciones_vinculadas = table_inputs.get("operaciones_vinculadas", [])

    for i, op in enumerate(operations):
        text_marker = op.get("text_marker")

        # Si hay datos para esta operación en la tabla, usar el tipo de operación
        if i < len(operaciones_vinculadas):
            op_data = operaciones_vinculadas[i]
            tipo_operacion = op_data.get("tipo_operacion", "")

            if tipo_operacion.strip():
                context[text_marker] = tipo_operacion
            else:
                context[text_marker] = ""
        else:
            # No hay datos, dejar vacío
            context[text_marker] = ""

    return context


def build_full_context(
    cfg_simple: dict,
    cfg_cond: dict,
    cfg_tab: dict,
    simple_inputs: dict,
    condition_inputs: dict,
    table_inputs: dict
) -> Tuple[dict, list]:
    """
    Construye el contexto completo para la generación del documento Word.

    Args:
        cfg_simple: Configuración de variables simples
        cfg_cond: Configuración de condiciones
        cfg_tab: Configuración de tablas
        simple_inputs: Datos de variables simples
        condition_inputs: Datos de condiciones
        table_inputs: Datos de tablas

    Returns:
        Tupla con:
        - context: Diccionario completo de contexto
        - docs_to_insert: Lista de bloques Word condicionales a insertar
    """
    context = {}

    # 1. Variables simples
    context.update(build_simple_context(cfg_simple, simple_inputs))

    # 2. Condiciones (marcadores + lista de docs)
    cond_markers, docs_to_insert = build_conditions_context(cfg_cond, condition_inputs)
    context.update(cond_markers)

    # 3. Operaciones vinculadas (texto)
    context.update(build_operations_list_context(cfg_simple, table_inputs))

    # 4. Las tablas se manejarán directamente por el motor de Word
    # Pasamos los datos completos
    context["_table_inputs"] = table_inputs
    context["_simple_inputs"] = simple_inputs

    return context, docs_to_insert


def format_number(value: Any, format_type: str = "number") -> str:
    """
    Formatea un número según el tipo especificado.

    Args:
        value: Valor a formatear
        format_type: Tipo de formato (number, percent, integer)

    Returns:
        String formateado
    """
    if value is None or value == "":
        return ""

    try:
        num_value = float(value)

        if format_type == "percent":
            return f"{num_value:.2f}%"
        elif format_type == "integer":
            return str(int(num_value))
        else:  # number
            return f"{num_value:,.2f}"

    except (ValueError, TypeError):
        return str(value)


def validate_inputs(cfg_simple: dict, simple_inputs: dict) -> List[str]:
    """
    Valida que las entradas requeridas estén presentes.

    Args:
        cfg_simple: Configuración de variables simples
        simple_inputs: Entradas del usuario

    Returns:
        Lista de errores (vacía si todo está correcto)
    """
    errors = []

    for var in cfg_simple.get("simple_variables", []):
        var_id = var["id"]
        label = var["label"]

        # Verificar si es requerido (por defecto, todas son requeridas excepto ejercicio_anterior)
        if var.get("marker") is not None:  # Si tiene marcador, es requerido
            value = simple_inputs.get(var_id)

            if value is None or (isinstance(value, str) and not value.strip()):
                errors.append(f"El campo '{label}' es requerido")

    return errors


def export_data_to_json(simple_inputs: dict, condition_inputs: dict, table_inputs: dict, table_format_config: dict) -> dict:
    """
    Exporta todos los datos a un diccionario para guardar como JSON.

    Args:
        simple_inputs: Datos de variables simples
        condition_inputs: Datos de condiciones
        table_inputs: Datos de tablas
        table_format_config: Configuración de formato de tablas

    Returns:
        Diccionario con todos los datos
    """
    return {
        "version": "1.0",
        "export_date": datetime.now().isoformat(),
        "simple_inputs": simple_inputs,
        "condition_inputs": condition_inputs,
        "table_inputs": table_inputs,
        "table_format_config": table_format_config
    }


def import_data_from_json(data: dict) -> Tuple[dict, dict, dict, dict]:
    """
    Importa datos desde un diccionario JSON.

    Args:
        data: Diccionario con los datos exportados

    Returns:
        Tupla con (simple_inputs, condition_inputs, table_inputs, table_format_config)
    """
    simple_inputs = data.get("simple_inputs", {})
    condition_inputs = data.get("condition_inputs", {})
    table_inputs = data.get("table_inputs", {})
    table_format_config = data.get("table_format_config", {})

    return simple_inputs, condition_inputs, table_inputs, table_format_config


def generate_filename(nombre_empresa: str, ejercicio: str) -> str:
    """
    Genera un nombre de archivo con formato: nombre_empresa_ejercicio_XXXX.json
    donde XXXX son 4 dígitos basados en el timestamp.

    Args:
        nombre_empresa: Nombre de la empresa
        ejercicio: Ejercicio fiscal

    Returns:
        Nombre de archivo generado
    """
    # Limpiar el nombre de la empresa (eliminar caracteres especiales)
    nombre_limpio = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in nombre_empresa)
    nombre_limpio = nombre_limpio.replace(' ', '_')

    # Obtener 4 dígitos del timestamp (últimos 4 dígitos de milisegundos)
    timestamp = int(datetime.now().timestamp() * 1000)
    timestamp_4digits = str(timestamp)[-4:]

    # Generar nombre de archivo
    filename = f"{nombre_limpio}_{ejercicio}_{timestamp_4digits}.json"

    return filename
