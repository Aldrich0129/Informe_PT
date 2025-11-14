"""
UI principal que orquesta todas las secciones de la aplicación.
"""
import streamlit as st
from pathlib import Path

from ui.sections_simple_vars import render_simple_vars_section
from ui.sections_conditions import render_conditions_section
from ui.sections_tables import render_tables_section
from ui.sections_table_format import render_table_format_section


def render_main_ui(cfg_simple: dict, cfg_cond: dict, cfg_tab: dict):
    """
    Renderiza la UI principal con pestañas para cada sección.

    Args:
        cfg_simple: Configuración de variables simples
        cfg_cond: Configuración de condiciones
        cfg_tab: Configuración de tablas

    Returns:
        Tupla con (simple_inputs, condition_inputs, table_inputs)
    """
    # Configuración de la página
    st.set_page_config(
        page_title="Generador de Informes PT",
        page_icon="📄",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Título principal
    st.title("📄 Generador de Informes de Precios de Transferencia")
    st.markdown("---")

    # Sidebar con información
    with st.sidebar:
        st.header("ℹ️ Información")
        st.markdown(
            """
            Esta aplicación genera informes Word personalizados
            a partir de plantillas y datos estructurados.

            **Pasos:**
            1. Completa las variables simples
            2. Selecciona bloques condicionales
            3. Rellena las tablas
            4. Genera el informe

            **Características:**
            - ✅ Reemplazo de variables con formato
            - ✅ Inserción de tablas dinámicas
            - ✅ Bloques condicionales
            - ✅ Limpieza automática
            """
        )

        st.divider()
        st.markdown("**Desarrollado con Streamlit + Python-docx**")

    # Usar pestañas para organizar las secciones (Tablas antes de Condiciones)
    tabs = st.tabs(["📝 Variables Simples", "📊 Tablas", "🎨 Formato de Tablas", "🔀 Condiciones"])

    # Inicializar variables
    simple_inputs = {}
    condition_inputs = {}
    table_inputs = {}
    table_format_config = {}

    # Pestaña 1: Variables Simples
    with tabs[0]:
        simple_inputs = render_simple_vars_section(cfg_simple)

    # Pestaña 2: Tablas
    with tabs[1]:
        table_inputs = render_tables_section(cfg_tab, simple_inputs)

    # Pestaña 3: Formato de Tablas
    with tabs[2]:
        table_format_config = render_table_format_section()

    # Pestaña 4: Condiciones
    with tabs[3]:
        condition_inputs = render_conditions_section(cfg_cond)

    return simple_inputs, condition_inputs, table_inputs, table_format_config


def render_generation_section():
    """
    Renderiza la sección de generación del documento.

    Returns:
        True si se debe generar el documento, False en caso contrario
    """
    st.markdown("---")
    st.header("🚀 Generar Informe")

    col1, col2, col3 = st.columns([2, 1, 2])

    with col2:
        generate_button = st.button(
            "📄 Generar Informe Word",
            type="primary",
            use_container_width=True
        )

    if generate_button:
        return True

    return False


def show_validation_errors(errors: list):
    """
    Muestra errores de validación.

    Args:
        errors: Lista de mensajes de error
    """
    if errors:
        st.error("❌ Se encontraron los siguientes errores:")
        for error in errors:
            st.markdown(f"- {error}")


def show_success_message():
    """Muestra mensaje de éxito."""
    st.success("✅ ¡Informe generado correctamente!")


def show_processing_spinner(message: str = "Generando informe..."):
    """
    Muestra un spinner de procesamiento.

    Args:
        message: Mensaje a mostrar

    Returns:
        Context manager para el spinner
    """
    return st.spinner(message)
