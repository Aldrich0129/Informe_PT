"""
UI principal que orquesta todas las secciones de la aplicación.
"""
import streamlit as st
from pathlib import Path
import json

from ui.sections_simple_vars import render_simple_vars_section
from ui.sections_conditions import render_conditions_section
from ui.sections_tables import render_tables_section
from ui.sections_table_format import render_table_format_section
from modules.utils import export_data_to_json, import_data_from_json, generate_filename


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

        # Sección de guardar/cargar datos
        st.header("💾 Datos")

        # Botón para cargar datos
        uploaded_file = st.file_uploader(
            "Cargar datos guardados",
            type=["json"],
            help="Carga un archivo JSON con datos previamente guardados"
        )

        if uploaded_file is not None:
            try:
                data = json.load(uploaded_file)
                simple_inputs_loaded, condition_inputs_loaded, table_inputs_loaded, table_format_loaded = import_data_from_json(data)

                # Guardar en session_state
                st.session_state.loaded_simple_inputs = simple_inputs_loaded
                st.session_state.loaded_condition_inputs = condition_inputs_loaded
                st.session_state.loaded_table_inputs = table_inputs_loaded
                st.session_state.loaded_table_format = table_format_loaded
                st.session_state.data_loaded = True

                st.success("✅ Datos cargados correctamente")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al cargar datos: {e}")

        st.divider()
        st.markdown("**Desarrollado con Streamlit + Python-docx**")

    # Aplicar datos cargados desde JSON a las keys de los widgets
    if st.session_state.get("data_loaded", False):
        # Copiar variables simples a las keys de los widgets
        if "loaded_simple_inputs" in st.session_state:
            for key, value in st.session_state.loaded_simple_inputs.items():
                st.session_state[f"simple_{key}"] = value

        # Copiar inputs de condiciones
        if "loaded_condition_inputs" in st.session_state:
            for key, value in st.session_state.loaded_condition_inputs.items():
                st.session_state[f"cond_{key}"] = value

        # Copiar configuración de formato de tabla
        if "loaded_table_format" in st.session_state:
            st.session_state.table_format = st.session_state.loaded_table_format

        # Copiar inputs de tablas
        if "loaded_table_inputs" in st.session_state:
            for table_key, table_data in st.session_state.loaded_table_inputs.items():
                # Para cada tabla, guardar sus datos en session_state
                st.session_state[f"table_{table_key}"] = table_data

                # Para las operaciones vinculadas, también necesitamos copiar cada campo
                if table_key == "operaciones_vinculadas":
                    # table_data es una lista de operaciones
                    if isinstance(table_data, list):
                        for i, op in enumerate(table_data):
                            st.session_state[f"op_{i}_tipo"] = op.get("tipo_operacion", "")
                            st.session_state[f"op_{i}_entidad"] = op.get("entidad_vinculada", "")
                            st.session_state[f"op_{i}_ingreso"] = op.get("ingreso_local_file", 0.0)
                            st.session_state[f"op_{i}_gasto"] = op.get("gasto_local_file", 0.0)
                        st.session_state.num_operaciones = len(table_data)

                # Para análisis indirecto global (TNMM)
                elif table_key == "analisis_indirecto_global":
                    if "rango_tnmm" in table_data:
                        rango = table_data["rango_tnmm"]
                        st.session_state["tnmm_global_min"] = rango.get("min", 1.0)
                        st.session_state["tnmm_global_lq"] = rango.get("lq", 2.0)
                        st.session_state["tnmm_global_med"] = rango.get("med", 3.0)
                        st.session_state["tnmm_global_uq"] = rango.get("uq", 4.0)
                        st.session_state["tnmm_global_max"] = rango.get("max", 5.0)

                # Para análisis indirecto por operación (TNMM por operación)
                elif table_key.startswith("analisis_indirecto_operacion_"):
                    # Extraer el número de operación
                    n = table_key.split("_")[-1]
                    # Las keys son tnmm_op_{n}_{col_id}
                    st.session_state[f"tnmm_op_{n}_min"] = table_data.get("min", 1.0)
                    st.session_state[f"tnmm_op_{n}_lq"] = table_data.get("lq", 2.0)
                    st.session_state[f"tnmm_op_{n}_med"] = table_data.get("med", 3.0)
                    st.session_state[f"tnmm_op_{n}_uq"] = table_data.get("uq", 4.0)
                    st.session_state[f"tnmm_op_{n}_max"] = table_data.get("max", 5.0)

        # Resetear bandera para no repetir el proceso
        st.session_state.data_loaded = False

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

    # Sección de descarga de datos (después de las pestañas)
    st.markdown("---")
    st.subheader("💾 Guardar Datos")
    st.markdown("Descarga todos los datos rellenados para recuperarlos en próximo uso.")

    col1, col2, col3 = st.columns([2, 1, 2])

    with col2:
        if st.button("💾 Descargar Datos", type="secondary", use_container_width=True):
            # Generar nombre de archivo
            nombre_empresa = simple_inputs.get("nombre_compania", "Empresa")
            ejercicio = simple_inputs.get("ejercicio_completo", "2023")
            filename = generate_filename(nombre_empresa, ejercicio)

            # Exportar datos
            data = export_data_to_json(simple_inputs, condition_inputs, table_inputs, table_format_config)

            # Convertir a JSON
            json_data = json.dumps(data, indent=2, ensure_ascii=False)

            # Botón de descarga
            st.download_button(
                label="📥 Descargar JSON",
                data=json_data,
                file_name=filename,
                mime="application/json",
                type="primary",
                use_container_width=True
            )

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
