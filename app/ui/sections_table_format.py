"""
Componente UI de Streamlit para configuración de formato de tablas.
"""
import streamlit as st
from typing import Dict


def render_table_format_section() -> dict:
    """
    Renderiza la sección de configuración de formato de tablas.

    Returns:
        Diccionario con la configuración de formato de tablas
    """
    st.header("🎨 Formato de Tablas")
    st.markdown("Configura el estilo visual de las tablas en el informe generado.")

    # Inicializar valores por defecto en session state
    if "table_format" not in st.session_state:
        st.session_state.table_format = {
            "show_borders": True,
            "border_style": "single",
            "header_bg_color": "#4472C4",
            "header_text_color": "#FFFFFF",
            "header_bold": True,
            "header_font_size": 11,
            "data_font_size": 10,
            "alternate_rows": False,
            "alternate_row_color": "#F2F2F2"
        }

    format_config = {}

    # Sección de bordes
    st.subheader("Bordes de la Tabla")
    col1, col2 = st.columns(2)

    with col1:
        show_borders = st.checkbox(
            "Mostrar bordes",
            value=st.session_state.table_format["show_borders"],
            key="format_show_borders",
            help="Activar/desactivar los bordes de la tabla"
        )
        format_config["show_borders"] = show_borders

    with col2:
        if show_borders:
            border_style = st.selectbox(
                "Estilo de bordes",
                options=["single", "double", "dashed", "dotted"],
                index=["single", "double", "dashed", "dotted"].index(
                    st.session_state.table_format["border_style"]
                ),
                key="format_border_style",
                help="Estilo de línea de los bordes"
            )
            format_config["border_style"] = border_style
        else:
            format_config["border_style"] = "none"

    # Sección de encabezado
    st.subheader("Formato del Encabezado")

    col1, col2 = st.columns(2)

    with col1:
        header_bg_color = st.color_picker(
            "Color de fondo del encabezado",
            value=st.session_state.table_format["header_bg_color"],
            key="format_header_bg_color",
            help="Color de fondo de la primera fila (encabezado)"
        )
        format_config["header_bg_color"] = header_bg_color

    with col2:
        header_text_color = st.color_picker(
            "Color de texto del encabezado",
            value=st.session_state.table_format["header_text_color"],
            key="format_header_text_color",
            help="Color del texto en el encabezado"
        )
        format_config["header_text_color"] = header_text_color

    col3, col4 = st.columns(2)

    with col3:
        header_bold = st.checkbox(
            "Texto en negrita",
            value=st.session_state.table_format["header_bold"],
            key="format_header_bold",
            help="Aplicar negrita al texto del encabezado"
        )
        format_config["header_bold"] = header_bold

    with col4:
        header_font_size = st.number_input(
            "Tamaño de fuente (pt)",
            min_value=8,
            max_value=24,
            value=st.session_state.table_format["header_font_size"],
            key="format_header_font_size",
            help="Tamaño de fuente del encabezado en puntos"
        )
        format_config["header_font_size"] = header_font_size

    # Sección de filas de datos
    st.subheader("Formato de Filas de Datos")

    col1, col2 = st.columns(2)

    with col1:
        data_font_size = st.number_input(
            "Tamaño de fuente (pt)",
            min_value=8,
            max_value=24,
            value=st.session_state.table_format["data_font_size"],
            key="format_data_font_size",
            help="Tamaño de fuente de las filas de datos"
        )
        format_config["data_font_size"] = data_font_size

    with col2:
        alternate_rows = st.checkbox(
            "Alternar color de filas",
            value=st.session_state.table_format["alternate_rows"],
            key="format_alternate_rows",
            help="Aplicar color alternado a las filas para mejor legibilidad"
        )
        format_config["alternate_rows"] = alternate_rows

    if alternate_rows:
        alternate_row_color = st.color_picker(
            "Color de filas alternadas",
            value=st.session_state.table_format["alternate_row_color"],
            key="format_alternate_row_color",
            help="Color de fondo para las filas pares"
        )
        format_config["alternate_row_color"] = alternate_row_color
    else:
        format_config["alternate_row_color"] = None

    # Botón para restablecer valores por defecto
    st.divider()
    if st.button("🔄 Restablecer valores por defecto"):
        st.session_state.table_format = {
            "show_borders": True,
            "border_style": "single",
            "header_bg_color": "#4472C4",
            "header_text_color": "#FFFFFF",
            "header_bold": True,
            "header_font_size": 11,
            "data_font_size": 10,
            "alternate_rows": False,
            "alternate_row_color": "#F2F2F2"
        }
        st.rerun()

    # Vista previa del formato
    st.divider()
    st.subheader("Vista Previa")

    # Crear una tabla de ejemplo con el formato aplicado
    preview_html = f"""
    <style>
        .preview-table {{
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
        }}
        .preview-table th {{
            background-color: {header_bg_color};
            color: {header_text_color};
            font-weight: {'bold' if header_bold else 'normal'};
            font-size: {header_font_size}pt;
            padding: 8px;
            text-align: left;
            border: {f'1px {border_style} #ddd' if show_borders else 'none'};
        }}
        .preview-table td {{
            font-size: {data_font_size}pt;
            padding: 8px;
            border: {f'1px {border_style} #ddd' if show_borders else 'none'};
        }}
        .preview-table tr:nth-child(even) {{
            background-color: {alternate_row_color if alternate_rows else 'transparent'};
        }}
    </style>
    <table class="preview-table">
        <thead>
            <tr>
                <th>Columna 1</th>
                <th>Columna 2</th>
                <th>Columna 3</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>Dato 1.1</td>
                <td>Dato 1.2</td>
                <td>Dato 1.3</td>
            </tr>
            <tr>
                <td>Dato 2.1</td>
                <td>Dato 2.2</td>
                <td>Dato 2.3</td>
            </tr>
            <tr>
                <td>Dato 3.1</td>
                <td>Dato 3.2</td>
                <td>Dato 3.3</td>
            </tr>
        </tbody>
    </table>
    """

    st.markdown(preview_html, unsafe_allow_html=True)

    # Actualizar session state
    st.session_state.table_format.update(format_config)

    return format_config
