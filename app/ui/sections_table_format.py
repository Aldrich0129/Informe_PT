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
            "border_color": "#000000",
            "header_bg_color": "#4472C4",
            "header_text_color": "#FFFFFF",
            "header_bold": True,
            "header_font_size": 11,
            "data_font_size": 10,
            "alternate_rows": False,
            "alternate_row_color": "#F2F2F2",
            "first_column_bold": False,
            "first_column_bg_color": None,
            "first_column_text_color": None
        }

    format_config = {}

    # Sección de bordes
    st.subheader("Bordes de la Tabla")
    col1, col2, col3 = st.columns(3)

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
            # Obtener el valor actual de border_style del session state
            current_border_style = st.session_state.table_format.get("border_style", "single")

            # Lista de estilos válidos
            valid_styles = ["single", "double", "dashed", "dotted"]

            # Si el valor actual no es válido (ej: "none"), usar "single" por defecto
            if current_border_style not in valid_styles:
                current_border_style = "single"

            border_style = st.selectbox(
                "Estilo de bordes",
                options=valid_styles,
                index=valid_styles.index(current_border_style),
                key="format_border_style",
                help="Estilo de línea de los bordes"
            )
            format_config["border_style"] = border_style
        else:
            # No actualizar border_style en session_state cuando está desactivado
            # Mantener el último valor válido para cuando se vuelva a activar
            format_config["border_style"] = st.session_state.table_format.get("border_style", "single")

    with col3:
        if show_borders:
            border_color = st.color_picker(
                "Color de bordes",
                value=st.session_state.table_format["border_color"],
                key="format_border_color",
                help="Color de las líneas de borde"
            )
            format_config["border_color"] = border_color
        else:
            format_config["border_color"] = "#000000"

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

    # Sección de primera columna
    st.subheader("Formato de Primera Columna")
    st.markdown("Configura el estilo de las celdas de la primera columna (excluyendo el encabezado)")

    col1, col2, col3 = st.columns(3)

    with col1:
        first_column_bold = st.checkbox(
            "Texto en negrita",
            value=st.session_state.table_format.get("first_column_bold", False),
            key="format_first_column_bold",
            help="Aplicar negrita al texto de la primera columna"
        )
        format_config["first_column_bold"] = first_column_bold

    with col2:
        use_first_col_bg = st.checkbox(
            "Color de fondo personalizado",
            value=st.session_state.table_format.get("first_column_bg_color") is not None,
            key="use_first_col_bg",
            help="Activar color de fondo personalizado para la primera columna"
        )
        if use_first_col_bg:
            first_column_bg_color = st.color_picker(
                "Color de fondo",
                value=st.session_state.table_format.get("first_column_bg_color", "#E7E6E6"),
                key="format_first_column_bg_color",
                help="Color de fondo de la primera columna"
            )
            format_config["first_column_bg_color"] = first_column_bg_color
        else:
            format_config["first_column_bg_color"] = None

    with col3:
        use_first_col_text = st.checkbox(
            "Color de texto personalizado",
            value=st.session_state.table_format.get("first_column_text_color") is not None,
            key="use_first_col_text",
            help="Activar color de texto personalizado para la primera columna"
        )
        if use_first_col_text:
            first_column_text_color = st.color_picker(
                "Color de texto",
                value=st.session_state.table_format.get("first_column_text_color", "#000000"),
                key="format_first_column_text_color",
                help="Color de texto de la primera columna"
            )
            format_config["first_column_text_color"] = first_column_text_color
        else:
            format_config["first_column_text_color"] = None

    # Botón para restablecer valores por defecto
    st.divider()
    if st.button("🔄 Restablecer valores por defecto"):
        st.session_state.table_format = {
            "show_borders": True,
            "border_style": "single",
            "border_color": "#000000",
            "header_bg_color": "#4472C4",
            "header_text_color": "#FFFFFF",
            "header_bold": True,
            "header_font_size": 11,
            "data_font_size": 10,
            "alternate_rows": False,
            "alternate_row_color": "#F2F2F2",
            "first_column_bold": False,
            "first_column_bg_color": None,
            "first_column_text_color": None
        }
        st.rerun()

    # Vista previa del formato
    st.divider()
    st.subheader("Vista Previa")

    # Crear una tabla de ejemplo con el formato aplicado
    border_color_value = format_config.get("border_color", "#000000")
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
            border: {f'1px {border_style} {border_color_value}' if show_borders else 'none'};
        }}
        .preview-table td {{
            font-size: {data_font_size}pt;
            padding: 8px;
            border: {f'1px {border_style} {border_color_value}' if show_borders else 'none'};
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
