"""
Componente UI de Streamlit para configuración de formato de tablas.
"""
import streamlit as st
import json
from typing import Dict


def render_table_format_section(table_custom_design: dict = None) -> dict:
    """
    Renderiza la sección de configuración de formato de tablas.

    Args:
        table_custom_design: Diccionario con {id_tabla: usar_diseño_personalizado}

    Returns:
        Diccionario con la configuración de formato de tablas
    """
    if table_custom_design is None:
        table_custom_design = {}

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

    # Nueva sección: Color de fondo por columna
    st.subheader("Color de Fondo por Columna")
    st.markdown("Configura colores de fondo específicos para columnas individuales (no afecta la primera fila/encabezado)")

    # Inicializar column_colors en session_state si no existe
    if "column_colors" not in st.session_state.table_format:
        st.session_state.table_format["column_colors"] = []

    # Número de colores de columna configurados
    if "num_column_colors" not in st.session_state:
        st.session_state.num_column_colors = 0

    # Botones para agregar/quitar colores de columna
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("➕ Añadir color de columna"):
            st.session_state.num_column_colors += 1
            st.rerun()

    with col2:
        if st.button("➖ Quitar último color"):
            if st.session_state.num_column_colors > 0:
                st.session_state.num_column_colors -= 1
                st.rerun()

    # Renderizar controles de color de columna
    column_colors = []
    for i in range(st.session_state.num_column_colors):
        col1, col2 = st.columns([1, 2])
        with col1:
            column_number = st.number_input(
                f"Número de columna",
                key=f"column_color_{i}_number",
                min_value=1,
                max_value=20,
                value=i + 2,  # Por defecto, comenzar desde la columna 2
                help="Número de la columna (1 es la primera columna)"
            )
        with col2:
            column_color = st.color_picker(
                f"Color de fondo",
                key=f"column_color_{i}_color",
                value="#F0F0F0",
                help="Color de fondo para esta columna"
            )

        column_colors.append({
            "column": column_number,
            "color": column_color
        })

    format_config["column_colors"] = column_colors

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
            "first_column_text_color": None,
            "column_colors": []
        }
        st.session_state.num_column_colors = 0
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

    # Sección de diseño personalizado para tablas individuales
    st.divider()
    st.header("🎨 Diseño Personalizado de Tablas Individuales")
    st.markdown("Configura el diseño específico para cada tabla marcada como 'Personalizar diseño'")

    # Obtener tablas que tienen diseño personalizado activado
    custom_tables = {table_id: custom for table_id, custom in table_custom_design.items() if custom}

    if not custom_tables:
        st.info("ℹ️ No hay tablas con diseño personalizado activado. Marca la casilla '🎨 Personalizar diseño de esta tabla' en la sección de Tablas para configurar diseños individuales.")
    else:
        # Inicializar custom_table_formats en session state si no existe
        if "custom_table_formats" not in st.session_state:
            st.session_state.custom_table_formats = {}

        # Mostrar configuración para cada tabla personalizada
        custom_table_formats = {}

        for table_id in custom_tables.keys():
            with st.expander(f"📊 {table_id.replace('_', ' ').title()}", expanded=False):
                st.markdown(f"**Configuración específica para la tabla: {table_id}**")

                # Inicializar config para esta tabla si no existe
                if table_id not in st.session_state.custom_table_formats:
                    st.session_state.custom_table_formats[table_id] = format_config.copy()

                table_config = {}

                # Bordes
                st.subheader("Bordes")
                col1, col2, col3 = st.columns(3)

                with col1:
                    show_borders = st.checkbox(
                        "Mostrar bordes",
                        value=st.session_state.custom_table_formats.get(table_id, {}).get("show_borders", True),
                        key=f"custom_{table_id}_show_borders"
                    )
                    table_config["show_borders"] = show_borders

                with col2:
                    if show_borders:
                        border_style = st.selectbox(
                            "Estilo de bordes",
                            options=["single", "double", "dashed", "dotted"],
                            index=0,
                            key=f"custom_{table_id}_border_style"
                        )
                        table_config["border_style"] = border_style
                    else:
                        table_config["border_style"] = "single"

                with col3:
                    if show_borders:
                        border_color = st.color_picker(
                            "Color de bordes",
                            value="#000000",
                            key=f"custom_{table_id}_border_color"
                        )
                        table_config["border_color"] = border_color
                    else:
                        table_config["border_color"] = "#000000"

                # Encabezado
                st.subheader("Encabezado")
                col1, col2, col3, col4 = st.columns(4)

                with col1:
                    header_bg_color = st.color_picker(
                        "Color de fondo",
                        value="#4472C4",
                        key=f"custom_{table_id}_header_bg_color"
                    )
                    table_config["header_bg_color"] = header_bg_color

                with col2:
                    header_text_color = st.color_picker(
                        "Color de texto",
                        value="#FFFFFF",
                        key=f"custom_{table_id}_header_text_color"
                    )
                    table_config["header_text_color"] = header_text_color

                with col3:
                    header_bold = st.checkbox(
                        "Negrita",
                        value=True,
                        key=f"custom_{table_id}_header_bold"
                    )
                    table_config["header_bold"] = header_bold

                with col4:
                    header_font_size = st.number_input(
                        "Tamaño fuente (pt)",
                        min_value=8,
                        max_value=24,
                        value=11,
                        key=f"custom_{table_id}_header_font_size"
                    )
                    table_config["header_font_size"] = header_font_size

                # Filas de datos
                st.subheader("Filas de Datos")
                col1, col2, col3 = st.columns(3)

                with col1:
                    data_font_size = st.number_input(
                        "Tamaño fuente (pt)",
                        min_value=8,
                        max_value=24,
                        value=10,
                        key=f"custom_{table_id}_data_font_size"
                    )
                    table_config["data_font_size"] = data_font_size

                with col2:
                    alternate_rows = st.checkbox(
                        "Alternar color de filas",
                        value=False,
                        key=f"custom_{table_id}_alternate_rows"
                    )
                    table_config["alternate_rows"] = alternate_rows

                with col3:
                    if alternate_rows:
                        alternate_row_color = st.color_picker(
                            "Color filas alternadas",
                            value="#F2F2F2",
                            key=f"custom_{table_id}_alternate_row_color"
                        )
                        table_config["alternate_row_color"] = alternate_row_color
                    else:
                        table_config["alternate_row_color"] = None

                # Primera columna
                st.subheader("Primera Columna")
                col1, col2, col3 = st.columns(3)

                with col1:
                    first_column_bold = st.checkbox(
                        "Negrita",
                        value=False,
                        key=f"custom_{table_id}_first_column_bold"
                    )
                    table_config["first_column_bold"] = first_column_bold

                with col2:
                    use_first_col_bg = st.checkbox(
                        "Color de fondo",
                        value=False,
                        key=f"custom_{table_id}_use_first_col_bg"
                    )
                    if use_first_col_bg:
                        first_column_bg_color = st.color_picker(
                            "Color",
                            value="#E7E6E6",
                            key=f"custom_{table_id}_first_column_bg_color"
                        )
                        table_config["first_column_bg_color"] = first_column_bg_color
                    else:
                        table_config["first_column_bg_color"] = None

                with col3:
                    use_first_col_text = st.checkbox(
                        "Color de texto",
                        value=False,
                        key=f"custom_{table_id}_use_first_col_text"
                    )
                    if use_first_col_text:
                        first_column_text_color = st.color_picker(
                            "Color",
                            value="#000000",
                            key=f"custom_{table_id}_first_column_text_color"
                        )
                        table_config["first_column_text_color"] = first_column_text_color
                    else:
                        table_config["first_column_text_color"] = None

                # Color de fondo por columna (específico para esta tabla)
                st.subheader("Color de Fondo por Columna")
                st.markdown("Configura colores para columnas específicas (no afecta el encabezado)")

                # Inicializar num_column_colors para esta tabla si no existe
                num_colors_key = f"num_column_colors_{table_id}"
                if num_colors_key not in st.session_state:
                    st.session_state[num_colors_key] = 0

                col1, col2 = st.columns([1, 1])
                with col1:
                    if st.button("➕ Añadir color", key=f"add_col_color_{table_id}"):
                        st.session_state[num_colors_key] += 1
                        st.rerun()

                with col2:
                    if st.button("➖ Quitar último", key=f"remove_col_color_{table_id}"):
                        if st.session_state[num_colors_key] > 0:
                            st.session_state[num_colors_key] -= 1
                            st.rerun()

                # Renderizar controles de color de columna para esta tabla
                column_colors = []
                for i in range(st.session_state[num_colors_key]):
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        column_number = st.number_input(
                            f"Columna",
                            key=f"custom_{table_id}_column_color_{i}_number",
                            min_value=1,
                            max_value=20,
                            value=i + 2
                        )
                    with col2:
                        column_color = st.color_picker(
                            f"Color",
                            key=f"custom_{table_id}_column_color_{i}_color",
                            value="#F0F0F0"
                        )

                    column_colors.append({
                        "column": column_number,
                        "color": column_color
                    })

                table_config["column_colors"] = column_colors

                # Guardar configuración de esta tabla
                custom_table_formats[table_id] = table_config
                st.session_state.custom_table_formats[table_id] = table_config

        # Agregar las configuraciones personalizadas al formato general
        format_config["custom_table_formats"] = custom_table_formats

    # Sección de guardar/cargar diseño de tablas
    st.divider()
    st.header("💾 Guardar/Cargar Diseño de Tablas")
    st.markdown("Guarda y carga configuraciones de diseño de tablas para reutilizarlas en futuros informes.")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📥 Cargar Diseño")
        uploaded_design_file = st.file_uploader(
            "Cargar archivo de diseño",
            type=["json"],
            help="Carga un archivo JSON con diseño de tablas previamente guardado",
            key="upload_table_design"
        )

        if uploaded_design_file is not None:
            try:
                design_data = json.load(uploaded_design_file)

                # Actualizar session_state con el diseño cargado
                if "table_format" in design_data:
                    st.session_state.table_format = design_data["table_format"]

                if "custom_table_formats" in design_data:
                    st.session_state.custom_table_formats = design_data["custom_table_formats"]

                if "num_column_colors" in design_data:
                    st.session_state.num_column_colors = design_data["num_column_colors"]

                # Actualizar también los contadores de colores de columna para tablas personalizadas
                if "column_color_counts" in design_data:
                    for table_id, count in design_data["column_color_counts"].items():
                        st.session_state[f"num_column_colors_{table_id}"] = count

                st.success("✅ Diseño de tablas cargado correctamente")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al cargar diseño: {e}")

    with col2:
        st.subheader("📤 Guardar Diseño")
        if st.button("💾 Descargar Diseño de Tablas", use_container_width=True):
            # Preparar datos para guardar
            design_data = {
                "table_format": st.session_state.get("table_format", format_config),
                "custom_table_formats": st.session_state.get("custom_table_formats", {}),
                "num_column_colors": st.session_state.get("num_column_colors", 0)
            }

            # Guardar también los contadores de colores de columna para tablas personalizadas
            column_color_counts = {}
            for key in st.session_state.keys():
                if key.startswith("num_column_colors_"):
                    table_id = key.replace("num_column_colors_", "")
                    column_color_counts[table_id] = st.session_state[key]

            design_data["column_color_counts"] = column_color_counts

            # Convertir a JSON
            json_data = json.dumps(design_data, indent=2, ensure_ascii=False)

            # Botón de descarga
            st.download_button(
                label="📥 Descargar JSON de Diseño",
                data=json_data,
                file_name="diseno_tablas.json",
                mime="application/json",
                type="primary",
                use_container_width=True
            )

    return format_config
