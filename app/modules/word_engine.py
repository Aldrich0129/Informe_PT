"""
Motor de generación de documentos Word usando python-docx.
Maneja reemplazo de variables, inserción de tablas y bloques condicionales.
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path
from typing import Dict, List, Any, Optional
import re
from copy import deepcopy


class WordEngine:
    """Motor para generar documentos Word desde plantillas."""

    def __init__(self, template_path: Path):
        """
        Inicializa el motor con una plantilla.

        Args:
            template_path: Ruta a la plantilla Word (.docx)
        """
        self.template_path = Path(template_path)

        if not self.template_path.exists():
            raise FileNotFoundError(f"Plantilla no encontrada: {template_path}")

        self.doc = Document(self.template_path)

    def replace_variables(self, context: dict):
        """
        Reemplaza todas las variables <<marcador>> en el documento manteniendo estilos.
        Si el valor está vacío, elimina el marcador del documento.

        Args:
            context: Diccionario con {marcador: valor}
        """
        # Filtrar valores vacíos - los marcadores con valores vacíos se eliminarán
        context_filtered = {}
        for marker, value in context.items():
            if value is None or value == "" or (isinstance(value, str) and not value.strip()):
                # Valor vacío, no reemplazar (se limpiará después)
                continue
            context_filtered[marker] = value

        # Reemplazar en párrafos
        for paragraph in self.doc.paragraphs:
            self._replace_in_paragraph(paragraph, context_filtered)

        # Reemplazar en tablas
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_in_paragraph(paragraph, context_filtered)

        # Reemplazar en headers y footers
        for section in self.doc.sections:
            # Header
            header = section.header
            for paragraph in header.paragraphs:
                self._replace_in_paragraph(paragraph, context_filtered)

            # Footer
            footer = section.footer
            for paragraph in footer.paragraphs:
                self._replace_in_paragraph(paragraph, context_filtered)

    def _replace_in_paragraph(self, paragraph, context: dict):
        """
        Reemplaza marcadores en un párrafo manteniendo el formato.

        Args:
            paragraph: Párrafo de python-docx
            context: Diccionario de reemplazos
        """
        # Construir el texto completo del párrafo
        full_text = paragraph.text

        # Buscar todos los marcadores en el texto
        for marker, value in context.items():
            if marker in full_text:
                # Reemplazar manteniendo el formato
                self._replace_marker_in_paragraph(paragraph, marker, str(value))
                full_text = paragraph.text  # Actualizar el texto completo

    def _replace_marker_in_paragraph(self, paragraph, marker: str, value: str):
        """
        Reemplaza un marcador específico manteniendo el formato del primer run.

        Args:
            paragraph: Párrafo de python-docx
            marker: Marcador a buscar (ej: "<<Variable>>")
            value: Valor de reemplazo
        """
        # Buscar el marcador en los runs
        full_text = ""
        runs_text = []

        for run in paragraph.runs:
            runs_text.append(run.text)
            full_text += run.text

        if marker not in full_text:
            return

        # Encontrar la posición del marcador
        marker_start = full_text.find(marker)
        marker_end = marker_start + len(marker)

        # Reconstruir el párrafo
        current_pos = 0
        new_runs = []

        for i, run in enumerate(paragraph.runs):
            run_text = runs_text[i]
            run_start = current_pos
            run_end = current_pos + len(run_text)

            # Verificar si este run contiene parte del marcador
            if run_end <= marker_start or run_start >= marker_end:
                # Este run no toca el marcador, mantenerlo
                new_runs.append((run_text, run))
            elif run_start <= marker_start and run_end >= marker_end:
                # El marcador está completamente dentro de este run
                before = run_text[:marker_start - run_start]
                after = run_text[marker_end - run_start:]
                new_runs.append((before + value + after, run))
            elif run_start < marker_start < run_end:
                # El marcador comienza en este run
                before = run_text[:marker_start - run_start]
                new_runs.append((before + value, run))
            elif run_start < marker_end <= run_end:
                # El marcador termina en este run
                after = run_text[marker_end - run_start:]
                if after:
                    new_runs.append((after, run))
            # else: este run está completamente dentro del marcador, se omite

            current_pos = run_end

        # Limpiar runs existentes
        for run in paragraph.runs:
            run.text = ""

        # Recrear runs con el nuevo texto
        if new_runs:
            for new_text, original_run in new_runs:
                if new_text:
                    # Copiar formato del run original
                    new_run = paragraph.add_run(new_text)
                    new_run.bold = original_run.bold
                    new_run.italic = original_run.italic
                    new_run.underline = original_run.underline
                    new_run.font.size = original_run.font.size
                    new_run.font.name = original_run.font.name
                    if original_run.font.color.rgb:
                        new_run.font.color.rgb = original_run.font.color.rgb

    def insert_tables(self, tables_data: dict, cfg_tab: dict, format_config: dict = None):
        """
        Inserta tablas en los marcadores correspondientes.

        Args:
            tables_data: Datos de las tablas construidas
            cfg_tab: Configuración de tablas
            format_config: Configuración de formato de tablas (opcional)
        """
        if format_config is None:
            format_config = {}

        for marker, table_data in tables_data.items():
            self._insert_table_at_marker(marker, table_data, format_config)

    def _insert_table_at_marker(self, marker: str, table_data: dict, format_config: dict = None):
        """
        Inserta una tabla en la posición del marcador.

        Args:
            marker: Marcador donde insertar la tabla
            table_data: Datos de la tabla (rows, columns, etc.)
            format_config: Configuración de formato de tablas (opcional)
        """
        if format_config is None:
            format_config = {}

        # Buscar el marcador en el documento
        for i, paragraph in enumerate(self.doc.paragraphs):
            if marker in paragraph.text:
                # Crear la tabla justo después de este párrafo
                self._create_table_after_paragraph(i, table_data, format_config)

                # Limpiar el marcador
                paragraph.text = paragraph.text.replace(marker, "")
                return

    def _create_table_after_paragraph(self, para_index: int, table_data: dict, format_config: dict = None):
        """
        Crea una tabla después de un párrafo específico.

        Args:
            para_index: Índice del párrafo
            table_data: Datos de la tabla
            format_config: Configuración de formato de tablas (opcional)
        """
        if format_config is None:
            format_config = {}

        columns = table_data.get("columns", [])
        rows = table_data.get("rows", [])
        footer_rows = table_data.get("footer_rows", [])
        headers = table_data.get("headers", {})

        if not columns or not rows:
            return

        # Calcular número de filas (header + data + footer)
        num_rows = 1 + len(rows) + len(footer_rows)
        num_cols = len(columns)

        # Encontrar el elemento XML del párrafo
        para = self.doc.paragraphs[para_index]
        para_element = para._element

        # Crear la tabla
        from docx.oxml import parse_xml
        from docx.table import Table

        table = self.doc.add_table(rows=num_rows, cols=num_cols)

        # Aplicar estilo de tabla según configuración
        show_borders = format_config.get("show_borders", True)
        if show_borders:
            # Intentar aplicar estilo con bordes
            try:
                table.style = 'Light Grid Accent 1'
            except KeyError:
                try:
                    table.style = 'Light Grid'
                except KeyError:
                    try:
                        table.style = 'Table Grid'
                    except KeyError:
                        pass
        else:
            # Sin bordes - usar estilo sin líneas
            try:
                table.style = 'Table Normal'
            except KeyError:
                pass

        # Insertar la tabla después del párrafo
        table_element = table._element
        para_element.addnext(table_element)

        # Obtener configuración de formato
        header_bg_color = format_config.get("header_bg_color", "#4472C4")
        header_text_color = format_config.get("header_text_color", "#FFFFFF")
        header_bold = format_config.get("header_bold", True)
        header_font_size = format_config.get("header_font_size", 11)
        data_font_size = format_config.get("data_font_size", 10)
        alternate_rows = format_config.get("alternate_rows", False)
        alternate_row_color = format_config.get("alternate_row_color", "#F2F2F2")
        border_color = format_config.get("border_color", "#000000")
        first_column_bold = format_config.get("first_column_bold", False)
        first_column_bg_color = format_config.get("first_column_bg_color", None)
        first_column_text_color = format_config.get("first_column_text_color", None)

        # Aplicar bordes personalizados a toda la tabla si show_borders es True
        if show_borders:
            self._apply_table_borders(table, border_color)
        else:
            # Eliminar todos los bordes si show_borders es False
            self._remove_table_borders(table)

        # Llenar encabezados
        header_row = table.rows[0]
        for j, col in enumerate(columns):
            cell = header_row.cells[j]
            header_text = col.get("header", "")

            # Reemplazar placeholders en el header si hay headers dinámicos
            if headers:
                for key, value in headers.items():
                    header_text = header_text.replace(f"{{{key}}}", str(value))

            cell.text = header_text

            # Aplicar formato de encabezado personalizado
            self._apply_cell_shading(cell, header_bg_color)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if header_bold:
                        run.bold = True
                    run.font.size = Pt(header_font_size)
                    # Aplicar color de texto del encabezado
                    run.font.color.rgb = self._hex_to_rgb(header_text_color)

        # Llenar filas de datos
        for i, row_data in enumerate(rows):
            row = table.rows[i + 1]

            # Aplicar color alternado a filas pares si está configurado
            if alternate_rows and i % 2 == 0:
                for cell in row.cells:
                    self._apply_cell_shading(cell, alternate_row_color)

            for j, col in enumerate(columns):
                cell = row.cells[j]
                col_id = col["id"]
                value = row_data.get(col_id, "")

                # Formatear según el tipo
                col_type = col.get("type", "text")
                formatted_value = self._format_cell_value(value, col_type)

                cell.text = formatted_value

                # Aplicar estilo de primera columna si está configurado y es la primera columna
                is_first_column = (j == 0)
                if is_first_column:
                    # Aplicar color de fondo si está configurado
                    if first_column_bg_color:
                        self._apply_cell_shading(cell, first_column_bg_color)

                # Aplicar tamaño de fuente a las celdas de datos
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(data_font_size)

                        # Aplicar estilos de primera columna
                        if is_first_column:
                            if first_column_bold:
                                run.bold = True
                            if first_column_text_color:
                                run.font.color.rgb = self._hex_to_rgb(first_column_text_color)

        # Llenar filas de footer
        if footer_rows:
            for i, footer_data in enumerate(footer_rows):
                row_index = 1 + len(rows) + i
                row = table.rows[row_index]

                for j, col in enumerate(columns):
                    cell = row.cells[j]
                    col_id = col["id"]
                    value = footer_data.get(col_id, "")

                    col_type = col.get("type", "text")
                    formatted_value = self._format_cell_value(value, col_type)

                    cell.text = formatted_value

                    # Aplicar formato de footer (negrita y tamaño de fuente)
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True
                            run.font.size = Pt(data_font_size)

    def _format_cell_value(self, value: Any, col_type: str) -> str:
        """
        Formatea el valor de una celda según su tipo.

        Args:
            value: Valor a formatear
            col_type: Tipo de columna

        Returns:
            Valor formateado como string
        """
        if value is None or value == "":
            return ""

        try:
            if col_type == "percent":
                # Convertir a número si es posible
                if isinstance(value, str):
                    # Eliminar el símbolo % si ya existe
                    value = value.replace('%', '').strip()
                    try:
                        value = float(value)
                    except ValueError:
                        return str(value)

                if isinstance(value, (int, float)):
                    return f"{value:.2f}%"
                return str(value)
            elif col_type == "number":
                if isinstance(value, (int, float)):
                    return f"{value:,.2f}"
                elif isinstance(value, str):
                    try:
                        value = float(value)
                        return f"{value:,.2f}"
                    except ValueError:
                        return str(value)
                return str(value)
            elif col_type == "integer":
                if isinstance(value, (int, float)):
                    return str(int(value))
                elif isinstance(value, str):
                    try:
                        value = int(float(value))
                        return str(value)
                    except ValueError:
                        return str(value)
                return str(value)
            else:
                return str(value)
        except:
            return str(value)

    def insert_conditional_blocks(self, docs_to_insert: list, config_dir: Path):
        """
        Inserta bloques condicionales de otros documentos Word.

        Args:
            docs_to_insert: Lista de {marker, file} a insertar
            config_dir: Directorio base para los archivos Word
        """
        for doc_info in docs_to_insert:
            marker = doc_info["marker"]
            word_file = doc_info["file"]

            file_path = config_dir.parent / word_file

            if file_path.exists():
                self._insert_document_at_marker(marker, file_path)
            else:
                print(f"Advertencia: Archivo no encontrado: {file_path}")
                # Limpiar el marcador
                self.replace_variables({marker: ""})

    def _insert_document_at_marker(self, marker: str, doc_path: Path):
        """
        Inserta el contenido de otro documento en la posición del marcador.

        Args:
            marker: Marcador donde insertar
            doc_path: Ruta al documento a insertar
        """
        # Cargar el documento a insertar
        doc_to_insert = Document(doc_path)

        # Buscar el marcador
        for i, paragraph in enumerate(self.doc.paragraphs):
            if marker in paragraph.text:
                # Limpiar el marcador
                paragraph.text = paragraph.text.replace(marker, "")

                # Insertar el contenido del documento
                para_element = paragraph._element

                # Copiar párrafos del documento a insertar
                for insert_para in doc_to_insert.paragraphs:
                    # Crear un nuevo párrafo con el mismo formato
                    new_para = deepcopy(insert_para._element)
                    para_element.addnext(new_para)
                    para_element = new_para

                # Copiar tablas del documento a insertar
                for insert_table in doc_to_insert.tables:
                    new_table = deepcopy(insert_table._element)
                    para_element.addnext(new_table)
                    para_element = new_table

                return

    def clean_unused_markers(self):
        """
        Limpia todos los marcadores no utilizados del documento.
        Si un párrafo contiene SOLO un marcador (o marcador con puntuación/numeración),
        elimina el párrafo completo. Si hay más contenido, solo elimina el marcador.
        """
        marker_pattern = re.compile(r'<<[^>]+>>')

        # Limpiar en párrafos - eliminar líneas completas si solo contienen marcadores
        paragraphs_to_delete = []
        for i, paragraph in enumerate(self.doc.paragraphs):
            if marker_pattern.search(paragraph.text):
                # Verificar si el párrafo solo contiene marcador y elementos comunes (puntos, números, guiones, espacios)
                text_without_markers = marker_pattern.sub('', paragraph.text).strip()
                # Eliminar también puntuación común, números, guiones, viñetas
                text_cleaned = re.sub(r'^[\d\.\-\)\(\s•·◦▪▫○●\*]+$', '', text_without_markers)

                if not text_cleaned:
                    # El párrafo solo contiene marcador + elementos decorativos, marcarlo para eliminación
                    paragraphs_to_delete.append(paragraph)
                else:
                    # Hay contenido real además del marcador, solo eliminar el marcador
                    paragraph.text = marker_pattern.sub('', paragraph.text)

        # Eliminar los párrafos marcados
        for paragraph in paragraphs_to_delete:
            p_element = paragraph._element
            p_element.getparent().remove(p_element)

        # Limpiar en tablas - solo eliminar el marcador, no la celda
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if marker_pattern.search(paragraph.text):
                            paragraph.text = marker_pattern.sub('', paragraph.text)

    def _hex_to_rgb(self, hex_color: str) -> RGBColor:
        """
        Convierte un color hexadecimal a RGBColor.

        Args:
            hex_color: Color en formato hexadecimal (ej: "#4472C4" o "4472C4")

        Returns:
            Objeto RGBColor
        """
        # Eliminar el '#' si existe
        hex_color = hex_color.lstrip('#')

        # Convertir a RGB
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)

        return RGBColor(r, g, b)

    def _apply_cell_shading(self, cell, hex_color: str):
        """
        Aplica un color de fondo a una celda de tabla.

        Args:
            cell: Celda de la tabla
            hex_color: Color en formato hexadecimal (ej: "#4472C4")
        """
        # Eliminar el '#' si existe
        hex_color = hex_color.lstrip('#')

        # Obtener el elemento XML de la celda
        cell_properties = cell._element.get_or_add_tcPr()

        # Crear elemento de sombreado
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), hex_color)

        # Añadir el sombreado a las propiedades de la celda
        cell_properties.append(shading)

    def _apply_table_borders(self, table, hex_color: str):
        """
        Aplica bordes con color personalizado a todas las celdas de una tabla.

        Args:
            table: Tabla de python-docx
            hex_color: Color en formato hexadecimal (ej: "#000000")
        """
        # Eliminar el '#' si existe
        hex_color = hex_color.lstrip('#')

        # Aplicar bordes a cada celda
        for row in table.rows:
            for cell in row.cells:
                tc = cell._element
                tcPr = tc.get_or_add_tcPr()

                # Crear bordes
                tcBorders = OxmlElement('w:tcBorders')

                # Definir cada borde (top, left, bottom, right)
                for border_name in ['top', 'left', 'bottom', 'right']:
                    border = OxmlElement(f'w:{border_name}')
                    border.set(qn('w:val'), 'single')
                    border.set(qn('w:sz'), '4')  # Tamaño del borde
                    border.set(qn('w:space'), '0')
                    border.set(qn('w:color'), hex_color)
                    tcBorders.append(border)

                tcPr.append(tcBorders)

    def _remove_table_borders(self, table):
        """
        Elimina todos los bordes de una tabla.

        Args:
            table: Tabla de python-docx
        """
        # Eliminar bordes de cada celda
        for row in table.rows:
            for cell in row.cells:
                tc = cell._element
                tcPr = tc.get_or_add_tcPr()

                # Crear bordes vacíos (sin líneas)
                tcBorders = OxmlElement('w:tcBorders')

                # Definir cada borde como "none"
                for border_name in ['top', 'left', 'bottom', 'right']:
                    border = OxmlElement(f'w:{border_name}')
                    border.set(qn('w:val'), 'none')
                    border.set(qn('w:sz'), '0')
                    border.set(qn('w:space'), '0')
                    border.set(qn('w:color'), 'auto')
                    tcBorders.append(border)

                tcPr.append(tcBorders)

    def clean_empty_paragraphs(self):
        """
        Elimina párrafos vacíos del documento.
        """
        # Nota: Eliminar párrafos en python-docx es complicado
        # Por ahora, solo los marcamos como vacíos
        for paragraph in self.doc.paragraphs:
            if not paragraph.text.strip():
                # Opcionalmente reducir el espaciado
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)

    def process_table_of_contents(self):
        """
        Procesa el índice (tabla de contenidos) del documento usando marcadores numéricos.

        Busca contenidos entre <<Indice>> y <<fin Indice>>, extrae los marcadores numéricos
        (<<1>>, <<2>>, etc.) de cada entrada del índice, localiza estos marcadores en el
        documento, inserta saltos de página antes de ellos, calcula los números de página
        y actualiza el índice.

        Al finalizar, elimina todos los marcadores numéricos del documento.
        """
        # Buscar los marcadores de inicio y fin del índice
        toc_start_idx = None
        toc_end_idx = None

        for i, paragraph in enumerate(self.doc.paragraphs):
            text = paragraph.text.strip()
            if "<<Indice>>" in text:
                toc_start_idx = i
            elif "<<fin Indice>>" in text:
                toc_end_idx = i
                break

        # Si no se encuentran los marcadores, no hacer nada
        if toc_start_idx is None or toc_end_idx is None:
            return

        # Extraer las entradas del índice con sus marcadores numéricos
        toc_entries = []  # Lista de {paragraph, title, marker}

        marker_pattern = re.compile(r'<<(\d+)>>')

        for i in range(toc_start_idx + 1, toc_end_idx):
            para = self.doc.paragraphs[i]
            text = para.text.strip()

            if text:  # Solo procesar párrafos no vacíos
                # Buscar el marcador numérico en el texto
                match = marker_pattern.search(text)
                if match:
                    marker_num = match.group(1)
                    marker = f"<<{marker_num}>>"

                    # Extraer el título (texto antes del marcador)
                    title = marker_pattern.sub('', text).strip()

                    toc_entries.append({
                        'paragraph': para,
                        'title': title,
                        'marker': marker,
                        'marker_num': marker_num
                    })

        # Si no hay entradas con marcadores, salir
        if not toc_entries:
            # Limpiar solo los marcadores de inicio y fin
            start_para = self.doc.paragraphs[toc_start_idx]
            start_para.text = start_para.text.replace("<<Indice>>", "").strip()
            end_para = self.doc.paragraphs[toc_end_idx]
            end_para.text = end_para.text.replace("<<fin Indice>>", "").strip()
            return

        # Fase 1: Buscar cada marcador en el documento e insertar saltos de página
        for entry in toc_entries:
            marker = entry['marker']
            self._insert_page_break_before_marker(marker, toc_end_idx)

        # Fase 2: Calcular números de página para cada marcador
        marker_to_page = {}

        for entry in toc_entries:
            marker = entry['marker']
            page_num = self._find_marker_page_number(marker, toc_end_idx)
            if page_num is not None:
                marker_to_page[marker] = page_num

        # Fase 3: Actualizar el índice con los números de página
        for entry in toc_entries:
            para = entry['paragraph']
            title = entry['title']
            marker = entry['marker']

            if marker in marker_to_page:
                page_num = marker_to_page[marker]

                # Eliminar cualquier número de página existente al final
                clean_title = re.sub(r'[\.\s]+\d+$', '', title).strip()

                # Calcular el número de puntos necesarios
                dots_length = max(3, 80 - len(clean_title) - len(str(page_num)) - 2)
                dots = '.' * dots_length

                # Actualizar el párrafo (sin el marcador)
                para.text = f"{clean_title} {dots} {page_num}"
            else:
                # Si no se encontró el marcador, al menos limpiar el marcador del título
                para.text = title

        # Fase 4: Eliminar todos los marcadores numéricos del documento
        self._remove_numeric_markers()

        # Eliminar el marcador <<Indice>>
        start_para = self.doc.paragraphs[toc_start_idx]
        start_para.text = start_para.text.replace("<<Indice>>", "").strip()

        # Eliminar el marcador <<fin Indice>>
        end_para = self.doc.paragraphs[toc_end_idx]
        end_para.text = end_para.text.replace("<<fin Indice>>", "").strip()

        # Asegurar que el índice empiece en una nueva página
        if toc_start_idx > 0:
            prev_para = self.doc.paragraphs[toc_start_idx - 1]
            run = prev_para.add_run()
            run.add_break(WD_BREAK.PAGE)

    def _insert_page_break_before_marker(self, marker: str, toc_end_idx: int):
        """
        Busca un marcador numérico en el documento e inserta un salto de página antes de él.

        Args:
            marker: Marcador numérico a buscar (ej: "<<1>>")
            toc_end_idx: Índice del final del índice (para empezar búsqueda después)
        """
        # Buscar el marcador en el documento (después del índice)
        for i in range(toc_end_idx + 1, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]

            if marker in para.text:
                # Verificar si este párrafo ya tiene un salto de página al inicio
                has_page_break = False
                if para.runs:
                    for run in para.runs:
                        if self._has_page_break(run):
                            has_page_break = True
                            break

                # Si no tiene salto de página, insertar uno al principio del párrafo
                if not has_page_break:
                    # Insertar un salto de página al principio del párrafo
                    # Crear un nuevo run al principio
                    if para.runs:
                        first_run = para.runs[0]
                        # Insertar el salto de página en el primer run
                        first_run.add_break(WD_BREAK.PAGE)
                    else:
                        # Si no hay runs, crear uno nuevo
                        run = para.add_run()
                        run.add_break(WD_BREAK.PAGE)

                # Solo procesar la primera ocurrencia
                break

    def _find_marker_page_number(self, marker: str, start_search_idx: int) -> Optional[int]:
        """
        Encuentra el número de página donde aparece un marcador numérico.

        Args:
            marker: Marcador numérico a buscar (ej: "<<1>>")
            start_search_idx: Índice desde donde empezar la búsqueda (después del índice)

        Returns:
            Número de página donde se encuentra el marcador, o None si no se encuentra
        """
        # Contador de páginas
        page_count = 1

        # Buscar el marcador en el documento (después del índice)
        for i in range(start_search_idx + 1, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]

            # Verificar si hay un salto de página en este párrafo
            for run in para.runs:
                if self._has_page_break(run):
                    page_count += 1

            # Verificar si este párrafo contiene el marcador
            if marker in para.text:
                return page_count

        # Si no se encuentra el marcador, devolver None
        return None

    def _remove_numeric_markers(self):
        """
        Elimina todos los marcadores numéricos (<<1>>, <<2>>, etc.) del documento.
        """
        marker_pattern = re.compile(r'<<\d+>>')

        # Eliminar de párrafos
        for paragraph in self.doc.paragraphs:
            if marker_pattern.search(paragraph.text):
                paragraph.text = marker_pattern.sub('', paragraph.text)

        # Eliminar de tablas
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if marker_pattern.search(paragraph.text):
                            paragraph.text = marker_pattern.sub('', paragraph.text)

        # Eliminar de headers y footers
        for section in self.doc.sections:
            # Header
            header = section.header
            for paragraph in header.paragraphs:
                if marker_pattern.search(paragraph.text):
                    paragraph.text = marker_pattern.sub('', paragraph.text)

            # Footer
            footer = section.footer
            for paragraph in footer.paragraphs:
                if marker_pattern.search(paragraph.text):
                    paragraph.text = marker_pattern.sub('', paragraph.text)

    def _has_page_break(self, run) -> bool:
        """
        Verifica si un run contiene un salto de página.

        Args:
            run: Run de python-docx

        Returns:
            True si el run contiene un salto de página, False en caso contrario
        """
        # Obtener el elemento XML del run
        run_element = run._element

        # Buscar elementos de salto de página (w:br con w:type="page")
        for br in run_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br'):
            break_type = br.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
            if break_type == 'page':
                return True

        return False

    def save(self, output_path: Path):
        """
        Guarda el documento generado.

        Args:
            output_path: Ruta donde guardar el documento
        """
        self.doc.save(output_path)

    def get_document_bytes(self) -> bytes:
        """
        Obtiene el documento como bytes (para descarga en Streamlit).

        Returns:
            Bytes del documento
        """
        from io import BytesIO

        buffer = BytesIO()
        self.doc.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()
