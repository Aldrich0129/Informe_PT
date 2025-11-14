"""
Aplicación principal para generar informes de Precios de Transferencia.

Punto de entrada: ejecutar con `streamlit run app.py`
"""
import sys
from pathlib import Path

# Añadir el directorio app al path
app_dir = Path(__file__).parent
sys.path.insert(0, str(app_dir))

import streamlit as st
from modules.config_loader import ConfigLoader
from modules.utils import build_full_context, validate_inputs
from modules.tables import TableBuilder
from modules.word_engine import WordEngine
from modules.simple_vars import validate_simple_vars
from modules.conditions import validate_conditions
from ui.main_ui import (
    render_main_ui,
    render_generation_section,
    show_validation_errors,
    show_success_message,
    show_processing_spinner
)


def main():
    """Función principal de la aplicación."""

    # Cargar configuraciones
    try:
        config_dir = app_dir / "config"
        loader = ConfigLoader(config_dir)
        cfg_simple, cfg_cond, cfg_tab = loader.load_all_configs()

    except Exception as e:
        st.error(f"❌ Error al cargar las configuraciones: {e}")
        st.stop()

    # Renderizar UI principal
    simple_inputs, condition_inputs, table_inputs, table_format_config = render_main_ui(
        cfg_simple, cfg_cond, cfg_tab
    )

    # Sección de generación
    if render_generation_section():
        # Validar entradas
        errors = []

        # Validar variables simples
        errors.extend(validate_simple_vars(cfg_simple, simple_inputs))

        # Validar condiciones
        errors.extend(validate_conditions(cfg_cond, condition_inputs))

        if errors:
            show_validation_errors(errors)
            st.stop()

        # Generar el documento
        try:
            with show_processing_spinner("Generando informe..."):
                # 1. Construir tablas
                table_builder = TableBuilder(cfg_tab, simple_inputs)
                tables_data = table_builder.build_all_tables(table_inputs)

                # 2. Construir contexto completo
                context, docs_to_insert = build_full_context(
                    cfg_simple,
                    cfg_cond,
                    cfg_tab,
                    simple_inputs,
                    condition_inputs,
                    table_inputs
                )

                # 3. Cargar plantilla y crear motor de Word
                template_path = config_dir / "Plantilla.docx"

                if not template_path.exists():
                    st.error(f"❌ Plantilla no encontrada: {template_path}")
                    st.stop()

                engine = WordEngine(template_path)

                # 4. Reemplazar variables simples
                engine.replace_variables(context)

                # 5. Insertar tablas
                engine.insert_tables(tables_data, cfg_tab, table_format_config)

                # 6. Insertar bloques condicionales
                engine.insert_conditional_blocks(docs_to_insert, config_dir)

                # 7. Procesar índice (tabla de contenidos)
                engine.process_table_of_contents()

                # 8. Limpieza final
                engine.clean_unused_markers()
                engine.clean_empty_paragraphs()

                # 9. Obtener el documento como bytes
                doc_bytes = engine.get_document_bytes()

                # 10. Mostrar botón de descarga
                show_success_message()

                nombre_empresa = simple_inputs.get("nombre_compania", "Empresa")
                ejercicio = simple_inputs.get("ejercicio_completo", "2023")

                # Limpiar nombre para archivo
                nombre_archivo = f"Informe_PT_{nombre_empresa.replace(' ', '_')}_{ejercicio}.docx"

                st.download_button(
                    label="📥 Descargar Informe",
                    data=doc_bytes,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )

                st.balloons()

        except Exception as e:
            st.error(f"❌ Error al generar el informe: {e}")
            st.exception(e)


if __name__ == "__main__":
    main()
