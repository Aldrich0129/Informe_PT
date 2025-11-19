# Generador de Informes de Precios de Transferencia

Aplicación completa basada en Streamlit + Python para generar informes Word a partir de plantillas usando marcadores.

## 📋 Características

- ✅ Carga dinámica de configuraciones desde archivos YAML
- ✅ UI dinámica generada automáticamente desde YAML
- ✅ Reemplazo de variables manteniendo formato Word
- ✅ Inserción de tablas dinámicas con formato
- ✅ Bloques condicionales (insertar archivos Word según condiciones)
- ✅ Limpieza automática de marcadores
- ✅ Soporte para múltiples tipos de datos (texto, número, porcentaje, email, etc.)
- ✅ Tablas con filas dinámicas y cálculos automáticos
- ✅ **Conservación automática de imágenes de fondo** en portada y páginas finales
- ✅ **Exportación dual**: Descarga en formato Word (.docx) y PDF

## 🏗️ Arquitectura

```
/app
   /config
      variables_simples.yaml
      variables_condicionales.yaml
      tablas.yaml
      Plantilla.docx

   /modules
      config_loader.py       # Carga y valida YAMLs
      simple_vars.py         # Manejo de variables simples
      conditions.py          # Manejo de condiciones
      tables.py              # Construcción de tablas
      word_engine.py         # Motor de generación Word
      utils.py               # Utilidades y construcción de contexto

   /ui
      main_ui.py             # UI principal y orquestación
      sections_simple_vars.py    # Sección de variables simples
      sections_conditions.py     # Sección de condiciones
      sections_tables.py         # Sección de tablas

   /condiciones               # Archivos Word para bloques condicionales
      (archivos .docx)

   app.py                     # Punto de entrada
   requirements.txt
   README.md
```

## 🚀 Instalación

1. **Instalar dependencias:**

```bash
pip install -r requirements.txt
```

2. **Instalar Pandoc (requerido para exportación a PDF):**

**En Linux (Ubuntu/Debian):**
```bash
sudo apt-get update
sudo apt-get install -y pandoc texlive-xetex texlive-fonts-recommended
```

**En macOS:**
```bash
brew install pandoc
brew install --cask basictex
```

**En Windows:**
- Descargar e instalar Pandoc desde: https://pandoc.org/installing.html
- Descargar e instalar MiKTeX desde: https://miktex.org/download

3. **Verificar estructura:**

Asegúrate de que los archivos YAML y la plantilla estén en `/config`:
- variables_simples.yaml
- variables_condicionales.yaml
- tablas.yaml
- Plantilla.docx

4. **Archivos condicionales:**

Coloca los archivos Word de bloques condicionales en `/condiciones`:
- comentario_inicial_formal.docx
- desarrollo_comentario_formal.docx
- etc.

## ▶️ Ejecución

```bash
streamlit run app.py
```

La aplicación se abrirá en tu navegador en `http://localhost:8501`

## 📖 Uso

1. **Variables Simples:** Completa los datos generales del informe
2. **Condiciones:** Selecciona qué bloques incluir (Sí/No)
3. **Tablas:** Rellena todas las tablas requeridas
4. **Generar:** Haz clic en "Generar Informe Word"
5. **Descargar:** Elige entre dos formatos:
   - 📥 **Word (.docx)**: Formato editable para modificaciones posteriores
   - 📑 **PDF**: Formato final para presentación y distribución

## 🖼️ Imágenes de Fondo en Plantillas

El sistema preserva automáticamente todas las imágenes de fondo de la plantilla Word original:

- **Imágenes de portada**: Se mantienen intactas
- **Imágenes de página final**: Se conservan automáticamente
- **Imágenes en headers/footers**: Se preservan sin modificación
- **Formas y elementos gráficos**: Permanecen en sus posiciones originales

**Cómo agregar imágenes de fondo a la plantilla:**

1. Abre `Plantilla.docx` en Microsoft Word
2. Ve a `Diseño` → `Marca de agua` → `Marca de agua personalizada`
3. O inserta una imagen y configúrala como "Detrás del texto"
4. Guarda la plantilla

El motor de generación (`python-docx`) solo modifica el contenido de texto y tablas, manteniendo intactos todos los elementos gráficos de la plantilla original.

## 🔧 Tipos de Tablas Soportadas

- **TNMM Global:** Análisis indirecto con rangos estadísticos
- **TNMM por Operación:** Análisis por cada operación vinculada
- **Partidas Contables:** Con cabeceras dinámicas y cálculos automáticos
- **Operaciones Vinculadas:** Tabla con filas dinámicas y totales
- **Cumplimiento Formal:** Local File y Master File (inicial y detallado)
- **Riesgos PT:** Tabla de evaluación de riesgos

## 📝 Personalización

Para personalizar la aplicación:

1. **Modificar variables:** Edita `config/variables_simples.yaml`
2. **Añadir condiciones:** Edita `config/variables_condicionales.yaml`
3. **Configurar tablas:** Edita `config/tablas.yaml`
4. **Plantilla Word:** Actualiza `config/Plantilla.docx` con tus marcadores

📚 ¿Necesitas más detalle? Consulta la guía paso a paso en [`GUIA_PERSONALIZACION_YAML.md`](GUIA_PERSONALIZACION_YAML.md) para ver ejemplos prácticos de cómo extender variables, condiciones y tablas.

## 🔍 Marcadores en la Plantilla

Los marcadores siguen el formato: `<<Nombre del Marcador>>`

Ejemplos:
- `<<Ejercicio completo>>`
- `<<Nombre de la Compañía>>`
- `<<Tabla operaciones vinculadas>>`
- `<<Comentario inicial formal>>`

## ⚙️ Tecnologías Utilizadas

- **Streamlit:** Framework de UI
- **python-docx:** Manipulación de documentos Word
- **PyYAML:** Parsing de configuraciones
- **Pandas:** Manipulación de datos tabulares
- **pypandoc:** Conversión de documentos a PDF
- **Pandoc + XeLaTeX:** Motor de conversión a PDF con soporte Unicode

## 📄 Licencia

Desarrollado para Mazars - Informes de Precios de Transferencia

## 🐛 Solución de Problemas

**Error: Plantilla no encontrada**
- Verifica que `Plantilla.docx` esté en `/config`

**Error: Archivo condicional no encontrado**
- Verifica que los archivos .docx de condiciones estén en `/condiciones`

**Error al cargar YAML**
- Verifica la sintaxis YAML
- Asegúrate de que los archivos tengan codificación UTF-8

**Tablas no se insertan correctamente**
- Verifica que los marcadores en la plantilla coincidan exactamente con los del YAML
- Los marcadores son case-sensitive

**No se puede generar PDF**
- Asegúrate de que Pandoc esté instalado: `pandoc --version`
- Instala XeLaTeX: `sudo apt-get install texlive-xetex` (Linux)
- Si falla, descarga el archivo Word y conviértelo manualmente

**Las imágenes de fondo no aparecen**
- Las imágenes deben estar en la plantilla original `Plantilla.docx`
- Verifica que las imágenes estén configuradas como "Detrás del texto" en Word
- python-docx preserva automáticamente las imágenes existentes en la plantilla
