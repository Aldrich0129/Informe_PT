# Guía Completa de Personalización YAML

## Índice
1. [Introducción](#introducción)
2. [variables_simples.yaml](#variables_simplesyaml)
3. [variables_condicionales.yaml](#variables_condicionalesyaml)
4. [tablas.yaml](#tablasyaml)
5. [Ejemplos Prácticos](#ejemplos-prácticos)
6. [Solución de Problemas](#solución-de-problemas)

---

## Introducción

Este documento proporciona una guía completa para personalizar el contenido de la aplicación de generación de informes de Precios de Transferencia mediante la modificación de tres archivos YAML principales:

- **variables_simples.yaml**: Define variables de texto y números que se reemplazan en la plantilla
- **variables_condicionales.yaml**: Define bloques de contenido condicionales (Sí/No)
- **tablas.yaml**: Define la estructura y formato de las tablas dinámicas

---

## variables_simples.yaml

### Estructura General

Este archivo contiene dos secciones principales:
1. **simple_variables**: Variables simples de texto, número, porcentaje, etc.
2. **operations**: Definición de operaciones vinculadas

### 1. Variables Simples

#### Formato Básico

```yaml
simple_variables:
  - id: nombre_variable           # Identificador único (sin espacios)
    label: "Etiqueta en la UI"    # Texto que verá el usuario
    marker: "<<Marcador>>"        # Marcador en la plantilla Word
    type: "text"                  # Tipo de dato
```

#### Tipos de Datos Disponibles

| Tipo | Descripción | Ejemplo de Uso |
|------|-------------|----------------|
| `text` | Texto corto (una línea) | Nombre de empresa, año |
| `long_text` | Texto largo (textarea) | Descripciones, comentarios |
| `number` | Número decimal | Cantidades, importes |
| `percent` | Porcentaje | 15.5% |
| `email` | Dirección de correo electrónico | usuario@empresa.com |
| `integer` | Número entero | 2023, 100 |

#### Campos de Configuración

| Campo | Requerido | Descripción |
|-------|-----------|-------------|
| `id` | Sí | Identificador único de la variable (usado internamente) |
| `label` | Sí | Etiqueta que aparece en la interfaz de usuario |
| `marker` | No | Marcador en la plantilla Word. Si es `null`, la variable solo se usa internamente |
| `type` | Sí | Tipo de dato (ver tabla anterior) |

#### Ejemplo Completo

```yaml
simple_variables:
  # Variable de texto simple
  - id: nombre_compania
    label: "Nombre de la Compañía"
    marker: "<<Nombre de la Compañía>>"
    type: "text"

  # Variable de texto largo
  - id: actividad_principal
    label: "Actividad principal de la Compañía"
    marker: "<<Actividad principal>>"
    type: "long_text"

  # Variable de porcentaje
  - id: info_pt_ingresos_pct
    label: "Porcentaje de ingresos vinculados sobre la actividad de la Compañía"
    marker: "<<Tablas Información PT Compañía1>>"
    type: "percent"

  # Variable sin marcador (solo para uso interno)
  - id: ejercicio_anterior
    label: "Ejercicio comparativo anterior (p.ej. 2022)"
    marker: null
    type: "text"
```

#### Cómo Agregar una Nueva Variable Simple

1. **Añadir la definición** al archivo `variables_simples.yaml`:
```yaml
- id: nueva_variable
  label: "Nueva Variable"
  marker: "<<Nueva Variable>>"
  type: "text"
```

2. **Agregar el marcador** en la plantilla Word donde desee que aparezca:
```
<<Nueva Variable>>
```

3. La variable aparecerá automáticamente en la interfaz de usuario

### 2. Operaciones Vinculadas

#### Estructura de Operations

```yaml
operations:
  base_text_marker_pattern: "<<Operación {n}>>"
  base_tnmm_table_marker_pattern: "<<Tabla Operación {n}>>"
  max_operations: 10

  items:
    - id: operacion_1
      index: 1
      default_label: "Operación 1"
      text_marker: "<<Operación 1>>"
      tnmm_table_marker: "<<Tabla Operación 1>>"
      vinculadas_row_hint: 1
```

#### Campos de Configuración de Operations

| Campo | Descripción |
|-------|-------------|
| `base_text_marker_pattern` | Patrón de marcador de texto para operaciones |
| `base_tnmm_table_marker_pattern` | Patrón de marcador de tabla TNMM |
| `max_operations` | Número máximo de operaciones permitidas |
| `id` | Identificador de la operación |
| `index` | Número de la operación |
| `default_label` | Etiqueta por defecto |
| `text_marker` | Marcador de texto en la plantilla |
| `tnmm_table_marker` | Marcador de tabla TNMM |
| `vinculadas_row_hint` | Posición sugerida en tabla de operaciones vinculadas |

#### Cómo Agregar una Nueva Operación

Para agregar una operación número 11:

```yaml
- id: operacion_11
  index: 11
  default_label: "Operación 11"
  text_marker: "<<Operación 11>>"
  tnmm_table_marker: "<<Tabla Operación 11>>"
  vinculadas_row_hint: 11
```

**Importante**: También debe actualizar `max_operations: 11`

---

## variables_condicionales.yaml

### Estructura General

Define bloques de contenido que se insertan condicionalmente basándose en una respuesta Sí/No.

#### Formato Básico

```yaml
conditions:
  - id: nombre_condicion
    label: "Etiqueta de la condición"
    marker: "<<Marcador>>"
    type: "boolean_include_doc"
    question: "¿Incluir este contenido?"
    yes_no_values: ["Sí", "No"]
    word_file: "ruta/al/archivo.docx"
```

#### Campos de Configuración

| Campo | Requerido | Descripción |
|-------|-----------|-------------|
| `id` | Sí | Identificador único |
| `label` | Sí | Etiqueta en la UI |
| `marker` | Sí | Marcador en la plantilla donde se insertará el contenido |
| `type` | Sí | Siempre debe ser `"boolean_include_doc"` |
| `question` | Sí | Pregunta que aparecerá en la UI |
| `yes_no_values` | Sí | Valores de las opciones (normalmente `["Sí", "No"]`) |
| `word_file` | Sí | Ruta relativa al archivo Word con el contenido a insertar |

#### Ejemplo Completo

```yaml
conditions:
  - id: comentario_inicial_formal
    label: "Comentario inicial formal"
    marker: "<<Comentario inicial formal>>"
    type: "boolean_include_doc"
    question: "¿Incluir comentario inicial formal?"
    yes_no_values: ["Sí", "No"]
    word_file: "condiciones/comentario_inicial_formal.docx"
```

#### Cómo Agregar una Nueva Condición

1. **Crear el archivo Word** con el contenido en la carpeta `app/condiciones/`:
   - Ejemplo: `app/condiciones/nuevo_comentario.docx`

2. **Agregar la definición** en `variables_condicionales.yaml`:
```yaml
- id: nuevo_comentario
  label: "Nuevo Comentario"
  marker: "<<Nuevo Comentario>>"
  type: "boolean_include_doc"
  question: "¿Incluir nuevo comentario?"
  yes_no_values: ["Sí", "No"]
  word_file: "condiciones/nuevo_comentario.docx"
```

3. **Agregar el marcador** en la plantilla Word:
```
<<Nuevo Comentario>>
```

4. La condición aparecerá automáticamente como una pregunta Sí/No en la UI

#### Funcionamiento

- Si el usuario selecciona **"Sí"**: El contenido del archivo Word se inserta en el marcador
- Si el usuario selecciona **"No"**: El marcador se elimina del documento

---

## tablas.yaml

### Estructura General

Define la estructura, columnas y formato de todas las tablas dinámicas del informe.

#### Formato Básico de una Tabla

```yaml
tables:
  nombre_tabla:
    marker: "<<Marcador Tabla>>"
    title: "Título de la Tabla"
    ui_component: "grid"
    remove_empty_rows: false
    columns:
      - id: columna_id
        header: "Encabezado"
        type: "text"
        editable: true
    rows:
      - id: fila_id
        label: "Etiqueta de Fila"
```

### Tipos de Tablas (ui_component)

| Tipo | Descripción | Uso |
|------|-------------|-----|
| `grid` | Tabla de múltiples filas fijas | Tablas estándar con filas predefinidas |
| `single_row_grid` | Tabla de una sola fila | Análisis TNMM global |
| `grid_dynamic_rows` | Tabla con filas dinámicas | Operaciones vinculadas (usuario puede agregar/quitar filas) |

### Tipos de Columnas

| Tipo | Descripción | Formato de Salida |
|------|-------------|-------------------|
| `text` | Texto libre | Sin formato especial |
| `number` | Número decimal | 1,234.56 |
| `integer` | Número entero | 1234 |
| `percent` | Porcentaje | 15.50% |
| `choice` | Lista de opciones | Una de las opciones predefinidas |

### Configuración de Columnas

#### Columna Simple

```yaml
- id: nombre_columna
  header: "Encabezado"
  type: "text"
  editable: true
```

#### Columna con Opciones (Choice)

```yaml
- id: cumplimiento
  header: "Cumplimiento"
  type: "choice"
  editable: true
  options: ["Sí", "No", "Ver comentario"]
```

#### Columna Condicional

```yaml
- id: comentario
  header: "Comentario"
  type: "text"
  editable: true
  conditional_on:
    field: "cumplimiento"
    value: "Ver comentario"
```

Esta columna solo se muestra si el campo `cumplimiento` tiene el valor "Ver comentario".

#### Columna con Encabezado Dinámico

```yaml
- id: ejercicio_actual
  header_template: "Ejercicio {ejercicio_actual}"
  type: "number"
  editable: true
```

El `{ejercicio_actual}` se reemplaza con el valor de la variable correspondiente.

### Configuración de Filas

#### Fila Manual (Input Mode)

```yaml
- id: cifra_negocios
  label: "Cifra de negocios"
  input_mode: "manual"
  calculate_variacion: true
```

El usuario introduce los datos manualmente.

#### Fila Calculada

```yaml
- id: operating_margin_om
  label: "Operating Margin (OM)"
  input_mode: "calculated"
  ejercicio_actual_formula: "ebit.ejercicio_actual / cifra_negocios.ejercicio_actual"
  ejercicio_anterior_formula: "ebit.ejercicio_anterior / cifra_negocios.ejercicio_anterior"
  calculate_variacion: true
```

Los valores se calculan automáticamente usando las fórmulas especificadas.

### Filas de Footer

```yaml
footer_rows:
  - id: total
    label: "Total"
    row_type: "sum"
    sum_columns: ["ingreso_local_file", "gasto_local_file"]

  - id: pesos
    label: "Pesos"
    row_type: "percent_of_total"
```

#### Tipos de Footer

| Tipo | Descripción |
|------|-------------|
| `sum` | Suma las columnas especificadas |
| `percent_of_total` | Calcula porcentajes del total |

### Encabezados Dinámicos

```yaml
dynamic_headers:
  ejercicio_actual_var: "<<Ejercicio completo>>"
  ejercicio_anterior_var: "ejercicio_anterior"
```

Permite usar valores de variables simples en los encabezados de columna.

### Configuración Avanzada

#### Tabla con Marcador Parametrizado

```yaml
analisis_indirecto_operacion:
  marker_pattern: "<<Tabla Operación {n}>>"
  title: "Análisis TNMM por operación"
  parameters:
    n:
      type: integer
      min: 1
      max: 10
```

Esto permite crear múltiples tablas usando un patrón (ej: `<<Tabla Operación 1>>`, `<<Tabla Operación 2>>`, etc.)

### Ejemplos Completos de Tablas

#### Tabla Simple de Una Fila

```yaml
analisis_indirecto_global:
  marker: "<<Tabla análisis indirecto>>"
  title: "Análisis indirecto TNMM"
  ui_component: "single_row_grid"
  remove_empty_rows: false
  columns:
    - id: min
      header: "Min"
      type: "percent"
      editable: true
    - id: lq
      header: "LQ"
      type: "percent"
      editable: true
    - id: med
      header: "Med"
      type: "percent"
      editable: true
  rows:
    - id: rango_tnmm
      label: "Rango TNMM"
      editable: true
```

#### Tabla con Filas Dinámicas

```yaml
operaciones_vinculadas:
  marker: "<<Tabla operaciones vinculadas>>"
  title: "Operaciones vinculadas"
  ui_component: "grid_dynamic_rows"
  remove_empty_rows: true
  columns:
    - id: tipo_operacion
      header: "Tipo de operación vinculada"
      type: "text"
      editable: true
    - id: entidad_vinculada
      header: "Entidad vinculada"
      type: "text"
      editable: true
  footer_rows:
    - id: total
      label: "Total"
      row_type: "sum"
      sum_columns: ["ingreso_local_file", "gasto_local_file"]
```

#### Tabla con Columnas Condicionales

```yaml
cumplimiento_formal_LF:
  marker: "<<Tabla de cumplimiento formal LF>>"
  title: "Cumplimiento formal detallado – Local File"
  ui_component: "grid"
  columns:
    - id: requisito
      header: "Requisito (art. 16 RIS)"
      type: "text"
      editable: false
    - id: cumplimiento
      header: "Cumplimiento"
      type: "choice"
      editable: true
      options: ["Sí", "No", "Ver comentario"]
    - id: comentario
      header: "Comentario"
      type: "text"
      editable: true
      conditional_on:
        field: "cumplimiento"
        value: "Ver comentario"
```

#### Cómo Agregar una Nueva Tabla

1. **Definir la tabla** en `tablas.yaml`:

```yaml
mi_nueva_tabla:
  marker: "<<Mi Nueva Tabla>>"
  title: "Mi Nueva Tabla"
  ui_component: "grid"
  remove_empty_rows: false
  columns:
    - id: columna1
      header: "Columna 1"
      type: "text"
      editable: true
    - id: columna2
      header: "Columna 2"
      type: "number"
      editable: true
  rows:
    - id: fila1
      label: "Fila 1"
    - id: fila2
      label: "Fila 2"
```

2. **Agregar el marcador** en la plantilla Word:
```
<<Mi Nueva Tabla>>
```

3. **Actualizar el código UI** en `app/ui/sections_tables.py` para renderizar la tabla en la interfaz

---

## Ejemplos Prácticos

### Ejemplo 1: Agregar un Campo de Email Adicional

**Paso 1**: Editar `variables_simples.yaml`

```yaml
simple_variables:
  # ... variables existentes ...

  - id: email_contacto_adicional
    label: "Email de Contacto Adicional"
    marker: "<<Email Contacto Adicional>>"
    type: "email"
```

**Paso 2**: Agregar en la plantilla Word
```
Contacto adicional: <<Email Contacto Adicional>>
```

### Ejemplo 2: Agregar un Bloque Condicional Nuevo

**Paso 1**: Crear el archivo `app/condiciones/comentario_personalizado.docx` con el contenido deseado

**Paso 2**: Editar `variables_condicionales.yaml`

```yaml
conditions:
  # ... condiciones existentes ...

  - id: comentario_personalizado
    label: "Comentario Personalizado"
    marker: "<<Comentario Personalizado>>"
    type: "boolean_include_doc"
    question: "¿Incluir comentario personalizado?"
    yes_no_values: ["Sí", "No"]
    word_file: "condiciones/comentario_personalizado.docx"
```

**Paso 3**: Agregar en la plantilla Word
```
<<Comentario Personalizado>>
```

### Ejemplo 3: Agregar una Columna a una Tabla Existente

**Editar** `tablas.yaml`, buscar la tabla y agregar una columna:

```yaml
operaciones_vinculadas:
  # ... configuración existente ...
  columns:
    # ... columnas existentes ...

    - id: nueva_columna
      header: "Nueva Información"
      type: "text"
      editable: true
```

---

## Solución de Problemas

### Problema: La variable no aparece en el documento generado

**Soluciones**:
1. Verificar que el `marker` coincida exactamente con el texto en la plantilla Word (incluyendo los símbolos `<< >>`)
2. Asegurarse de que la variable tiene un valor (no está vacía)
3. Verificar que no hay espacios adicionales en el marcador

### Problema: El contenido condicional no se inserta

**Soluciones**:
1. Verificar que el archivo Word especificado en `word_file` existe en la ruta correcta
2. Asegurarse de que la ruta es relativa al directorio `app/`
3. Verificar que el marcador está en la plantilla principal

### Problema: La tabla no se genera correctamente

**Soluciones**:
1. Verificar que el marcador de tabla está en un párrafo separado (no dentro de otro texto)
2. Asegurarse de que todas las columnas tienen un `id` único
3. Verificar que el `ui_component` es válido (`grid`, `single_row_grid`, o `grid_dynamic_rows`)

### Problema: Los valores calculados no funcionan

**Soluciones**:
1. Verificar que las fórmulas usan los IDs correctos de las filas
2. Asegurarse de que las filas referenciadas existen y tienen valores
3. Verificar que la sintaxis de la fórmula es correcta (usar `.ejercicio_actual` o `.ejercicio_anterior`)

### Problema: Los encabezados dinámicos no se reemplazan

**Soluciones**:
1. Verificar que las variables usadas en `dynamic_headers` existen en `variables_simples.yaml`
2. Asegurarse de que los placeholders en `header_template` coinciden con los nombres en `dynamic_headers`
3. Verificar que las variables tienen valores antes de generar el informe

---

## Referencia Rápida

### Tipos de Datos para Variables Simples
- `text`: Texto corto
- `long_text`: Texto largo (textarea)
- `number`: Número decimal
- `percent`: Porcentaje
- `email`: Correo electrónico
- `integer`: Número entero

### Tipos de Componentes UI para Tablas
- `grid`: Tabla estándar con filas fijas
- `single_row_grid`: Tabla de una sola fila
- `grid_dynamic_rows`: Tabla con filas dinámicas (añadir/quitar)

### Tipos de Columnas para Tablas
- `text`: Texto libre
- `number`: Número decimal
- `integer`: Número entero
- `percent`: Porcentaje
- `choice`: Lista de opciones predefinidas

### Estructura de Marcadores
- Variables simples: `<<Nombre de Variable>>`
- Condiciones: `<<Nombre de Condición>>`
- Tablas: `<<Nombre de Tabla>>`
- Tablas parametrizadas: `<<Tabla Operación {n}>>`

---

## Notas Importantes

1. **Los marcadores son case-sensitive**: `<<Variable>>` es diferente de `<<variable>>`
2. **No use caracteres especiales** en los IDs (solo letras, números y guiones bajos)
3. **Mantenga copias de seguridad**: Siempre haga una copia de los archivos YAML antes de editarlos
4. **Validación de YAML**: Use un validador YAML online para verificar la sintaxis antes de guardar
5. **Orden de procesamiento**: Las variables simples se procesan primero, luego las tablas, y finalmente las condiciones
6. **Marcadores no utilizados**: Los marcadores sin valor se eliminan automáticamente del documento final

---

## Recursos Adicionales

- **Validador YAML Online**: http://www.yamllint.com/
- **Tutorial YAML**: https://yaml.org/
- **Documentación python-docx**: https://python-docx.readthedocs.io/

---

**Última actualización**: 2025-11-15
**Versión**: 2.0
