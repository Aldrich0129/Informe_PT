# YAML 个性化操作指南

本指南面向需要扩展或定制 Transfer Pricing 报告的用户，帮助你安全地修改 `app/config` 目录下的 YAML 配置文件。示例均以实际项目中的文件结构为基础，可直接复制后调整。

## 1. 熟悉目录结构

```
app/
 ├─ config/
 │   ├─ variables_simples.yaml          # 文本与数字变量
 │   ├─ variables_condicionales.yaml    # 条件/开关块
 │   ├─ tablas.yaml                     # 所有表格定义
 │   └─ Plantilla.docx                  # Word 模板（含 <<Marcador>> ）
 └─ condiciones/                        # 条件块对应的 Word 片段
```

> 修改 YAML 时，务必保证 `marker` 名称与 `Plantilla.docx` 中的标记完全一致（区分大小写）。

## 2. 个性化变量 `variables_simples.yaml`

### 2.1 新增或修改字段

```yaml
simple_variables:
  - id: nombre_compania
    label: "Nombre de la Compañía"
    marker: "<<Nombre de la Compañía>>"
    type: "text"

  - id: descripcion_resumen
    label: "Resumen ejecutivo"
    marker: "<<Resumen ejecutivo>>"
    type: "long_text"
    optional: true        # 留空时自动清除对应 marcador
```

**操作步骤**
1. 在 YAML 中添加条目（如上例中的 `descripcion_resumen`）。
2. 打开 `Plantilla.docx` 并插入同名 marcador：`<<Resumen ejecutivo>>`。
3. 运行应用并在 UI 中填写数据，导出的 Word 会自动替换内容并保留原有格式。

### 2.2 控制展示类型

- `type: text` → 单行输入框。
- `type: long_text` → 多行文本区域（换行将被保留）。
- `type: percent` → UI 中输入 0~1 数值，生成文档自动转成百分比（如 `0.35 → 35.00%`）。
- `optional: true` → 允许留空；生成文档时会删除 marcador 和空行。

## 3. 个性化条件 `variables_condicionales.yaml`

用于插入或隐藏 Word 片段，例如正式意见、附录等。

```yaml
conditions:
  - id: comentario_formal
    label: "¿Incluir comentario formal?"
    question: "¿Desea incorporar el bloque formal?"
    marker: "<<Comentario formal>>"
    word_file: "condiciones/comentario_formal.docx"
```

**自定义技巧**
1. 复制一个现有 `.docx` 条件文件并编辑内容，例如 `condiciones/nuevo_anexo.docx`。
2. 在 YAML 中新增一条 `condition`，`marker` 必须与主模板中的 `<<...>>` 完全一致。
3. 运行应用时选择 “Sí”，系统会在 marcador 位置插入该 Word 片段；若选择 “No”，该段会被整体删除。

## 4. 个性化表格 `tablas.yaml`

每张表由列定义、默认值、格式与可选的计算项组成。

```yaml
- id: operaciones_vinculadas
  marker: "<<Tabla operaciones vinculadas>>"
  columns:
    - id: tipo_operacion
      label: "Tipo de operación"
      width: 2000
    - id: ingreso
      label: "Ingresos (€)"
      type: number
      format: "{:,.2f}"
  totals:
    ingreso: sum
```

**添加新列的步骤**
1. 在表的 `columns` 列表末尾增加新列（如 `margen_pto`）。
2. 如果需要自动合计/平均值，可在 `totals` 中添加 `margen_pto: avg` 等规则。
3. 确保 Word 模板中对应 marcador（如 `<<Tabla operaciones vinculadas>>`）存在，系统会重建整张表。

## 5. 组合案例：新增 “Anexo ESG”

**需求**：在 índice 中显示 “Anexo ESG <<6>>”，正文新增章节，并允许用户填写文本。

1. **模板修改**
   - 在目录中加入 `Anexo ESG <<6>>`。
   - 在正文相应位置添加 `<<6>> Anexo ESG` 标题和 `<<Detalle Anexo ESG>>` marcador。
2. **YAML 配置**
   ```yaml
   simple_variables:
     - id: detalle_anexo_esg
       label: "Detalle Anexo ESG"
       marker: "<<Detalle Anexo ESG>>"
       type: "long_text"
   ```
3. **结果**
   - 用户在 UI 中输入内容，生成的 Word 会在正文章节中展示文本。
   - 如果正文中缺少 `<<6>>` marcador，新版目录逻辑会自动移除该行，避免目录与正文不一致。

## 6. QA 建议

| 检查点 | 做法 |
| --- | --- |
| Marcador 正确性 | 使用 Word 的 “Buscar” 功能确认 `<<Nombre>>` 是否匹配 YAML 中的 `marker`。 |
| YAML 语法 | 推荐使用 VS Code + YAML 插件，保存前执行一次格式化。 |
| 多语言/重音 | 若 marcador 含重音（如 `<<Información>>`），YAML 也必须包含重音，且编码保持 UTF-8。 |
| 版本管理 | 每次更改 YAML/Plantilla 后提交到 Git，方便回滚。 |

按照以上流程即可快速个性化 YAML 内容，并保持目录、正文与模板三者之间的同步。
