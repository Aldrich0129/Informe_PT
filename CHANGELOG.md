# Changelog - 更新日志

## 2025-11-17 - 修复列布局和添加图片背景功能

### 新增功能

#### 1. 保留文档列布局 (Column Layout Preservation)
- **问题**: 插入条件变量时，双列布局会被转换为单列布局
- **解决方案**: 改进了 `_insert_document_at_marker` 方法
  - 自动检测源文档的列配置（单列/双列/多列）
  - 在插入内容时保留原始的列布局设置
  - 插入后自动恢复目标文档的原始列配置
  - 使用 Section (sectPr) 和 Columns (cols) XML 元素实现

#### 2. 图片背景插入功能 (Background Image Insertion)
- **新增目录**: `app/config/images/` - 用于存储首页和末页的背景图片
- **支持格式**: PNG, JPEG, JPG, BMP
- **侧边栏新增功能**:
  - 首页图片选择/上传
  - 末页图片选择/上传
  - 三种模式：
    1. 无图片
    2. 从现有图片中选择
    3. 上传新图片

#### 3. 智能图片处理
- 自动调整图片大小以适配页面尺寸
- 自动检测页面方向（横向/纵向）
- 图片作为背景处理（置于所有内容后方）
- 使用 anchor 定位替代 inline 定位
- 图片居中对齐

### 技术实现

#### WordEngine 更新 (`app/modules/word_engine.py`)

1. **增强的 `_insert_document_at_marker` 方法** (行 557-644)
   - 检测源文档的 section 配置
   - 提取 `w:cols` XML 元素
   - 在插入前创建 section break 以保留列布局
   - 在插入后恢复原始列配置

2. **新增 `insert_background_image` 方法** (行 1503-1617)
   - 参数:
     - `image_path`: 图片文件路径
     - `page_type`: "first" (首页) 或 "last" (末页)
   - 功能:
     - 获取页面尺寸和方向
     - 插入图片到指定位置
     - 转换 inline 为 anchor 定位
     - 设置 `behindDoc='1'` 将图片置于文本后方
     - 配置居中对齐

#### UI 更新 (`app/ui/main_ui.py`)

**新增侧边栏图片部分** (行 92-178)
- 图片选择界面
- 文件上传功能
- 图片列表显示
- Session state 管理:
  - `first_page_image_path`
  - `last_page_image_path`

#### 应用集成 (`app/app.py`)

**图片插入集成** (行 117-128)
- 在文档生成的最后阶段插入图片
- 检查 session state 中的图片路径
- 调用 `insert_background_image` 方法

### XML 结构说明

#### 列布局 (Columns)
```xml
<w:sectPr>
  <w:cols w:num="2"/>  <!-- 双列布局 -->
</w:sectPr>
```

#### 图片定位 (Anchor)
```xml
<wp:anchor behindDoc="1" allowOverlap="1">
  <wp:simplePos x="0" y="0"/>
  <wp:positionH relativeFrom="page">
    <wp:align>center</wp:align>
  </wp:positionH>
  <wp:positionV relativeFrom="page">
    <wp:align>center</wp:align>
  </wp:positionV>
  <!-- 图片数据 -->
</wp:anchor>
```

### 文件更改摘要

| 文件 | 更改类型 | 说明 |
|------|---------|------|
| `app/modules/word_engine.py` | 修改/新增 | 改进列布局保留，添加图片背景功能 |
| `app/ui/main_ui.py` | 新增 | 侧边栏图片选择/上传界面 |
| `app/app.py` | 修改 | 集成图片插入到文档生成流程 |
| `app/config/images/` | 新增 | 图片存储目录 |
| `app/config/images/README.md` | 新增 | 图片目录使用说明 |

### 使用指南

#### 添加背景图片

1. **准备图片**
   - 将图片文件（PNG, JPEG, BMP）放入 `app/config/images/` 目录
   - 或通过应用界面直接上传

2. **选择图片**
   - 打开应用侧边栏
   - 找到 "🖼️ Imágenes de Fondo" 部分
   - 选择首页或末页图片

3. **图片选项**
   - **无图片**: 不插入背景图片
   - **选择现有**: 从已上传的图片中选择
   - **上传新图片**: 上传新的背景图片

4. **生成文档**
   - 正常填写表单
   - 点击 "生成报告"
   - 图片将自动插入到首页/末页作为背景

### 注意事项

1. **列布局保留**
   - 仅适用于通过 `insert_conditional_blocks` 插入的内容
   - 直接文本替换不受影响

2. **图片背景**
   - 图片会被拉伸以填充整个页面
   - 建议使用与页面尺寸相匹配的图片
   - 图片将位于所有文本和元素后方

3. **性能考虑**
   - 大型图片会增加文档文件大小
   - 建议优化图片尺寸和压缩

### 兼容性

- ✅ Microsoft Word 2010+
- ✅ LibreOffice Writer 6.0+
- ✅ Google Docs (部分功能)
- ✅ PDF 导出

### 下一步计划

- [ ] 添加图片预览功能
- [ ] 支持多页背景图片
- [ ] 图片透明度控制
- [ ] 批量图片管理

---

## 历史版本

### 2025-11-17 (之前)
- 初始版本
- 基本的文档生成功能
- 变量替换、表格插入、条件块
