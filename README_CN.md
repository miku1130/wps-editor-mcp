<div align="center">

# 📝 WPS Editor MCP

**让AI能够编辑Word、Excel、PowerPoint文档的MCP服务器**

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-Protocol-purple.svg)](https://modelcontextprotocol.io/)

[English](README.md) • [功能特性](#-功能特性) • [快速开始](#-快速开始) • [工具列表](#-mcp工具) • [使用示例](#-使用示例)

</div>

---

## 📖 简介

WPS Editor MCP 是一个基于 Model Context Protocol (MCP) 的服务器，让AI助手能够直接读写和编辑 Microsoft Office 及 WPS Office 文档。支持两种工作模式，无需安装Office即可处理现代文档格式，也可通过COM接口控制WPS应用实现更强大的功能。

## ✨ 功能特性

### 🔄 双模式架构

| 模式 | 说明 | 依赖 | 支持格式 |
|:----:|:-----|:-----|:---------|
| **file** | 直接解析Office Open XML | python-docx, openpyxl, python-pptx | .docx, .xlsx, .pptx |
| **com** | 通过COM控制WPS应用 | pywin32 + WPS Office | .doc, .xls, .ppt, .docx, .xlsx, .pptx |

### 🎯 核心能力

- **📄 Word处理**: 读写文档、设置字体样式、插入图片、段落格式化
- **📊 Excel处理**: 读写单元格、表格操作、样式设置、合并单元格
- **📽️ PPT处理**: 读写幻灯片、插入图片、自动生成演示文稿
- **🎨 专业模板**: 7种精美PPT模板（ocean/forest/sunset/royal/minimal/blue/green）
- **🔄 格式转换**: Word转PPT、PPT转Word总结、导出PDF
- **⚙️ WPS自动化**: 运行宏、控制WPS应用

### 🌟 亮点功能

| 功能 | 说明 |
|:-----|:-----|
| `doc_to_ppt` | 根据Word文档自动生成专业PPT演示文稿 |
| `ppt_to_doc` | 读取PPT生成结构化Word总结文档 |
| `export_pdf` | 将Office文档导出为PDF |
| `run_macro` | 运行WPS宏脚本 |

---

## 🚀 快速开始

### 安装依赖

```bash
# 克隆仓库
git clone https://github.com/miku1130/wps-editor-mcp.git
cd wps-editor-mcp

# 安装依赖
pip install -r requirements.txt
```

### 启动服务器

```bash
python server.py
```

### 配置MCP客户端

在你的MCP客户端配置文件中添加：

```json
{
  "mcpServers": {
    "wps-editor": {
      "command": "python",
      "args": ["path/to/server.py"]
    }
  }
}
```

---

## 🛠️ MCP工具

### 基础操作

| 工具 | 说明 | 参数 |
|:-----|:-----|:-----|
| `read_document` | 读取文档内容 | file_path, mode, sheet_name |
| `write_document` | 写入文档内容 | file_path, content, mode, append |
| `edit_cell` | 编辑Excel单元格 | file_path, row, column, value |
| `get_sheet_names` | 获取工作表列表 | file_path |

### 表格操作

| 工具 | 说明 |
|:-----|:-----|
| `add_rows` | 添加行 |
| `add_columns` | 添加列 |
| `merge_cells` | 合并单元格 |
| `unmerge_cells` | 取消合并 |
| `set_column_width` | 设置列宽 |
| `set_row_height` | 设置行高 |

### 样式设置

| 工具 | 说明 | 支持属性 |
|:-----|:-----|:---------|
| `set_cell_style` | 设置单元格样式 | 字体、字号、加粗、斜体、下划线、颜色、背景色、对齐 |
| `set_paragraph_style` | 设置段落样式 | 字体、字号、加粗、斜体、下划线、颜色、对齐 |

### 高级功能

| 工具 | 说明 |
|:-----|:-----|
| `insert_image` | 插入图片 |
| `export_pdf` | 导出PDF |
| `run_macro` | 运行WPS宏 |
| `doc_to_ppt` | Word转PPT |
| `ppt_to_doc` | PPT转Word总结 |

### 批量操作（新增）

| 工具 | 说明 |
|:-----|:-----|
| `batch_read` | 批量读取多个段落或单元格 |
| `batch_write` | 批量写入多个Excel单元格 |
| `batch_set_style` | 批量设置多个元素的样式 |

### 质量检查与模板（新增）

| 工具 | 说明 |
|:-----|:-----|
| `check_quality` | 检查文档质量并提供改进建议 |
| `extract_template` | 从文档提取模板 |
| `apply_template` | 应用模板到文档 |

---

## 💡 使用示例

### 示例1: 创建Word文档

```python
# 写入内容
write_document(
    file_path="report.docx",
    content="这是标题\n\n这是正文内容。",
    mode="file"
)

# 设置标题样式（宋体、22号、加粗、居中）
set_paragraph_style(
    file_path="report.docx",
    para_index=0,
    font_name="SimSun",
    font_size=22,
    bold=True,
    align="center"
)
```

### 示例2: 操作Excel表格

```python
# 创建表格
write_document(
    file_path="data.xlsx",
    content='[["姓名", "年龄"], ["张三", 25], ["李四", 30]]',
    mode="file"
)

# 设置表头样式
set_cell_style(
    file_path="data.xlsx",
    row=1, column=1,
    font_name="SimHei",
    font_size=12,
    bold=True,
    bg_color="4472C4",
    color="FFFFFF"
)

# 合并单元格
merge_cells(
    file_path="data.xlsx",
    start_row=1, start_col=1,
    end_row=1, end_col=2
)
```

### 示例3: 生成专业PPT

```python
# 从Word文档生成PPT
doc_to_ppt(
    doc_path="paper.docx",
    ppt_path="presentation.pptx",
    template="ocean"  # 可选: ocean/forest/sunset/royal/minimal/blue/green
)

# 从PPT生成Word总结
ppt_to_doc(
    ppt_path="presentation.pptx",
    doc_path="summary.docx"
)
```

### 示例4: COM模式（需要安装WPS）

```python
# 读取旧版.doc文件
read_document(
    file_path="old_file.doc",
    mode="com"
)

# 导出PDF
export_pdf(
    file_path="document.docx",
    output_path="output.pdf"
)
```

---

## 🎨 PPT模板预览

| 模板 | 风格 | 适用场景 |
|:-----|:-----|:---------|
| `ocean` | 🌊 深海蓝 | 商务汇报、技术分享 |
| `forest` | 🌲 森林绿 | 环保主题、自然科学 |
| `sunset` | 🌅 日落橙 | 创意展示、艺术设计 |
| `royal` | 👑 皇家紫 | 高端场合、正式汇报 |
| `minimal` | ⬜ 极简白 | 学术报告、简洁风格 |
| `blue` | 💙 经典蓝 | 通用场景、专业汇报 |
| `green` | 💚 清新绿 | 科技主题、创新展示 |

---

## ⚙️ 配置说明

### 环境要求

- Python 3.8+
- Windows系统（COM模式需要）

### 依赖包

```
# Office文档处理
python-docx>=1.1.0      # Word文档
openpyxl>=3.1.0          # Excel文档
python-pptx>=0.6.21      # PowerPoint文档

# COM自动化（可选，Windows）
pywin32>=306
comtypes>=1.4.1
```

---

## 📁 项目结构

```
wps-editor-mcp/
├── server.py              # MCP服务器主程序
├── requirements.txt       # 依赖列表
├── test_server.py         # 单元测试
├── README.md              # 英文文档
├── README_CN.md           # 中文文档
└── LICENSE                # MIT许可证
```

---

## 🧪 运行测试

```bash
python test_server.py
```

测试输出：
```
============================================================
WPS Editor MCP Unit Tests
============================================================

=== Test MCP Protocol ===
[OK] MCP init success
[OK] Tools list correct, 17 tools

=== Test Word Basic ===
[OK] Write Word success
[OK] Read Word success

=== Test Excel Basic ===
[OK] Write Excel success
[OK] Read Excel success
[OK] Edit cell success

All tests passed!
```

---

## 📊 工具总览

共 **23个** MCP工具：

| 类别 | 数量 | 工具 |
|:-----|:----:|:-----|
| 基础操作 | 4 | read_document, write_document, edit_cell, get_sheet_names |
| 表格操作 | 6 | add_rows, add_columns, merge_cells, unmerge_cells, set_column_width, set_row_height |
| 样式设置 | 2 | set_cell_style, set_paragraph_style |
| 图片插入 | 1 | insert_image |
| 文档转换 | 2 | doc_to_ppt, ppt_to_doc |
| 批量操作 | 3 | batch_read, batch_write, batch_set_style |
| 质量与模板 | 3 | check_quality, extract_template, apply_template |
| 高级功能 | 2 | export_pdf, run_macro |

---

## 🤝 贡献

欢迎提交Issue和Pull Request！

1. Fork本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 创建Pull Request

---

## 📄 许可证

本项目基于 MIT 许可证开源 - 详见 [LICENSE](LICENSE) 文件

---

## 🙏 致谢

- [Model Context Protocol](https://modelcontextprotocol.io/) - MCP协议规范
- [python-docx](https://python-docx.readthedocs.io/) - Word文档处理
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel文档处理
- [python-pptx](https://python-pptx.readthedocs.io/) - PowerPoint文档处理

---

<div align="center">

**如果这个项目对你有帮助，请给个 ⭐ Star 支持一下！**

Made with ❤️

</div>
