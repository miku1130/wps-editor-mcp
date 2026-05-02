<div align="center">

# 📝 WPS Editor MCP

**An MCP server that enables AI to edit Word, Excel, and PowerPoint documents**

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-Protocol-purple.svg)](https://modelcontextprotocol.io/)

**English** | [中文文档](README_CN.md)

[Features](#-features) • [Quick Start](#-quick-start) • [Tools](#-mcp-tools) • [Examples](#-examples) • [Configuration](#-configuration)

</div>

---

## 📖 Introduction

WPS Editor MCP is a Model Context Protocol (MCP) server that enables AI assistants to directly read, write, and edit Microsoft Office and WPS Office documents. It supports two operating modes: processing modern document formats without installing Office, or controlling WPS applications via COM interface for more powerful features.

## ✨ Features

### 🔄 Dual Mode Architecture

| Mode | Description | Dependencies | Supported Formats |
|:----:|:------------|:-------------|:------------------|
| **file** | Direct Office Open XML parsing | python-docx, openpyxl, python-pptx | .docx, .xlsx, .pptx |
| **com** | WPS automation via COM | pywin32 + WPS Office | .doc, .xls, .ppt, .docx, .xlsx, .pptx |

### 🎯 Core Capabilities

- **📄 Word Processing**: Read/write documents, set font styles, insert images, paragraph formatting
- **📊 Excel Processing**: Read/write cells, table operations, style settings, merge cells
- **📽️ PPT Processing**: Read/write slides, insert images, auto-generate presentations
- **🎨 Professional Templates**: 7 beautiful PPT templates (ocean/forest/sunset/royal/minimal/blue/green)
- **🔄 Format Conversion**: Word to PPT, PPT to Word summary, export to PDF
- **⚙️ WPS Automation**: Run macros, control WPS applications

### 🌟 Highlight Features

| Feature | Description |
|:--------|:------------|
| `doc_to_ppt` | Auto-generate professional PPT from Word documents |
| `ppt_to_doc` | Generate structured Word summary from PPT |
| `export_pdf` | Export Office documents to PDF |
| `run_macro` | Run WPS macro scripts |

---

## 🚀 Quick Start

### Install Dependencies

```bash
# Clone the repository
git clone https://github.com/miku1130/wps-editor-mcp.git
cd wps-editor-mcp

# Install dependencies
pip install -r requirements.txt
```

### Start Server

```bash
python server.py
```

### Configure MCP Client

Add to your MCP client configuration file:

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

## 🛠️ MCP Tools

### Basic Operations

| Tool | Description | Parameters |
|:-----|:------------|:-----------|
| `read_document` | Read document content | file_path, mode, sheet_name |
| `write_document` | Write document content | file_path, content, mode, append |
| `edit_cell` | Edit Excel cell | file_path, row, column, value |
| `get_sheet_names` | Get worksheet names | file_path |

### Table Operations

| Tool | Description |
|:-----|:------------|
| `add_rows` | Add rows |
| `add_columns` | Add columns |
| `merge_cells` | Merge cells |
| `unmerge_cells` | Unmerge cells |
| `set_column_width` | Set column width |
| `set_row_height` | Set row height |

### Style Settings

| Tool | Description | Supported Properties |
|:-----|:------------|:---------------------|
| `set_cell_style` | Set cell style | font, size, bold, italic, underline, color, background, alignment |
| `set_paragraph_style` | Set paragraph style | font, size, bold, italic, underline, color, alignment |

### Advanced Features

| Tool | Description |
|:-----|:------------|
| `insert_image` | Insert image |
| `export_pdf` | Export to PDF |
| `run_macro` | Run WPS macro |
| `doc_to_ppt` | Word to PPT |
| `ppt_to_doc` | PPT to Word summary |

### Batch Operations (New)

| Tool | Description |
|:-----|:------------|
| `batch_read` | Batch read multiple paragraphs or cells |
| `batch_write` | Batch write multiple Excel cells |
| `batch_set_style` | Batch set styles for multiple elements |

### Quality & Templates (New)

| Tool | Description |
|:-----|:------------|
| `check_quality` | Check document quality and get suggestions |
| `extract_template` | Extract template from document |
| `apply_template` | Apply template to document |

---

## 💡 Examples

### Example 1: Create Word Document

```python
# Write content
write_document(
    file_path="report.docx",
    content="Title\n\nBody content here.",
    mode="file"
)

# Set title style (SimSun, 22pt, bold, centered)
set_paragraph_style(
    file_path="report.docx",
    para_index=0,
    font_name="SimSun",
    font_size=22,
    bold=True,
    align="center"
)
```

### Example 2: Excel Table Operations

```python
# Create table
write_document(
    file_path="data.xlsx",
    content='[["Name", "Age"], ["Alice", 25], ["Bob", 30]]',
    mode="file"
)

# Set header style
set_cell_style(
    file_path="data.xlsx",
    row=1, column=1,
    font_name="SimHei",
    font_size=12,
    bold=True,
    bg_color="4472C4",
    color="FFFFFF"
)

# Merge cells
merge_cells(
    file_path="data.xlsx",
    start_row=1, start_col=1,
    end_row=1, end_col=2
)
```

### Example 3: Generate Professional PPT

```python
# Generate PPT from Word document
doc_to_ppt(
    doc_path="paper.docx",
    ppt_path="presentation.pptx",
    template="ocean"  # Options: ocean/forest/sunset/royal/minimal/blue/green
)

# Generate Word summary from PPT
ppt_to_doc(
    ppt_path="presentation.pptx",
    doc_path="summary.docx"
)
```

### Example 4: COM Mode (Requires WPS Installation)

```python
# Read legacy .doc file
read_document(
    file_path="old_file.doc",
    mode="com"
)

# Export to PDF
export_pdf(
    file_path="document.docx",
    output_path="output.pdf"
)
```

---

## 🎨 PPT Templates

| Template | Style | Use Cases |
|:---------|:------|:----------|
| `ocean` | 🌊 Deep Blue | Business reports, tech sharing |
| `forest` | 🌲 Forest Green | Environmental, natural science |
| `sunset` | 🌅 Sunset Orange | Creative displays, art design |
| `royal` | 👑 Royal Purple | High-end occasions, formal reports |
| `minimal` | ⬜ Minimal White | Academic reports, clean style |
| `blue` | 💙 Classic Blue | General purpose, professional |
| `green` | 💚 Fresh Green | Tech themes, innovation |

---

## ⚙️ Configuration

### Requirements

- Python 3.8+
- Windows (for COM mode)

### Dependencies

```
# Office document processing
python-docx>=1.1.0      # Word documents
openpyxl>=3.1.0          # Excel documents
python-pptx>=0.6.21      # PowerPoint documents

# COM automation (optional, Windows)
pywin32>=306
comtypes>=1.4.1
```

---

## 📁 Project Structure

```
wps-editor-mcp/
├── server.py              # MCP server main program
├── requirements.txt       # Dependencies
├── test_server.py         # Unit tests
├── README.md              # English documentation
├── README_CN.md           # Chinese documentation
└── LICENSE                # MIT License
```

---

## 🧪 Run Tests

```bash
python test_server.py
```

Test output:
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

## 📊 Tools Overview

Total **23** MCP tools:

| Category | Count | Tools |
|:---------|:-----:|:------|
| Basic Operations | 4 | read_document, write_document, edit_cell, get_sheet_names |
| Table Operations | 6 | add_rows, add_columns, merge_cells, unmerge_cells, set_column_width, set_row_height |
| Style Settings | 2 | set_cell_style, set_paragraph_style |
| Image Insert | 1 | insert_image |
| Document Conversion | 2 | doc_to_ppt, ppt_to_doc |
| Batch Operations | 3 | batch_read, batch_write, batch_set_style |
| Quality & Templates | 3 | check_quality, extract_template, apply_template |
| Advanced Features | 2 | export_pdf, run_macro |

---

## 🤝 Contributing

Issues and Pull Requests are welcome!

1. Fork the repository
2. Create feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add AmazingFeature'`)
4. Push to branch (`git push origin feature/AmazingFeature`)
5. Create Pull Request

---

## 📄 License

This project is licensed under the MIT License - see [LICENSE](LICENSE) for details

---

## 🙏 Acknowledgments

- [Model Context Protocol](https://modelcontextprotocol.io/) - MCP protocol specification
- [python-docx](https://python-docx.readthedocs.io/) - Word document processing
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel document processing
- [python-pptx](https://python-pptx.readthedocs.io/) - PowerPoint document processing

---

## 💬 Community

Join our QQ group for discussion and support:

**AI编程开源技术交流群**

[![QQ Group](https://img.shields.io/badge/QQ_Group-Join-12B7F5?style=for-the-badge&logo=tencent-qq&logoColor=white)](https://qm.qq.com/q/kjM22sjGU2)

🔗 Direct link: https://qm.qq.com/q/kjM22sjGU2

---

<div align="center">

**If this project helps you, please give a ⭐ Star!**

Made with ❤️

</div>
