"""
WPS文件编辑MCP服务器
兼容Python 3.8+，不依赖官方MCP SDK
支持两种模式：
1. file_mode: 直接解析Office Open XML格式（不需要安装WPS）
2. com_mode: 通过COM接口控制WPS应用（需要安装WPS）

MCP协议实现基于JSON-RPC 2.0 over stdio
"""

import os
import sys
import json
from enum import Enum
from typing import Optional, Any, Dict, List, Tuple
from pathlib import Path


class EditMode(str, Enum):
    FILE = "file"    # 直接文件解析
    COM = "com"      # WPS COM自动化


def get_mode(mode: Optional[str] = None) -> EditMode:
    """获取编辑模式，默认为file模式"""
    if mode is None:
        return EditMode.FILE
    return EditMode(mode)


# ==================== 文件格式解析模式 ====================

def _file_read_docx(file_path: str) -> dict:
    """读取Word文档内容"""
    from docx import Document
    doc = Document(file_path)
    paragraphs = [p.text for p in doc.paragraphs]
    tables = []
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            table_data.append([cell.text for cell in row.cells])
        tables.append(table_data)
    return {
        "paragraphs": paragraphs,
        "tables": tables,
        "paragraph_count": len(paragraphs),
        "table_count": len(tables)
    }


def _file_write_docx(file_path: str, content: str, append: bool = False) -> str:
    """写入Word文档"""
    from docx import Document
    if append and os.path.exists(file_path):
        doc = Document(file_path)
    else:
        doc = Document()
    
    # 按换行符分割，每一行作为一个段落
    lines = content.split('\n')
    for line in lines:
        if line.strip():  # 非空行
            doc.add_paragraph(line)
        else:  # 空行也添加空段落
            doc.add_paragraph('')
    
    doc.save(file_path)
    return f"已写入: {file_path}"


def _file_read_xlsx(file_path: str, sheet_name: Optional[str] = None) -> dict:
    """读取Excel文件内容"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb.active
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append(list(row))
    return {
        "sheet_name": ws.title,
        "data": data,
        "rows": ws.max_row,
        "columns": ws.max_column
    }


def _file_write_xlsx(file_path: str, sheet_name: Optional[str], data: list) -> str:
    """写入Excel文件"""
    from openpyxl import Workbook, load_workbook
    if os.path.exists(file_path):
        wb = load_workbook(file_path)
    else:
        wb = Workbook()
    if sheet_name:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(sheet_name)
    else:
        ws = wb.active
    for row_idx, row_data in enumerate(data, 1):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    wb.save(file_path)
    return f"已写入: {file_path}"


def _file_read_pptx(file_path: str) -> dict:
    """读取PowerPoint文件内容"""
    from pptx import Presentation
    prs = Presentation(file_path)
    slides_data = []
    for slide in prs.slides:
        slide_content = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                slide_content.append(shape.text)
        slides_data.append(slide_content)
    return {
        "slide_count": len(prs.slides),
        "slides": slides_data
    }


def _file_write_pptx(file_path: str, content: str) -> str:
    """写入PowerPoint文件"""
    from pptx import Presentation
    from pptx.util import Inches
    if os.path.exists(file_path):
        prs = Presentation(file_path)
    else:
        prs = Presentation()
    slide_layout = prs.slide_layouts[6]  # 空白布局
    slide = prs.slides.add_slide(slide_layout)
    left = Inches(1)
    top = Inches(1)
    width = Inches(8)
    height = Inches(1)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    textbox.text = content
    prs.save(file_path)
    return f"已写入: {file_path}"


# ==================== Excel表格操作 ====================

def _file_add_rows(file_path: str, sheet_name: Optional[str], count: int, after_row: Optional[int] = None) -> str:
    """添加行"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    if after_row is None:
        after_row = ws.max_row
    ws.insert_rows(after_row + 1, count)
    wb.save(file_path)
    return f"已在第{after_row}行后添加{count}行"


def _file_add_columns(file_path: str, sheet_name: Optional[str], count: int, after_col: Optional[int] = None) -> str:
    """添加列"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    if after_col is None:
        after_col = ws.max_column
    ws.insert_cols(after_col + 1, count)
    wb.save(file_path)
    return f"已在第{after_col}列后添加{count}列"


def _file_merge_cells(file_path: str, sheet_name: Optional[str], start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """合并单元格"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    ws.merge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)
    wb.save(file_path)
    return f"已合并单元格: ({start_row},{start_col}) -> ({end_row},{end_col})"


def _file_unmerge_cells(file_path: str, sheet_name: Optional[str], start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """取消合并单元格"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    ws.unmerge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)
    wb.save(file_path)
    return f"已取消合并: ({start_row},{start_col}) -> ({end_row},{end_col})"


def _file_set_column_width(file_path: str, sheet_name: Optional[str], column: str, width: float) -> str:
    """设置列宽"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    ws.column_dimensions[column].width = width
    wb.save(file_path)
    return f"已设置列{column}宽度为{width}"


def _file_set_row_height(file_path: str, sheet_name: Optional[str], row: int, height: float) -> str:
    """设置行高"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    ws.row_dimensions[row].height = height
    wb.save(file_path)
    return f"已设置行{row}高度为{height}"


# ==================== 图片插入 ====================

def _file_insert_image_docx(file_path: str, image_path: str, width: Optional[float] = None, height: Optional[float] = None) -> str:
    """在Word中插入图片"""
    from docx import Document
    from docx.shared import Inches
    doc = Document(file_path) if os.path.exists(file_path) else Document()
    if width and height:
        doc.add_picture(image_path, width=Inches(width), height=Inches(height))
    elif width:
        doc.add_picture(image_path, width=Inches(width))
    else:
        doc.add_picture(image_path)
    doc.save(file_path)
    return f"已插入图片: {image_path}"


def _file_insert_image_xlsx(file_path: str, sheet_name: Optional[str], image_path: str, anchor: str = "A1") -> str:
    """在Excel中插入图片"""
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    img = Image(image_path)
    ws.add_image(img, anchor)
    wb.save(file_path)
    return f"已插入图片到{anchor}"


def _file_insert_image_pptx(file_path: str, image_path: str, slide_index: int = -1, left: float = 1, top: float = 1, width: float = 4) -> str:
    """在PowerPoint中插入图片"""
    from pptx import Presentation
    from pptx.util import Inches
    prs = Presentation(file_path) if os.path.exists(file_path) else Presentation()
    if slide_index == -1:
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
    else:
        slide = prs.slides[slide_index]
    slide.shapes.add_picture(image_path, Inches(left), Inches(top), Inches(width))
    prs.save(file_path)
    return f"已插入图片: {image_path}"


# ==================== 文档转换 ====================

def _set_slide_bg(slide, color_hex):
    """设置幻灯片背景颜色"""
    from pptx.dml.color import RGBColor
    
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(color_hex)


def _add_decorative_shapes(slide, accent_color_hex, slide_type="content"):
    """添加装饰形状"""
    from pptx.util import Inches
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.dml.color import RGBColor
    
    accent_color = RGBColor.from_string(accent_color_hex)
    
    if slide_type == "title":
        # Top accent bar
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), 
                                        Inches(13.333), Inches(0.12))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()
        
        # Bottom accent bar
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.2), 
                                        Inches(13.333), Inches(0.3))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()
        
        # Left decorative bar
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(2.8), 
                                        Inches(0.06), Inches(2))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()
    
    elif slide_type == "content":
        # Top thin line
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.35), 
                                        Inches(11.7), Inches(0.025))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()
        
        # Bottom bar
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.2), 
                                        Inches(13.333), Inches(0.3))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()
        
        # Accent dot
        shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.6), Inches(0.45), 
                                        Inches(0.12), Inches(0.12))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()


def _add_page_number(slide, page_num, total_pages, color_hex):
    """添加页码"""
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    
    left = Inches(12)
    top = Inches(7.22)
    width = Inches(1)
    height = Inches(0.25)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    p = tf.paragraphs[0]
    p.text = f"{page_num} / {total_pages}"
    p.font.size = Pt(9)
    p.font.color.rgb = RGBColor.from_string(color_hex)
    p.alignment = 2  # Right


def _doc_to_ppt(doc_path: str, ppt_path: str, template: str = "ocean") -> dict:
    """根据Word文档生成专业PPT"""
    from docx import Document
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE
    
    # Read Word document
    doc = Document(doc_path)
    
    # Extract structure
    title = ""
    subtitle = ""
    sections = []
    current_section = None
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        
        if not title:
            title = text
            continue
        
        if not subtitle and len(text) < 100:
            subtitle = text
            continue
        
        is_header = False
        if para.runs and para.runs[0].font.bold:
            is_header = True
        elif text and text[0].isdigit() and '. ' in text[:5]:
            is_header = True
        elif text.startswith('关键词') or text.startswith('Keywords'):
            is_header = True
        
        if is_header:
            current_section = {"title": text, "content": []}
            sections.append(current_section)
        elif current_section:
            current_section["content"].append(text)
        else:
            sections.append({"title": "", "content": [text]})
    
    # Create PowerPoint
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # Professional color themes
    themes = {
        "ocean": {
            "bg": "0A1628", "accent": "00B4D8", "accent2": "0077B6",
            "title": "FFFFFF", "text": "CAF0F8", "light": "90E0EF",
            "card": "1A2D4A"
        },
        "forest": {
            "bg": "1B2D1B", "accent": "52B788", "accent2": "40916C",
            "title": "FFFFFF", "text": "D8F3DC", "light": "95D5B2",
            "card": "2D4A2D"
        },
        "sunset": {
            "bg": "2D1B2E", "accent": "FF6B6B", "accent2": "EE5A24",
            "title": "FFFFFF", "text": "FFEAA7", "light": "FFB8B8",
            "card": "4A2D4A"
        },
        "royal": {
            "bg": "1A1A2E", "accent": "E94560", "accent2": "0F3460",
            "title": "FFFFFF", "text": "EAEAEA", "light": "A8A8A8",
            "card": "16213E"
        },
        "minimal": {
            "bg": "F8F9FA", "accent": "2D3436", "accent2": "636E72",
            "title": "2D3436", "text": "636E72", "light": "B2BEC3",
            "card": "FFFFFF"
        },
        "blue": {
            "bg": "0D1B2A", "accent": "00BFFF", "accent2": "1E90FF",
            "title": "FFFFFF", "text": "E0F0FF", "light": "87CEEB",
            "card": "1B2838"
        },
        "green": {
            "bg": "0D2818", "accent": "00FF7F", "accent2": "32CD32",
            "title": "FFFFFF", "text": "E0FFE0", "light": "90EE90",
            "card": "1A3D1A"
        }
    }
    
    theme = themes.get(template, themes["ocean"])
    slide_layout = prs.slide_layouts[6]  # Blank
    total_slides = len(sections) + 3  # Title + TOC + Content + Thank You
    
    # ============ SLIDE 1: Title ============
    slide = prs.slides.add_slide(slide_layout)
    _set_slide_bg(slide, theme["bg"])
    _add_decorative_shapes(slide, theme["accent"], "title")
    
    # Main title
    left, top = Inches(1.5), Inches(2.5)
    width, height = Inches(10), Inches(1.8)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = RGBColor.from_string(theme["title"])
    p.alignment = PP_ALIGN.LEFT
    
    # Subtitle
    if subtitle:
        left, top = Inches(1.5), Inches(4.5)
        width, height = Inches(10), Inches(1.2)
        textbox = slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        if len(subtitle) > 150:
            subtitle = subtitle[:150] + "..."
        p.text = subtitle
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor.from_string(theme["light"])
        p.alignment = PP_ALIGN.LEFT
    
    # ============ SLIDE 2: Table of Contents ============
    slide = prs.slides.add_slide(slide_layout)
    _set_slide_bg(slide, theme["bg"])
    _add_decorative_shapes(slide, theme["accent"], "content")
    _add_page_number(slide, 2, total_slides, theme["light"])
    
    # TOC Title
    left, top = Inches(1), Inches(0.8)
    width, height = Inches(11), Inches(0.7)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    p = tf.paragraphs[0]
    p.text = "CONTENTS"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = RGBColor.from_string(theme["accent"])
    
    # TOC items
    left, top = Inches(1), Inches(1.8)
    width, height = Inches(11), Inches(5)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.word_wrap = True
    
    toc_items = [s for s in sections if s["title"]][:8]
    for i, section in enumerate(toc_items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        
        run1 = p.add_run()
        run1.text = f"  {i+1:02d}  "
        run1.font.size = Pt(26)
        run1.font.bold = True
        run1.font.color.rgb = RGBColor.from_string(theme["accent"])
        
        run2 = p.add_run()
        title_text = section["title"]
        if len(title_text) > 50:
            title_text = title_text[:50] + "..."
        run2.text = title_text
        run2.font.size = Pt(20)
        run2.font.color.rgb = RGBColor.from_string(theme["text"])
        
        p.space_after = Pt(14)
    
    # ============ CONTENT SLIDES ============
    for i, section in enumerate(sections):
        if not section["title"] and not section["content"]:
            continue
        
        slide_num = i + 3
        slide = prs.slides.add_slide(slide_layout)
        _set_slide_bg(slide, theme["bg"])
        _add_decorative_shapes(slide, theme["accent"], "content")
        _add_page_number(slide, slide_num, total_slides, theme["light"])
        
        # Section number badge
        if section["title"]:
            left, top = Inches(0.6), Inches(0.5)
            width, height = Inches(0.9), Inches(0.9)
            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(theme["accent"])
            shape.line.fill.background()
            
            textbox = slide.shapes.add_textbox(left, top, width, height)
            tf = textbox.text_frame
            p = tf.paragraphs[0]
            p.text = f"{i+1:02d}"
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = RGBColor.from_string(theme["bg"])
            p.alignment = PP_ALIGN.CENTER
            p.space_before = Pt(12)
            
            # Section title
            left, top = Inches(1.8), Inches(0.6)
            width, height = Inches(10), Inches(0.8)
            textbox = slide.shapes.add_textbox(left, top, width, height)
            tf = textbox.text_frame
            p = tf.paragraphs[0]
            title_text = section["title"]
            if len(title_text) > 60:
                title_text = title_text[:60] + "..."
            p.text = title_text
            p.font.size = Pt(30)
            p.font.bold = True
            p.font.color.rgb = RGBColor.from_string(theme["title"])
        
        # Content card
        content_top = Inches(1.8) if section["title"] else Inches(0.5)
        
        left, top = Inches(0.6), content_top
        width, height = Inches(12), Inches(5.1)
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(theme["card"])
        shape.line.color.rgb = RGBColor.from_string(theme["accent"])
        shape.line.width = Pt(1)
        
        # Content bullets
        left, top = Inches(1), content_top + Inches(0.3)
        width, height = Inches(11.2), Inches(4.6)
        textbox = slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        tf.word_wrap = True
        
        for j, line in enumerate(section["content"][:5]):
            if j == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            
            run1 = p.add_run()
            run1.text = "  ●  "
            run1.font.size = Pt(12)
            run1.font.color.rgb = RGBColor.from_string(theme["accent"])
            
            run2 = p.add_run()
            if len(line) > 120:
                line = line[:120] + "..."
            run2.text = line
            run2.font.size = Pt(17)
            run2.font.color.rgb = RGBColor.from_string(theme["text"])
            
            p.space_after = Pt(12)
    
    # ============ FINAL SLIDE: Thank You ============
    slide = prs.slides.add_slide(slide_layout)
    _set_slide_bg(slide, theme["bg"])
    _add_decorative_shapes(slide, theme["accent"], "title")
    
    left, top = Inches(2), Inches(2.5)
    width, height = Inches(9), Inches(1.5)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    p = tf.paragraphs[0]
    p.text = "THANK YOU"
    p.font.size = Pt(56)
    p.font.bold = True
    p.font.color.rgb = RGBColor.from_string(theme["title"])
    p.alignment = PP_ALIGN.CENTER
    
    left, top = Inches(3), Inches(4.2)
    width, height = Inches(7), Inches(0.8)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    p = tf.paragraphs[0]
    p.text = "感谢聆听 | Questions & Discussion"
    p.font.size = Pt(22)
    p.font.color.rgb = RGBColor.from_string(theme["light"])
    p.alignment = PP_ALIGN.CENTER
    
    # Decorative line
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5), Inches(5.3), 
                                    Inches(3), Inches(0.04))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(theme["accent"])
    shape.line.fill.background()
    
    # Save
    prs.save(ppt_path)
    
    return {
        "title": title,
        "slides_count": len(prs.slides),
        "sections_count": len(sections),
        "template": template,
        "output": ppt_path
    }


def _ppt_to_doc(ppt_path: str, doc_path: str) -> dict:
    """读取PPT生成Word总结文档"""
    from pptx import Presentation
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    prs = Presentation(ppt_path)
    
    doc = Document()
    
    # Title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run("Presentation Summary")
    run.font.size = Pt(28)
    run.font.bold = True
    
    doc.add_paragraph()
    
    # Metadata
    meta_para = doc.add_paragraph()
    meta_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = meta_para.add_run(f"Total Slides: {len(prs.slides)}")
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(100, 100, 100)
    
    doc.add_paragraph()
    doc.add_paragraph("=" * 50)
    doc.add_paragraph()
    
    # Extract content from each slide
    all_content = []
    
    for i, slide in enumerate(prs.slides):
        slide_content = {"index": i + 1, "texts": []}
        
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if text:
                        slide_content["texts"].append(text)
        
        if slide_content["texts"]:
            all_content.append(slide_content)
    
    # Generate summary structure
    # Find title (usually first text of first slide)
    main_title = ""
    if all_content and all_content[0]["texts"]:
        main_title = all_content[0]["texts"][0]
    
    if main_title:
        h1 = doc.add_heading(main_title, level=1)
    
    doc.add_heading("Overview", level=2)
    doc.add_paragraph(f"This presentation contains {len(prs.slides)} slides covering the following topics:")
    
    # List all section titles
    doc.add_heading("Table of Contents", level=2)
    for i, slide_data in enumerate(all_content):
        if slide_data["texts"]:
            first_text = slide_data["texts"][0]
            if len(first_text) < 80:  # Likely a title
                doc.add_paragraph(f"Slide {slide_data['index']}: {first_text}", style='List Number')
    
    doc.add_paragraph()
    
    # Detailed content
    doc.add_heading("Detailed Content", level=2)
    
    for slide_data in all_content:
        if not slide_data["texts"]:
            continue
        
        # First text as slide heading
        slide_title = slide_data["texts"][0]
        doc.add_heading(f"Slide {slide_data['index']}: {slide_title}", level=3)
        
        # Rest as bullet points
        for text in slide_data["texts"][1:]:
            if len(text) > 5:  # Skip very short texts
                doc.add_paragraph(text, style='List Bullet')
        
        doc.add_paragraph()
    
    # Conclusion
    doc.add_heading("Key Takeaways", level=2)
    doc.add_paragraph("Based on the presentation content, the main points are:")
    
    # Extract key points from content
    key_points = []
    for slide_data in all_content:
        for text in slide_data["texts"]:
            if len(text) > 20 and len(text) < 150:
                key_points.append(text)
                if len(key_points) >= 5:
                    break
        if len(key_points) >= 5:
            break
    
    for point in key_points[:5]:
        doc.add_paragraph(point, style='List Bullet')
    
    # Save
    doc.save(doc_path)
    
    return {
        "slides_processed": len(all_content),
        "output": doc_path,
        "main_title": main_title
    }


# ==================== 批量操作 ====================

def _batch_read_docx(file_path: str, para_indices: list) -> list:
    """批量读取Word段落"""
    from docx import Document
    doc = Document(file_path)
    results = []
    for idx in para_indices:
        if 0 <= idx < len(doc.paragraphs):
            para = doc.paragraphs[idx]
            results.append({
                "index": idx,
                "text": para.text,
                "style": para.style.name if para.style else None
            })
        else:
            results.append({"index": idx, "error": "Index out of range"})
    return results


def _batch_read_xlsx(file_path: str, cells: list, sheet_name: str = None) -> list:
    """批量读取Excel单元格"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    results = []
    for cell_ref in cells:
        try:
            cell = ws[cell_ref]
            results.append({
                "cell": cell_ref,
                "value": cell.value,
                "type": type(cell.value).__name__ if cell.value is not None else "null"
            })
        except Exception as e:
            results.append({"cell": cell_ref, "error": str(e)})
    return results


def _batch_write_xlsx(file_path: str, data: list, sheet_name: str = None) -> str:
    """批量写入Excel单元格
    data格式: [{"cell": "A1", "value": "xxx"}, ...]
    """
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    count = 0
    for item in data:
        try:
            cell_ref = item.get("cell")
            value = item.get("value")
            if cell_ref and value is not None:
                ws[cell_ref] = value
                count += 1
        except:
            pass
    wb.save(file_path)
    return f"已批量写入{count}个单元格"


def _batch_set_style_docx(file_path: str, operations: list) -> str:
    """批量设置Word段落样式
    operations格式: [{"para_index": 0, "font_name": "SimSun", "bold": true}, ...]
    """
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from lxml import etree
    
    doc = Document(file_path)
    count = 0
    
    for op in operations:
        para_index = op.get("para_index")
        if para_index is None or para_index >= len(doc.paragraphs):
            continue
        
        para = doc.paragraphs[para_index]
        
        for run in para.runs:
            if "font_name" in op:
                rPr = run._element.get_or_add_rPr()
                for rf in rPr.findall(qn('w:rFonts')):
                    rPr.remove(rf)
                rFonts = etree.SubElement(rPr, qn('w:rFonts'))
                rFonts.set(qn('w:ascii'), op["font_name"])
                rFonts.set(qn('w:hAnsi'), op["font_name"])
                rFonts.set(qn('w:eastAsia'), op["font_name"])
                rFonts.set(qn('w:cs'), op["font_name"])
                rFonts.set(qn('w:hint'), 'eastAsia')
            
            if "font_size" in op:
                run.font.size = Pt(op["font_size"])
            if "bold" in op:
                run.font.bold = op["bold"]
            if "italic" in op:
                run.font.italic = op["italic"]
            if "underline" in op:
                run.font.underline = op["underline"]
            if "color" in op:
                r, g, b = int(op["color"][0:2], 16), int(op["color"][2:4], 16), int(op["color"][4:6], 16)
                run.font.color.rgb = RGBColor(r, g, b)
        
        count += 1
    
    doc.save(file_path)
    return f"已批量设置{count}个段落样式"


def _batch_set_style_xlsx(file_path: str, operations: list) -> str:
    """批量设置Excel单元格样式
    operations格式: [{"cell": "A1", "font_name": "SimSun", "bold": true}, ...]
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Color
    
    wb = load_workbook(file_path)
    ws = wb.active
    count = 0
    
    for op in operations:
        cell_ref = op.get("cell")
        if not cell_ref:
            continue
        
        try:
            cell = ws[cell_ref]
            
            font_kwargs = {}
            if "font_name" in op:
                font_kwargs['name'] = op["font_name"]
            if "font_size" in op:
                font_kwargs['size'] = op["font_size"]
            if "bold" in op:
                font_kwargs['bold'] = op["bold"]
            if "italic" in op:
                font_kwargs['italic'] = op["italic"]
            if "underline" in op and op["underline"]:
                font_kwargs['underline'] = 'single'
            if "color" in op:
                font_kwargs['color'] = Color(rgb=op["color"])
            
            if font_kwargs:
                cell.font = Font(**font_kwargs)
            
            if "bg_color" in op:
                cell.fill = PatternFill(start_color=op["bg_color"], end_color=op["bg_color"], fill_type="solid")
            
            count += 1
        except:
            pass
    
    wb.save(file_path)
    return f"已批量设置{count}个单元格样式"


# ==================== 文档质量检查 ====================

def _check_docx_quality(file_path: str) -> dict:
    """检查Word文档质量"""
    from docx import Document
    doc = Document(file_path)
    
    issues = []
    stats = {
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
        "images": 0,
        "empty_paragraphs": 0,
        "long_paragraphs": 0,
        "short_paragraphs": 0
    }
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        
        if not text:
            stats["empty_paragraphs"] += 1
            continue
        
        if len(text) > 500:
            stats["long_paragraphs"] += 1
            issues.append(f"段落{i}: 内容过长({len(text)}字符)")
        
        if len(text) < 5 and text not in ["", " ", "。", "，", ".", ","]:
            stats["short_paragraphs"] += 1
            issues.append(f"段落{i}: 内容过短({text})")
    
    # Count images
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            stats["images"] += 1
    
    # Quality score
    score = 100
    if stats["empty_paragraphs"] > 5:
        score -= 10
        issues.append(f"空段落过多({stats['empty_paragraphs']}个)")
    if stats["long_paragraphs"] > 3:
        score -= 10
        issues.append(f"长段落过多({stats['long_paragraphs']}个)")
    if stats["images"] == 0 and stats["paragraphs"] > 10:
        score -= 5
        issues.append("缺少图片")
    
    return {
        "score": max(0, score),
        "stats": stats,
        "issues": issues[:10],  # 最多返回10个问题
        "suggestions": _generate_suggestions(stats, issues)
    }


def _check_xlsx_quality(file_path: str) -> dict:
    """检查Excel文档质量"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path, data_only=True)
    
    issues = []
    stats = {
        "sheets": len(wb.sheetnames),
        "total_cells": 0,
        "empty_cells": 0,
        "merged_cells": 0,
        "formulas": 0
    }
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        stats["total_cells"] += ws.max_row * ws.max_column
        stats["merged_cells"] += len(ws.merged_cells.ranges)
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    stats["empty_cells"] += 1
                elif isinstance(cell.value, str) and cell.value.startswith("="):
                    stats["formulas"] += 1
    
    # Quality score
    score = 100
    empty_ratio = stats["empty_cells"] / max(stats["total_cells"], 1)
    if empty_ratio > 0.8:
        score -= 20
        issues.append(f"空单元格比例过高({empty_ratio:.1%})")
    if stats["formulas"] == 0 and stats["total_cells"] > 50:
        score -= 10
        issues.append("缺少公式")
    
    return {
        "score": max(0, score),
        "stats": stats,
        "issues": issues[:10],
        "suggestions": _generate_suggestions(stats, issues)
    }


def _generate_suggestions(stats: dict, issues: list) -> list:
    """生成改进建议"""
    suggestions = []
    
    if stats.get("empty_paragraphs", 0) > 5:
        suggestions.append("建议删除多余的空段落")
    if stats.get("long_paragraphs", 0) > 3:
        suggestions.append("建议将长段落拆分为多个短段落")
    if stats.get("images", 0) == 0:
        suggestions.append("建议添加图片以增强可读性")
    if stats.get("empty_cells", 0) > stats.get("total_cells", 0) * 0.5:
        suggestions.fill("建议填充空单元格或删除空行/列")
    if stats.get("formulas", 0) == 0:
        suggestions.append("考虑使用公式进行数据计算")
    
    return suggestions[:5]  # 最多返回5个建议


# ==================== 模板管理 ====================

def _extract_docx_template(file_path: str) -> dict:
    """提取Word文档模板"""
    from docx import Document
    from docx.shared import Pt
    
    doc = Document(file_path)
    
    template = {
        "styles": [],
        "structure": [],
        "fonts": set(),
        "sizes": set()
    }
    
    # Extract paragraph styles
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip():
            style_info = {
                "index": i,
                "text_preview": para.text[:50],
                "style_name": para.style.name if para.style else "Normal"
            }
            
            for run in para.runs:
                if run.font.name:
                    template["fonts"].add(run.font.name)
                if run.font.size:
                    template["sizes"].add(str(run.font.size))
            
            template["structure"].append(style_info)
    
    # Extract table info
    for i, table in enumerate(doc.tables):
        template["structure"].append({
            "type": "table",
            "index": i,
            "rows": len(table.rows),
            "columns": len(table.columns)
        })
    
    template["fonts"] = list(template["fonts"])
    template["sizes"] = list(template["sizes"])
    
    return template


def _apply_template(file_path: str, template: dict) -> str:
    """应用模板到文档"""
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from lxml import etree
    
    doc = Document(file_path)
    
    # Apply font to all paragraphs
    target_font = template.get("font_name", "SimSun")
    target_size = template.get("font_size", 12)
    
    for para in doc.paragraphs:
        for run in para.runs:
            rPr = run._element.get_or_add_rPr()
            for rf in rPr.findall(qn('w:rFonts')):
                rPr.remove(rf)
            rFonts = etree.SubElement(rPr, qn('w:rFonts'))
            rFonts.set(qn('w:ascii'), target_font)
            rFonts.set(qn('w:hAnsi'), target_font)
            rFonts.set(qn('w:eastAsia'), target_font)
            rFonts.set(qn('w:cs'), target_font)
            rFonts.set(qn('w:hint'), 'eastAsia')
            run.font.size = Pt(target_size)
    
    doc.save(file_path)
    return f"已应用模板: {target_font} {target_size}pt"


# ==================== 样式设置 ====================

def _file_set_cell_style_xlsx(
    file_path: str, sheet_name: Optional[str], row: int, column: int,
    font_name: Optional[str] = None, font_size: Optional[int] = None,
    bold: Optional[bool] = None, italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color: Optional[str] = None, bg_color: Optional[str] = None,
    align_h: Optional[str] = None, align_v: Optional[str] = None
) -> str:
    """设置Excel单元格样式"""
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Color
    
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    cell = ws.cell(row=row, column=column)
    
    if font_name or font_size or bold is not None or italic is not None or underline is not None or color:
        font_kwargs = {}
        if font_name:
            font_kwargs['name'] = font_name
        if font_size:
            font_kwargs['size'] = font_size
        if bold is not None:
            font_kwargs['bold'] = bold
        if italic is not None:
            font_kwargs['italic'] = italic
        if underline is not None and underline:
            font_kwargs['underline'] = 'single'
        if color:
            font_kwargs['color'] = Color(rgb=color)
        cell.font = Font(**font_kwargs)
    
    if bg_color:
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    
    if align_h or align_v:
        align_kwargs = {}
        if align_h:
            align_kwargs['horizontal'] = align_h
        if align_v:
            align_kwargs['vertical'] = align_v
        cell.alignment = Alignment(**align_kwargs)
    
    wb.save(file_path)
    return f"已设置单元格({row},{column})样式"


def _file_set_para_style_docx(
    file_path: str, para_index: int,
    font_name: Optional[str] = None, font_size: Optional[int] = None,
    bold: Optional[bool] = None, italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color: Optional[str] = None, align: Optional[str] = None
) -> str:
    """设置Word段落样式"""
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from lxml import etree
    
    doc = Document(file_path)
    if para_index >= len(doc.paragraphs):
        return f"段落索引{para_index}超出范围"
    
    para = doc.paragraphs[para_index]
    
    if align:
        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        para.alignment = align_map.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    
    for run in para.runs:
        if font_name:
            # 直接操作XML设置所有字体类型
            rPr = run._element.get_or_add_rPr()
            # 移除现有的rFonts
            for rf in rPr.findall(qn('w:rFonts')):
                rPr.remove(rf)
            # 创建新的rFonts，设置所有字体类型
            rFonts = etree.SubElement(rPr, qn('w:rFonts'))
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:eastAsia'), font_name)
            rFonts.set(qn('w:cs'), font_name)
            rFonts.set(qn('w:hint'), 'eastAsia')
        if font_size:
            run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
        if italic is not None:
            run.font.italic = italic
        if underline is not None:
            run.font.underline = underline
        if color:
            r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
            run.font.color.rgb = RGBColor(r, g, b)
    
    doc.save(file_path)
    return f"已设置段落{para_index}样式"


# ==================== WPS表格操作（COM） ====================

def _get_wps_app(app_type: str):
    """获取WPS COM应用对象"""
    import win32com.client
    
    prog_ids = {
        "word": ["Kwps.Application", "kwps.Application", "wps.Application", "Word.Application"],
        "excel": ["Ket.Application", "et.Application", "Excel.Application"],
        "powerpoint": ["Kwpp.Application", "kwpp.Application", "WPP.Application", "PowerPoint.Application"]
    }
    
    for prog_id in prog_ids.get(app_type, []):
        try:
            app = win32com.client.Dispatch(prog_id)
            return app
        except Exception:
            continue
    raise Exception(f"无法启动WPS {app_type}应用，请检查WPS是否已安装")


def _com_read_docx(file_path: str) -> dict:
    """通过COM读取Word文档"""
    app = _get_wps_app("word")
    try:
        app.Visible = False
        doc = app.Documents.Open(os.path.abspath(file_path))
        content = doc.Content.Text
        doc.Close()
        return {"content": content}
    finally:
        app.Quit()


def _com_write_docx(file_path: str, content: str) -> str:
    """通过COM写入Word文档"""
    app = _get_wps_app("word")
    try:
        app.Visible = False
        doc = app.Documents.Add()
        doc.Content.Text = content
        doc.SaveAs(os.path.abspath(file_path))
        doc.Close()
        return f"已保存: {file_path}"
    finally:
        app.Quit()


def _com_read_xlsx(file_path: str, sheet_name: Optional[str] = None) -> dict:
    """通过COM读取Excel"""
    app = _get_wps_app("excel")
    try:
        app.Visible = False
        wb = app.Workbooks.Open(os.path.abspath(file_path))
        if sheet_name:
            ws = wb.Sheets(sheet_name)
        else:
            ws = wb.ActiveSheet
        used_range = ws.UsedRange
        data = used_range.Value
        sheet_title = ws.Name
        wb.Close()
        return {"sheet_name": sheet_title, "data": data}
    finally:
        app.Quit()


def _com_write_xlsx(file_path: str, sheet_name: Optional[str], data: list) -> str:
    """通过COM写入Excel"""
    app = _get_wps_app("excel")
    try:
        app.Visible = False
        wb = app.Workbooks.Add()
        ws = wb.ActiveSheet
        if sheet_name:
            ws.Name = sheet_name
        for row_idx, row_data in enumerate(data, 1):
            for col_idx, value in enumerate(row_data, 1):
                ws.Cells(row_idx, col_idx).Value = value
        wb.SaveAs(os.path.abspath(file_path))
        wb.Close()
        return f"已保存: {file_path}"
    finally:
        app.Quit()


def _com_read_pptx(file_path: str) -> dict:
    """通过COM读取PowerPoint"""
    app = _get_wps_app("powerpoint")
    try:
        prs = app.Presentations.Open(os.path.abspath(file_path))
        slides_data = []
        for slide in prs.Slides:
            slide_content = []
            for shape in slide.Shapes:
                if shape.HasTextFrame:
                    slide_content.append(shape.TextFrame.TextRange.Text)
            slides_data.append(slide_content)
        prs.Close()
        return {"slides": slides_data}
    finally:
        app.Quit()


def _com_write_pptx(file_path: str, content: str) -> str:
    """通过COM写入PowerPoint"""
    app = _get_wps_app("powerpoint")
    try:
        prs = app.Presentations.Add()
        slide = prs.Slides.Add(1, 12)  # ppLayoutBlank
        shape = slide.Shapes.AddTextbox(1, 100, 100, 600, 100)
        shape.TextFrame.TextRange.Text = content
        prs.SaveAs(os.path.abspath(file_path))
        prs.Close()
        return f"已保存: {file_path}"
    finally:
        app.Quit()


def _com_insert_image_word(file_path: str, image_path: str) -> str:
    """通过COM在Word中插入图片"""
    app = _get_wps_app("word")
    try:
        app.Visible = False
        doc = app.Documents.Open(os.path.abspath(file_path))
        selection = app.Selection
        selection.InlineShapes.AddPicture(os.path.abspath(image_path))
        doc.Save()
        doc.Close()
        return f"已插入图片: {image_path}"
    finally:
        app.Quit()


def _com_insert_image_excel(file_path: str, sheet_name: Optional[str], image_path: str, range_addr: str) -> str:
    """通过COM在Excel中插入图片"""
    app = _get_wps_app("excel")
    try:
        app.Visible = False
        wb = app.Workbooks.Open(os.path.abspath(file_path))
        ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
        rng = ws.Range(range_addr)
        ws.Shapes.AddPicture(os.path.abspath(image_path), True, True, rng.Left, rng.Top, rng.Width, rng.Height)
        wb.Save()
        wb.Close()
        return f"已插入图片到{range_addr}"
    finally:
        app.Quit()


def _com_run_macro(file_path: str, macro_name: str, args: Optional[list] = None) -> str:
    """通过COM运行WPS宏"""
    app = _get_wps_app("word")
    try:
        app.Visible = False
        doc = app.Documents.Open(os.path.abspath(file_path))
        if args:
            result = app.Run(macro_name, *args)
        else:
            result = app.Run(macro_name)
        doc.Save()
        doc.Close()
        return f"宏执行结果: {result}"
    finally:
        app.Quit()


def _com_export_pdf(file_path: str, output_path: Optional[str] = None) -> str:
    """通过COM导出PDF"""
    ext = Path(file_path).suffix.lower()
    if ext in ['.doc', '.docx']:
        app = _get_wps_app("word")
        try:
            app.Visible = False
            doc = app.Documents.Open(os.path.abspath(file_path))
            if not output_path:
                output_path = str(Path(file_path).with_suffix('.pdf'))
            doc.ExportAsFixedFormat(output_path, 17)  # wdExportFormatPDF
            doc.Close()
            return f"已导出PDF: {output_path}"
        finally:
            app.Quit()
    elif ext in ['.xls', '.xlsx']:
        app = _get_wps_app("excel")
        try:
            app.Visible = False
            wb = app.Workbooks.Open(os.path.abspath(file_path))
            if not output_path:
                output_path = str(Path(file_path).with_suffix('.pdf'))
            wb.ExportAsFixedFormat(0, output_path)  # xlTypePDF
            wb.Close()
            return f"已导出PDF: {output_path}"
        finally:
            app.Quit()
    elif ext in ['.ppt', '.pptx']:
        app = _get_wps_app("powerpoint")
        try:
            prs = app.Presentations.Open(os.path.abspath(file_path))
            if not output_path:
                output_path = str(Path(file_path).with_suffix('.pdf'))
            prs.SaveAs(os.path.abspath(output_path), 32)  # ppSaveAsPDF
            prs.Close()
            return f"已导出PDF: {output_path}"
        finally:
            app.Quit()
    else:
        raise ValueError(f"不支持的文件格式: {ext}")


# ==================== 工具定义 ====================

TOOLS = [
    {
        "name": "read_document",
        "description": "读取Office文档内容（Word/Excel/PowerPoint）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"},
                "mode": {"type": "string", "enum": ["file", "com"], "description": "编辑模式", "default": "file"},
                "sheet_name": {"type": "string", "description": "Excel工作表名称"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "write_document",
        "description": "写入Office文档内容",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"},
                "content": {"type": "string", "description": "写入内容"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"},
                "sheet_name": {"type": "string", "description": "Excel工作表名称"},
                "append": {"type": "boolean", "description": "是否追加", "default": False}
            },
            "required": ["file_path", "content"]
        }
    },
    {
        "name": "edit_cell",
        "description": "编辑Excel单元格",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "row": {"type": "integer", "description": "行号"},
                "column": {"type": "integer", "description": "列号"},
                "value": {"type": "string", "description": "单元格值"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "row", "column", "value"]
        }
    },
    {
        "name": "get_sheet_names",
        "description": "获取Excel工作表名称列表",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "add_rows",
        "description": "在Excel中添加行",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "count": {"type": "integer", "description": "添加行数"},
                "after_row": {"type": "integer", "description": "在第几行后添加"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "count"]
        }
    },
    {
        "name": "add_columns",
        "description": "在Excel中添加列",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "count": {"type": "integer", "description": "添加列数"},
                "after_col": {"type": "integer", "description": "在第几列后添加"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "count"]
        }
    },
    {
        "name": "merge_cells",
        "description": "合并Excel单元格",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "start_row": {"type": "integer", "description": "起始行"},
                "start_col": {"type": "integer", "description": "起始列"},
                "end_row": {"type": "integer", "description": "结束行"},
                "end_col": {"type": "integer", "description": "结束列"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "start_row", "start_col", "end_row", "end_col"]
        }
    },
    {
        "name": "unmerge_cells",
        "description": "取消合并Excel单元格",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "start_row": {"type": "integer", "description": "起始行"},
                "start_col": {"type": "integer", "description": "起始列"},
                "end_row": {"type": "integer", "description": "结束行"},
                "end_col": {"type": "integer", "description": "结束列"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "start_row", "start_col", "end_row", "end_col"]
        }
    },
    {
        "name": "set_column_width",
        "description": "设置Excel列宽",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "column": {"type": "string", "description": "列字母（如A、B）"},
                "width": {"type": "number", "description": "宽度"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "column", "width"]
        }
    },
    {
        "name": "set_row_height",
        "description": "设置Excel行高",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "row": {"type": "integer", "description": "行号"},
                "height": {"type": "number", "description": "高度"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "row", "height"]
        }
    },
    {
        "name": "insert_image",
        "description": "插入图片到Office文档",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Office文件路径"},
                "image_path": {"type": "string", "description": "图片文件路径"},
                "width": {"type": "number", "description": "宽度（英寸）"},
                "height": {"type": "number", "description": "高度（英寸）"},
                "anchor": {"type": "string", "description": "Excel锚点单元格（如A1）", "default": "A1"},
                "slide_index": {"type": "integer", "description": "PowerPoint幻灯片索引（-1新建）", "default": -1},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "image_path"]
        }
    },
    {
        "name": "set_cell_style",
        "description": "设置Excel单元格样式",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "row": {"type": "integer", "description": "行号"},
                "column": {"type": "integer", "description": "列号"},
                "font_name": {"type": "string", "description": "字体名称"},
                "font_size": {"type": "integer", "description": "字体大小"},
                "bold": {"type": "boolean", "description": "是否加粗"},
                "italic": {"type": "boolean", "description": "是否斜体"},
                "underline": {"type": "boolean", "description": "是否下划线"},
                "color": {"type": "string", "description": "字体颜色（如FF0000）"},
                "bg_color": {"type": "string", "description": "背景颜色"},
                "align_h": {"type": "string", "description": "水平对齐（left/center/right）"},
                "align_v": {"type": "string", "description": "垂直对齐（top/center/bottom）"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "row", "column"]
        }
    },
    {
        "name": "set_paragraph_style",
        "description": "设置Word段落样式",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Word文件路径"},
                "para_index": {"type": "integer", "description": "段落索引"},
                "font_name": {"type": "string", "description": "字体名称（如SimHei、SimSun、KaiTi、FangSong）"},
                "font_size": {"type": "integer", "description": "字体大小"},
                "bold": {"type": "boolean", "description": "是否加粗"},
                "italic": {"type": "boolean", "description": "是否斜体"},
                "underline": {"type": "boolean", "description": "是否下划线"},
                "color": {"type": "string", "description": "字体颜色（如FF0000红色、0000FF蓝色）"},
                "align": {"type": "string", "description": "对齐方式（left/center/right/justify）"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "para_index"]
        }
    },
    {
        "name": "export_pdf",
        "description": "导出Office文档为PDF",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Office文件路径"},
                "output_path": {"type": "string", "description": "输出PDF路径"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "run_macro",
        "description": "运行WPS宏",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Office文件路径"},
                "macro_name": {"type": "string", "description": "宏名称"},
                "args": {"type": "array", "description": "宏参数", "items": {}}
            },
            "required": ["file_path", "macro_name"]
        }
    },
    {
        "name": "doc_to_ppt",
        "description": "根据Word文档自动生成专业PPT演示文稿",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_path": {"type": "string", "description": "Word文档路径"},
                "ppt_path": {"type": "string", "description": "输出PPT路径"},
                "template": {"type": "string", "description": "模板样式", "enum": ["ocean", "forest", "sunset", "royal", "minimal", "blue", "green"], "default": "ocean"}
            },
            "required": ["doc_path", "ppt_path"]
        }
    },
    {
        "name": "ppt_to_doc",
        "description": "读取PPT生成Word总结文档",
        "inputSchema": {
            "type": "object",
            "properties": {
                "ppt_path": {"type": "string", "description": "PPT文件路径"},
                "doc_path": {"type": "string", "description": "输出Word路径"}
            },
            "required": ["ppt_path", "doc_path"]
        }
    },
    {
        "name": "batch_read",
        "description": "批量读取文档内容（多个段落或单元格）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"},
                "indices": {"type": "array", "description": "Word段落索引列表或Excel单元格列表（如['A1','B2']）", "items": {}},
                "sheet_name": {"type": "string", "description": "Excel工作表名称"}
            },
            "required": ["file_path", "indices"]
        }
    },
    {
        "name": "batch_write",
        "description": "批量写入Excel单元格",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "data": {"type": "array", "description": "数据列表，格式: [{\"cell\":\"A1\",\"value\":\"xxx\"}]", "items": {}},
                "sheet_name": {"type": "string", "description": "工作表名称"}
            },
            "required": ["file_path", "data"]
        }
    },
    {
        "name": "batch_set_style",
        "description": "批量设置样式（Word段落或Excel单元格）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"},
                "operations": {"type": "array", "description": "操作列表", "items": {}},
                "sheet_name": {"type": "string", "description": "Excel工作表名称"}
            },
            "required": ["file_path", "operations"]
        }
    },
    {
        "name": "check_quality",
        "description": "检查文档质量并提供改进建议",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "extract_template",
        "description": "提取文档模板（字体、样式、结构）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "apply_template",
        "description": "应用模板到文档",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"},
                "font_name": {"type": "string", "description": "目标字体"},
                "font_size": {"type": "integer", "description": "目标字号"}
            },
            "required": ["file_path"]
        }
    }
]


def handle_tool_call(name: str, arguments: dict) -> Any:
    """处理工具调用"""
    if name == "read_document":
        file_path = os.path.abspath(arguments["file_path"])
        mode = arguments.get("mode", "file")
        sheet_name = arguments.get("sheet_name")
        ext = Path(file_path).suffix.lower()
        edit_mode = get_mode(mode)
        
        if ext in ['.docx', '.doc']:
            if edit_mode == EditMode.COM:
                return _com_read_docx(file_path)
            return _file_read_docx(file_path)
        elif ext in ['.xlsx', '.xls']:
            if edit_mode == EditMode.COM:
                return _com_read_xlsx(file_path, sheet_name)
            return _file_read_xlsx(file_path, sheet_name)
        elif ext in ['.pptx', '.ppt']:
            if edit_mode == EditMode.COM:
                return _com_read_pptx(file_path)
            return _file_read_pptx(file_path)
        else:
            raise ValueError(f"不支持的文件格式: {ext}")
    
    elif name == "write_document":
        file_path = os.path.abspath(arguments["file_path"])
        content = arguments["content"]
        mode = arguments.get("mode", "file")
        sheet_name = arguments.get("sheet_name")
        append = arguments.get("append", False)
        ext = Path(file_path).suffix.lower()
        edit_mode = get_mode(mode)
        
        os.makedirs(os.path.dirname(file_path) or '.', exist_ok=True)
        
        if ext in ['.docx', '.doc']:
            if edit_mode == EditMode.COM:
                return _com_write_docx(file_path, content)
            return _file_write_docx(file_path, content, append)
        elif ext in ['.xlsx', '.xls']:
            data = json.loads(content)
            if edit_mode == EditMode.COM:
                return _com_write_xlsx(file_path, sheet_name, data)
            return _file_write_xlsx(file_path, sheet_name, data)
        elif ext in ['.pptx', '.ppt']:
            if edit_mode == EditMode.COM:
                return _com_write_pptx(file_path, content)
            return _file_write_pptx(file_path, content)
        else:
            raise ValueError(f"不支持的文件格式: {ext}")
    
    elif name == "edit_cell":
        file_path = os.path.abspath(arguments["file_path"])
        row = arguments["row"]
        column = arguments["column"]
        value = arguments["value"]
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        edit_mode = get_mode(mode)
        
        if edit_mode == EditMode.COM:
            app = _get_wps_app("excel")
            try:
                app.Visible = False
                wb = app.Workbooks.Open(file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                ws.Cells(row, column).Value = value
                wb.Save()
                wb.Close()
                return f"已更新单元格 [{row},{column}] = {value}"
            finally:
                app.Quit()
        else:
            from openpyxl import load_workbook
            wb = load_workbook(file_path)
            ws = wb[sheet_name] if sheet_name else wb.active
            ws.cell(row=row, column=column, value=value)
            wb.save(file_path)
            return f"已更新单元格 [{row},{column}] = {value}"
    
    elif name == "get_sheet_names":
        file_path = os.path.abspath(arguments["file_path"])
        mode = arguments.get("mode", "file")
        edit_mode = get_mode(mode)
        
        if edit_mode == EditMode.COM:
            app = _get_wps_app("excel")
            try:
                app.Visible = False
                wb = app.Workbooks.Open(file_path)
                names = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
                wb.Close()
                return names
            finally:
                app.Quit()
        else:
            from openpyxl import load_workbook
            wb = load_workbook(file_path, read_only=True)
            return wb.sheetnames
    
    elif name == "add_rows":
        file_path = os.path.abspath(arguments["file_path"])
        count = arguments["count"]
        after_row = arguments.get("after_row")
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app = _get_wps_app("excel")
            try:
                app.Visible = False
                wb = app.Workbooks.Open(file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                if after_row:
                    ws.Rows(f"{after_row + 1}:{after_row + count}").Insert()
                else:
                    ws.Rows(f"{ws.UsedRange.Rows.Count + 1}:{ws.UsedRange.Rows.Count + count}").Insert()
                wb.Save()
                wb.Close()
                return f"已添加{count}行"
            finally:
                app.Quit()
        else:
            return _file_add_rows(file_path, sheet_name, count, after_row)
    
    elif name == "add_columns":
        file_path = os.path.abspath(arguments["file_path"])
        count = arguments["count"]
        after_col = arguments.get("after_col")
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app = _get_wps_app("excel")
            try:
                app.Visible = False
                wb = app.Workbooks.Open(file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                col_letter = chr(64 + after_col + 1) if after_col else chr(64 + ws.UsedRange.Columns.Count + 1)
                ws.Columns(f"{col_letter}:{chr(ord(col_letter) + count - 1)}").Insert()
                wb.Save()
                wb.Close()
                return f"已添加{count}列"
            finally:
                app.Quit()
        else:
            return _file_add_columns(file_path, sheet_name, count, after_col)
    
    elif name == "merge_cells":
        file_path = os.path.abspath(arguments["file_path"])
        start_row = arguments["start_row"]
        start_col = arguments["start_col"]
        end_row = arguments["end_row"]
        end_col = arguments["end_col"]
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app = _get_wps_app("excel")
            try:
                app.Visible = False
                wb = app.Workbooks.Open(file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                ws.Range(ws.Cells(start_row, start_col), ws.Cells(end_row, end_col)).Merge()
                wb.Save()
                wb.Close()
                return f"已合并单元格"
            finally:
                app.Quit()
        else:
            return _file_merge_cells(file_path, sheet_name, start_row, start_col, end_row, end_col)
    
    elif name == "unmerge_cells":
        file_path = os.path.abspath(arguments["file_path"])
        start_row = arguments["start_row"]
        start_col = arguments["start_col"]
        end_row = arguments["end_row"]
        end_col = arguments["end_col"]
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app = _get_wps_app("excel")
            try:
                app.Visible = False
                wb = app.Workbooks.Open(file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                ws.Range(ws.Cells(start_row, start_col), ws.Cells(end_row, end_col)).UnMerge()
                wb.Save()
                wb.Close()
                return f"已取消合并"
            finally:
                app.Quit()
        else:
            return _file_unmerge_cells(file_path, sheet_name, start_row, start_col, end_row, end_col)
    
    elif name == "set_column_width":
        file_path = os.path.abspath(arguments["file_path"])
        column = arguments["column"]
        width = arguments["width"]
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app = _get_wps_app("excel")
            try:
                app.Visible = False
                wb = app.Workbooks.Open(file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                ws.Columns(column).ColumnWidth = width
                wb.Save()
                wb.Close()
                return f"已设置列{column}宽度"
            finally:
                app.Quit()
        else:
            return _file_set_column_width(file_path, sheet_name, column, width)
    
    elif name == "set_row_height":
        file_path = os.path.abspath(arguments["file_path"])
        row = arguments["row"]
        height = arguments["height"]
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app = _get_wps_app("excel")
            try:
                app.Visible = False
                wb = app.Workbooks.Open(file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                ws.Rows(row).RowHeight = height
                wb.Save()
                wb.Close()
                return f"已设置行{row}高度"
            finally:
                app.Quit()
        else:
            return _file_set_row_height(file_path, sheet_name, row, height)
    
    elif name == "insert_image":
        file_path = os.path.abspath(arguments["file_path"])
        image_path = os.path.abspath(arguments["image_path"])
        width = arguments.get("width")
        height = arguments.get("height")
        anchor = arguments.get("anchor", "A1")
        slide_index = arguments.get("slide_index", -1)
        mode = arguments.get("mode", "file")
        ext = Path(file_path).suffix.lower()
        
        if mode == "com":
            if ext in ['.doc', '.docx']:
                return _com_insert_image_word(file_path, image_path)
            elif ext in ['.xls', '.xlsx']:
                sheet_name = arguments.get("sheet_name")
                return _com_insert_image_excel(file_path, sheet_name, image_path, anchor)
            elif ext in ['.ppt', '.pptx']:
                return _file_insert_image_pptx(file_path, image_path, slide_index, width or 1, 1, width or 4)
        else:
            if ext in ['.docx', '.doc']:
                return _file_insert_image_docx(file_path, image_path, width, height)
            elif ext in ['.xlsx', '.xls']:
                sheet_name = arguments.get("sheet_name")
                return _file_insert_image_xlsx(file_path, sheet_name, image_path, anchor)
            elif ext in ['.pptx', '.ppt']:
                return _file_insert_image_pptx(file_path, image_path, slide_index, width or 1, 1, width or 4)
            else:
                raise ValueError(f"不支持的文件格式: {ext}")
    
    elif name == "set_cell_style":
        file_path = os.path.abspath(arguments["file_path"])
        row = arguments["row"]
        column = arguments["column"]
        font_name = arguments.get("font_name")
        font_size = arguments.get("font_size")
        bold = arguments.get("bold")
        italic = arguments.get("italic")
        underline = arguments.get("underline")
        color = arguments.get("color")
        bg_color = arguments.get("bg_color")
        align_h = arguments.get("align_h")
        align_v = arguments.get("align_v")
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app = _get_wps_app("excel")
            try:
                app.Visible = False
                wb = app.Workbooks.Open(file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                cell = ws.Cells(row, column)
                if font_name:
                    cell.Font.Name = font_name
                if font_size:
                    cell.Font.Size = font_size
                if bold is not None:
                    cell.Font.Bold = bold
                if italic is not None:
                    cell.Font.Italic = italic
                if underline is not None:
                    cell.Font.Underline = underline
                if color:
                    cell.Font.Color = int(color, 16)
                if bg_color:
                    cell.Interior.Color = int(bg_color, 16)
                if align_h:
                    align_map = {"left": -4131, "center": -4108, "right": -4152}
                    cell.HorizontalAlignment = align_map.get(align_h, -4131)
                if align_v:
                    align_map = {"top": -4160, "center": -4108, "bottom": -4107}
                    cell.VerticalAlignment = align_map.get(align_v, -4160)
                wb.Save()
                wb.Close()
                return f"已设置单元格样式"
            finally:
                app.Quit()
        else:
            return _file_set_cell_style_xlsx(
                file_path, sheet_name, row, column,
                font_name, font_size, bold, italic, underline, color, bg_color, align_h, align_v
            )
    
    elif name == "set_paragraph_style":
        file_path = os.path.abspath(arguments["file_path"])
        para_index = arguments["para_index"]
        font_name = arguments.get("font_name")
        font_size = arguments.get("font_size")
        bold = arguments.get("bold")
        italic = arguments.get("italic")
        underline = arguments.get("underline")
        color = arguments.get("color")
        align = arguments.get("align")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app = _get_wps_app("word")
            try:
                app.Visible = False
                doc = app.Documents.Open(file_path)
                para = doc.Paragraphs(para_index + 1)
                if font_name:
                    para.Range.Font.Name = font_name
                if font_size:
                    para.Range.Font.Size = font_size
                if bold is not None:
                    para.Range.Font.Bold = bold
                if italic is not None:
                    para.Range.Font.Italic = italic
                if underline is not None:
                    para.Range.Font.Underline = underline
                if color:
                    para.Range.Font.Color = int(color, 16)
                if align:
                    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
                    para.Alignment = align_map.get(align, 0)
                doc.Save()
                doc.Close()
                return f"已设置段落样式"
            finally:
                app.Quit()
        else:
            return _file_set_para_style_docx(file_path, para_index, font_name, font_size, bold, italic, underline, color, align)
    
    elif name == "export_pdf":
        file_path = os.path.abspath(arguments["file_path"])
        output_path = arguments.get("output_path")
        if output_path:
            output_path = os.path.abspath(output_path)
        return _com_export_pdf(file_path, output_path)
    
    elif name == "run_macro":
        file_path = os.path.abspath(arguments["file_path"])
        macro_name = arguments["macro_name"]
        args = arguments.get("args")
        return _com_run_macro(file_path, macro_name, args)
    
    elif name == "doc_to_ppt":
        doc_path = os.path.abspath(arguments["doc_path"])
        ppt_path = os.path.abspath(arguments["ppt_path"])
        template = arguments.get("template", "default")
        os.makedirs(os.path.dirname(ppt_path) or '.', exist_ok=True)
        return _doc_to_ppt(doc_path, ppt_path, template)
    
    elif name == "ppt_to_doc":
        ppt_path = os.path.abspath(arguments["ppt_path"])
        doc_path = os.path.abspath(arguments["doc_path"])
        os.makedirs(os.path.dirname(doc_path) or '.', exist_ok=True)
        return _ppt_to_doc(ppt_path, doc_path)
    
    elif name == "batch_read":
        file_path = os.path.abspath(arguments["file_path"])
        indices = arguments["indices"]
        sheet_name = arguments.get("sheet_name")
        ext = Path(file_path).suffix.lower()
        
        if ext in ['.docx', '.doc']:
            return _batch_read_docx(file_path, indices)
        elif ext in ['.xlsx', '.xls']:
            return _batch_read_xlsx(file_path, indices, sheet_name)
        else:
            raise ValueError(f"batch_read不支持: {ext}")
    
    elif name == "batch_write":
        file_path = os.path.abspath(arguments["file_path"])
        data = arguments["data"]
        sheet_name = arguments.get("sheet_name")
        return _batch_write_xlsx(file_path, data, sheet_name)
    
    elif name == "batch_set_style":
        file_path = os.path.abspath(arguments["file_path"])
        operations = arguments["operations"]
        sheet_name = arguments.get("sheet_name")
        ext = Path(file_path).suffix.lower()
        
        if ext in ['.docx', '.doc']:
            return _batch_set_style_docx(file_path, operations)
        elif ext in ['.xlsx', '.xls']:
            return _batch_set_style_xlsx(file_path, operations)
        else:
            raise ValueError(f"batch_set_style不支持: {ext}")
    
    elif name == "check_quality":
        file_path = os.path.abspath(arguments["file_path"])
        ext = Path(file_path).suffix.lower()
        
        if ext in ['.docx', '.doc']:
            return _check_docx_quality(file_path)
        elif ext in ['.xlsx', '.xls']:
            return _check_xlsx_quality(file_path)
        else:
            raise ValueError(f"check_quality不支持: {ext}")
    
    elif name == "extract_template":
        file_path = os.path.abspath(arguments["file_path"])
        return _extract_docx_template(file_path)
    
    elif name == "apply_template":
        file_path = os.path.abspath(arguments["file_path"])
        template = {
            "font_name": arguments.get("font_name", "SimSun"),
            "font_size": arguments.get("font_size", 12)
        }
        return _apply_template(file_path, template)
    
    else:
        raise ValueError(f"未知工具: {name}")


# ==================== MCP协议实现 ====================

class MCPServer:
    """简化版MCP服务器，基于JSON-RPC 2.0 over stdio"""
    
    def __init__(self):
        self.server_info = {
            "name": "wps-editor",
            "version": "2.0.0"
        }
        self.capabilities = {
            "tools": {}
        }
    
    def create_response(self, id: Any, result: Any) -> dict:
        return {
            "jsonrpc": "2.0",
            "id": id,
            "result": result
        }
    
    def create_error(self, id: Any, code: int, message: str) -> dict:
        return {
            "jsonrpc": "2.0",
            "id": id,
            "error": {
                "code": code,
                "message": message
            }
        }
    
    def handle_request(self, request: dict) -> dict:
        method = request.get("method")
        params = request.get("params", {})
        id = request.get("id")
        
        try:
            if method == "initialize":
                return self.create_response(id, {
                    "protocolVersion": "2024-11-05",
                    "capabilities": self.capabilities,
                    "serverInfo": self.server_info
                })
            
            elif method == "tools/list":
                return self.create_response(id, {
                    "tools": TOOLS
                })
            
            elif method == "tools/call":
                tool_name = params.get("name")
                arguments = params.get("arguments", {})
                result = handle_tool_call(tool_name, arguments)
                return self.create_response(id, {
                    "content": [
                        {
                            "type": "text",
                            "text": json.dumps(result, ensure_ascii=False, indent=2)
                        }
                    ]
                })
            
            elif method == "notifications/initialized":
                return None
            
            else:
                return self.create_error(id, -32601, f"方法不存在: {method}")
        
        except Exception as e:
            return self.create_error(id, -32000, str(e))
    
    def run(self):
        """运行MCP服务器（stdio模式）"""
        sys.stderr.write("WPS Editor MCP Server v2.0.0 启动中...\n")
        sys.stderr.flush()
        
        for line in sys.stdin:
            line = line.strip()
            if not line:
                continue
            
            try:
                request = json.loads(line)
                response = self.handle_request(request)
                if response:
                    output = json.dumps(response, ensure_ascii=False)
                    sys.stdout.write(output + "\n")
                    sys.stdout.flush()
            except json.JSONDecodeError as e:
                sys.stderr.write(f"JSON解析错误: {e}\n")
                sys.stderr.flush()
            except Exception as e:
                sys.stderr.write(f"处理请求错误: {e}\n")
                sys.stderr.flush()


# 入口点
def main():
    server = MCPServer()
    server.run()


if __name__ == "__main__":
    main()
