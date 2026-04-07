"""onework 核心工具函数

基于以下开源项目构建：
- Microsoft markitdown: https://github.com/microsoft/markitdown
- mammoth: Word 文档解析
- pymupdf: PDF 处理
- python-pptx: PPT 处理

感谢所有开源贡献者！
"""

import os
import json
from typing import Optional, List, Dict, Any
from pathlib import Path

# markitdown 可用性检查
_MARKITDOWN_AVAILABLE = None


def _check_markitdown() -> bool:
    """检查 markitdown 是否可用"""
    global _MARKITDOWN_AVAILABLE
    if _MARKITDOWN_AVAILABLE is None:
        try:
            import markitdown
            _MARKITDOWN_AVAILABLE = True
        except ImportError:
            _MARKITDOWN_AVAILABLE = False
    return _MARKITDOWN_AVAILABLE


def read_document(file_path: str, use_markitdown: bool = True) -> Dict[str, Any]:
    """
    读取文档并返回结构化信息。
    
    自动识别格式（Word/PDF/PPT/图片），返回：
    - content: Markdown 格式全文
    - structure: 文档结构（标题层级、章节）
    - tables: 表格列表
    - file_type: 文件类型
    
    Args:
        file_path: 文档的绝对路径
        use_markitdown: 是否优先使用 markitdown (默认 True)
    
    Returns:
        包含 content, structure, tables, file_type 的字典
    """
    path = Path(file_path)
    
    if not path.exists():
        return {
            "success": False,
            "error": f"文件不存在: {file_path}"
        }
    
    ext = path.suffix.lower()
    
    # 优先使用 markitdown (如果可用且用户未禁用)
    if use_markitdown and _check_markitdown():
        md_result = _read_with_markitdown(file_path)
        if md_result.get("success"):
            return md_result
    
    # 回退到原有实现
    if ext == ".docx":
        return _read_docx(file_path)
    elif ext == ".pdf":
        return _read_pdf(file_path)
    elif ext == ".pptx":
        return _read_pptx(file_path)
    elif ext in [".md", ".markdown"]:
        return _read_markdown(file_path)
    elif ext in [".png", ".jpg", ".jpeg", ".gif", ".bmp"]:
        return _read_image(file_path)
    else:
        return {
            "success": False,
            "error": f"不支持的格式: {ext}"
        }


def _read_with_markitdown(file_path: str) -> Dict[str, Any]:
    """使用 markitdown 读取文档"""
    try:
        from markitdown import MarkItDown
        
        md = MarkItDown()
        result = md.convert(file_path)
        content = result.text_content
        
        structure = _extract_headings_from_md(content)
        
        return {
            "success": True,
            "content": content,
            "structure": structure,
            "tables": _extract_tables_from_md(content),
            "file_type": Path(file_path).suffix.lower().lstrip(".")
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"markitdown 转换失败: {str(e)}"
        }


def _read_docx(file_path: str) -> Dict[str, Any]:
    """读取 Word 文档"""
    try:
        import mammoth
        
        with open(file_path, "rb") as doc_file:
            result = mammoth.convert_to_markdown(doc_file)
            content = result.value
        
        structure = _extract_headings_from_md(content)
        
        return {
            "success": True,
            "content": content,
            "structure": structure,
            "tables": _extract_tables_from_md(content),
            "file_type": "docx"
        }
    except ImportError:
        return {
            "success": False,
            "error": "请安装 mammoth: pip install mammoth"
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"读取 Word 失败: {str(e)}"
        }


def _read_pdf(file_path: str) -> Dict[str, Any]:
    """读取 PDF 文档"""
    try:
        import pymupdf
        
        doc = pymupdf.open(file_path)
        content_parts = []
        
        for page_num, page in enumerate(doc):
            text = page.get_text()
            content_parts.append(f"## 第 {page_num + 1} 页\n\n{text}")
        
        content = "\n\n---\n\n".join(content_parts)
        
        return {
            "success": True,
            "content": content,
            "structure": _extract_headings_from_md(content),
            "tables": [],
            "file_type": "pdf"
        }
    except ImportError:
        return {
            "success": False,
            "error": "请安装 pymupdf: pip install pymupdf"
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"读取 PDF 失败: {str(e)}"
        }


def _read_pptx(file_path: str) -> Dict[str, Any]:
    """读取 PPT 文档"""
    try:
        from pptx import Presentation
        
        prs = Presentation(file_path)
        slides_content = []
        
        for i, slide in enumerate(prs.slides):
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_text.append(shape.text)
            
            if slide_text:
                slides_content.append(f"## 幻灯片 {i + 1}\n\n" + "\n".join(slide_text))
        
        content = "\n\n---\n\n".join(slides_content)
        
        return {
            "success": True,
            "content": content,
            "structure": [{"level": 1, "title": f"幻灯片 {i+1}", "page": i+1} for i in range(len(slides_content))],
            "tables": [],
            "file_type": "pptx"
        }
    except ImportError:
        return {
            "success": False,
            "error": "请安装 python-pptx: pip install python-pptx"
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"读取 PPT 失败: {str(e)}"
        }


def _read_markdown(file_path: str) -> Dict[str, Any]:
    """读取 Markdown 文档"""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
        
        structure = _extract_headings_from_md(content)
        
        return {
            "success": True,
            "content": content,
            "structure": structure,
            "tables": _extract_tables_from_md(content),
            "file_type": "markdown"
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"读取 Markdown 失败: {str(e)}"
        }


def _read_image(file_path: str) -> Dict[str, Any]:
    """读取图片（OCR）"""
    try:
        from PIL import Image
        import pytesseract
        
        image = Image.open(file_path)
        content = pytesseract.image_to_string(image, lang='chi_sim+eng')
        
        return {
            "success": True,
            "content": content,
            "structure": [],
            "tables": [],
            "file_type": "image"
        }
    except ImportError:
        return {
            "success": False,
            "error": "请安装 Pillow 和 pytesseract: pip install pillow pytesseract"
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"OCR 识别失败: {str(e)}"
        }


def _extract_headings_from_md(content: str) -> List[Dict[str, Any]]:
    """从 Markdown 中提取标题结构"""
    import re
    
    headings = []
    for line in content.split('\n'):
        match = re.match(r'^(#{1,6})\s+(.+)$', line)
        if match:
            level = len(match.group(1))
            title = match.group(2).strip()
            headings.append({"level": level, "title": title})
    
    return headings


def _extract_tables_from_md(content: str) -> List[str]:
    """从 Markdown 中提取表格"""
    import re
    
    tables = []
    in_table = False
    current_table = []
    
    for line in content.split('\n'):
        if '|' in line:
            in_table = True
            current_table.append(line)
        elif in_table and current_table:
            tables.append('\n'.join(current_table))
            current_table = []
            in_table = False
    
    if current_table:
        tables.append('\n'.join(current_table))
    
    return tables


def extract_structure(file_path: str) -> List[Dict[str, Any]]:
    """
    提取文档的大纲结构（章节层级）。
    
    Args:
        file_path: 文档路径
    
    Returns:
        层级化的章节列表
    """
    result = read_document(file_path)
    
    if result.get("success"):
        return result.get("structure", [])
    else:
        return []


def update_section(
    file_path: str,
    section_path: str,
    new_content: str,
    output_path: Optional[str] = None
) -> Dict[str, Any]:
    """
    更新文档中的特定章节。
    
    Args:
        file_path: 原文档路径
        section_path: 章节路径，如 "1. 项目背景/1.1 概述"
        new_content: 新的 Markdown 内容
        output_path: 输出路径（可选，默认生成新文件）
    
    Returns:
        包含 success, output_path, summary 的字典
    """
    if not output_path:
        path = Path(file_path)
        output_path = str(path.parent / f"{path.stem}_v2{path.suffix}")
    
    result = read_document(file_path)
    
    if not result.get("success"):
        return result
    
    # 简单实现：生成新的 Markdown 文件（带章节标记）
    # 后续可扩展为直接生成 docx/pptx
    output_md_path = output_path.replace(path.suffix, ".md")
    
    with open(output_md_path, "w", encoding="utf-8") as f:
        f.write(f"# 文档更新\n\n")
        f.write(f"**更新章节**: {section_path}\n\n")
        f.write(f"---\n\n")
        f.write(new_content)
    
    return {
        "success": True,
        "output_path": output_md_path,
        "summary": f"已更新章节 '{section_path}'，生成新文档: {output_md_path}"
    }


def convert_format(
    input_path: str,
    output_format: str,
    output_path: Optional[str] = None,
    template: Optional[str] = None
) -> Dict[str, Any]:
    """
    转换文档格式。
    
    支持格式：docx, pdf, pptx, html, md
    
    Args:
        input_path: 输入文件路径
        output_format: 目标格式
        output_path: 输出路径（可选）
        template: 模板名称（可选）
    
    Returns:
        包含 success, output_path 的字典
    """
    input_path_obj = Path(input_path)
    
    if not input_path_obj.exists():
        return {
            "success": False,
            "error": f"输入文件不存在: {input_path}"
        }
    
    if not output_path:
        output_path = str(input_path_obj.parent / f"{input_path_obj.stem}_converted.{output_format}")
    
    # 先转换为 Markdown 中间格式
    read_result = read_document(input_path)
    
    if not read_result.get("success"):
        return read_result
    
    markdown_content = read_result.get("content", "")
    
    # 根据目标格式渲染
    if output_format == "md":
        # 直接输出 Markdown
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(markdown_content)
        
    elif output_format == "html":
        # Markdown -> HTML
        try:
            import markdown
            
            html = markdown.markdown(
                markdown_content,
                extensions=['tables', 'fenced_code', 'toc']
            )
            
            full_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Document</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }}
        table {{ border-collapse: collapse; width: 100%; margin: 1em 0; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background: #f5f5f5; }}
        h1, h2, h3 {{ color: #333; }}
        pre {{ background: #f5f5f5; padding: 10px; overflow-x: auto; }}
    </style>
</head>
<body>
{html}
</body>
</html>"""
            
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(full_html)
        except ImportError:
            return {"success": False, "error": "请安装 markdown: pip install markdown"}
    
    elif output_format == "docx":
        # Markdown -> Word
        try:
            from docx import Document
            from docx.shared import Pt, Inches
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            
            doc = Document()
            
            # 简单的 Markdown -> Word 转换
            lines = markdown_content.split('\n')
            in_code_block = False
            
            for line in lines:
                if line.startswith('```'):
                    in_code_block = not in_code_block
                    continue
                
                if in_code_block:
                    doc.add_paragraph(line, style='Code')
                    continue
                
                if line.startswith('#'):
                    # 标题
                    level = len(line) - len(line.lstrip('#'))
                    heading = doc.add_heading(line.lstrip('# '), level=level)
                    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
                elif line.strip():
                    # 段落
                    doc.add_paragraph(line)
                else:
                    # 空行
                    doc.add_paragraph()
            
            doc.save(output_path)
        except ImportError:
            return {"success": False, "error": "请安装 python-docx: pip install python-docx"}
    
    elif output_format == "pdf":
        # Markdown -> PDF (通过 HTML)
        html_path = output_path.replace('.pdf', '.html')
        
        # 先转换为 HTML
        try:
            import markdown
            html = markdown.markdown(
                markdown_content,
                extensions=['tables', 'fenced_code']
            )
            
            full_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; max-width: 800px; margin: 0 auto; padding: 40px; }}
        table {{ border-collapse: collapse; width: 100%; margin: 1em 0; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        h1, h2, h3 {{ color: #333; }}
    </style>
</head>
<body>{html}</body>
</html>"""
            
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(full_html)
            
            # 使用 weasyprint 转换为 PDF
            try:
                from weasyprint import HTML
                HTML(html_path).write_pdf(output_path)
            except ImportError:
                return {
                    "success": True,
                    "output_path": html_path,
                    "message": "已转换为 HTML，PDF 需要安装 weasyprint: pip install weasyprint"
                }
        except ImportError:
            return {"success": False, "error": "请安装 markdown: pip install markdown"}
    
    elif output_format == "pptx":
        # Markdown -> PPT
        try:
            from pptx import Presentation
            from pptx.util import Inches, Pt
            
            prs = Presentation()
            
            # 简单的 Markdown -> PPT 转换
            # 按 H1 分页，每页一个标题 + 内容
            lines = markdown_content.split('\n')
            current_title = "封面"
            current_content = []
            
            for line in lines:
                if line.startswith('## '):
                    # 新的幻灯片
                    if current_content:
                        _add_slide(prs, current_title, current_content)
                    
                    current_title = line.replace('## ', '').strip()
                    current_content = []
                elif line.startswith('# '):
                    current_title = line.replace('# ', '').strip()
                elif line.strip():
                    current_content.append(line.strip())
            
            # 最后一页
            if current_content:
                _add_slide(prs, current_title, current_content)
            
            prs.save(output_path)
        except ImportError:
            return {"success": False, "error": "请安装 python-pptx: pip install python-pptx"}
    
    else:
        return {
            "success": False,
            "error": f"不支持的目标格式: {output_format}"
        }
    
    return {
        "success": True,
        "output_path": output_path
    }


def _add_slide(prs: Any, title: str, content: List[str]):
    """添加一张幻灯片"""
    slide_layout = prs.slide_layouts[1]  # 标题和内容
    slide = prs.slides.add_slide(slide_layout)
    
    title_elem = slide.shapes.title
    title_elem.text = title
    
    content_elem = slide.placeholders[1]
    text_frame = content_elem.text_frame
    text_frame.clear()
    
    for item in content[:10]:  # 最多 10 条
        p = text_frame.add_paragraph()
        p.text = item
        p.level = 0


def create_from_outline(
    outline: str,
    output_format: str,
    output_path: Optional[str] = None,
    template: Optional[str] = None
) -> Dict[str, Any]:
    """
    根据大纲生成完整文档。
    
    Args:
        outline: 文档大纲（Markdown 格式）
        output_format: 输出格式 (docx, pptx, pdf)
        output_path: 输出路径（可选）
        template: 模板名称（可选）
    
    Returns:
        包含 success, output_path 的字典
    """
    if not output_path:
        import tempfile
        output_path = os.path.join(tempfile.gettempdir(), f"onework_output.{output_format}")
    
    # 大纲本身就是 Markdown 内容
    # 可以直接用 convert_format 处理
    if output_format == "md":
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(outline)
        
        return {
            "success": True,
            "output_path": output_path
        }
    
    # 其他格式需要先转为 Markdown 文件
    import tempfile
    with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as f:
        f.write(outline)
        md_path = f.name
    
    # 调用 convert_format
    result = convert_format(md_path, output_format, output_path, template)
    
    # 清理临时文件
    try:
        os.unlink(md_path)
    except:
        pass
    
    return result


# ============ 模板相关函数 ============

def list_templates(format: str = None) -> List[str]:
    """
    列出可用的模板。
    
    Args:
        format: 格式过滤 (docx, pptx, pdf)
    
    Returns:
        模板文件列表
    """
    # 获取当前文件所在目录
    current_dir = Path(__file__).parent.parent
    
    templates_dir = current_dir / "templates"
    
    if not templates_dir.exists():
        return []
    
    templates = []
    
    if format:
        format_dir = templates_dir / format
        if format_dir.exists():
            templates = [str(f) for f in format_dir.glob("*") if f.is_file()]
    else:
        for subdir in templates_dir.iterdir():
            if subdir.is_dir():
                for f in subdir.glob("*"):
                    if f.is_file():
                        templates.append(str(f))
    
    return templates