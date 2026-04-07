"""Internationalization support for onework"""

import os

# Default language
DEFAULT_LANG = "en"

# Get language from environment or use default
def get_lang():
    return os.getenv("ONEWORK_LANG", DEFAULT_LANG)


MESSAGES = {
    "en": {
        # Errors
        "file_not_found": "File not found: {path}",
        "unsupported_format": "Unsupported format: {ext}",
        "install_mammoth": "Please install mammoth: pip install mammoth",
        "install_pymupdf": "Please install pymupdf: pip install pymupdf",
        "install_python_pptx": "Please install python-pptx: pip install python-pptx",
        "install_markdown": "Please install markdown: pip install markdown",
        "install_python_docx": "Please install python-docx: pip install python-docx",
        "install_weasyprint": "Please install weasyprint: pip install weasyprint",
        "read_word_failed": "Failed to read Word: {error}",
        "read_pdf_failed": "Failed to read PDF: {error}",
        "read_ppt_failed": "Failed to read PPT: {error}",
        "read_md_failed": "Failed to read Markdown: {error}",
        "ocr_failed": "OCR failed: {error}",
        "convert_success": "Conversion successful: {path}",
        "section_updated": "Section '{section}' updated, new document: {path}",
    },
    "zh": {
        # 错误信息
        "file_not_found": "文件不存在: {path}",
        "unsupported_format": "不支持的格式: {ext}",
        "install_mammoth": "请安装 mammoth: pip install mammoth",
        "install_pymupdf": "请安装 pymupdf: pip install pymupdf",
        "install_python_pptx": "请安装 python-pptx: pip install python-pptx",
        "install_markdown": "请安装 markdown: pip install markdown",
        "install_python_docx": "请安装 python-docx: pip install python-docx",
        "install_weasyprint": "请安装 weasyprint: pip install weasyprint",
        "read_word_failed": "读取 Word 失败: {error}",
        "read_pdf_failed": "读取 PDF 失败: {error}",
        "read_ppt_failed": "读取 PPT 失败: {error}",
        "read_md_failed": "读取 Markdown 失败: {error}",
        "ocr_failed": "OCR 识别失败: {error}",
        "convert_success": "转换成功: {path}",
        "section_updated": "章节 '{section}' 已更新，新文档: {path}",
    }
}


def _(key, **kwargs):
    """Get translated message"""
    lang = get_lang()
    message = MESSAGES.get(lang, MESSAGES["en"]).get(key, key)
    return message.format(**kwargs) if kwargs else message
