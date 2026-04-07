"""onework - 对话式文档协作 Skill

支持多格式文档的读取、编辑、转换、生成，内置模板系统。
"""

__version__ = "1.0.0"
__author__ = "onework"

# 导出核心工具函数
from .tools import (
    read_document,
    extract_structure,
    update_section,
    convert_format,
    create_from_outline,
    list_templates,
)

__all__ = [
    "read_document",
    "extract_structure",
    "update_section",
    "convert_format",
    "create_from_outline",
    "list_templates",
]