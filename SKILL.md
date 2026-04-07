---
name: onework
description: "对话式文档协作 - 读取、优化、转换、生成 Word/PPT/PDF 文档"
description_en: "Conversational Document Collaboration - Read, optimize, convert, and generate Word/PPT/PDF"
description_zh: "对话式文档协作 - 使用 AI 读取、优化、转换、生成 Word/PPT/PDF 文档"
---

# onework

对话式文档协作 Skill / Conversational Document Collaboration

## 功能 | Features

- **read_document** - 读取文档并返回结构化信息（支持 Word/PDF/PPT/图片）
- **extract_structure** - 提取文档的大纲结构（章节层级）
- **update_section** - 更新文档中的特定章节
- **convert_format** - 格式转换（Word↔PDF↔PPT↔HTML 互转）
- **create_from_outline** - 根据大纲生成完整文档

## 使用示例 | Examples

```
# 读取文档
read_document("报告.docx")

# 提取结构
extract_structure("报告.docx")

# 格式转换
convert_format("报告.md", "pptx")

# 从大纲生成
create_from_outline("大纲内容", output_format="docx")
```

## 安装 | Installation

```bash
pip install -r requirements.txt
```

## 模板 | Templates

- `templates/docx/` - Word 模板
- `templates/pptx/` - PPT 模板
- `templates/pdf/` - PDF 模板