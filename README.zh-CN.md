# onework - 对话式文档协作

[English](README.md) | [简体中文](#概述)

---

## 概述

**onework** 是一个 AI 驱动的智能文档处理系统，作为各类文档格式与 AI 能力之间的桥梁。

核心理念：**任意格式 → Markdown 中间表示 → AI 处理 → 目标格式**

## 功能特性

| 功能 | 描述 |
|------|------|
| **读取** | Word/PDF/PPT/图片 → 结构化 Markdown |
| **结构** | 提取大纲、表格、关键信息 |
| **更新** | 精准编辑特定章节 |
| **转换** | Word ↔ PDF ↔ PPT ↔ HTML ↔ Markdown |
| **生成** | 根据大纲生成完整文档 |

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt

# Python 中使用
from onework import read_document, convert_format

# 读取任意文档
result = read_document("/path/to/report.docx")
print(result["content"])

# 转换为 PDF
result = convert_format("/path/to/report.docx", "pdf")
```

## 支持格式

### 输入
- **Word** (.docx) - 使用 mammoth
- **PDF** (.pdf) - 使用 PyMuPDF
- **PowerPoint** (.pptx) - 使用 python-pptx
- **图片** (.png, .jpg) - OCR 使用 PaddleOCR
- **Markdown** (.md)

### 输出
- **Word** (.docx) - python-docx
- **PDF** (.pdf) - weasyprint
- **PowerPoint** (.pptx) - python-pptx
- **HTML** (.html) - markdown
- **Markdown** (.md)

## 工作流程

```
┌─────────────┐    ┌──────────────┐    ┌──────────────┐    ┌─────────────┐
│    输入      │ →  │ Markdown IR │ →  │   AI 处理    │ →  │    输出     │
│ (任意格式)   │    │   (统一)     │    │  (提示词)    │    │  (目标格式)  │
└─────────────┘    └──────────────┘    └──────────────┘    └─────────────┘
```

## 工具函数

| 工具 | 描述 |
|------|------|
| `read_document` | 解析文档为带结构的 Markdown |
| `extract_structure` | 提取文档大纲 |
| `update_section` | 更新特定章节 |
| `convert_format` | 带模板的格式转换 |
| `create_from_outline` | 根据大纲生成文档 |

## 模板系统

将模板文件放入 `templates/` 目录：

```
templates/
├── docx/
│   ├── default.docx      # 默认模板
│   ├── report.docx       # 报告模板
│   └── contract.docx     # 合同模板
├── pptx/
│   ├── presentation.pptx # 演示模板
│   ├── pitch.pptx        # 路演模板
│   └── training.pptx     # 培训模板
└── pdf/
    └── default.html      # PDF 渲染模板
```

## 作为 Skill 安装

```bash
# 克隆仓库
git clone https://github.com/yourname/onework.git
cd onework

# 安装
pip install -r requirements.txt
```

然后在你的 AI 助手平台中配置使用。

## 使用示例

### 示例 1：Word 转 PPT

```python
from onework import read_document, convert_format

# 读取 Word 文档
doc = read_document("report.docx")

# 转换为 PPT 演示文稿
convert_format("report.docx", "pptx", template="presentation")
```

### 示例 2：根据大纲生成

```python
from onework import create_from_outline

outline = """
# 产品发布会

## 市场分析
- 目标用户
- 竞争格局

## 产品功能
- 核心能力
- 独特优势

## 下一步计划
"""

create_from_outline(outline, "pptx")
```

### 示例 3：对话式协作

```
用户：帮我把这份季度报告转成 PPT
AI：已读取报告，共 5 章。建议 PPT 结构：
    - 封面
    - 执行摘要（1页）
    - 核心发现（3页）
    - 建议方案（2页）
    - 下一步（1页）
    确认吗？

用户：执行摘要改成 2 页，加上数据图表
AI：好的，正在生成...
    ✅ PPT 已生成：report_presentation.pptx
    共 9 页，包含 3 个数据图表
```

## 依赖项

核心依赖：
- `mammoth` - Word 解析
- `pymupdf` - PDF 处理
- `python-pptx` - PowerPoint 处理
- `python-docx` - Word 生成
- `markdown` - Markdown 处理

可选：
- `weasyprint` - PDF 生成
- `paddleocr` - 图片 OCR

## 项目架构

```
onework/
├── skill.yaml              # Skill 配置
├── onework/
│   ├── __init__.py         # 包入口
│   └── tools.py            # 核心工具
├── templates/              # 模板文件夹
├── examples/               # 示例文件
├── requirements.txt        # 依赖
├── README.md               # 英文文档
└── README.zh-CN.md         # 中文文档
```

## 技术选型

| 功能 | 开源工具 | 许可证 |
|------|----------|--------|
| Word → MD | mammoth | BSD |
| PDF → MD | pymupdf | AGPL/商业 |
| PPT → MD | python-pptx | MIT |
| MD → Word | python-docx | MIT |
| MD → PDF | weasyprint | BSD |
| MD → PPT | python-pptx | MIT |

## 许可证

MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 贡献

欢迎贡献代码！请先阅读 [CONTRIBUTING.md](CONTRIBUTING.md)

---

<p align="center">
  <sub>用 ❤️ 打造，让文档协作更智能</sub>
</p>
