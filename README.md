# onework - Conversational Document Collaboration

[English](#overview) | [简体中文](README.zh-CN.md)

---

## Overview

**onework** is an AI-driven document processing system that serves as a bridge between various document formats and AI capabilities.

Core philosophy: **Any format → Markdown IR → AI processing → Target format**

## Features

| Feature | Description |
|---------|-------------|
| **Read** | Word/PDF/PPT/Images → Structured Markdown |
| **Structure** | Extract headings, tables, key information |
| **Update** | Edit specific sections with precision |
| **Convert** | Word ↔ PDF ↔ PPT ↔ HTML ↔ Markdown |
| **Generate** | Create documents from outlines |

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Use in Python
from onework import read_document, convert_format

# Read any document
result = read_document("/path/to/report.docx")
print(result["content"])

# Convert to PDF
result = convert_format("/path/to/report.docx", "pdf")
```

## Supported Formats

### Input
- **Word** (.docx) - via mammoth
- **PDF** (.pdf) - via PyMuPDF
- **PowerPoint** (.pptx) - via python-pptx
- **Images** (.png, .jpg) - OCR via PaddleOCR
- **Markdown** (.md)

### Output
- **Word** (.docx) - python-docx
- **PDF** (.pdf) - weasyprint
- **PowerPoint** (.pptx) - python-pptx
- **HTML** (.html) - markdown
- **Markdown** (.md)

## Workflow

```
┌─────────────┐    ┌──────────────┐    ┌──────────────┐    ┌─────────────┐
│   Input     │ →  │  Markdown IR │ →  │  AI Process  │ →  │   Output    │
│  (any fmt)  │    │  (unified)   │    │  (prompts)   │    │ (target fmt)│
└─────────────┘    └──────────────┘    └──────────────┘    └─────────────┘
```

## Tools

| Tool | Description |
|------|-------------|
| `read_document` | Parse document to Markdown with structure |
| `extract_structure` | Extract document outline |
| `update_section` | Update specific chapter/section |
| `convert_format` | Convert between formats with templates |
| `create_from_outline` | Generate document from outline |

## Templates

Place templates in `templates/` directory:

```
templates/
├── docx/
│   ├── default.docx
│   └── report.docx
├── pptx/
│   ├── presentation.pptx
│   └── pitch.pptx
└── pdf/
    └── default.html
```

## Installation as Skill

```bash
# Clone repository
git clone https://github.com/yourname/onework.git
cd onework

# Install
pip install -r requirements.txt
```

Then configure in your AI assistant platform.

## Examples

### Example 1: Word to PPT

```python
from onework import read_document, convert_format

# Read Word document
doc = read_document("report.docx")

# Convert to PPT presentation
convert_format("report.docx", "pptx", template="presentation")
```

### Example 2: Generate from Outline

```python
from onework import create_from_outline

outline = """
# Product Launch

## Market Analysis
- Target audience
- Competitive landscape

## Product Features
- Key capabilities
- Unique advantages

## Next Steps
"""

create_from_outline(outline, "pptx")
```

## Dependencies

Core dependencies:
- `mammoth` - Word parsing
- `pymupdf` - PDF processing
- `python-pptx` - PowerPoint handling
- `python-docx` - Word generation
- `markdown` - Markdown processing

Optional:
- `weasyprint` - PDF generation
- `paddleocr` - OCR for images

## License

MIT License - see [LICENSE](LICENSE) file

## Contributing

Contributions are welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) first.

---

<p align="center">
  <sub>Built with ❤️ for better document collaboration</sub>
</p>
