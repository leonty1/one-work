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
| **Web** | URL → Markdown (intelligent content extraction) |
| **Structure** | Extract headings, tables, key information |
| **Update** | Edit specific sections with precision |
| **Convert** | Word ↔ PDF ↔ PPT ↔ HTML ↔ Markdown |
| **Generate** | Create documents from outlines |

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Use in Python
from onework import read_document, read_url, convert_format

# Read any document
result = read_document("/path/to/report.docx")
print(result["content"])

# Read web page to Markdown
result = read_url("https://example.com/article")
print(result["title"])
print(result["content"])

# Convert to PDF
result = convert_format("/path/to/report.docx", "pdf")
```

## Supported Formats

### Input
- **Word** (.docx) - via mammoth
- **PDF** (.pdf) - via PyMuPDF
- **PowerPoint** (.pptx) - via python-pptx
- **Web pages** (.url) - via trafilatura/markdownify
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
- `markitdown` - Microsoft open source document to markdown (optional, auto-detected)

Web page support:
- `trafilatura` - Intelligent web content extraction (filters ads/navigation)
- `markdownify` - HTML to Markdown conversion
- `requests` - HTTP requests
- `beautifulsoup4` - HTML parsing

Optional:
- `weasyprint` - PDF generation
- `paddleocr` - OCR for images

---

## Acknowledgments

**onework** is built upon the shoulders of giants. We gratefully acknowledge the following open source projects:

| Project | Purpose | License |
|---------|---------|---------|
| [Microsoft markitdown](https://github.com/microsoft/markitdown) | Document to Markdown conversion (DOCX, PDF, PPTX, XLSX, images, audio) | MIT |
| [trafilatura](https://trafilatura.readthedocs.io/) | Web page content extraction (filters ads/nav) | GPL-3.0 |
| [markdownify](https://github.com/matthewwithanm/python-markdownify) | HTML to Markdown conversion | MIT |
| [mammoth](https://github.com/markdown/mammoth) | Word to Markdown | BSD |
| [PyMuPDF](https://pymupdf.readthedocs.io/) | PDF processing | AGPL/Commercial |
| [python-pptx](https://python-pptx.readthedocs.io/) | PowerPoint handling | MIT |
| [python-docx](https://python-docx.readthedocs.io/) | Word generation | MIT |
| [weasyprint](https://doc.courtbouillon.org/weasyprint/) | PDF generation | BSD |

Thank you to all the open source contributors!

---

## Technical Stack

| Feature | Open Source Tool | License |
|---------|------------------|---------|
| Document → MD (General) | [Microsoft markitdown](https://github.com/microsoft/markitdown) | MIT |
| Web → MD | [trafilatura](https://trafilatura.readthedocs.io/) + markdownify | GPL-3.0/MIT |
| Word → MD | mammoth | BSD |
| PDF → MD | pymupdf | AGPL/Commercial |
| PPT → MD | python-pptx | MIT |
| MD → Word | python-docx | MIT |
| MD → PDF | weasyprint | BSD |
| MD → PPT | python-pptx | MIT |

> **About Microsoft markitdown**
> 
> Starting from v0.2.0, onework integrates [Microsoft markitdown](https://github.com/microsoft/markitdown) as an enhanced document parsing engine.

> **About Web → Markdown**
> 
> Starting from v0.3.0, onework supports URL to Markdown conversion using trafilatura (intelligent content extraction) or markdownify (HTML conversion).
> Once installed, onework will automatically use markitdown for document processing, providing stronger format support.

---

## License

MIT License - see [LICENSE](LICENSE) file

## Contributing

Contributions are welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) first.

---

<p align="center">
  <sub>Built with ❤️ for better document collaboration</sub>
</p>
