# Markdown to Office Converter

A Python utility for converting Markdown files to multiple formats: **PDF**, **DOCX** (Word), **XLSX** (Excel), and **PPTX** (PowerPoint) with advanced features including table support, appending content, and replacing matching slides.

## Features

- **Multiple Output Formats**: PDF, DOCX (Word), XLSX (Excel), PPTX (PowerPoint)
- **Create Documents**: Convert standalone markdown files to any format
- **Append Mode**: Add markdown content to existing PDFs or presentations
- **Replace Mode**: Replace specific pages/slides when markdown headers match titles
- **Table Support**: Full support for markdown tables with styling
- **Smart Formatting**: Format-specific styling optimized for each output type

## Installation

\`\`\`bash
# Create virtual environment (recommended)
python -m venv .venv
source .venv/bin/activate  # On Linux/Mac
# or
.venv\Scripts\activate  # On Windows

# Install dependencies
pip install -r requirements.txt
\`\`\`

## Usage

### Basic Conversion

Convert markdown to any supported format:

\`\`\`bash
# PDF (default)
python convert.py input.md
python convert.py input.md -o output.pdf

# Word Document
python convert.py input.md --format docx
python convert.py input.md --format docx -o report.docx

# Excel Spreadsheet
python convert.py input.md --format xlsx
python convert.py input.md --format xlsx -o data.xlsx

# PowerPoint Presentation
python convert.py input.md --format pptx
python convert.py input.md --format pptx -o slides.pptx
\`\`\`

### Append to Existing Files

Add markdown content to the end of existing PDFs or presentations:

\`\`\`bash
# Append to PDF
python convert.py new_content.md --format pdf --append existing.pdf -o combined.pdf

# Append to PowerPoint
python convert.py new_slides.md --format pptx --append presentation.pptx -o updated.pptx
\`\`\`

**Use case**: Adding new slides to a presentation, appending reports, combining documents.

### Replace Matching Content

Replace specific pages/slides when markdown \`## headers\` match titles:

\`\`\`bash
# Replace in PDF (experimental)
python convert.py updates.md --format pdf --replace document.pdf -o updated.pdf

# Replace in PowerPoint (recommended)
python convert.py updates.md --format pptx --replace presentation.pptx -o final.pptx
\`\`\`

**How it works**:
1. Extracts titles from each page/slide of the existing file
2. Parses markdown file for \`## Header\` sections
3. Matches header text to titles (case-insensitive, normalized whitespace)
4. Replaces matched content while preserving all others

**Use case**: Updating specific slides in a presentation without regenerating the entire deck.

**Best for**: PowerPoint (PPTX) presentations where slide replacement is native and reliable. PDF replacement is experimental.

## Format Details

### PDF
- Landscape letter size with 0.5" margins
- Tables with borders and alternating row colors
- Green header backgrounds
- 8px font size for dense tables

### DOCX (Word)
- Headers (H1, H2, H3)
- Paragraphs with basic bold formatting
- Tables with "Light Grid Accent 1" style
- Professional document formatting

### XLSX (Excel)
- Markdown tables converted to Excel tables
- Section headers as bold text
- Green header row styling
- Auto-adjusted column widths
- Multiple tables supported

### PPTX (PowerPoint)
- Each \`## Header\` becomes a slide title
- Content below header becomes slide body
- Standard 10" x 7.5" slide size
- Title and Content layout
- Perfect for updating specific slides

## Examples

### Example 1: Create Documents in All Formats

\`\`\`bash
# Generate all formats from the same markdown
python convert.py report.md --format pdf -o report.pdf
python convert.py report.md --format docx -o report.docx
python convert.py report.md --format xlsx -o report.xlsx
python convert.py report.md --format pptx -o report.pptx
\`\`\`

### Example 2: Build a Presentation from Sections

Create \`presentation.md\`:
\`\`\`markdown
## Introduction
Welcome to our Q4 review...

## Financial Results
Revenue increased by 25%...

## Next Steps
Our roadmap for 2026...
\`\`\`

Generate presentation:
\`\`\`bash
python convert.py presentation.md --format pptx -o q4_review.pptx
\`\`\`

### Example 3: Update Specific Slides

Create \`financial_update.md\` with just the updated section:
\`\`\`markdown
## Financial Results
**UPDATED**: Revenue increased by 30% (revised figures)...
\`\`\`

Replace just that slide:
\`\`\`bash
python convert.py financial_update.md --format pptx --replace q4_review.pptx -o q4_review_final.pptx
\`\`\`

### Example 4: Generate Excel from Data Tables

Create \`data.md\`:
\`\`\`markdown
# Sales Report

## Q1 Results

| Region | Sales | Growth |
|--------|-------|--------|
| North  | 500K  | 12%    |
| South  | 450K  | 8%     |

## Q2 Results

| Region | Sales | Growth |
|--------|-------|--------|
| North  | 550K  | 10%    |
| South  | 480K  | 6.7%   |
\`\`\`

Generate spreadsheet:
\`\`\`bash
python convert.py data.md --format xlsx -o sales_report.xlsx
\`\`\`

## Markdown Features

All formats support:
- Headers (H1, H2, H3, etc.)
- Tables (with format-specific styling)
- Lists (ordered and unordered)
- Bold and italic text
- Paragraphs

Format-specific features:
- **PDF**: Full HTML styling support
- **DOCX**: Word document styles and formatting
- **XLSX**: Table-focused, headers become sections
- **PPTX**: Section-based, each \`##\` header becomes a slide

## Requirements

- Python 3.7+
- **PDF**: markdown2, xhtml2pdf, pypdf, pdfplumber
- **DOCX**: python-docx
- **XLSX**: openpyxl
- **PPTX**: python-pptx

See \`requirements.txt\` for specific versions.

## Dependencies

Install all dependencies with:
\`\`\`bash
pip install -r requirements.txt
\`\`\`

Or install format-specific dependencies:
\`\`\`bash
# PDF only
pip install markdown2 xhtml2pdf pypdf pdfplumber

# Office formats only
pip install python-docx openpyxl python-pptx

# Everything
pip install -r requirements.txt
\`\`\`

## Development

Current structure:
\`\`\`
MarkdownToPDF/
├── convert.py          # Main conversion script (all formats)
├── convert_to_pdf.py   # Legacy PDF-only script
├── requirements.txt    # Python dependencies
├── README.md          # This file
├── table.md           # Example markdown with tables
├── presentation.md    # Example presentation markdown
└── .venv/             # Virtual environment (not in git)
\`\`\`

## Scripts

- **convert.py** - New multi-format converter (recommended)
- **convert_to_pdf.py** - Legacy PDF-only script (still works)

## Troubleshooting

**"PDF support not available"**
\`\`\`bash
pip install markdown2 xhtml2pdf pypdf pdfplumber
\`\`\`

**"DOCX support not available"**
\`\`\`bash
pip install python-docx
\`\`\`

**"XLSX support not available"**
\`\`\`bash
pip install openpyxl
\`\`\`

**"PPTX support not available"**
\`\`\`bash
pip install python-pptx
\`\`\`

**No matching headers/slides found**
- Ensure markdown uses \`## Header\` format (level 2 headers)
- Check that header text closely matches slide/page titles
- Matching is case-insensitive and ignores extra whitespace
- For PPTX, this is very reliable; for PDF, it's experimental

## License

See project repository for license information.
