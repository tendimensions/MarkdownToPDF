# MarkdownToPDF

A Python utility for converting Markdown files to PDF format with advanced features including table support, appending to existing PDFs, and replacing matching slides.

## Features

- **Create PDF**: Convert standalone markdown files to PDF
- **Append Mode**: Add markdown content as new pages to existing PDFs
- **Replace Mode**: Replace specific PDF pages when markdown headers match slide titles
- **Table Support**: Full support for markdown tables with styling
- **Custom Styling**: Landscape layout optimized for tables

## Installation

```bash
# Create virtual environment (recommended)
python -m venv .venv
source .venv/bin/activate  # On Linux/Mac
# or
.venv\Scripts\activate  # On Windows

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Create New PDF

Convert a markdown file to PDF:

```bash
python convert_to_pdf.py input.md
python convert_to_pdf.py input.md -o custom_output.pdf
```

### Append to Existing PDF

Add markdown content as new pages to the end of an existing PDF:

```bash
python convert_to_pdf.py new_content.md --append existing.pdf -o combined.pdf
```

**Use case**: Adding new slides to a presentation, appending reports, combining documents.

### Replace Matching Slides

Replace specific pages in a PDF when markdown `## headers` match page titles:

```bash
python convert_to_pdf.py updates.md --replace presentation.pdf -o updated.pdf
```

**How it works**:
1. Extracts text from each page of the existing PDF
2. Identifies the title/heading at the top of each page
3. Parses markdown file for `## Header` sections
4. Matches header text to page titles (case-insensitive, normalized whitespace)
5. Replaces matched pages while preserving all others

**Use case**: Updating specific slides in a presentation without regenerating the entire deck.

**Note**: This feature works best with PDFs where each page has a distinct title/header at the top. If no matches are found, the tool will fall back to append mode and warn you.

## Examples

### Example 1: Create Simple PDF

```bash
# Create table.pdf from table.md (default)
python convert_to_pdf.py table.md
```

### Example 2: Append New Content

```bash
# Add quarterly_update.md to the end of annual_report.pdf
python convert_to_pdf.py quarterly_update.md --append annual_report.pdf -o annual_report_q4.pdf
```

### Example 3: Replace Specific Slides

Create `updates.md`:
```markdown
## Introduction
Updated introduction text...

## Q4 Results
New financial data for Q4...
```

Replace matching slides:
```bash
# This will replace only the "Introduction" and "Q4 Results" slides
python convert_to_pdf.py updates.md --replace presentation.pdf -o presentation_updated.pdf
```

## Markdown Features

- Headers (H1, H2, H3, etc.)
- Tables with automatic styling
- Lists (ordered and unordered)
- Bold, italic, code formatting
- Links

## PDF Styling

The generated PDFs include:
- Letter size, landscape orientation
- 0.5 inch margins
- Tables with borders and alternating row colors
- Green header backgrounds
- Arial sans-serif font
- Responsive font sizing (8px for dense tables)

## Requirements

- Python 3.7+
- markdown2: Markdown to HTML conversion
- xhtml2pdf: HTML to PDF rendering
- pypdf: PDF manipulation (merge, append)
- pdfplumber: PDF text extraction (for replace mode)

See `requirements.txt` for specific versions.

## Dependencies

Install all dependencies with:
```bash
pip install -r requirements.txt
```

## Development

Current structure:
```
MarkdownToPDF/
├── convert_to_pdf.py   # Main script
├── requirements.txt    # Python dependencies
├── table.md           # Example markdown file
├── README.md          # This file
└── .venv/             # Virtual environment (not in git)
```

## Troubleshooting

**"pypdf is not installed"**
```bash
pip install pypdf
```

**"pdfplumber is required for replace mode"**
```bash
pip install pdfplumber
```

**No matching headers found**
- Ensure markdown uses `## Header` format (level 2 headers)
- Check that header text closely matches PDF page titles
- Matching is case-insensitive and ignores extra whitespace

## License

See project repository for license information.
