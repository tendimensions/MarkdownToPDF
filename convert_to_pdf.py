import argparse
import io
import os
import re
import sys
import tempfile
from typing import List, Optional, Tuple

import markdown2
from xhtml2pdf import pisa
try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("Error: pypdf is not installed. Install it with: pip install pypdf")
    sys.exit(1)

try:
    import pdfplumber
except ImportError:
    pdfplumber = None  # Optional dependency for replace mode


def get_html_template(content: str) -> str:
    """Generate HTML template with styling for PDF conversion."""
    return f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        @page {{
            size: letter landscape;
            margin: 0.5in;
        }}
        body {{
            font-family: Arial, sans-serif;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            font-size: 8px;
        }}
        th, td {{
            border: 1px solid #333;
            padding: 6px;
            text-align: left;
        }}
        th {{
            background-color: #4CAF50;
            color: white;
            font-weight: bold;
        }}
        tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
    </style>
</head>
<body>
    {content}
</body>
</html>
"""


def markdown_to_pdf_bytes(md_content: str) -> bytes:
    """Convert markdown content to PDF bytes."""
    html_content = markdown2.markdown(md_content, extras=['tables'])
    full_html = get_html_template(html_content)
    
    # Create PDF in memory
    pdf_buffer = io.BytesIO()
    pisa_status = pisa.CreatePDF(full_html, dest=pdf_buffer)
    
    if pisa_status.err:
        raise RuntimeError("Error creating PDF from markdown")
    
    return pdf_buffer.getvalue()


def create_pdf(markdown_file: str, output_file: str) -> None:
    """Create a new PDF from markdown file."""
    with open(markdown_file, 'r', encoding='utf-8') as f:
        md_content = f.read()
    
    pdf_bytes = markdown_to_pdf_bytes(md_content)
    
    with open(output_file, 'wb') as f:
        f.write(pdf_bytes)
    
    print(f"PDF created successfully: {output_file}")


def append_to_pdf(markdown_file: str, existing_pdf: str, output_file: str) -> None:
    """Append markdown content as new pages to an existing PDF."""
    # Convert markdown to PDF
    with open(markdown_file, 'r', encoding='utf-8') as f:
        md_content = f.read()
    
    new_pdf_bytes = markdown_to_pdf_bytes(md_content)
    
    # Read existing PDF
    reader_existing = PdfReader(existing_pdf)
    reader_new = PdfReader(io.BytesIO(new_pdf_bytes))
    
    # Create writer and append all pages
    writer = PdfWriter()
    
    # Add all pages from existing PDF
    for page in reader_existing.pages:
        writer.add_page(page)
    
    # Add all pages from new PDF
    for page in reader_new.pages:
        writer.add_page(page)
    
    # Write output
    with open(output_file, 'wb') as f:
        writer.write(f)
    
    print(f"PDF created successfully: {output_file}")
    print(f"  Original pages: {len(reader_existing.pages)}")
    print(f"  Appended pages: {len(reader_new.pages)}")
    print(f"  Total pages: {len(reader_existing.pages) + len(reader_new.pages)}")


def extract_markdown_sections(md_content: str) -> List[Tuple[str, str]]:
    """Extract sections from markdown based on ## headers.
    
    Returns:
        List of (header_text, section_content) tuples
    """
    sections = []
    lines = md_content.split('\n')
    current_header = None
    current_content = []
    
    for line in lines:
        # Check for ## header (level 2)
        if line.strip().startswith('## '):
            # Save previous section if exists
            if current_header is not None:
                sections.append((current_header, '\n'.join(current_content)))
            
            # Start new section
            current_header = line.strip()[3:].strip()  # Remove '## ' prefix
            current_content = [line]  # Include the header in content
        else:
            if current_header is not None:
                current_content.append(line)
    
    # Add last section
    if current_header is not None:
        sections.append((current_header, '\n'.join(current_content)))
    
    return sections


def extract_page_titles(pdf_path: str) -> List[Tuple[int, str]]:
    """Extract title/heading from each page of a PDF.
    
    Returns:
        List of (page_number, title_text) tuples
    """
    if pdfplumber is None:
        print("Error: pdfplumber is required for replace mode.")
        print("Install it with: pip install pdfplumber")
        sys.exit(1)
    
    page_titles = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                # Try to extract first non-empty line as title
                lines = [line.strip() for line in text.split('\n') if line.strip()]
                if lines:
                    # Look for lines that might be headers (usually short, at the top)
                    title = lines[0]
                    page_titles.append((i, title))
    
    return page_titles


def normalize_text(text: str) -> str:
    """Normalize text for comparison (lowercase, remove extra whitespace)."""
    return re.sub(r'\s+', ' ', text.lower().strip())


def replace_pdf_pages(markdown_file: str, existing_pdf: str, output_file: str) -> None:
    """Replace PDF pages where markdown headers match page titles."""
    # Read markdown and extract sections
    with open(markdown_file, 'r', encoding='utf-8') as f:
        md_content = f.read()
    
    sections = extract_markdown_sections(md_content)
    if not sections:
        print("Warning: No ## headers found in markdown file")
        print("Using append mode instead...")
        append_to_pdf(markdown_file, existing_pdf, output_file)
        return
    
    # Extract titles from existing PDF
    page_titles = extract_page_titles(existing_pdf)
    
    # Create mapping of normalized title -> markdown section
    section_map = {normalize_text(header): content for header, content in sections}
    
    # Find matches
    matches = []
    for page_num, page_title in page_titles:
        normalized = normalize_text(page_title)
        if normalized in section_map:
            matches.append((page_num, page_title, section_map[normalized]))
    
    if not matches:
        print("Warning: No matching headers found between PDF and markdown")
        print("Using append mode instead...")
        append_to_pdf(markdown_file, existing_pdf, output_file)
        return
    
    print(f"Found {len(matches)} matching pages to replace:")
    for page_num, title, _ in matches:
        print(f"  Page {page_num + 1}: {title}")
    
    # Read existing PDF
    reader_existing = PdfReader(existing_pdf)
    writer = PdfWriter()
    
    # Process each page
    for page_num in range(len(reader_existing.pages)):
        # Check if this page should be replaced
        match = next((m for m in matches if m[0] == page_num), None)
        
        if match:
            # Generate PDF from markdown section
            _, title, section_content = match
            section_pdf_bytes = markdown_to_pdf_bytes(section_content)
            section_reader = PdfReader(io.BytesIO(section_pdf_bytes))
            
            # Add all pages from the section (might be multiple pages)
            for section_page in section_reader.pages:
                writer.add_page(section_page)
        else:
            # Keep original page
            writer.add_page(reader_existing.pages[page_num])
    
    # Write output
    with open(output_file, 'wb') as f:
        writer.write(f)
    
    print(f"\nPDF created successfully: {output_file}")
    print(f"  Original pages: {len(reader_existing.pages)}")
    print(f"  Pages replaced: {len(matches)}")


def main():
    parser = argparse.ArgumentParser(
        description='Convert Markdown to PDF with optional append/replace modes',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create new PDF from markdown
  python convert_to_pdf.py input.md
  python convert_to_pdf.py input.md -o output.pdf
  
  # Append markdown to existing PDF
  python convert_to_pdf.py input.md --append existing.pdf -o combined.pdf
  
  # Replace matching slides in PDF
  python convert_to_pdf.py input.md --replace existing.pdf -o updated.pdf
        """
    )
    
    parser.add_argument('markdown_file', help='Input markdown file')
    parser.add_argument('-o', '--output', help='Output PDF file (default: input name + .pdf)')
    parser.add_argument('--append', metavar='PDF', help='Append to existing PDF file')
    parser.add_argument('--replace', metavar='PDF', 
                       help='Replace matching pages in existing PDF (matches ## headers to page titles)')
    
    args = parser.parse_args()
    
    # Validate inputs
    if not os.path.exists(args.markdown_file):
        print(f"Error: Markdown file not found: {args.markdown_file}")
        sys.exit(1)
    
    if args.append and args.replace:
        print("Error: Cannot use both --append and --replace modes")
        sys.exit(1)
    
    # Determine output file
    if args.output:
        output_file = args.output
    else:
        base = os.path.splitext(args.markdown_file)[0]
        output_file = f"{base}.pdf"
    
    # Execute appropriate mode
    try:
        if args.append:
            if not os.path.exists(args.append):
                print(f"Error: PDF file not found: {args.append}")
                sys.exit(1)
            append_to_pdf(args.markdown_file, args.append, output_file)
        elif args.replace:
            if not os.path.exists(args.replace):
                print(f"Error: PDF file not found: {args.replace}")
                sys.exit(1)
            replace_pdf_pages(args.markdown_file, args.replace, output_file)
        else:
            create_pdf(args.markdown_file, output_file)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
