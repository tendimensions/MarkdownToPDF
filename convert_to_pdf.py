import markdown2
from xhtml2pdf import pisa

# Read the markdown file
with open('table.md', 'r', encoding='utf-8') as f:
    md_content = f.read()

# Convert markdown to HTML with table support
html_content = markdown2.markdown(md_content, extras=['tables'])

# Create a complete HTML document with styling
full_html = f"""
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
    {html_content}
</body>
</html>
"""

# Convert HTML to PDF
with open('table.pdf', 'wb') as pdf_file:
    pisa_status = pisa.CreatePDF(full_html, dest=pdf_file)
    
if pisa_status.err:
    print("Error creating PDF")
else:
    print("PDF created successfully: table.pdf")
