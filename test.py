# from docx2pdf import convert

# # Convert a single DOCX file to PDF
# convert("1.docx")


import os
import pdfkit
from docx import Document

def convert_docx_to_pdf(input_file, output_file):
    # Create a new document
    doc = Document(input_file)
    # Extract text (for demonstration purposes)
    doc_text = "\n".join([p.text for p in doc.paragraphs])
    
    # Save text to HTML format (as intermediate step for PDF conversion)
    html_file = "temp.html"
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(f"<html><body>{doc_text}</body></html>")

    try:
        # Convert HTML to PDF
        pdfkit.from_file(html_file, output_file)
        print(f"Converted {input_file} to {output_file}")
    except Exception as e:
        print(f"Conversion failed: {e}")
    finally:
        
        # Clean up temporary HTML file
        if os.path.exists(html_file):
            os.remove(html_file)

# Example usage
folderFinle = "/"
nameFile = "output_name"
docx_file = "1.docx"
pdf_file_path = os.path.join(folderFinle, f"{nameFile}.pdf")

if not os.path.exists(pdf_file_path):
    print(f"Starting conversion: {docx_file} to {pdf_file_path}")
    convert_docx_to_pdf(docx_file, pdf_file_path)
else:
    print(f"PDF already exists: {pdf_file_path}")
