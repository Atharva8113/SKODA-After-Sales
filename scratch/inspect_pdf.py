import pdfplumber
import sys

pdf_path = r"c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\SKODA After Sales\neewwww\IN sea 80554814.pdf"
output_path = r"c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\SKODA After Sales\scratch\pdf_content.txt"

with pdfplumber.open(pdf_path) as pdf:
    with open(output_path, "w", encoding="utf-8") as f:
        for i, page in enumerate(pdf.pages):
            f.write(f"--- Page {i+1} ---\n")
            text = page.extract_text()
            if text:
                f.write(text)
            f.write("\n" + "="*50 + "\n")
print(f"Content saved to {output_path}")
