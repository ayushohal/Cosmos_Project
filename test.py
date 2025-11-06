import fitz  # PyMuPDF
from bs4 import BeautifulSoup

def pdf_to_svg_texts(pdf_path):
    # Open PDF
    doc = fitz.open(pdf_path)

    for page_num in range(len(doc)):
        page = doc[page_num]

        # Convert page to SVG
        svg = page.get_svg_image()

        # Parse SVG XML
        soup = BeautifulSoup(svg, "xml")

        # Extract all <text> elements
        texts = [t.get_text() for t in soup.find_all("text")]

        print(f"\n--- Page {page_num+1} ---")
        for txt in texts:
            print(txt)

    doc.close()

# Example usage
pdf_to_svg_texts("123.pdf")
