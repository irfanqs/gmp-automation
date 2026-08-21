"""
GMP Automation - debug OCR tool
Menampilkan teks mentah hasil DeepSeek-OCR untuk satu file PDF.

Usage:
    python debug_ocr.py <path_pdf> <ngrok_url>
"""

import sys
from deepseek_ocr.client import ocr_pdf

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    pdf_path, endpoint = sys.argv[1], sys.argv[2]
    pages = ocr_pdf(pdf_path, endpoint)
    for i, page in enumerate(pages, start=1):
        print(f'==================== PAGE {i} ====================')
        print(page)
