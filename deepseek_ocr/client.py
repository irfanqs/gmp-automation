"""
GMP Automation System - DeepSeek-OCR Client
Sends PDF page images to a DeepSeek-OCR backend running on Kaggle (via ngrok)
and returns raw OCR text/markdown per page.
"""

import base64
import requests
from io import BytesIO
from pdf2image import convert_from_path

OCR_TIMEOUT = 300


def pdf_to_images(pdf_path, dpi=200):
    """Convert PDF pages to PIL Images."""
    return convert_from_path(pdf_path, dpi=dpi)


def image_to_base64(pil_image):
    """Convert PIL Image to base64 PNG string."""
    buffer = BytesIO()
    pil_image.save(buffer, format='PNG')
    return base64.standard_b64encode(buffer.getvalue()).decode('utf-8')


def call_deepseek_ocr(image_b64, endpoint_url, mode='markdown', resolution='gundam'):
    """Call the Kaggle DeepSeek-OCR /ocr endpoint for a single image. Returns raw text."""
    if not endpoint_url:
        raise ValueError("DeepSeek-OCR endpoint URL is required (Kaggle ngrok URL).")

    url = endpoint_url.rstrip('/') + '/ocr'
    payload = {'image_b64': image_b64, 'mode': mode, 'resolution': resolution}

    try:
        response = requests.post(url, json=payload, timeout=OCR_TIMEOUT)
    except requests.exceptions.RequestException as e:
        raise Exception(f"Failed to reach DeepSeek-OCR backend at {endpoint_url}: {e}")

    if response.status_code != 200:
        raise Exception(f"DeepSeek-OCR backend error {response.status_code}: {response.text}")

    return response.json().get('text', '')


def ocr_pdf(pdf_path, endpoint_url, mode='markdown', resolution='gundam', dpi=200):
    """Convert a PDF to images and OCR every page. Returns list of raw text, one per page."""
    images = pdf_to_images(pdf_path, dpi=dpi)
    pages_text = []
    for img in images:
        img_b64 = image_to_base64(img)
        text = call_deepseek_ocr(img_b64, endpoint_url, mode=mode, resolution=resolution)
        pages_text.append(text)
    return pages_text
