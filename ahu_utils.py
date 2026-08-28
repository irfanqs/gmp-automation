import os
import re


def extract_ahu_number(value, filename=''):
    """Normalize an OCR AHU value, then fall back to the uploaded filename."""
    value_text = str(value or '').strip()
    if re.fullmatch(r'\d+(?:\.0)?', value_text):
        number = int(float(value_text))
        if number > 0:
            return str(number)

    for text in (value_text, os.path.basename(filename)):
        compact = re.sub(r'\s+', '', text)
        match = re.search(r'(?:AHU|공조기)[_:\-–—−]*([1-9]\d*)', compact, re.IGNORECASE)
        if match:
            return match.group(1)
    return 'unknown'
