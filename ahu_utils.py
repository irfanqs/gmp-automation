import os
import re


def extract_ahu_number(value, filename='', default='unknown'):
    """Prefer an explicit filename AHU, then normalize the OCR value."""
    filename_text = re.sub(r'\s+', '', os.path.basename(filename))
    filename_match = re.search(
        r'(?:AHU|공조기)[_:\-–—−]*([1-9]\d*)',
        filename_text,
        re.IGNORECASE,
    )
    if filename_match:
        return filename_match.group(1)

    value_text = str(value or '').strip()
    if re.fullmatch(r'\d+(?:\.0)?', value_text):
        number = int(float(value_text))
        if number > 0:
            return str(number)

    compact = re.sub(r'\s+', '', value_text)
    match = re.search(r'(?:AHU|공조기)[_:\-–—−]*([1-9]\d*)', compact, re.IGNORECASE)
    if match:
        return match.group(1)
    return default
