"""
GMP Automation System - DeepSeek-OCR Output Parsers
Converts raw markdown/text produced by the DeepSeek-OCR backend (Kaggle) into
the same structured JSON schema that ocr_engine.py (Claude) produces, so
excel_generator.py does not need to change.

NOTE: DeepSeek-OCR is a raw OCR / doc-to-markdown model, not an instruction
model — it does not reliably follow "return this JSON schema" prompts.
These parsers use header-keyword matching + column position heuristics on
the markdown tables it produces. They were written from the *known* table
layout described in ocr_engine.py's Claude prompts, but have not been
calibrated against real DeepSeek-OCR output — once you have real output for
each document type, check `_find_col_indices` and the column-pairing logic
below and adjust to match what the model actually emits (e.g. exact header
text, merged-cell splitting, page-break duplication of headers).
"""

import re


# =============================================================================
# GENERIC FIELD / TABLE EXTRACTION HELPERS
# =============================================================================

def _join_pages(pages_text):
    return "\n".join(pages_text)


def extract_field(text, patterns, default=None):
    """Try a list of regexes in order, return the first captured group found."""
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            return m.group(1).strip()
    return default


def extract_ahu(text):
    return extract_field(
        text,
        [r'공조기\s*[-–]\s*(\d+)', r'해당\s*공조기[^\d]*(\d+)', r'공조기\s*[:\-]?\s*(\d+)'],
        default='unknown',
    )


def extract_date(text):
    return extract_field(
        text,
        [
            r'측정일자[^\d]*(\d{4}[.\-]\s*\d{1,2}[.\-]\s*\d{1,2})',
            r'(\d{4}\.\d{1,2}\.\d{1,2})',
        ],
        default='2025.08.01',
    )


def extract_result(text):
    return extract_field(text, [r'측정결과[^\S\n]*[:|]?\s*(적합|부적합)'], default='적합')


def extract_standard(text):
    return extract_field(text, [r'측정기준[^\S\n]*[:|]?\s*([\d.]+\s*%)'], default=None)


def extract_markdown_tables(text):
    """Return list of tables; each table is a list of rows (list of cell strings)."""
    tables = []
    current = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            cells = [c.strip() for c in stripped.strip('|').split('|')]
            # skip markdown separator rows like |---|---|
            if all(re.fullmatch(r':?-+:?', c) for c in cells):
                continue
            current.append(cells)
        else:
            if current:
                tables.append(current)
                current = []
    if current:
        tables.append(current)
    return tables


def _find_col(header_row, keywords):
    """Return index of first header cell containing any of the given keywords."""
    for i, cell in enumerate(header_row):
        for kw in keywords:
            if kw in cell:
                return i
    return None


def _to_number(s, as_int=False):
    if s is None:
        return None
    s = s.strip().replace(',', '').replace('%', '')
    if s in ('', '-', '—'):
        return None
    try:
        return int(float(s)) if as_int else float(s)
    except ValueError:
        return None


# =============================================================================
# A. AIRBORNE PARTICLE
# =============================================================================

def parse_airborne_particle(pages_text):
    text = _join_pages(pages_text)
    rooms = []

    for table in extract_markdown_tables(text):
        if len(table) < 2:
            continue
        header, rows = table[0], table[1:]

        grade_i = _find_col(header, ['청정등급', '등급'])
        room_no_i = _find_col(header, ['실번호'])
        room_name_i = _find_col(header, ['실명'])
        if room_no_i is None and room_name_i is None:
            continue  # not the measurement table

        used = {i for i in (grade_i, room_no_i, room_name_i) if i is not None}
        measure_cols = [i for i in range(len(header)) if i not in used]
        # measurement columns come in pairs: 0.5um, 5.0um per point
        pairs = [measure_cols[i:i + 2] for i in range(0, len(measure_cols) - 1, 2)]

        no_counter = 1
        for row in rows:
            if len(row) < len(header):
                continue
            room_number = row[room_no_i] if room_no_i is not None else ''
            room_name = row[room_name_i] if room_name_i is not None else ''
            grade = row[grade_i] if grade_i is not None else ''
            if not room_number and not room_name:
                continue

            measurements = []
            for point, (c05, c50) in enumerate(pairs, start=1):
                v05 = _to_number(row[c05], as_int=True) if c05 < len(row) else None
                v50 = _to_number(row[c50], as_int=True) if c50 < len(row) else None
                if v05 is None and v50 is None:
                    continue
                measurements.append({'point': point, 'value_05': v05 or 0, 'value_50': v50 or 0})

            if not measurements:
                continue

            rooms.append({
                'no_start': no_counter,
                'no_end': no_counter + len(measurements) - 1,
                'grade': grade,
                'room_number': room_number,
                'room_name': room_name,
                'measurements': measurements,
            })
            no_counter += len(measurements)

    return {
        'ahu': extract_ahu(text),
        'date': extract_date(text),
        'result': extract_result(text),
        'rooms': rooms,
    }


# =============================================================================
# B. AIR VELOCITY
# =============================================================================

def parse_air_velocity(pages_text):
    text = _join_pages(pages_text)
    machines = []

    for table in extract_markdown_tables(text):
        if len(table) < 2:
            continue
        header, rows = table[0], table[1:]

        grade_i = _find_col(header, ['청정등급', '등급'])
        room_no_i = _find_col(header, ['실번호'])
        name_i = _find_col(header, ['실명'])
        if room_no_i is None and name_i is None:
            continue

        used = {i for i in (grade_i, room_no_i, name_i) if i is not None}
        value_cols = [i for i in range(len(header)) if i not in used]

        no_counter = 1
        for row in rows:
            if len(row) < len(header):
                continue
            room_number = row[room_no_i] if room_no_i is not None else ''
            machine_name = row[name_i].replace('<br>', '\n') if name_i is not None else ''
            grade = row[grade_i] if grade_i is not None else ''
            if not room_number and not machine_name:
                continue

            measurements = []
            for point, c in enumerate(value_cols, start=1):
                v = _to_number(row[c]) if c < len(row) else None
                if v is None:
                    continue
                measurements.append({'point': point, 'value': v})

            if not measurements:
                continue

            machines.append({
                'no_start': no_counter,
                'no_end': no_counter + len(measurements) - 1,
                'grade': grade,
                'room_number': room_number,
                'machine_name': machine_name,
                'measurements': measurements,
            })
            no_counter += len(measurements)

    return {
        'ahu': extract_ahu(text),
        'date': extract_date(text),
        'result': extract_result(text),
        'machines': machines,
    }


# =============================================================================
# C. AIR CHANGE RATE (ACH)
# =============================================================================

def parse_air_change_rate(pages_text):
    text = _join_pages(pages_text)
    rooms = []
    no_counter = 1

    for table in extract_markdown_tables(text):
        if len(table) < 2:
            continue
        header, rows = table[0], table[1:]

        grade_i = _find_col(header, ['청정등급', '등급'])
        room_no_i = _find_col(header, ['실번호'])
        room_name_i = _find_col(header, ['실명'])
        volume_i = _find_col(header, ['체적'])
        total_i = _find_col(header, ['합계'])
        ach_i = _find_col(header, ['환기횟수', 'ACH', '회/hr'])
        if room_no_i is None and room_name_i is None:
            continue

        used = {i for i in (grade_i, room_no_i, room_name_i, volume_i, total_i, ach_i) if i is not None}
        flow_cols = [i for i in range(len(header)) if i not in used]

        for row in rows:
            if len(row) < len(header):
                continue
            room_number = row[room_no_i] if room_no_i is not None else ''
            room_name = row[room_name_i] if room_name_i is not None else ''
            if not room_number and not room_name:
                continue

            air_flow_measurements = []
            for point, c in enumerate(flow_cols, start=1):
                v = _to_number(row[c]) if c < len(row) else None
                if v is None:
                    continue
                air_flow_measurements.append({'point': point, 'air_flow': v})

            total_air_flow = _to_number(row[total_i]) if total_i is not None and total_i < len(row) else None
            if total_air_flow is None and air_flow_measurements:
                total_air_flow = round(sum(m['air_flow'] for m in air_flow_measurements), 1)

            rooms.append({
                'no': no_counter,
                'grade': row[grade_i] if grade_i is not None else '',
                'room_number': room_number,
                'room_name': room_name,
                'volume': _to_number(row[volume_i]) if volume_i is not None and volume_i < len(row) else None,
                'air_flow_measurements': air_flow_measurements,
                'total_air_flow': total_air_flow,
                'ach': _to_number(row[ach_i], as_int=True) if ach_i is not None and ach_i < len(row) else None,
            })
            no_counter += 1

    return {
        'ahu': extract_ahu(text),
        'date': extract_date(text),
        'result': extract_result(text),
        'rooms': rooms,
    }


# =============================================================================
# D. HEPA FILTER
# =============================================================================

def parse_hepa_filter(pages_text):
    text = _join_pages(pages_text)
    items = []

    for table in extract_markdown_tables(text):
        if len(table) < 2:
            continue
        header, rows = table[0], table[1:]

        room_no_i = _find_col(header, ['실번호'])
        name_i = _find_col(header, ['실명'])
        if room_no_i is None and name_i is None:
            continue

        used = {i for i in (room_no_i, name_i) if i is not None}
        value_cols = [i for i in range(len(header)) if i not in used]

        no_counter = 1
        for row in rows:
            if len(row) < len(header):
                continue
            room_number = row[room_no_i] if room_no_i is not None else ''
            item_name = row[name_i] if name_i is not None else ''
            if not room_number and not item_name:
                continue

            measurements = []
            for point, c in enumerate(value_cols, start=1):
                v = _to_number(row[c]) if c < len(row) else None
                if v is None:
                    continue
                measurements.append({'point': point, 'value': v})

            if not measurements:
                continue

            items.append({
                'no_start': no_counter,
                'no_end': no_counter + len(measurements) - 1,
                'room_number': room_number,
                'item_name': item_name,
                'measurements': measurements,
            })
            no_counter += len(measurements)

    return {
        'ahu': extract_ahu(text),
        'date': extract_date(text),
        'result': extract_result(text),
        'standard': extract_standard(text),
        'items': items,
    }


# =============================================================================
# E. AIRFLOW PATTERN (field-based, one item per page)
# =============================================================================

def parse_airflow_pattern(pages_text):
    items = []
    for page_text in pages_text:
        name = extract_field(page_text, [r'측정대상[^\S\n]*[:|]?\s*(.+)'])
        if not name:
            continue
        date = extract_field(page_text, [r'측정일자[^\S\n]*[:|]?\s*([\d.]+)'], default='')
        criteria = extract_field(page_text, [r'측정기준[^\S\n]*[:|]?\s*([\s\S]+?)(?:동영상|판정결과|$)'], default='')
        video = extract_field(page_text, [r'동영상\s*첨부[^\S\n]*[:|]?\s*(\S+)'], default='')
        judgment = extract_field(page_text, [r'판정결과[^\S\n]*[:|]?\s*(적합|부적합)'], default='')

        items.append({
            'name': name.strip(),
            'date': date.strip(),
            'criteria': criteria.strip(),
            'video_attached': video.strip(),
            'judgment': judgment.strip(),
        })

    ahu = extract_ahu(_join_pages(pages_text))
    return {'ahu': ahu, 'items': items}
