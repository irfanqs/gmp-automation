"""
GMP Automation System - DeepSeek-OCR Output Parsers
Converts raw text/markdown/HTML produced by the DeepSeek-OCR backend (Kaggle) into
the same structured JSON schema that ocr_engine.py (Claude) produces, so
excel_generator.py does not need to change.

DeepSeek-OCR emits documents as HTML <table> blocks (with rowspan/colspan for
merged cells) wrapped in <|ref|> markers, e.g.:

    <|ref|>table<|/ref|><|det|>[[23, 66, 958, 831]]<|/det|>
    <table><tr><td rowspan="2">NO.</td> ... </tr></table>

The parsers below therefore:
  1. Extract every <table> block and expand rowspan/colspan into a flat grid.
  2. Merge multi-row (group + sub) headers into one header row.
  3. Locate identity columns (grade / room no / room name / volume / total / ACH)
     and treat every remaining column as measurement values.
  4. Fall back to classic markdown |...| tables if the model ever emits those.
"""

import re
from html.parser import HTMLParser


# =============================================================================
# GENERIC FIELD / TABLE EXTRACTION HELPERS
# =============================================================================

def _join_pages(pages_text):
    return "\n".join(pages_text)


# -----------------------------------------------------------------------------
# HTML table extraction (rowspan/colspan aware)
# -----------------------------------------------------------------------------

class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self._table = None
        self._row = None
        self._cell = None
        self._rowspan = 1
        self._colspan = 1

    def handle_starttag(self, tag, attrs):
        if tag == 'table':
            self._table = []
        elif tag == 'tr' and self._table is not None:
            self._row = []
        elif tag in ('td', 'th') and self._row is not None:
            self._cell = []
            d = dict(attrs)
            self._rowspan = int(d.get('rowspan', 1) or 1)
            self._colspan = int(d.get('colspan', 1) or 1)
        elif tag == 'br' and self._cell is not None:
            self._cell.append(' ')

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag):
        if tag == 'table':
            if self._table:
                self.tables.append(self._table)
            self._table = None
        elif tag == 'tr':
            if self._row is not None and self._table is not None:
                self._table.append(self._row)
            self._row = None
        elif tag in ('td', 'th') and self._cell is not None and self._row is not None:
            text = ' '.join(''.join(self._cell).split())
            self._row.append({'text': text, 'rowspan': self._rowspan, 'colspan': self._colspan})
            self._cell = None


def _expand_table(rows):
    """Expand rowspan/colspan cells into a rectangular grid of plain strings."""
    result = []
    pending = {}  # col -> (text, remaining rows after current)
    for row in rows:
        out_row = []
        col = 0
        cell_idx = 0
        while cell_idx < len(row):
            if col in pending:
                text, remaining = pending[col]
                out_row.append(text)
                if remaining <= 1:
                    del pending[col]
                else:
                    pending[col] = (text, remaining - 1)
                col += 1
                continue
            cell = row[cell_idx]
            cell_idx += 1
            for k in range(cell['colspan']):
                out_row.append(cell['text'])
                if cell['rowspan'] > 1:
                    pending[col + k] = (cell['text'], cell['rowspan'] - 1)
            col += cell['colspan']
        result.append(out_row)
    return result


def _to_number(s, as_int=False):
    if s is None:
        return None
    s = str(s).strip().replace(',', '').replace('%', '').replace(' ', '')
    if s in ('', '-', '—'):
        return None
    m = re.match(r'-?\d+\.?\d*', s)
    if not m:
        return None
    try:
        v = float(m.group(0))
        return int(v) if as_int else v
    except ValueError:
        return None


def _is_data_row(row):
    """A data row starts with an integer NO. and has at least one non-numeric
    cell in the following columns (grade letter, room number, room name...)."""
    if not row:
        return False
    if not re.fullmatch(r'\d+', str(row[0]).strip()):
        return False
    return any(re.search(r'[A-Za-z가-힣]', str(c)) for c in row[1:6])


def _merge_header_rows(table):
    """Split a (possibly multi-row) header from the data rows and merge header
    cells so downstream code sees a single flat header row."""
    data_start = 0
    for idx, row in enumerate(table):
        if _is_data_row(row):
            data_start = idx
            break
    if data_start == 0:
        return table
    header_rows = table[:data_start]
    width = max(len(r) for r in header_rows)
    header = []
    for i in range(width):
        parts = [r[i] for r in header_rows if i < len(r) and r[i]]
        header.append(' '.join(parts))
    return [header] + table[data_start:]


def extract_html_tables(text):
    """All <table> blocks in text as list of tables; each table is a list of
    rows, each row a list of cell strings (merged cells repeated)."""
    parser = _TableParser()
    parser.feed(text)
    tables = []
    for raw in parser.tables:
        tables.append(_merge_header_rows(_expand_table(raw)))
    return tables


def extract_markdown_tables(text):
    """Fallback: classic markdown |...| tables. Each table is a list of rows."""
    tables = []
    current = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            cells = [c.strip() for c in stripped.strip('|').split('|')]
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


def extract_tables(text):
    """All tables (HTML from DeepSeek-OCR, then markdown fallback)."""
    return extract_html_tables(text) + extract_markdown_tables(text)


# -----------------------------------------------------------------------------
# Text normalization for key:value field matching
# -----------------------------------------------------------------------------

class _TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []

    def handle_starttag(self, tag, attrs):
        if tag in ('br', 'tr', 'td', 'th', 'p', 'div', 'table'):
            self.parts.append(' ')

    def handle_data(self, data):
        self.parts.append(data)


def html_to_text(text):
    """HTML -> readable text (tags dropped, whitespace collapsed)."""
    parser = _TextExtractor()
    parser.feed(text)
    return ' '.join(''.join(parser.parts).split())


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
    return extract_field(text, [r'측정\s*결과[^\S\n]*[:|]?\s*(적합|부적합)'], default='적합')


def extract_standard(text):
    return extract_field(text, [r'측정\s*기준[^\S\n]*[:|]?\s*([\d.]+\s*%)'], default=None)


# -----------------------------------------------------------------------------
# Shared column helpers
# -----------------------------------------------------------------------------

def _find_col(header_row, keywords):
    """Return index of first header cell containing any keyword (whitespace-agnostic)."""
    for i, cell in enumerate(header_row):
        norm = re.sub(r'\s+', '', str(cell))
        for kw in keywords:
            if kw in norm:
                return i
    return None


_ID_COLS_KEYWORDS = {
    'grade': ['청정등급', '등급'],
    'room_no': ['실번호'],
    'room_name': ['실명', '설비명', '기기명'],
    'volume': ['체적'],
    'total': ['합계'],
    'ach': ['환기횟수', 'ACH', '회/hr'],
    'no': ['NO', 'No', 'no'],
    'point': ['측정번호', '측정횟수', '측정점'],
}


def _align_row(row, header_width):
    """Fix rows that are longer than the header because the OCR emitted a
    colspan artifact (a merged cell duplicated). Drop one duplicate cell."""
    if len(row) <= header_width:
        return row
    for i in range(1, len(row)):
        if row[i] == row[i - 1]:
            return row[:i] + row[i + 1:]
    return row


def _used_cols(header, *groups):
    """Set of header indices belonging to any of the given identity groups."""
    used = set()
    for group in groups:
        i = _find_col(header, _ID_COLS_KEYWORDS[group])
        if i is not None:
            used.add(i)
    return used


# =============================================================================
# A. AIRBORNE PARTICLE
# =============================================================================

def parse_airborne_particle(pages_text):
    text = _join_pages(pages_text)
    readable = html_to_text(text)
    rooms = []
    no_counter = 1

    for table in extract_tables(text):
        if len(table) < 2:
            continue
        header, rows = table[0], table[1:]

        room_no_i = _find_col(header, _ID_COLS_KEYWORDS['room_no'])
        room_name_i = _find_col(header, _ID_COLS_KEYWORDS['room_name'])
        grade_i = _find_col(header, _ID_COLS_KEYWORDS['grade'])
        if room_no_i is None and room_name_i is None:
            continue  # not the measurement table

        used = _used_cols(header, 'grade', 'room_no', 'room_name', 'no', 'point')
        measure_cols = [i for i in range(len(header)) if i not in used]
        # measurement columns come in pairs: 0.5um, 5.0um per point
        pairs = [measure_cols[i:i + 2] for i in range(0, len(measure_cols) - 1, 2)]

        entries = []
        for row in rows:
            row = _align_row(row, len(header))
            if len(row) < len(header):
                continue
            room_number = row[room_no_i] if room_no_i is not None else ''
            room_name = row[room_name_i] if room_name_i is not None else ''
            grade = row[grade_i] if grade_i is not None else ''
            if not room_number and not room_name:
                continue

            measurements = []
            for (c05, c50) in pairs:
                v05 = _to_number(row[c05], as_int=True) if c05 < len(row) else None
                v50 = _to_number(row[c50], as_int=True) if c50 < len(row) else None
                if v05 is None and v50 is None:
                    continue
                measurements.append({'point': len(measurements) + 1, 'value_05': v05 or 0, 'value_50': v50 or 0})

            if not measurements:
                continue
            entries.append({
                'grade': grade,
                'room_number': room_number,
                'room_name': room_name,
                'measurements': measurements,
            })

        # group consecutive entries for the same room into one room object
        i = 0
        while i < len(entries):
            e = entries[i]
            j = i
            while j + 1 < len(entries) and \
                    entries[j + 1]['room_number'] == e['room_number'] and \
                    entries[j + 1]['room_name'] == e['room_name'] and \
                    entries[j + 1]['grade'] == e['grade']:
                j += 1
            group = []
            for en in entries[i:j + 1]:
                for m in en['measurements']:
                    group.append(dict(m, point=len(group) + 1))
            rooms.append({
                'no_start': no_counter,
                'no_end': no_counter + len(group) - 1,
                'grade': e['grade'],
                'room_number': e['room_number'],
                'room_name': e['room_name'],
                'measurements': group,
            })
            no_counter += len(group)
            i = j + 1

    return {
        'ahu': extract_ahu(readable),
        'date': extract_date(readable),
        'result': extract_result(readable),
        'rooms': rooms,
    }


# =============================================================================
# B. AIR VELOCITY
# =============================================================================

def parse_air_velocity(pages_text):
    text = _join_pages(pages_text)
    readable = html_to_text(text)
    machines = []
    no_counter = 1

    for table in extract_tables(text):
        if len(table) < 2:
            continue
        header, rows = table[0], table[1:]

        room_no_i = _find_col(header, _ID_COLS_KEYWORDS['room_no'])
        name_i = _find_col(header, _ID_COLS_KEYWORDS['room_name'])
        grade_i = _find_col(header, _ID_COLS_KEYWORDS['grade'])
        if room_no_i is None and name_i is None:
            continue

        used = _used_cols(header, 'grade', 'room_no', 'room_name', 'no', 'point')
        value_cols = [i for i in range(len(header)) if i not in used]

        entries = []
        for row in rows:
            row = _align_row(row, len(header))
            if len(row) < len(header):
                continue
            room_number = row[room_no_i] if room_no_i is not None else ''
            machine_name = row[name_i] if name_i is not None else ''
            grade = row[grade_i] if grade_i is not None else ''
            if not room_number and not machine_name:
                continue

            measurements = []
            for c in value_cols:
                v = _to_number(row[c]) if c < len(row) else None
                if v is None:
                    continue
                measurements.append({'point': len(measurements) + 1, 'value': v})

            if not measurements:
                continue
            entries.append({
                'grade': grade,
                'room_number': room_number,
                'machine_name': machine_name,
                'measurements': measurements,
            })

        # group consecutive entries for the same machine into one machine object
        i = 0
        while i < len(entries):
            e = entries[i]
            j = i
            while j + 1 < len(entries) and \
                    entries[j + 1]['room_number'] == e['room_number'] and \
                    entries[j + 1]['machine_name'] == e['machine_name'] and \
                    entries[j + 1]['grade'] == e['grade']:
                j += 1
            group = []
            for en in entries[i:j + 1]:
                for m in en['measurements']:
                    group.append(dict(m, point=len(group) + 1))
            machines.append({
                'no_start': no_counter,
                'no_end': no_counter + len(group) - 1,
                'grade': e['grade'],
                'room_number': e['room_number'],
                'machine_name': e['machine_name'],
                'measurements': group,
            })
            no_counter += len(group)
            i = j + 1

    return {
        'ahu': extract_ahu(readable),
        'date': extract_date(readable),
        'result': extract_result(readable),
        'machines': machines,
    }


# =============================================================================
# C. AIR CHANGE RATE (ACH)
# =============================================================================

def parse_air_change_rate(pages_text):
    text = _join_pages(pages_text)
    readable = html_to_text(text)
    rooms = []
    no_counter = 1

    for table in extract_tables(text):
        if len(table) < 2:
            continue
        header, rows = table[0], table[1:]

        room_no_i = _find_col(header, _ID_COLS_KEYWORDS['room_no'])
        room_name_i = _find_col(header, _ID_COLS_KEYWORDS['room_name'])
        grade_i = _find_col(header, _ID_COLS_KEYWORDS['grade'])
        if room_no_i is None and room_name_i is None:
            continue

        used = _used_cols(header, 'grade', 'room_no', 'room_name', 'volume', 'total', 'ach', 'no', 'point')
        flow_cols = [i for i in range(len(header)) if i not in used]

        for row in rows:
            row = _align_row(row, len(header))
            if len(row) < len(header):
                continue
            room_number = row[room_no_i] if room_no_i is not None else ''
            room_name = row[room_name_i] if room_name_i is not None else ''
            if not room_number and not room_name:
                continue

            air_flow_measurements = []
            for c in flow_cols:
                v = _to_number(row[c]) if c < len(row) else None
                if v is None:
                    continue
                air_flow_measurements.append({'point': len(air_flow_measurements) + 1, 'air_flow': v})

            total_i = _find_col(header, _ID_COLS_KEYWORDS['total'])
            total_air_flow = _to_number(row[total_i]) if total_i is not None and total_i < len(row) else None
            if total_air_flow is None and air_flow_measurements:
                total_air_flow = round(sum(m['air_flow'] for m in air_flow_measurements), 1)

            volume_i = _find_col(header, _ID_COLS_KEYWORDS['volume'])
            ach_i = _find_col(header, _ID_COLS_KEYWORDS['ach'])
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
        'ahu': extract_ahu(readable),
        'date': extract_date(readable),
        'result': extract_result(readable),
        'rooms': rooms,
    }


# =============================================================================
# D. HEPA FILTER
# =============================================================================

def parse_hepa_filter(pages_text):
    text = _join_pages(pages_text)
    readable = html_to_text(text)
    items = []
    no_counter = 1

    for table in extract_tables(text):
        if len(table) < 2:
            continue
        header, rows = table[0], table[1:]

        room_no_i = _find_col(header, _ID_COLS_KEYWORDS['room_no'])
        name_i = _find_col(header, _ID_COLS_KEYWORDS['room_name'])
        if room_no_i is None and name_i is None:
            continue

        used = _used_cols(header, 'room_no', 'room_name', 'no', 'point')
        value_cols = [i for i in range(len(header)) if i not in used]

        entries = []
        for row in rows:
            row = _align_row(row, len(header))
            if len(row) < len(header):
                continue
            room_number = row[room_no_i] if room_no_i is not None else ''
            item_name = row[name_i] if name_i is not None else ''
            if not room_number and not item_name:
                continue

            measurements = []
            for c in value_cols:
                v = _to_number(row[c]) if c < len(row) else None
                if v is None:
                    continue
                measurements.append({'point': len(measurements) + 1, 'value': v})

            if not measurements:
                continue
            entries.append({
                'room_number': room_number,
                'item_name': item_name,
                'measurements': measurements,
            })

        # group consecutive entries for the same item into one item object
        i = 0
        while i < len(entries):
            e = entries[i]
            j = i
            while j + 1 < len(entries) and \
                    entries[j + 1]['room_number'] == e['room_number'] and \
                    entries[j + 1]['item_name'] == e['item_name']:
                j += 1
            group = []
            for en in entries[i:j + 1]:
                for m in en['measurements']:
                    group.append(dict(m, point=len(group) + 1))
            items.append({
                'no_start': no_counter,
                'no_end': no_counter + len(group) - 1,
                'room_number': e['room_number'],
                'item_name': e['item_name'],
                'measurements': group,
            })
            no_counter += len(group)
            i = j + 1

    return {
        'ahu': extract_ahu(readable),
        'date': extract_date(readable),
        'result': extract_result(readable),
        'standard': extract_standard(readable),
        'items': items,
    }


# =============================================================================
# E. AIRFLOW PATTERN (field-based, one item per page)
# =============================================================================

def parse_airflow_pattern(pages_text):
    items = []
    for page_text in pages_text:
        readable = html_to_text(page_text)
        name = extract_field(readable, [r'측정\s*대상[^\S\n]*[:|]?\s*(.+)'])
        if not name:
            continue
        date = extract_field(readable, [r'측정\s*일자[^\S\n]*[:|]?\s*([\d.]+)'], default='')
        criteria = extract_field(readable, [r'측정\s*기준[^\S\n]*[:|]?\s*([\s\S]+?)(?:동영상|판정결과|$)'], default='')
        video = extract_field(readable, [r'동영상\s*첨부[^\S\n]*[:|]?\s*(\S+)'], default='')
        judgment = extract_field(readable, [r'판정\s*결과[^\S\n]*[:|]?\s*(적합|부적합)'], default='')

        items.append({
            'name': name.strip(),
            'date': date.strip(),
            'criteria': criteria.strip(),
            'video_attached': video.strip(),
            'judgment': judgment.strip(),
        })

    ahu = extract_ahu(html_to_text(_join_pages(pages_text)))
    return {'ahu': ahu, 'items': items}
