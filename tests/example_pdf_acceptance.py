"""Compare the published PDF's six pages and vector chart with the saved Summary."""

import argparse
from collections import Counter
from pathlib import Path
import re
from unicodedata import normalize

from openpyxl import load_workbook
import pdfplumber
from pypdf import PdfReader
from pypdf.generic import ArrayObject, DictionaryObject, IndirectObject, TextStringObject


PRIVATE_TEXT = re.compile(r'(?:/users/|/home/|[a-z]:[\\/]|macintosh hd|[\w.+-]+@[\w.-]+\.[a-z]{2,})', re.I)
CHART_LABELS = 'Portfolio balance by age Age Middle 80% of balances Base projection Simulated median'


def words(text):
    return re.findall(r'[\w$%]+|[^\w\s]', normalize('NFKC', text))


def printed_value(cell):
    value = cell.value
    if isinstance(value, str):
        return value
    formats = {'"$"#,##0': '${:,.0f}', '0.00%': '{:.2%}', '0': '{:.0f}', 'General': '{:g}'}
    if cell.number_format not in formats:
        raise AssertionError(f'Unrecognized Summary number format at {cell.coordinate}')
    return formats[cell.number_format].format(value)


def check_document(reader):
    assert len(reader.pages) == 6 and set(reader.metadata) == {'/Title', '/Subject', '/Producer'}
    assert reader.root_object.get('/Lang') == 'en-US' and reader.root_object['/MarkInfo'].get('/Marked')
    assert '/StructTreeRoot' in reader.root_object and len(reader.outline) == 6
    assert not PRIVATE_TEXT.search(str(reader.metadata)), 'Private text in document properties'
    visited, queue = set(), [reader.trailer]
    while queue:
        obj = queue.pop()
        if isinstance(obj, IndirectObject):
            key = obj.idnum, obj.generation
            if key not in visited:
                visited.add(key)
                queue.append(obj.get_object())
        elif isinstance(obj, DictionaryObject):
            assert not set(obj) & {'/URI', '/JS', '/EmbeddedFiles', '/EF', '/AA', '/OpenAction', '/AcroForm'}
            assert obj.get('/S') not in ('/Launch', '/GoToR', '/JavaScript', '/SubmitForm')
            assert obj.get('/Type') != '/EmbeddedFile'
            if obj.get('/S') == '/Figure':
                assert obj.get('/Alt'), 'The portfolio chart needs a text description'
            queue.extend(obj.values())
        elif isinstance(obj, ArrayObject):
            queue.extend(obj)
        elif isinstance(obj, TextStringObject):
            assert not PRIVATE_TEXT.search(obj), 'Private text in a PDF property'


def check_chart(page, data):
    """Use the printed dollar axis to convert each cached balance to its PDF height."""
    count = int(data['Q2'].value)
    values = data.iter_rows(min_row=2, max_row=count + 1, min_col=31, max_col=34, values_only=True)
    lines = [line for line in page.lines if line['width'] > page.width / 2 and line['height'] < .01]
    left = min(line['x0'] for line in lines)
    grid = sorted({round(line['top'], 3) for line in lines})
    ticks = sorted(int(w['text'][1:].replace(',', '')) for w in page.extract_words()
                   if re.fullmatch(r'\$[\d,]+', w['text']) and w['x1'] <= left
                   and grid[0] - 10 <= w['top'] <= grid[-1])
    assert len(ticks) == len(grid) >= 2 and ticks[0] == 0 and ticks[-1] > 0
    assert ticks == list(range(0, ticks[-1] + 1, ticks[1])), 'Unexpected chart axis intervals'
    zero = grid[-1]
    scale = (grid[0] - zero) / ticks[-1]
    paths = [curve['pts'] for curve in page.curves if len(curve['pts']) >= count]
    assert [len(path) for path in paths] == [2 * count, count, count], 'Expected the band and both vector lines'
    band, base, median = paths
    high = band[:count]
    low = list(reversed(band[count:]))
    max_delta = 0.
    for index, (baseline, floor, width, middle) in enumerate(values):
        for path, value in ((base, baseline), (median, middle), (low, floor), (high, floor + width)):
            delta = abs(path[index][1] - (zero + value * scale))
            max_delta = max(max_delta, delta)
    assert max_delta < .25, f'Chart differs from Report_Data by {max_delta:.3f} points'
    return 4 * count, max_delta


def check(workbook, pdf):
    book = load_workbook(workbook, read_only=True, data_only=True)
    try:
        summary = book['Summary']
        rows = list(summary.iter_rows(min_row=1, max_row=190, min_col=2, max_col=9))
        reader = PdfReader(pdf)
        check_document(reader)
        cell_count = 0
        with pdfplumber.open(pdf) as document:
            for number, page in enumerate(document.pages, 1):
                assert (round(page.width), round(page.height)) in ((612,792), (595,842)), 'Expected Letter or A4'
                lines = page.extract_text(x_tolerance=1).splitlines()
                assert lines[-1].startswith('Fictional example · ') and lines[-1].endswith(f'Page {number} of 6')
                content = '\n'.join(lines[:-1])
                assert not PRIVATE_TEXT.search(content) and '#' not in content
                start = (number - 1) * 32
                cells = [cell for row in rows[start:start + 32] for cell in row if cell.value not in (None,'')]
                expected = Counter(token for cell in cells for token in words(printed_value(cell)))
                actual = Counter(words(content))
                assert not expected - actual, f'Page {number} omits saved Summary content'
                if number == 1:
                    chart = check_chart(page, book['Report_Data'])
                    extras = actual - expected
                    allowed = set(words(CHART_LABELS))
                    assert all(t in allowed or re.fullmatch(r'\$?\d+|,', t) for t in extras), 'Unexpected chart text'
                else:
                    assert actual == expected, f'Page {number} contains unexpected content'
                for row in {2: range(41,54), 3: range(71,74)}.get(number, ()):
                    row_text = ' '.join(printed_value(c) for c in rows[row - 1] if c.value not in (None,''))
                    assert ' '.join(row_text.split()) in content, f'Values differ on Summary row {row}'
                cell_count += len(cells)
        print(f'PASS six pages, {cell_count} Summary cells, {chart[0]} chart values. '
              f'Largest chart difference {chart[1]:.3f} PDF points.')
    finally:
        book.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('pdf', type=Path)
    args = parser.parse_args()
    check(args.workbook, args.pdf)
