"""Prepare a tagged fictional report from a locally exported Calc PDF."""

import argparse
from collections import defaultdict
from math import isfinite
from pathlib import Path
import re
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
import pdfplumber
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, DictionaryObject, NameObject, TextStringObject

from example_pdf_acceptance import check, words
from prepare_report_print import audit_sample_workbook, visible_ages


HEADINGS = ('B3', 'B34', 'B66', 'B98', 'B130', 'B162')
TABLE_HEADERS = {
    2: (('Year', 'Age', 'Income used', 'Savings used', 'Other spending', 'Healthcare', 'Tax due', 'Unpaid'),),
    3: (('Market outcome', 'Ending portfolio', 'First shortfall age', 'Unpaid that year'),),
    4: (('Test', 'End age', 'Paid all costs', 'First shortfall age', 'Unpaid in that year'),) * 2,
}


def cached_description(workbook: Path) -> tuple[tuple[str, ...], str, tuple[str, ...]]:
    """Derive the headings and chart alternative from the saved example."""
    book = load_workbook(workbook, read_only=True, data_only=True)
    try:
        report = book['Summary']
        headings = tuple(report[cell].value for cell in HEADINGS)
        count = book['Report_Data']['Q2'].value
        if type(count) is not int or not 2 <= count <= 126:
            raise ValueError('The saved report does not have a complete chart.')
        base, low, spread, median = next(book['Report_Data'].iter_rows(
            min_row=count + 1, max_row=count + 1, min_col=31, max_col=34, values_only=True))
        if not all(isinstance(amount, (int, float)) and isfinite(amount)
                   for amount in (base, low, spread, median)):
            raise ValueError('The saved ending chart values are missing or invalid.')
        if any(not isinstance(title, str) or not title.strip() for title in headings):
            raise ValueError('The saved report headings are incomplete.')
        start, end = (book['Inputs'][cell].value for cell in ('B8', 'B53'))
        unit = report['K18'].value
        if (type(start) is not int or type(end) is not int or end - start + 1 != count
                or unit not in ('Today’s dollars', 'Future dollars')):
            raise ValueError('The saved chart has an unexpected age range or dollar view.')
        alternative = (f'Portfolio balances from age {start} through {end} in {unit.lower()}. '
                       f'At age {end} the base projection is ${base:,.0f} and the simulated median '
                       f'is ${median:,.0f}. The middle 80% of simulated balances at age {end} range '
                       f'from ${low:,.0f} to ${low + spread:,.0f}. The band is calculated separately '
                       'at each age. Balances are before estimated tax on traditional savings.')
        return headings, alternative, visible_ages(start, end)
    finally:
        book.close()


def kids(node: DictionaryObject) -> ArrayObject:
    """Read this exporter's array of child references."""
    value = node.get('/K', ArrayObject()).get_object()
    if not isinstance(value, ArrayObject):
        raise ValueError('The Calc export has an unexpected structure child.')
    return value


def attach(node: DictionaryObject, child_refs: list) -> None:
    """Set the hierarchy while leaving original marked-content entries intact."""
    node[NameObject('/K')] = ArrayObject(child_refs)
    for child in child_refs:
        child.get_object()[NameObject('/P')] = node.indirect_reference


def element(pdf: PdfWriter, role: str, page_ref) -> DictionaryObject:
    """Create a new structure element on an existing PDF page."""
    node = DictionaryObject({NameObject('/Type'): NameObject('/StructElem'),
                             NameObject('/S'): NameObject(role), NameObject('/Pg'): page_ref})
    return pdf._add_object(node).get_object()


def cell_refs(row_ref) -> ArrayObject:
    """Return the original cells of a Calc table row."""
    row = row_ref.get_object()
    cells = kids(row)
    if row.get('/S') != '/TR' or any(cell.get_object().get('/S') != '/TD' for cell in cells):
        raise ValueError('The Calc page table has an unexpected row or cell.')
    return cells


def paragraph_refs(row_ref) -> list:
    """Read the text leaves of a layout row, in their original order."""
    paragraphs = [p for cell in cell_refs(row_ref) for p in kids(cell.get_object())]
    if any(p.get_object().get('/S') != '/P' for p in paragraphs):
        raise ValueError('The Calc layout row contains unexpected marked content.')
    return paragraphs


def normalized(value: str) -> str:
    """Normalize extracted words without changing the page content."""
    return ' '.join(words(value))


def page_text(source: Path) -> list[dict[int, str]]:
    """Index actual printed glyphs by each page's marked-content identifier."""
    with pdfplumber.open(source) as doc:
        pages = []
        for page in doc.pages:
            indexed = defaultdict(list)
            for char in page.chars:
                if char['mcid'] is not None:
                    indexed[char['mcid']].append(char['text'])
            pages.append({key: normalized(''.join(chars)) for key, chars in indexed.items()})
        return pages


def paragraph_text(ref, text: dict[int, str]) -> str:
    """Get a text leaf's content from the existing PDF marking."""
    ids = ref.get_object().get('/K')
    if isinstance(ids, int):
        ids = (ids,)
    if not ids or any(not isinstance(key, int) or key not in text for key in ids):
        raise ValueError('A report paragraph has no readable marked text.')
    return ' '.join(text[key] for key in ids)


def table_range(rows: list, text: dict[int, str], labels: tuple[str, ...], first: int) -> tuple[int, int]:
    """Find a real data-table header and its following financial rows."""
    expected = [normalized(label) for label in labels]
    for index in range(first, len(rows)):
        cells = cell_refs(rows[index])
        actual = [' '.join(paragraph_text(p, text) for p in kids(cell.get_object())) for cell in cells]
        if actual == expected:
            last = index + 1
            while last < len(rows) and len(cell_refs(rows[last])) >= len(labels) - 1:
                last += 1
            if last == index + 1:
                break
            return index, last
    raise ValueError(f'The report table beginning with {labels[0]} could not be identified.')


def data_table(pdf: PdfWriter, page, rows: list, columns: int) -> object:
    """Preserve data rows and identify the real column headers."""
    table = element(pdf, '/Table', page)
    header = cell_refs(rows[0])
    for ref in header:
        cell = ref.get_object()
        cell[NameObject('/S')] = NameObject('/TH')
        cell[NameObject('/A')] = DictionaryObject({NameObject('/O'): NameObject('/Table'),
                                                   NameObject('/Scope'): NameObject('/Column')})
    for ref in rows[1:]:
        cells = list(cell_refs(ref))
        for cell_ref in cells:
            # Calc reports physical worksheet column spans on merged cells.
            # Each of these cells occupies one column in the financial table.
            cell_ref.get_object().pop('/A', None)
        if len(cells) == columns - 1 and columns in (4, 5):
            # The two risk tables leave the final unpaid-amount field blank
            # when the row has no shortfall.
            cells.append(element(pdf, '/TD', page).indirect_reference)
            attach(ref.get_object(), cells)
        elif len(cells) != columns:
            raise ValueError('A report data row does not match its column headers.')
    attach(table, rows)
    return table.indirect_reference


def page_section(pdf: PdfWriter, table: DictionaryObject, text: dict[int, str],
                 number: int, title: str, figure) -> object:
    """Read layout rows in page order and retain only the meaningful tables."""
    page = table.raw_get('/Pg')
    rows = list(kids(table))
    section = element(pdf, '/Sect', page)
    children = []
    next_row = 0
    ranges = {}
    for labels in TABLE_HEADERS.get(number, ()):
        first, last = table_range(rows, text, labels, next_row)
        ranges[first] = (last, len(labels))
        next_row = last
    index = matches = 0
    inserted_figure = False
    while index < len(rows):
        if index in ranges:
            last, columns = ranges[index]
            children.append(data_table(pdf, page, rows[index:last], columns))
            index = last
            continue
        for paragraph in paragraph_refs(rows[index]):
            line = paragraph_text(paragraph, text)
            if line == normalized(title):
                paragraph.get_object()[NameObject('/S')] = NameObject('/H1' if number == 1 else '/H2')
                matches += 1
            if (figure is not None and not inserted_figure
                    and re.match(r'(Today|Future).*dollars use', line)):
                children.append(figure)
                inserted_figure = True
            children.append(paragraph)
        index += 1
    if matches != 1 or (figure is not None and not inserted_figure):
        raise ValueError(f'The exact saved heading or figure context is missing on page {number}.')
    attach(section, children)
    return section.indirect_reference


def modify_structure(pdf: PdfWriter, text: list[dict[int, str]], headings: tuple[str, ...],
                     alternative: str) -> None:
    """Replace Calc's layout tables with headings and true data tables."""
    tree = pdf.root_object.get('/StructTreeRoot')
    if tree is None or '/ParentTree' not in tree.get_object() or len(pdf.pages) != 6:
        raise ValueError('Use a tagged Calc export of the six-page Summary.')
    try:
        workbook_ref, = kids(tree.get_object())
        worksheet_ref, = kids(workbook_ref.get_object())
    except ValueError as error:
        raise ValueError('The tagged export must contain only the Summary worksheet.') from error
    worksheet = worksheet_ref.get_object()
    items = list(kids(worksheet))
    if tuple(ref.get_object().get('/S') for ref in items) != ('/Table', '/Figure', *('/Table',) * 5):
        raise ValueError('The tagged export has unexpected Summary tables or figures.')
    figure = items.pop(1)
    figure.get_object()[NameObject('/Alt')] = TextStringObject(alternative)
    sections = []
    for number, (ref, glyphs, heading) in enumerate(zip(items, text, headings, strict=True), 1):
        table = ref.get_object()
        if table.raw_get('/Pg') != pdf.pages[number - 1].indirect_reference:
            raise ValueError('The PDF tag and page orders disagree.')
        sections.append(page_section(pdf, table, glyphs, number, heading, figure if number == 1 else None))
    attach(worksheet, sections)


def verify_structure(pdf: PdfReader, headings: tuple[str, ...], alternative: str) -> None:
    """Check the serialized heading, table and figure associations."""
    workbook_ref, = kids(pdf.root_object['/StructTreeRoot'])
    worksheet_ref, = kids(workbook_ref.get_object())
    sections = [ref.get_object() for ref in kids(worksheet_ref.get_object())]
    if len(sections) != len(headings) or [item.title for item in pdf.outline] != list(headings):
        raise ValueError('The exported report does not have a bookmark for each section.')
    for number, section in enumerate(sections, 1):
        elements = [ref.get_object() for ref in kids(section)]
        expected_heading = '/H1' if number == 1 else '/H2'
        tables = [node for node in elements if node.get('/S') == '/Table']
        expected_tables = TABLE_HEADERS.get(number, ())
        if (section.get('/S') != '/Sect' or section.raw_get('/Pg') != pdf.pages[number - 1].indirect_reference
                or sum(node.get('/S') == expected_heading for node in elements) != 1
                or len(tables) != len(expected_tables)):
            raise ValueError(f'The serialized semantics on page {number} differ from the report.')
        for table, labels in zip(tables, expected_tables, strict=True):
            rows = kids(table)
            cells = [kids(row.get_object()) for row in rows]
            if any(len(row) != len(labels) for row in cells) or any(
                ref.get_object().get('/S') != '/TH' or ref.get_object()['/A'].get('/Scope') != '/Column'
                for ref in cells[0]
            ):
                raise ValueError(f'A serialized table on page {number} lacks scoped column headers.')
        figures = [node.get('/Alt') for node in elements if node.get('/S') == '/Figure']
        if figures != ([alternative] if number == 1 else []):
            raise ValueError('The chart text alternative is missing or misplaced.')


def verify_age_axis(path: Path, ages: tuple[str, ...]) -> None:
    """Verify the printed age labels and reject digits wrapped onto different lines."""
    with pdfplumber.open(path) as report:
        page = report.pages[0]
        grid = [line['top'] for line in page.lines if line['width'] > page.width / 2 and line['height'] < .01]
        if not grid:
            raise ValueError('The printed portfolio chart has no horizontal axis.')
        axis = max(grid)
        labels = defaultdict(list)
        for char in page.chars:
            if char['text'].isdecimal() and 0 < char['top'] - axis < 19:
                labels[char['mcid']].append(char)
        if tuple(''.join(char['text'] for char in label) for label in labels.values()) != ages:
            raise ValueError('The printed portfolio chart has unexpected age labels.')
        if any(char['x0'] - previous['x0'] < .5 for label in labels.values()
               for previous, char in zip(label, label[1:])):
            raise ValueError('The printed portfolio chart contains a wrapped age label.')


def prepare(workbook: Path, source: Path, output: Path) -> None:
    """Write a verified publication copy and leave both source files untouched."""
    if output.exists() or output.is_symlink():
        raise ValueError('Use a new report output path.')
    audit_sample_workbook(workbook)
    headings, alternative, ages = cached_description(workbook)
    source_pdf = PdfReader(source)
    pdf = PdfWriter(clone_from=source_pdf)
    pdf.pdf_header = source_pdf.pdf_header
    original_streams = [page.get_contents().get_data() for page in pdf.pages]
    modify_structure(pdf, page_text(source), headings, alternative)
    pdf.root_object.pop('/Metadata', None)
    pdf.root_object.pop('/Outlines', None)
    pdf.metadata = None
    pdf.add_metadata({'/Title': 'Retirement plan | Fictional example',
                      '/Subject': 'Six-page US federal and California retirement plan example',
                      '/Producer': 'Financial Model'})
    for number, heading in enumerate(headings):
        pdf.add_outline_item(heading, number)
    with TemporaryDirectory(dir=output.parent) as work:
        staged = Path(work) / 'sample.pdf'
        pdf.write(staged)
        candidate = PdfReader(staged)
        if [page.get_contents().get_data() for page in candidate.pages] != original_streams:
            raise ValueError('PDF preparation changed the original page contents.')
        verify_structure(candidate, headings, alternative)
        verify_age_axis(staged, ages)
        check(workbook, staged)
        output.hardlink_to(staged)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path, help='Saved fictional workbook that supplied the report')
    parser.add_argument('source', type=Path, help='Raw tagged six-page Calc export')
    parser.add_argument('output', type=Path, help='New verified PDF path')
    args = parser.parse_args()
    prepare(args.workbook, args.source, args.output)
