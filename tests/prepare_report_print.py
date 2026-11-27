"""Create a macro-free print projection of a verified fictional workbook."""

import argparse
from datetime import datetime
from hashlib import file_digest
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from lxml import etree

from xlsx_read import Workbook, q, sheet_parts

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from check_workbook_commit import audit_sample_workbook
from prepare_native_cache import prepare_cache_copy


MACRO_TYPE = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
STANDARD_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
PAPER_SIZES = {'letter': '1', 'a4': '9'}
DOLLAR_VIEWS = {'today': 'Today’s dollars', 'future': 'Future dollars'}
CHART = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
DRAWING = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
PORTFOLIO_CHART = 'xl/charts/chart1.xml'


def report_context(source: Path) -> tuple[str, str]:
    """Read the saved preparation line and corresponding fictional footer date."""
    with Workbook(source) as book:
        context = book['Summary']['B6'].value
    if not isinstance(context, str) or 'Prepared ' not in context:
        raise ValueError('The saved fictional report has no preparation date.')
    date = datetime.strptime(context.rsplit('Prepared ', 1)[1], '%b %d, %Y')
    return context, f'{date:%B} {date.day}, {date:%Y}'


def project(source: Path, output: Path, title: str, paper: str,
            dollar_view: str | None, fictional_footer: bool, *, calc_export: bool = False) -> None:
    """Keep the original data, chart and print ranges in an isolated workbook copy."""
    if calc_export and (title != 'Summary' or dollar_view is not None):
        raise ValueError('Calc export requires Summary and the saved dollar view. Save the desired view before exporting.')
    if source.suffix.lower() != '.xlsm' or output.suffix.lower() != '.xlsx' or output.exists() or output.is_symlink():
        raise ValueError('Use a fictional macro-enabled source and a new .xlsx output path.')
    audit_sample_workbook(source)
    context = report_context(source) if fictional_footer else None
    reader = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)
    with TemporaryDirectory(dir=output.parent) as temp:
        staged = Path(temp) / 'report.xlsx'
        print_source = source
        if calc_export:
            print_source = Path(temp) / 'verified-records.xlsm'
            with Workbook(source) as book:
                count, signature = (book['Risk_Data'][cell].value for cell in ('Z10', 'Z8'))
            checked = prepare_cache_copy(source, print_source, count, signature)
            print(f'Checked {checked:,} repeated assumption records for Calc export.')
        with ZipFile(print_source) as original, ZipFile(staged, 'x') as projected:
            parts = sheet_parts(original)
            if title not in parts or 'Summary' not in parts:
                raise ValueError('The workbook lacks the requested print worksheet.')
            workbook = etree.fromstring(original.read('xl/workbook.xml'), reader)
            relationships = etree.fromstring(original.read('xl/_rels/workbook.xml.rels'), reader)
            types = etree.fromstring(original.read('[Content_Types].xml'), reader)
            for index, sheet in enumerate(workbook.find(q('sheets'))):
                if sheet.get('name') == title:
                    selected = index
                    sheet.set('state', 'visible')
                elif sheet.get('state') != 'veryHidden':
                    sheet.set('state', 'hidden')
            for view in workbook.findall(f'.//{q("workbookView")}'):
                view.attrib.update({'activeTab': str(selected), 'firstSheet': str(selected)})
            for item in list(relationships):
                if item.get('Type', '').endswith('/vbaProject'):
                    relationships.remove(item)
            for item in list(types):
                mime = item.get('ContentType', '')
                if mime == MACRO_TYPE:
                    item.set('ContentType', STANDARD_TYPE)
                elif 'vbaProject' in mime:
                    types.remove(item)
            replacements = {'xl/workbook.xml': workbook, 'xl/_rels/workbook.xml.rels': relationships,
                            '[Content_Types].xml': types}
            for name, part in parts.items():
                sheet = etree.fromstring(original.read(part), reader)
                if apply_view(sheet, name, title, paper, dollar_view, context):
                    replacements[part] = sheet
            if calc_export:
                chart = etree.fromstring(original.read(PORTFOLIO_CHART), reader)
                sparse_age_labels(chart)
                replacements[PORTFOLIO_CHART] = chart
            for part in original.infolist():
                if part.filename == 'xl/vbaProject.bin':
                    continue
                tree = replacements.get(part.filename)
                data = etree.tostring(tree, encoding='UTF-8') if tree is not None else original.read(part)
                projected.writestr(part, data)
        output.hardlink_to(staged)


def sparse_age_labels(chart: etree._Element) -> None:
    """Use the saved series in the Calc print copy, with sparse category labels."""
    references = chart.findall(f'.//{CHART}cat/{CHART}strRef')
    if len(references) != 4 or any(ref.findtext(CHART + 'f') != "'Report_Data'!RMChartPortfolioAge"
                                   for ref in references):
        raise ValueError('The report chart has unexpected age category references.')
    categories = [ref.findall(f'{CHART}strCache/{CHART}pt') for ref in references]
    ages = [point.findtext(CHART + 'v') for point in categories[0]]
    if (not 2 <= len(ages) <= 126 or any(not age or not age.isdecimal() for age in ages)
            or any([point.findtext(CHART + 'v') for point in series] != ages for series in categories)):
        raise ValueError('The saved report chart has inconsistent age categories.')
    first, last = int(ages[0]), int(ages[-1])
    if ages != [str(age) for age in range(first, last + 1)]:
        raise ValueError('The saved report age categories are not in year order.')
    visible = visible_ages(first, last)
    labels = [age if age in visible else '\u00a0' for age in ages]
    for reference, series in zip(references, categories, strict=True):
        for point, label in zip(series, labels, strict=True):
            point.find(CHART + 'v').text = label
        cache = reference.find(CHART + 'strCache')
        cache.tag = CHART + 'strLit'
        reference.getparent().replace(reference, cache)
    values = chart.findall(f'.//{CHART}val/{CHART}numRef')
    if len(values) != 4:
        raise ValueError('The report chart does not have four saved value series.')
    for reference in values:
        cache = reference.find(CHART + 'numCache')
        if cache is None or len(cache.findall(CHART + 'pt')) != len(labels):
            raise ValueError('The saved chart values are incomplete.')
        cache.tag = CHART + 'numLit'
        reference.getparent().replace(reference, cache)
    for skip in chart.findall(f'.//{CHART}catAx/{CHART}tickLblSkip'):
        skip.set('val', '1')
    axis_text = chart.find(f'.//{CHART}catAx/{CHART}txPr/{DRAWING}bodyPr')
    if axis_text is None:
        raise ValueError('The report chart has no age-axis text settings.')
    axis_text.attrib.update({'rot': '-2700000', 'wrap': 'none'})


def visible_ages(first: int, last: int) -> tuple[str, ...]:
    """Label the first and last ages and each intervening 15-year step."""
    return tuple(str(age) for age in (*range(first, last, 15), last))


def apply_view(sheet: etree._Element, name: str, title: str, paper: str,
               dollar_view: str | None, context: tuple[str, str] | None) -> bool:
    """Change only requested tab and print settings on a disposable worksheet."""
    changed = False
    for view in sheet.findall(f'{q("sheetViews")}/{q("sheetView")}'):
        if name == title or view.get('tabSelected') is not None:
            view.set('tabSelected', '1' if name == title else '0')
            changed = True
    if name == title:
        setup = sheet.find(q('pageSetup'))
        if setup is None:
            raise ValueError('The selected worksheet has no print layout.')
        setup.set('paperSize', PAPER_SIZES[paper])
        if context is not None:
            footer = sheet.find(f'{q("headerFooter")}/{q("oddFooter")}')
            if footer is None:
                raise ValueError('The selected worksheet has no print footer.')
            footer.text = (f'&L&"Arial,Regular"&9Fictional example · {context[1]} · US federal and California taxes.'
                           '&R&"Arial,Regular"&9Page &P of &N')
            if name == 'Summary':
                cell = sheet.find(f'.//{q("c")}[@r="B6"]')
                if cell is None:
                    raise ValueError('The saved preparation line is missing.')
                inline_text(cell, context[0])
        changed = True
    if name == 'Summary' and dollar_view is not None:
        cell = sheet.find(f'.//{q("c")}[@r="K18"]')
        if cell is None or cell.find(q('f')) is not None:
            raise ValueError('The Summary dollar selector is missing or has changed.')
        inline_text(cell, DOLLAR_VIEWS[dollar_view])
        changed = True
    return changed


def inline_text(cell: etree._Element, value: str) -> None:
    """Write a display-only string on the disposable print sheet."""
    cell.set('t', 'inlineStr')
    for child in list(cell):
        cell.remove(child)
    text = etree.SubElement(etree.SubElement(cell, q('is')), q('t'))
    text.text = value
    if value != value.strip():
        text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--sheet', choices=('Summary', 'Annual Detail'), default='Summary')
    parser.add_argument('--paper', choices=tuple(PAPER_SIZES), default='letter')
    parser.add_argument('--dollar-view', choices=tuple(DOLLAR_VIEWS),
                        help='Override the selector for a recalculated native Excel preview')
    parser.add_argument('--fictional-footer', action='store_true', help='Show the saved example preparation date in print')
    parser.add_argument('--calc-export', action='store_true', help='Encode verified repeated records for a Calc print copy')
    args = parser.parse_args()
    project(args.source, args.output, args.sheet, args.paper, args.dollar_view, args.fictional_footer,
            calc_export=args.calc_export)
    with args.output.open('rb') as exported:
        print(f'Fictional print projection SHA-256 {file_digest(exported, "sha256").hexdigest()}')
