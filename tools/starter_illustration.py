"""A five-input saving illustration confined to Start Here."""

from lxml import etree as ET

from example_inputs import STARTER_EXAMPLE


EDITABLE_CELLS = tuple(STARTER_EXAMPLE)
# The two cases share the same pay, saving rate, return and retirement age.
INPUTS = (
    ('Current age', 'whole', 'Your age in whole years.'),
    ('Annual pay before tax', 'money', 'Use today’s dollars.'),
    ('Share of pay saved', 'percent_input', 'Enter a percentage of annual pay.'),
    ('Annual investment return', 'percent_input', 'Enter a return after inflation and fees.'),
    ('Retirement age', 'whole', 'The age when saving stops.'),
)


def configure(put):
    def write(address, value, style):
        put('README', address, value, style)

    write('B18', 'Try the saving illustration', 'report_link')
    write('B22', 'Start investing', 'report_section')
    write('B23', 'Compare starting now with waiting five years. Both accounts start at $0.', 'report_body')
    write('B24', 'Change the five shaded values. These figures apply only to this illustration.', 'report_caption')
    for row, (label, style, guide) in enumerate(INPUTS, 26):
        write(f'B{row}', label, 'report_label')
        write(f'E{row}', STARTER_EXAMPLE[f'E{row}'], style)
        write(f'G{row}', guide, 'report_caption_gutter')

    checks = (
        ('COUNT(E26:E30)<>5', 'Enter a number in each of the five shaded cells.'),
        ('OR(E26<>INT(E26),E26<0,E26>125)', 'Enter a current age from 0 to 125 in whole years.'),
        ('E27<0', 'Enter annual pay of zero or more.'),
        ('OR(E28<0,E28>1)', 'Enter a saving percentage from 0% to 100%.'),
        ('E29<-1', 'Enter an annual return of at least −100%.'),
        ('OR(E30<>INT(E30),E30<E26,E30>125)', 'Enter a whole retirement age from your current age through 125.'),
        ('OR(ISERROR(E27*E28*(E30-E26)),ISERROR(IF(OR(E29=-1,E30=E26,E27*E28=0),0,FV(E29,E30-E26,-(E27*E28),0,0))))', 'The projected balance exceeds Excel’s range. Review the figures.'),
    )
    status = '""'
    for condition, message in reversed(checks):
        status = f'IF({condition},"{message}",{status})'
    write('B31', f'=IFERROR({status},"Check the five shaded cells for an invalid entry.")', 'report_caption')

    def result(address, formula, style='report_money'):
        write(address, f'=IF($B$31="",{formula},"")', style)

    write('B32', 'Saved each year', 'report_label')
    result('E32', '$E$27*$E$28')
    for address, label in (('B34', 'At retirement'), ('F34', 'Start now'), ('H34', 'Wait five years')):
        write(address, label, 'report_table_head')
    for row, label in enumerate(('Years of saving', 'Total contributions', 'Investment gain or loss', 'Ending balance'), 35):
        write(f'B{row}', label, 'report_section' if row == 38 else 'report_label_alt' if row == 36 else 'report_label')
    result('F35', '$E$30-$E$26', 'report_row')
    result('H35', 'MAX(0,F35-5)', 'report_row')
    for col in ('F', 'H'):
        result(f'{col}36', f'{col}35*$E$32', 'report_money_alt')
        result(f'{col}37', f'{col}38-{col}36')
        # A 100% loss each year leaves only the last year-end contribution.
        result(f'{col}38', f'IF(OR({col}35=0,$E$32=0),0,IF($E$29=-1,$E$32,FV($E$29,{col}35,-$E$32,0,0)))')

    write('B40', 'Ending balance in today’s dollars', 'report_section')
    for row, title, source in ((41, 'Start now', 'F38'), (42, 'Wait five years', 'H38')):
        write(f'B{row}', title, 'report_label')
        result(f'E{row}', source)
    write('B44', 'Each deposit is made at year end. Pay and annual savings keep pace with inflation. '
          'The illustration assumes no taxes, withdrawals or account contribution limits.', 'report_caption')
    write('B46', 'Open Inputs to plan for spending, taxes and retirement income', 'report_link')


def layout(root, styles, numbers):
    """Apply the established report style and two bars on one dollar scale."""
    from report_layout import X, _link, _merges, _replace, _rows

    merges = [f'B{r}:I{r}' for r in (3, 5, 7, 8, 10, 11, 12, 14, 15, 16, 18, 19,
                                      22, 23, 24, 31, 40, 44, 46)]
    for r in range(26, 31):
        merges.extend((f'B{r}:D{r}', f'E{r}:F{r}', f'G{r}:I{r}'))
    merges.extend(('B32:D32', 'E32:F32'))
    for r in range(34, 39):
        merges.extend((f'B{r}:E{r}', f'F{r}:G{r}', f'H{r}:I{r}'))
    for r in (41, 42):
        merges.extend((f'B{r}:D{r}', f'E{r}:I{r}'))
    _merges(root, merges)
    _rows(root, 46, 18, {3: 39, 5: 35, 8: 26, 11: 26, 15: 27, 22: 32,
                          23: 28, 24: 24, 31: 26, 32: 25, 34: 28, 38: 29,
                          40: 28, 41: 22, 42: 22, 44: 45, 46: 28,
                          **dict.fromkeys(range(26, 31), 29)})
    root.find(X + 'dimension').set('ref', 'A1:I46')
    _link(root, {'B12': 'Inputs!A1', 'B16': 'Summary!B3', 'B18': "'Start Here'!B22", 'B46': 'Inputs!A1'})

    for address in EDITABLE_CELLS:
        cell = root.find(f".//{X}sheetData/{X}row/{X}c[@r='{address}']")
        code = numbers[int(cell.get('s', '0'))]
        cell.set('s', styles['input_value', code])

    following = ('hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter',
                 'rowBreaks', 'colBreaks', 'drawing', 'legacyDrawing', 'tableParts', 'extLst')
    validations = _replace(root, 'dataValidations', following, count='5')
    rules = (
        ('whole', 'between', '0', '125', 'Enter a whole age from 0 to 125.'),
        ('decimal', 'greaterThanOrEqual', '0', None, 'Enter annual pay of zero or more.'),
        ('decimal', 'between', '0', '1', 'Enter a saving percentage from 0% to 100%.'),
        ('decimal', 'greaterThanOrEqual', '-1', None, 'Enter a return of at least −100%.'),
        ('whole', 'between', '$E$26', '125', 'Enter a whole age from your current age through 125.'),
    )
    for address, (kind, operator, lower, upper, message) in zip(EDITABLE_CELLS, rules, strict=True):
        rule = ET.SubElement(validations, X + 'dataValidation', sqref=address, type=kind,
                             operator=operator, showErrorMessage='1', errorTitle='Check this entry', error=message)
        ET.SubElement(rule, X + 'formula1').text = lower
        if upper is not None:
            ET.SubElement(rule, X + 'formula2').text = upper

    block = _replace(root, 'conditionalFormatting', ('dataValidations', *following), sqref='E41:E42')
    bars = ET.SubElement(ET.SubElement(block, X + 'cfRule', type='dataBar', priority='1'),
                         X + 'dataBar', minLength='0', maxLength='100', showValue='0')
    ET.SubElement(bars, X + 'cfvo', type='num', val='0')
    ET.SubElement(bars, X + 'cfvo', type='max')
    ET.SubElement(bars, X + 'color', rgb='FF126761')
