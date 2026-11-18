"""Describe Excel entry rules for the workbook's editable assumptions."""

from lxml import etree as ET

X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def _between(cells, low, high, message, *, whole=False, optional=False):
    return cells, 'whole' if whole else 'decimal', 'between', low, high, message, optional


def _minimum(cells, low, message, *, strict=False, optional=False):
    return cells, 'decimal', 'greaterThan' if strict else 'greaterThanOrEqual', low, None, message, optional


def _custom(cells, formula, message):
    return cells, 'custom', None, formula, None, message, False


def _ages(cells, high=125):
    return _between(cells, 0, high, f'Enter a whole age from 0 to {high}.', whole=True)


def _share(cells):
    return _between(cells, 0, 1, 'Enter a percentage from 0% to 100%.')


def _rows(*rows):
    return ' '.join(f'D{row}:F{row}' for row in rows)


_INPUT_RULES = (
    _between('B4', 1, 3, 'Choose scenario 1, 2 or 3.', whole=True),
    _ages('D8:F10 ' + _rows(45, 50, 53, 67, 87, 88)),
    _between(_rows(11, 28, 48, 97, 121), 1, 2, 'Enter 1 or 2 as shown in the row note.', whole=True),
    _minimum('D13:F13', -1, 'Enter wage growth of at least −100%.'),
    _minimum('D14:F14', -1, 'Enter inflation above −100%.', strict=True),
    _between('D15:F16', 1, 3, 'Choose investment mix 1, 2 or 3.', whole=True),
    _between(_rows(29, 92, 93, 94), 0, '$B$130', 'Enter the first tax year, up to the opening year, or 0.', whole=True),
    _between('D39:F39', 0, 9999, 'Enter the final calendar year or 0.', whole=True),
    _between('D40:F40', 0, 12, 'Enter a whole number of months from 0 to 12.', whole=True),
    _custom(_rows(54, 58), 'AND(ISNUMBER(D54),D54=INT(D54),OR(AND(D54>=0,D54<=125),D54=999))',
            'Enter a whole age from 0 to 125 or use 999 to turn this assumption off.'),
    _between('D60:F60', 0, 125, 'Enter a whole number of years from 0 to 125.', whole=True),
    _between(_rows(68, 98, 101, 116, 117), 0, 1, 'Enter 0 or 1 as shown in the row note.', whole=True),
    _ages('D95:F96 D104:F105', 126),
    _between('D102:F103', 0, 4, 'Enter a coverage code from 0 to 4.', whole=True),
    _share(_rows(55, 61, 69, 86, 89, 90, 91, 124, 126)),
    _minimum(_rows(12, *range(25, 28), *range(30, 39), 41, 42, 43, 44, 46, 47, 49, 51, 52, 56, 57, 59, 62,
                   *range(106, 116), *range(118, 121), 127, 128), 0, 'Enter zero or a positive number.'),
    _minimum('D122:F123', -1, 'Enter the income amount or −1 if unknown.'),
    _between('B75:B82', 1900, 9999, 'Enter a four-digit calendar year.', whole=True, optional=True),
    _minimum('C75:C82', 0, 'Enter zero or a positive expense.', optional=True),
    _between('B130', 2026, 9874, 'Enter a whole opening year from 2026 to 9874.', whole=True),
    _custom('B131', 'AND(ISNUMBER(B131),B131>=0,B131<1)', 'Enter a fee from 0% to less than 100%.'),
    _minimum('B132', -1, 'Enter inflation above −100%, or leave this blank.', strict=True, optional=True),
)

RULES = {
    'Inputs': _INPUT_RULES,
    'Asset_Allocation': (
        _minimum('B6:B10', -1, 'Enter an expected return above −100%.', strict=True),
        _minimum('C6:E10', 0, 'Enter zero or a positive percentage.'),
        _between('B16:F20', -1, 1, 'Enter a correlation from −1 to 1.'),
        _share('B26:E30 B35:E39 B44:E48'),
    ),
    'Income_Targeting': (_ages('B5'), _minimum('B21 B36', 0, 'Enter zero or a positive amount.')),
    'Risk_Settings': (
        _between('B5', 0, 125, 'Enter a whole number of years from 0 to 125.', whole=True),
        _share('B6:B8'), _between('B9', 1, 2, 'Enter return convention 1 or 2.', whole=True),
        _minimum('B14:B26', -1, 'Enter a return of at least −100%.'),
        _between('B29', 1, 1000, 'Enter a whole number of paths from 1 to 1,000.', whole=True),
        _between('B30', 0, 4294967295, 'Enter a whole seed from 0 to 4,294,967,295.', whole=True),
        _minimum('B31 B33', 0, 'Enter zero or a positive number.'),
        _between('B32', 0, 1, 'Enter 0 for fixed inflation or 1 to vary prices.', whole=True),
        _custom('B34', 'AND(ISNUMBER(B34),B34>=0,B34<1)', 'Enter persistence from 0% to less than 100%.'),
        _minimum('B35:B36', -1, 'Enter inflation above −100%.', strict=True),
    ),
    'Risk_Trials': (
        _between('B3', -1, 2207, 'Enter a whole trial number from the list on this sheet.', whole=True),
    ),
}


def apply(name, root):
    rules = RULES.get(name)
    if not rules:
        return
    validations = root.find(X + 'dataValidations')
    if validations is None:
        validations = ET.Element(X + 'dataValidations')
        later = ('hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter',
                 'rowBreaks', 'colBreaks', 'drawing', 'legacyDrawing', 'tableParts', 'extLst')
        following = next((node for node in root if ET.QName(node).localname in later), None)
        if following is None:
            root.append(validations)
        else:
            following.addprevious(validations)
    for cells, kind, operator, first, last, message, optional in rules:
        item = ET.SubElement(validations, X + 'dataValidation', sqref=cells, type=kind,
                             allowBlank='1' if optional else '0', showErrorMessage='1', errorStyle='stop',
                             errorTitle='Check this input', error=message)
        if operator:
            item.set('operator', operator)
        ET.SubElement(item, X + 'formula1').text = str(first)
        if last is not None:
            ET.SubElement(item, X + 'formula2').text = str(last)
    validations.set('count', str(len(validations)))


def core_input_checks():
    """Mirror entry rules in the plan status because pasting can bypass Excel validation."""
    from openpyxl.utils.cell import range_boundaries

    checks = [('OR(NOT(ISTEXT(B7)),LEN(TRIM(B7))=0)', 'Enter a scenario name on Inputs')]
    for areas, kind, operator, first, last, _, _ in _INPUT_RULES:
        # Later account, healthcare and tax inputs have their own model checks.
        rows = [r for area in areas.split() if area.startswith('D')
                for r in range(range_boundaries(area)[1], min(69, range_boundaries(area)[3]) + 1)]
        expressions = []
        for row in rows:
            # B55 is the survivor benefit in dollars. The selected scenario stores the share.
            ref = 'INDEX(D55:F55,1,B4)' if row == 55 else f'B{row}'
            if kind == 'custom':
                expression = first.replace('D54', ref)
            else:
                conditions = [f'ISNUMBER({ref})']
                if kind == 'whole':
                    conditions.append(f'{ref}=INT({ref})')
                sign = '>' if operator == 'greaterThan' else '>='
                conditions.append(f'{ref}{sign}{first}')
                if last is not None:
                    conditions.append(f'{ref}<={last}')
                expression = 'AND(' + ','.join(conditions) + ')'
            if row in (12, 36, 37):
                # The selected phase requires pay and spending only in years that use them.
                # Check the input itself: a reference to an empty scenario cell may display as zero.
                expression = f'OR(ISBLANK(INDEX(D{row}:F{row},1,B4)),{expression})'
            expressions.append(expression)
        if expressions:
            # Each IFERROR isolates bad pasted text or Excel errors to the relevant inputs.
            failing = 'NOT(IFERROR(AND(' + ','.join(expressions) + '),FALSE()))'
            first_row = rows[0]
            if len(rows) == 1:
                issue = f'Check Inputs row {first_row}'
            else:
                issue = _GROUP_MESSAGES.get(first_row, f'Check Inputs rows {first_row}–{rows[-1]}')
            checks.append((failing, issue))
    return checks


_GROUP_MESSAGES = {
    8: 'Check both ages, Social Security, the conversion end age and the plan horizon',
    11: 'Check the number of earners, Roth account type and withdrawal order on Inputs',
    12: 'Check for text or a negative amount in pay, savings, spending and income on Inputs',
    29: 'Check the first Roth tax year on Inputs',
    54: 'Check the age for survivor or long-term care costs on Inputs',
    55: 'Check the survivor benefit, tax rate and early withdrawal percentages on Inputs',
}


def expense_check():
    result = '"OK"'
    for row in reversed(range(75, 83)):
        year, cost = f'B{row}', f'C{row}'
        year_ok = f'AND(ISNUMBER({year}),{year}=INT({year}),{year}>=1900,{year}<=9999)'
        cost_ok = f'OR({cost}="",AND(ISNUMBER({cost}),{cost}>=0))'
        valid = (f'IF({year}="",OR({cost}="",AND(ISNUMBER({cost}),{cost}=0)),'
                 f'AND({year_ok},{cost_ok}))')
        result = f'IF(IFERROR({valid},FALSE()),{result},"Check the year and cost on Inputs row {row}")'
    return f'=IF(COUNTA(B83:C83)>0,"Move one-time expenses to Inputs rows 75–82",{result})'
