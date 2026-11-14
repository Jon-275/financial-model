"""Direct Roth IRA limits for the filing statuses the model supports."""

import math


def configure(model):
    sheet = 'Tax_Tables'
    model.put(sheet, 'A253', 'DIRECT ROTH IRA CONTRIBUTIONS', 'band')
    amounts = (
        ('2027 IRA contribution limit', 7500),
        ('2027 IRA catch-up, age 50+', 1100),
        ('2026 joint Roth phase-out starts', 242000),
        ('2026 joint Roth phase-out ends', 252000),
        ('2026 single Roth phase-out starts', 153000),
        ('2026 single Roth phase-out ends', 168000),
        ('2027 joint Roth phase-out starts', 250000),
        ('2027 joint Roth phase-out ends', 260000),
        ('2027 single Roth phase-out starts', 158000),
        ('2027 single Roth phase-out ends', 173000),
    )
    for row, (label, value) in enumerate(amounts, 254):
        model.put(sheet, f'A{row}', label, 'label')
        model.put(sheet, f'B{row}', value, 'money')
    model.put(sheet, 'A264', 'Sources: IRS Notice 2025-67 and Notice 2026-66. Later years estimate limits using inflation.', 'note')
    model.put(sheet, 'A265', 'https://www.irs.gov/pub/irs-drop/n-26-66.pdf', 'note')


def formula_capacity(year, tax_index, age, single, magi):
    def table(row):
        return f'Tax_Tables!$B${row}'

    def updated(old, new):
        # Roth limits follow the tax year even when other tax tables use the baseline.
        return f'IF({year}>=2027,{table(new)}*{tax_index}/(1+Inputs!$B$14),{table(old)}*{tax_index})'

    limit = f'({updated(61, 254)}+IF({age}>=50,{updated(62, 255)},0))'
    low = f'IF({single},{updated(258, 262)},{updated(256, 260)})'
    high = f'IF({single},{updated(259, 263)},{updated(257, 261)})'
    return f'IF({magi}>={high},0,IF({magi}<={low},{limit},MAX(200,ROUNDUP({limit}*({high}-{magi})/({high}-{low})/10,0)*10)))'


def annual_limits(tax, year, index, inflation, single):
    """Return the annual cap, catch-up and phase-out bounds, if supported."""
    if 'B254' not in tax:
        return None
    if year >= 2027:
        start = 262 if single else 260
        factor = index / (1 + inflation)
        return tuple(tax[f'B{r}'] * factor for r in (254, 255, start, start + 1))
    start = 258 if single else 256
    return tuple(tax[f'B{r}'] * index for r in (61, 62, start, start + 1))


def reduced_limit(limit, magi, bounds):
    if bounds is None:
        return limit
    low, high = bounds
    if magi >= high:
        return 0
    if magi <= low:
        return limit
    return max(200, math.ceil(limit * (high - magi) / (high - low) / 10 - 1e-12) * 10)
