"""Fictional household assumptions for the distributed workbook."""

BASE_SCENARIO = {
    8: 35, 9: 33, 10: 60, 11: 2, 12: 95_000, 13: 0.035, 14: 0.025,
    15: 1, 16: 2,
    25: 15_000, 26: 20_000, 27: 8_000, 28: 1, 29: 2022, 30: 5_000,
    31: 0, 32: 5_000, 33: 2_000, 34: 0, 35: 4_000,
    36: 58_000, 37: 42_000, 38: 0, 39: 0, 40: 0, 41: 0.04,
    42: 12_000, 43: 8_000, 44: 40_000, 45: 67, 46: 0, 47: 0,
    48: 1, 49: 0, 50: 74, 51: 15_000, 52: 0, 53: 95, 54: 999,
    55: 0.6, 56: 0.7, 57: 0.6, 58: 999, 59: 0, 60: 0,
    61: 0.2, 62: 84_600, 67: 75, 68: 0, 69: 0.045,
    86: 1, 87: 0, 88: 125, 89: 1, 90: 0.5, 91: 0.5,
    92: 2022, 93: 0, 94: 0, 95: 126, 96: 126, 97: 1, 98: 1,
    101: 0, 102: 0, 103: 0, 104: 65, 105: 65, 106: 0, 107: 0,
    108: 0, 109: 0, 110: 0, 111: 0, 112: 0, 113: 0, 114: 0, 115: 0,
    116: 0, 117: 0, 118: 0, 119: 0, 120: 0, 121: 1, 122: -1,
    123: -1, 124: 0, 126: 0.5, 127: 0, 128: 0,
}


def example_cells():
    """Return all scenario inputs and clear the shared expense entry area."""
    inputs = {'B4': 1, 'B130': 2026, 'B131': 0.0025, 'B132': ''}
    for column, name, changes in (
        ('D', 'Example plan', {}),
        ('E', 'Retire at 65', {10: 65}),
        ('F', 'Spend 10% less', {37: round(BASE_SCENARIO[37] * 0.9, 2)}),
    ):
        inputs[f'{column}7'] = name
        inputs.update({f'{column}{row}': value for row, value in (BASE_SCENARIO | changes).items()})
    inputs.update({f'{column}{row}': '' for row in range(75, 83) for column in 'ABC'})
    inputs.update({'B83': '', 'C83': ''})
    return inputs


EXAMPLE_NOTICE = 'This workbook uses fictional assumptions. Save a private copy before entering your own figures.'
INSPECTION_NOTE = 'Trace draws in Funding_Calc, conversion trials in Targeting_Calc, and Summary charts in Report_Data.'
