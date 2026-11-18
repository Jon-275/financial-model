"""Reproducible 32-bit market draws shared by Excel and the comparison engine.

The xoshiro128** 1.1 transition follows Blackman and Vigna's public-domain
reference. Trial seeds use the public-domain MurmurHash3 finalizer on distinct
counter values. All Excel integer operations stay below 2**53.
"""
from math import cos, log, pi, sin, sqrt
from openpyxl.utils.cell import get_column_letter as column_letter

MASK = 2**32 - 1
MODULUS = 2**32
SEED_STEP = 0x9E3779B9
MIX_A = 0x85EBCA6B
MIX_B = 0xC2B2AE35


def _rotl(value, bits):
    return ((value << bits) | (value >> (32 - bits))) & MASK


def _mix(value):
    value ^= value >> 16
    value = value * MIX_A & MASK
    value ^= value >> 13
    value = value * MIX_B & MASK
    return value ^ (value >> 16)


def seed_state(seed, trial):
    if type(seed) is not int or not 0 <= seed <= MASK:
        raise ValueError('Seed must be a whole number from 0 to 4,294,967,295.')
    if type(trial) is not int or not 1 <= trial <= 1000:
        raise ValueError('Trial must be a whole number from 1 to 1,000.')
    return tuple(_mix((seed + (4 * (trial - 1) + word) * SEED_STEP) & MASK)
                 for word in range(1, 5))


def next_word(state):
    a, b, c, d = state
    result = _rotl(b * 5 & MASK, 7) * 9 & MASK
    shifted = b << 9 & MASK
    c ^= a
    d ^= b
    b ^= c
    a ^= d
    c ^= shifted
    d = _rotl(d, 11)
    return result, (a, b, c, d)


def normal_draws(seed, trial, years):
    """Yield six independent normal shocks per year, in stable column order."""
    if type(years) is not int or not 0 <= years <= 125:
        raise ValueError('Years must be a whole number from 0 to 125.')
    state = seed_state(seed, trial)
    for _ in range(years):
        normals = []
        for _ in range(3):
            first, state = next_word(state)
            second, state = next_word(state)
            radius = sqrt(-2 * log((first + .5) / MODULUS))
            angle = 2 * pi * (second + .5) / MODULUS
            normals.extend((radius * cos(angle), radius * sin(angle)))
        yield tuple(normals)


def excel_rotl(value, bits):
    return f'(MOD({value},{2**(32-bits)})*{2**bits}+INT({value}/{2**(32-bits)}))'


def excel_multiply(value, multiplier):
    """Exact low 32 bits, using 16-bit partial products that fit a double."""
    low, high = multiplier & 0xFFFF, multiplier >> 16
    return (f'MOD(MOD({value},65536)*{low}+'
            f'MOD(INT({value}/65536)*{low}+MOD({value},65536)*{high},65536)*65536,{MODULUS})')


def configure(put):
    """Author every seed, state transition and uniform draw as visible cells."""
    sheet = 'Random_Numbers'
    emit = lambda address, value, style=None: put(sheet, address, value, style)
    emit('A1', 'Random number calculation', 'detail_title')
    emit('A2', 'Choose a trial on Risk_Trials to inspect its draws. The base view displays trial 1.', 'report_caption')
    emit('A3', 'Trial shown', 'label')
    emit('B3', '=IF(AND(Risk_Trials!B3>=1,Risk_Trials!B3<=2000),MOD(Risk_Trials!B3-1,1000)+1,1)', 'whole')
    emit('T1', 'Trial seed calculation', 'detail_title')
    emit('T2', 'Counter mixing initializes four 32-bit state words for each stable trial number.', 'report_caption')
    seed = 'Risk_Settings!$B$30'
    valid_seed = f'IF(ISNUMBER({seed}),AND({seed}=INT({seed}),{seed}>=0,{seed}<={MASK}),FALSE())'
    emit('T3', f'=IF({valid_seed},"Seed is valid","Enter a whole-number seed from 0 to 4,294,967,295")', 'report_caption')
    headers = ('Draw', 'Year offset', 'Draw in year', 'Opening word 1', 'Opening word 2',
               'Opening word 3', 'Opening word 4', 'Random integer', 'Uniform draw',
               'Shifted word', 'Combined word 3', 'Combined word 4', 'Closing word 1',
               'Closing word 2', 'Closing word 3', 'Closing word 4')
    for number, title in enumerate(headers, 1):
        emit(f'{column_letter(number)}4', title, 'header')
    emit('T4', 'Trial', 'header')
    for word in range(4):
        for step, title in enumerate(('Counter', 'Mix 1', 'Multiply 1', 'Mix 2', 'Multiply 2', 'Seed word')):
            emit(f'{column_letter(21+word*6+step)}4', f'{title} {word+1}', 'header')
    for row in range(5, 1005):
        emit(f'T{row}', row - 4, 'whole')
        for word in range(4):
            start = 21 + word * 6
            addresses = [f'{column_letter(start+offset)}{row}' for offset in range(6)]
            a,b,c,d,e,_ = addresses
            expressions = (
                f'MOD({seed}+(4*($T{row}-1)+{word+1})*{SEED_STEP},{MODULUS})',
                f'_xlfn.BITXOR({a},INT({a}/65536))', excel_multiply(b,MIX_A),
                f'_xlfn.BITXOR({c},INT({c}/8192))', excel_multiply(d,MIX_B),
                f'_xlfn.BITXOR({e},INT({e}/65536))',
            )
            for address, expression in zip(addresses, expressions):
                emit(address, f'=IF({valid_seed},{expression},"")', 'whole')
    for row in range(5, 755):
        draw = row - 4
        valid = f'AND({draw}<=MAX(0,Market_Paths!$D$3-1)*6,{valid_seed})'
        expressions = {'A':str(draw),'B':f'INT(({draw}-1)/6)+1','C':f'MOD({draw}-1,6)+1'}
        for letter, previous, seeded in zip('DEFG', 'MNOP', ('Z','AF','AL','AR')):
            expressions[letter] = (f'INDEX(${seeded}$5:${seeded}$1004,$B$3)' if row==5
                                   else f'{previous}{row-1}')
        expressions.update({
            'H':f'MOD({excel_rotl(f"MOD(E{row}*5,{MODULUS})",7)}*9,{MODULUS})',
            'I':f'(H{row}+0.5)/{MODULUS}', 'J':f'MOD(E{row},{2**23})*512',
            'K':f'_xlfn.BITXOR(F{row},D{row})','L':f'_xlfn.BITXOR(G{row},E{row})',
            'M':f'_xlfn.BITXOR(D{row},L{row})','N':f'_xlfn.BITXOR(E{row},K{row})',
            'O':f'_xlfn.BITXOR(K{row},J{row})','P':excel_rotl(f'L{row}',11),
        })
        for column, expression in expressions.items():
            emit(f'{column}{row}', f'=IF({valid},{expression},"")', 'decimal' if column=='I' else 'whole')
