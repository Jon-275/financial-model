"""Encode duplicate assumption records only in the build's temporary Calc copy."""

from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree

import native_capture as trial
import native_funding_paths as funding
from native_results import _status
from xlsx_values import X, sheet_parts


def prepare_cache_copy(source, destination, count, context):
    """Use one short key after checking every stored record against the full value.

    Calc's XLSX exporter is very slow when thousands of cells contain the
    same 15,000-character string. The record is only tested for equality at
    these positions. Keep the long live and completed records intact and
    encode the verified duplicates just in the disposable calculation input.
    The release retains the full strings, original formulas and comparison caches.
    """
    checked_key = 'verified-completed-assumptions'
    expected_rows = {
        trial.SHEET: (range(trial.FIRST, trial.FIRST + 2 * count + 33), trial.SIGNATURE),
        funding.DATA_SHEET: (range(funding.FIRST_ROW, funding.LAST_ROW + 1), funding.SIGNATURE),
    }
    with ZipFile(source) as original:
        parts = sheet_parts(original)
        string_part = 'xl/sharedStrings.xml'
        strings = etree.fromstring(original.read(string_part))
        entries = strings.findall(X + 'si')
        matching = [i for i, entry in enumerate(entries)
                    if len(entry) == 1 and entry[0].tag == X + 't' and entry[0].text == context]
        if len(matching) != 1:
            raise ValueError('The captured batch is missing its unique shared assumption record.')
        record = str(matching[0])
        replacements = {}
        encoded_index = str(len(entries))
        for entry in entries:
            if ''.join(entry.itertext()) == checked_key:
                raise ValueError('The temporary dictionary key is already in use.')
        encoded_entry = etree.Element(X + 'si')
        etree.SubElement(encoded_entry, X + 't').text = checked_key
        strings.insert(len(entries), encoded_entry)
        strings.set('uniqueCount', str(len(entries) + 1))
        count_verified = 0
        for sheet, (rows, last_column) in expected_rows.items():
            expected = {f'{column}{row}' for row in rows for column in ('B', last_column)}
            part = parts[sheet]
            root = etree.fromstring(original.read(part))
            for cell in root.iter(X + 'c'):
                coordinate = cell.get('r')
                if coordinate not in expected:
                    continue
                value = cell.find(X + 'v')
                if (cell.get('t') != 's' or cell.find(X + 'f') is not None
                        or value is None or value.text != record):
                    raise ValueError(f'{sheet}!{coordinate} has a different captured assumption record.')
                value.text = encoded_index
                expected.remove(coordinate)
                count_verified += 1
            if expected:
                raise ValueError(f'{sheet} is missing a captured assumption record.')
            replacements[part] = etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)
        result_part = parts['Risk_Data']
        results = etree.fromstring(original.read(result_part))
        metadata = results.find(f".//{X}c[@r='Z8']")
        if (metadata is None or metadata.get('t') != 's' or metadata.findtext(X + 'v') != record
                or metadata.find(X+'f') is not None):
            raise ValueError('The completed run metadata differs from the captured assumptions.')
        # Z8 and all other cells retain their exact text. The full live record
        # is compared again in Calc before accepting the encoded trial rows.
        originals, encoded = {}, {}
        _status(lambda cell, formula, style: originals.update({cell: formula[1:]}))
        comparator = f'IF(EXACT($Z$8,{trial.CONTEXT}),"{checked_key}","different-assumptions")'
        _status(lambda cell, formula, style: encoded.update({cell: formula[1:]}), stored_context=comparator)
        for coordinate, original_formula in originals.items():
            formula = results.find(f".//{X}c[@r='{coordinate}']/{X}f")
            if formula is None or formula.text != original_formula:
                raise ValueError(f'Risk_Data!{coordinate} does not match the release calculation.')
            formula.text = encoded[coordinate]
        for part, tree in ((result_part, results), (string_part, strings)):
            replacements[part] = etree.tostring(tree, encoding='UTF-8', xml_declaration=True, standalone=True)
        with ZipFile(destination, 'x', ZIP_DEFLATED, compresslevel=5) as output:
            for part in original.namelist():
                output.writestr(part, replacements[part] if part in replacements else original.read(part))
    return count_verified
