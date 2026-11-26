"""Verify the disposable quote uses current candidate economics and no solver bank."""

from graphlib import TopologicalSorter
import unittest

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import get_column_letter, range_boundaries

from native_funding_witness_formulas import witness_formulas


class NativeFundingWitnessFormulaTests(unittest.TestCase):
    def test_quote_reads_prefunding_cash_and_previous_year_state(self) -> None:
        """Use the same funding inputs while excluding funded cash and solver outputs."""
        generated = witness_formulas((126,))
        self.assertIn('Trial_Cash_Flow!$AB126', generated['Z126'])
        self.assertNotIn('Trial_Cash_Flow!$C126', generated['Z126'])
        self.assertIn('SUM(H126:J126)', generated['Z126'])
        self.assertIn('Trial_Projection!$BE125', generated['K126'])
        self.assertIn('G126=1', generated['K126'])
        self.assertIn('Trial_Inputs!$B$47', generated['K126'])
        self.assertIn('Trial_Projection!$BW125', generated['AD126'])
        self.assertIn('Trial_Projection!$BY125', generated['AF126'])
        self.assertIn('Trial_Healthcare!$M126', generated['AB126'])
        self.assertIn('Trial_Healthcare!$O126', generated['AB126'])
        self.assertIn('Tax_Tables!', generated['AB126'])

    def test_breakpoints_and_quote_dependencies_stay_in_the_aligned_year(self) -> None:
        """Exercise the first, trace and last rows, including the former grid control row."""
        rows = (6, 78, 126, 130)
        generated = witness_formulas(rows)
        for row in rows:
            with self.subTest(row=row):
                self.assertIn(f'$AM{row}', generated[f'H{row}'])
                self.assertIn(f'$AN{row}', generated[f'I{row}'])
                self.assertIn(f'$AO{row}', generated[f'I{row}'])
                self.assertIn(f'Trial_Cash_Flow!$H{row}', generated[f'AN{row}'])
                self.assertIn(f'Trial_Cash_Flow!$I{row}', generated[f'AO{row}'])
                self.assertIn(f'{row}-4<=Trial_Inputs!$B$70', generated[f'D{row}'])
                self.assertEqual(generated[f'G{row}'], f'=IF(ISNUMBER($D{row}),$C{row},"")')
                self.assertNotIn('ROW()', generated[f'Z{row}'])
        sources = {'Trial_Inputs', 'Trial_Projection', 'Trial_Cash_Flow', 'Trial_Healthcare', 'Tax_Tables'}
        used = {item.value.rsplit('!', 1)[0] for formula in generated.values() for item in Tokenizer(formula).items
                if item.type == 'OPERAND' and item.subtype == 'RANGE' and '!' in item.value}
        self.assertEqual(used, sources)

    def test_all_local_references_are_acyclic_and_defined_in_the_quote_row(self) -> None:
        """Prevent a copied candidate from retaining a hidden grid or another year's value."""
        row = 126
        generated = witness_formulas((row,))
        available = generated.keys() | {f'{column}{row}' for column in 'ABC'}
        graph = TopologicalSorter()
        for address, formula in generated.items():
            graph.add(address)
            for item in Tokenizer(formula).items:
                if item.type != 'OPERAND' or item.subtype != 'RANGE' or '!' in item.value:
                    continue
                left, top, right, bottom = range_boundaries(item.value)
                self.assertEqual((top, bottom), (row, row), (address, item.value))
                for column in range(left, right + 1):
                    dependency = f'{get_column_letter(column)}{row}'
                    self.assertIn(dependency, available, address)
                    graph.add(address, dependency)
        self.assertEqual(set(graph.static_order()), available)

    def test_disposition_only_refutes_supported_funded_candidates(self) -> None:
        """Keep uncertain branches, near-cent cases and failed quotes inconclusive."""
        generated = witness_formulas((126,))
        draw = generated['D126']
        self.assertIn('EXACT($A126,"Funded")', draw)
        self.assertIn('ISNUMBER($C126)', draw)
        self.assertIn('OR($C126=0,$C126=1)', draw)
        self.assertIn('$B126<=Trial_Cash_Flow!$K126', draw)
        self.assertIn('$B126-($B126-0.01)>0.005', draw)
        self.assertIn('$B126-0.01', draw)
        report = generated['AP126']
        self.assertIn('EXACT($A126,"Unfunded"),"grid-required"', report)
        self.assertIn('ISNUMBER($D126),ISNUMBER($Z126),$Z126>0.005', report)
        self.assertIn('"earlier-feasible","inconclusive"', report)
        self.assertNotIn('"accepted"', report)

    def test_rejects_opening_and_beyond_capacity_rows(self) -> None:
        """Only the workbook's active annual row range has funding quotes."""
        for row in (5, 131, 6.0, True):
            with self.subTest(row=row), self.assertRaisesRegex(ValueError, '6 through 130'):
                witness_formulas((row,))
