"""Regression for a flexible trial containing the preceding trial's last years."""
import unittest

from native_capture_acceptance import validate_uncut_pairs
from xlsx_read import Cell, Sheet, address_parts


class UncutPairTests(unittest.TestCase):
    def setUp(self):
        self.data = Sheet('Risk_Trials')
        self.positions = {59: 66, 1059: 166}
        for row in (66, 166):
            for col, value in [('I', 0), ('CO', 5329330), ('CP', 5615108), ('G', 5615108), ('AB', row + 40000)]:
                self.write(f'{col}{row}', value)

    def write(self, address, value):
        row, col = address_parts(address)
        self.data.rows.setdefault(row, {})[col] = Cell(address, value, 'n')

    def test_no_cut_pair_matches_even_across_a_refresh_date_change(self):
        validate_uncut_pairs(self.data, self.positions)

    def test_previous_paths_balance_is_rejected_even_if_the_ending_amount_matches(self):
        self.write('CO166', 3618645)
        with self.assertRaisesRegex(AssertionError, '1059 differs from fixed trial 59 in CO166'):
            validate_uncut_pairs(self.data, self.positions)

    def test_actual_spending_cuts_allow_a_different_balance(self):
        self.write('I166', 2)
        self.write('G166', 5500000)
        validate_uncut_pairs(self.data, self.positions)


if __name__ == '__main__':
    unittest.main()
