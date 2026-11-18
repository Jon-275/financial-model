"""Regression for a flexible trial containing the preceding trial's last years."""
from math import inf, nan, nextafter
import unittest

from native_capture_acceptance import expected_cut, validate_cut_decisions, validate_uncut_pairs
from native_excel_acceptance import check_reuse_note
import native_capture as trial
from xlsx_read import Cell, Sheet, address_parts, col_letter


class UncutPairTests(unittest.TestCase):
    def setUp(self):
        self.data = Sheet('Risk_Trials')
        self.positions = {59: 66, 1059: 166}
        for row in (66, 166):
            for col, value in [('B', 'test context'), (col_letter(trial.SIGNATURE_COLUMN), 'test context'),
                               ('I', 0), ('J', 0), ('CO', 5329330), ('CP', 5615108), ('G', 5615108), ('AB', row + 40000)]:
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
        self.write('J166', 300)
        self.write('G166', 5500000)
        validate_uncut_pairs(self.data, self.positions)

    def test_missing_cut_metadata_cannot_classify_a_path_as_cut_or_uncut(self):
        for column, value in (('I', None), ('I', '0'), ('I', False), ('I', 0.1), ('J', nan), ('J', 1)):
            with self.subTest(column=column, value=value):
                self.setUp()
                self.write(f'{column}166', value)
                with self.assertRaises(AssertionError):
                    validate_uncut_pairs(self.data, self.positions)

    def test_uncut_and_cut_pairs_require_current_opening_and_closing_context(self):
        for count in (0, 1):
            for col, row in (('B',66), ('B',166), (col_letter(trial.SIGNATURE_COLUMN),66),
                             (col_letter(trial.SIGNATURE_COLUMN),166)):
                with self.subTest(count=count, col=col, row=row):
                    self.setUp()
                    self.write('I166', count)
                    self.write('J166', 100 * count)
                    self.write(f'{col}{row}', 'stale')
                    with self.assertRaisesRegex(AssertionError, 'assumption record'):
                        validate_uncut_pairs(self.data, self.positions)


class FlexibleDecisionTests(unittest.TestCase):
    def test_retirement_opens_below_but_never_at_threshold(self):
        years = [(100, False), (100, True), (0, True)]
        for fixed, expected in ((nextafter(80, -inf), True), (80, False), (nextafter(80, inf), False)):
            with self.subTest(fixed=fixed):
                self.assertEqual(expected_cut([0, fixed, 0], years, .8, .15, .3), expected)

    def test_policy_off_and_ineligible_working_or_mortgage_years(self):
        for settings in ((0, .15, .3), (.8, 0, .3), (.8, .15, 0)):
            self.assertFalse(expected_cut([0], [(100, True)], *settings))
        self.assertFalse(expected_cut([0, 0], [(100, False), (100, False)], .8, .15, .3))
        self.assertTrue(expected_cut([0], [(100, True)], .8, .15, .3))

    def test_unknown_amounts_do_not_support_any_equivalence_claim(self):
        for bad in (True, '', '80', None, '#VALUE!', -1, inf, nan):
            with self.subTest(bad=bad), self.assertRaises(AssertionError):
                expected_cut([50, bad], [(100, True), (100, True)], .8, .15, .3)
        with self.assertRaises(ValueError):
            expected_cut([50], [(100, True), (100, True)], .8, .15, .3)

    def test_native_run_note_matches_visible_status_and_cannot_exceed_uncut_pairs(self):
        prefix = 'Updated 1,000 paired paths and 33 stress cases. '
        for value, expected in (('0', 0), ('515', 515), ('1,013', 1013), ('1\u202f013', 1013)):
            text = prefix + value + ' flexible paths used their verified fixed results.'
            self.assertEqual(check_reuse_note(text, text, expected), expected)
        for paired in ('1 paired path', '10 paired paths'):
            singular = f'Updated {paired} and 33 stress cases. 1 flexible path used its verified fixed result.'
            self.assertEqual(check_reuse_note(singular, singular, 1, expected=1), 1)
        malformed = prefix + '1 flexible path used their verified fixed results.'
        with self.assertRaises(AssertionError):
            check_reuse_note(malformed, malformed, 1)
        text = prefix + '516 flexible paths used their verified fixed results.'
        for visible, ceiling in ((text, 515), ('stale message', 1013)):
            with self.assertRaises(AssertionError):
                check_reuse_note(text, visible, ceiling)
        with self.assertRaisesRegex(AssertionError, 'Expected 515'):
            check_reuse_note(prefix + '0 flexible paths used their verified fixed results.',
                             prefix + '0 flexible paths used their verified fixed results.', 515, expected=515)

    def test_first_opening_last_opening_and_excluded_final_balance(self):
        def write(sheet, column, row, value):
            coord = f'{column}{row}'
            number = address_parts(coord)[1]
            sheet.rows.setdefault(row, {})[number] = Cell(coord, value, 'n')
        book = {name: Sheet(name) for name in ('Inputs','Projection','Risk_Settings')}
        inputs, projection, settings = book.values()
        for row, number in ((70,4), (37,300), (38,100), (56,.7)):
            write(inputs,'B',row,number)
        for row, number in ((6,.8), (7,.15), (8,.3)):
            write(settings,'B',row,number)
        for row in range(5,9):
            write(projection,'BH',row,100)
            write(projection,'E',row,'Medicare')
            write(projection,'F',row,'MFJ')
        data = Sheet('Risk_Trials')
        fixed_row, flexible_row = 10, 20
        for pair, paths, expected_count in [
                ((1,1001), [79,100,100,100], 1), ((1000,2000), [100,100,79,100], 1),
                ((2001,2101), [100,100,100,79], 0), ((2013,2113), [80,80,80,80], 0)]:
            with self.subTest(pair=pair):
                positions = dict(zip(pair,(fixed_row,flexible_row)))
                for offset, value in enumerate(paths):
                    write(data,col_letter(trial.ANNUAL_FIRST + trial.ANNUAL_CAPACITY + offset),fixed_row,value)
                write(data,'I',flexible_row,expected_count)
                validate_cut_decisions(book,data,positions)
                write(data,'I',flexible_row,1 - expected_count)
                with self.assertRaisesRegex(AssertionError, 'disagrees'):
                    validate_cut_decisions(book,data,positions)


if __name__ == '__main__':
    unittest.main()
