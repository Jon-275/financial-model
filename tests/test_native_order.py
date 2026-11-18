"""The dependency check must reject a missing stage and unsafe dynamic addresses."""
import unittest
from pathlib import Path
from native_order_acceptance import annual_reference, public_input_link, references, solver_references, verify_funding_cleanup, walk_order


class OrderedCalculationTests(unittest.TestCase):
    def test_a_range_resolves_its_internal_dependencies(self):
        graph={('market',1):set(),('market',2):{('market',1)},('projection',1):{('market',2)}}
        walk_order(graph,('market','projection'))

    def test_earlier_range_must_recalculate_after_new_cross_sheet_precedent(self):
        graph={('projection',1):set(),('funding',1):{('projection',1)},('projection',2):{('funding',1)}}
        with self.assertRaisesRegex(ValueError,'unresolved'):
            walk_order(graph,('projection','funding'))
        walk_order(graph,('projection','funding','projection'))

    def test_cyclic_and_unknown_cell_references_do_not_pass(self):
        with self.assertRaisesRegex(ValueError,'unresolved'):
            walk_order({('projection',1):{('cash',1)},('cash',1):{('projection',1)}},('projection','cash')*2)
        for formula in ('INDIRECT(A1)', 'OFFSET(A1,1,0)', 'UnclassifiedSourceRange', '[other.xlsx]Inputs!B8'):
            with self.subTest(formula=formula), self.assertRaises(ValueError):
                list(references(formula,'Trial_Projection'))

    def test_future_year_rejected_and_previous_year_already_settled(self):
        current, = references('Trial_Projection!BH7', 'Trial_Checks')
        fields = {('Trial_Projection', 60)}
        self.assertEqual(annual_reference(current, 7, fields), fields)
        self.assertEqual(annual_reference(current, 8, fields), set())
        with self.assertRaisesRegex(ValueError, 'Forward financial dependency'):
            annual_reference(current, 6, fields)

    def test_trial_expenses_must_reach_the_same_public_input(self):
        with self.assertRaisesRegex(ValueError, 'Missing Trial_Inputs!C75'):
            public_input_link(3, 75, {})
        with self.assertRaisesRegex(ValueError, 'does not read its public input'):
            public_input_link(3, 75, {(3, 75): 'Inputs!C76'})
        public_input_link(3, 75, {(3, 75): 'IF(ISBLANK(Inputs!C75),"",Inputs!C75)'})

    def test_shared_solver_formula_cannot_read_a_future_year_at_its_far_end(self):
        with self.assertRaisesRegex(ValueError, 'reads outside its year'):
            list(solver_references(2, 2, 'A3', 'B2:B124', 2, 41))
        local = list(solver_references(2, 2, 'A2', 'B2:B124', 2, 41))
        self.assertEqual({ref.top for ref in local}, {2, 124})

    def test_cumulative_index_endpoints_are_both_visible(self):
        rr=list(references('MAX(INDEX($G$5:G12,2):G12)','Trial_Projection'))
        self.assertEqual([(r.left,r.top,r.right,r.bottom) for r in rr],[(7,5,7,12),(7,12,7,12)])

    def test_inactive_solver_years_follow_the_model_horizon_and_full_table_width(self):
        vba = (Path(__file__).resolve().parents[1] / 'tools/vba/RunMarketRisk.bas').read_text()
        horizon, = references('Trial_Inputs!$B$70', 'Trial_Funding_Calc')
        verify_funding_cleanup(vba, horizon, 41)
        for actual in (None, next(references('Inputs!$B$70', 'Trial_Funding_Calc'))):
            with self.subTest(horizon=actual), self.assertRaisesRegex(ValueError, 'horizon'):
                verify_funding_cleanup(vba, actual, 41)
        for old, new in (('FUNDING_ROWS_PER_YEAR As Long = 123', 'FUNDING_ROWS_PER_YEAR As Long = 122'),
                         ('":AO" & lastRow', '":AN" & lastRow'),
                         ('If years > 1 Then firstRow', 'If years > 2 Then firstRow')):
            with self.subTest(change=old), self.assertRaisesRegex(ValueError, 'Inactive funding'):
                verify_funding_cleanup(vba.replace(old, new), horizon, 41)

if __name__=='__main__':
    unittest.main()
