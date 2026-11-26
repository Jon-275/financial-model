"""Hand-calculated tax and funding cases independent of workbook caches."""

from dataclasses import replace
import unittest

from funding_selection_witness import earlier_feasible_draw
from risk_engine.funding import FundingInputs, _FundingSearch, solve_funding
from risk_engine.tax import TaxInputs, TaxTables


def joint_2027(primary_age=60, spouse_age=60):
    cells = {
        "B219": 1, "B231": 33200, "B233": 1700, "B235": 144700,
        "B237": 1031900, "B239": 252500, "B240": 192300,
        "B241": 11800, "B243": 316, "B244": 158, "B190": 3000,
        "B135": 32000, "B136": 44000, "B137": 0.5, "B138": 0.85,
        "B74": 0.038, "B75": 250000, "B82": 0.5, "B83": 0.26, "B84": 0.28,
        "B44": 0.01, "B45": 1000000, "B192": 504411, "B194": 2500, "B195": 6,
        "B46": 0.013, "B50": 0.062, "B52": 0.0145, "B53": 0.009, "B54": 250000,
    }
    schedules = (
        (6, (0, 25600, 104050, 218250, 416650, 529100, 793650),
         (0.10, 0.02, 0.10, 0.02, 0.08, 0.03, 0.02)),
        (17, (0, 102100, 633600), (0, 0.15, 0.05)),
        (30, (0, 22912, 54314, 85722, 118996, 150394, 768218, 921854, 1536426),
         (0.01, 0.01, 0.02, 0.02, 0.02, 0.013, 0.01, 0.01, 0.01)),
    )
    for first, floors, slopes in schedules:
        for row, (floor, slope) in enumerate(zip(floors, slopes), first):
            cells[f"E{row}"] = floor
            cells[f"D{row}"] = slope
    bands = (
        (1, 1.33, .0215, .0215), (1.33, 1.5, .0323, .0430), (1.5, 2, .0430, .0678),
        (2, 2.5, .0678, .0866), (2.5, 3, .0866, .1022), (3, 4, .1022, .1022),
    )
    for row, band in enumerate(bands, 221):
        cells.update({f"{column}{row}": value for column, value in zip("ABCD", band)})
    return TaxTables.from_cells(cells).for_year(2027, 0.0, 1.0, False, primary_age, spouse_age)


class CandidateTaxTests(unittest.TestCase):
    def test_wage_cap_per_earner(self):
        payroll, sdi = joint_2027().payroll_taxes(270000, 30000)
        self.assertAlmostEqual(payroll, 18582.60)
        self.assertAlmostEqual(sdi, 3900)

    def test_social_security_hsa_and_separate_capital_losses(self):
        year = joint_2027(primary_age=65)
        tax = year.evaluate(TaxInputs(
            35000, qualified_dividends=1000, tax_exempt_interest=2000,
            social_security=20000, hsa_contribution=2000, opening_ca_loss=4000,
        ), 0, 0, -6000)
        self.assertEqual(year.federal_deduction, 34900)
        self.assertEqual(tax.provisional_income, 45000)
        self.assertEqual(tax.taxable_ss, 6850)
        self.assertEqual(tax.agi, 39850)
        self.assertEqual(tax.aca_magi, 55000)
        self.assertEqual(tax.california_agi, 35000)
        self.assertEqual(tax.federal_loss, 3000)
        self.assertEqual(tax.ca_loss, 7000)
        self.assertEqual(tax.california_exemption_credit, 474)
        phaseout = year.evaluate(TaxInputs(521561 + 2500), 0, 0, 0)
        self.assertEqual(phaseout.california_exemption_credit, 456)

    def test_loss_below_deduction_remains_available(self):
        tax = joint_2027().evaluate(TaxInputs(10000), 0, 0, -6000)
        self.assertEqual(tax.total_tax, 0)
        self.assertEqual(tax.federal_loss, 6000)
        self.assertEqual(tax.ca_loss, 6000)

    def test_gains_are_capped_at_all_ordinary_tax(self):
        tax = joint_2027().evaluate(TaxInputs(136200), 0, 0, 500)
        self.assertAlmostEqual(tax.federal_ordinary, 11848)
        self.assertAlmostEqual(tax.federal_gains, 60)

    def test_amt_retains_preferential_rates_and_niit_uses_net_investment(self):
        tax = joint_2027().evaluate(TaxInputs(350000), 0, 0, 850000)
        self.assertAlmostEqual(tax.federal_ordinary, 60750)
        self.assertAlmostEqual(tax.federal_gains, 154160)
        self.assertAlmostEqual(tax.amt_tentative, 230128)
        self.assertAlmostEqual(tax.amt, 15218)
        self.assertAlmostEqual(tax.niit, 32300)


class FundingSearchTests(unittest.TestCase):
    def setUp(self):
        self.year = joint_2027()
        self.cash = FundingInputs(
            cash_income=0, required_spending=0, hsa_paid=0, opening_tax_payable=0,
            taxable_available=0, traditional_available=0, roth_available=0, mandatory=0,
            headroom=0, opening_traditional=0, basis_ratio=0, requested_conversion=0,
            order=1, conversion_allowed=False,
        )

    def test_mandatory_draw_can_cover_the_year_without_extra_sales(self):
        result = solve_funding(self.year, TaxInputs(0), replace(
            self.cash, required_spending=5, mandatory=10, traditional_available=100, opening_traditional=100,
        ))
        self.assertEqual(result.requested_draw, 0)
        self.assertEqual(result.traditional_draw, 10)
        self.assertEqual(result.margin, 5)

    def test_unfunded_year_liquidates_each_available_account_exactly(self):
        # These amounts leave a positive residual when summed capacities are
        # subtracted in the opposite order. A fully sold account must close.
        for order in (1, 2):
            for mandatory in (0, 1234.56789):
                for roth in (0, 1999.17):
                    with self.subTest(order=order, mandatory=mandatory, roth=roth):
                        cash = replace(
                            self.cash, required_spending=1000000,
                            taxable_available=4246.77694910289, traditional_available=50000.123,
                            roth_available=roth, mandatory=mandatory, headroom=1000000,
                            opening_traditional=50000.123, basis_ratio=2, order=order,
                        )
                        result = solve_funding(self.year, TaxInputs(0), cash)
                        self.assertEqual(result.status, 'Unfunded')
                        self.assertEqual(result.taxable_draw, cash.taxable_available)
                        self.assertEqual(result.traditional_draw, cash.traditional_available)
                        self.assertEqual(result.roth_draw, cash.roth_available)

    def test_preferred_boundary_uses_the_exact_available_balance(self):
        # Adding these separately rounded portions yields one ULP more than the
        # available traditional balance, enough to leave brokerage cost basis.
        cash = replace(
            self.cash, traditional_available=45195.9618, mandatory=9657.0105,
            taxable_available=41752.7438, roth_available=3125.71,
            headroom=50000, order=2,
        )
        search = _FundingSearch(self.year, TaxInputs(0), cash)
        traditional = cash.traditional_available
        brokerage = cash.taxable_available
        for draw, expected in (
            (0, (0, cash.mandatory, 0)),
            (traditional, (0, traditional, 0)),
            (traditional+brokerage, (brokerage, traditional, 0)),
            (search.capacity, (brokerage, traditional, cash.roth_available)),
        ):
            with self.subTest(draw=draw):
                self.assertEqual(search._draws(draw, False)[:3], expected)

    def test_preferred_draw_pays_its_own_tax(self):
        result = solve_funding(self.year, TaxInputs(33200), replace(
            self.cash, required_spending=36, taxable_available=50, traditional_available=70,
            roth_available=20, mandatory=10, headroom=20, opening_traditional=70, basis_ratio=1, order=2,
        ))
        self.assertAlmostEqual(result.requested_draw, 39)
        self.assertAlmostEqual(result.traditional_draw, 30)
        self.assertAlmostEqual(result.taxable_draw, 9)
        self.assertAlmostEqual(result.tax.total_tax, 3)

    def test_aca_can_fund_at_interior_draw_even_if_capacity_fails(self):
        cash = replace(self.cash, required_spending=365, taxable_available=350,
                       aca_actual=100, aca_benchmark=150, aca_fpl=100)
        result = solve_funding(self.year, TaxInputs(90), cash)
        self.assertEqual(result.status, "Funded")
        self.assertAlmostEqual(result.requested_draw, 265)
        self.assertEqual(result.premium_credit, 100)
        shortfall = solve_funding(self.year, TaxInputs(90), replace(cash, required_spending=470))
        self.assertEqual(shortfall.status, "Unfunded")
        self.assertAlmostEqual(shortfall.requested_draw, 310)
        self.assertAlmostEqual(shortfall.margin, -60)

    def test_search_preserves_better_credit_before_133_percent_jump(self):
        result = solve_funding(self.year, TaxInputs(10000), replace(
            self.cash, required_spending=13000, taxable_available=40000,
            aca_actual=10000, aca_benchmark=10000, aca_fpl=10000,
        ))
        self.assertAlmostEqual(result.requested_draw, 3215 / 0.9785)
        self.assertLess(result.tax.aca_magi, 13300)

    def test_loss_sale_can_enter_aca_eligibility_at_400_percent(self):
        result = solve_funding(self.year, TaxInputs(405), replace(
            self.cash, required_spending=105, taxable_available=30, basis_ratio=2,
            aca_actual=100, aca_benchmark=150, aca_fpl=100,
        ))
        self.assertEqual(result.status, "Funded")
        self.assertAlmostEqual(result.requested_draw, 5)
        self.assertAlmostEqual(result.realized_gain, -5)
        self.assertAlmostEqual(result.tax.aca_magi, 400)
        self.assertEqual(result.premium_credit, 100)

    def test_credit_includes_100_percent_endpoint(self):
        result = solve_funding(self.year, TaxInputs(104), replace(
            self.cash, required_spending=104, taxable_available=50, basis_ratio=2,
            aca_actual=100, aca_benchmark=150, aca_fpl=100,
        ))
        self.assertEqual(result.status, "Funded")
        self.assertAlmostEqual(result.requested_draw, 4)
        self.assertEqual(result.tax.aca_magi, 100)
        self.assertEqual(result.premium_credit, 100)

    def test_feasible_conversion_has_priority_over_a_smaller_no_conversion_draw(self) -> None:
        """Hold the eligible conversion choice fixed when evaluating a lower draw."""
        income = TaxInputs(33_200)
        cash = replace(self.cash, cash_income=33_200, required_spending=34_200,
                       taxable_available=20_000, basis_ratio=1, traditional_available=10_000,
                       opening_traditional=10_000, requested_conversion=5_000, conversion_allowed=True)
        selected = solve_funding(self.year, income, cash)
        no_conversion = solve_funding(self.year, income, replace(cash, conversion_allowed=False))
        # The $1,000 cash gap increases by $500 of federal tax on the conversion.
        self.assertEqual((selected.status, selected.conversion, selected.tax.total_tax), ('Funded', 5_000, 500))
        self.assertAlmostEqual(selected.requested_draw, 1_500)
        self.assertAlmostEqual(no_conversion.requested_draw, 1_000)
        quote = _FundingSearch(self.year, income, cash)
        lower = selected.requested_draw - .01
        self.assertAlmostEqual(quote._quote(lower, True)[0], -.01)
        self.assertAlmostEqual(quote._quote(lower, False)[0], 499.99)
        self.assertIsNone(earlier_feasible_draw(selected.requested_draw, lambda draw: quote._quote(draw, True)[0]))
        self.assertIsNotNone(earlier_feasible_draw(selected.requested_draw,
                                                  lambda draw: quote._quote(draw, False)[0]))

    def test_same_mode_lower_draw_can_restore_conversion_when_the_claim_converts_zero(self) -> None:
        """Distinguish zero converted at the draw from an ineligible conversion."""
        income = TaxInputs(33_200)
        base = replace(self.cash, cash_income=33_200, required_spending=37_700,
                       requested_conversion=1_000, conversion_allowed=True)
        for cash in (
            replace(base, traditional_available=5_000, opening_traditional=5_000),
            replace(base, traditional_available=10_000, opening_traditional=10_000,
                    headroom=5_000, order=2, taxable_available=1_000, basis_ratio=1),
        ):
            with self.subTest(order=cash.order):
                # The selected traditional draw exhausts the balance in order 1
                # and fills the preferred bracket headroom in order 2.
                prior = _FundingSearch(self.year, income, cash)
                selected = prior.solve(exact=False)
                self.assertTrue(prior.conversion)
                self.assertGreaterEqual(max(prior._grid(True).values()), 0)
                self.assertEqual((selected.status, selected.traditional_draw, selected.conversion),
                                 ('Funded', 5_000, 0))
                self.assertAlmostEqual(selected.requested_draw, 5_000)
                lower = selected.requested_draw - .01
                self.assertAlmostEqual(prior._draws(lower, True)[3], .01)
                self.assertEqual(prior._draws(lower, False)[3], 0)
                self.assertAlmostEqual(prior._quote(lower, True)[0], -.01)
                self.assertIsNone(earlier_feasible_draw(selected.requested_draw,
                                                       lambda draw: prior._quote(draw, True)[0]))

                # With $100 less spending, the old gross draw still converts
                # zero, while a lower same-mode quote funds and converts again.
                current = _FundingSearch(self.year, income, replace(cash, required_spending=37_600))
                actual = current.solve(exact=False)
                self.assertEqual(actual.status, 'Funded')
                self.assertAlmostEqual(actual.requested_draw, 4_900)
                self.assertAlmostEqual(actual.conversion, 100)
                self.assertEqual(current._draws(selected.requested_draw, True)[3], 0)
                self.assertAlmostEqual(current._draws(lower, True)[3], .01)
                witness = earlier_feasible_draw(selected.requested_draw, lambda draw: current._quote(draw, True)[0])
                self.assertIsNotNone(witness)
                assert witness is not None
                self.assertAlmostEqual(witness[0], lower)
                self.assertAlmostEqual(witness[1], 99.99)

    def _aca_cutoff_search(self, spending: float, capacity: float = 410) -> _FundingSearch:
        """Use small fictional amounts to isolate the tax-free ACA grid."""
        return _FundingSearch(self.year, TaxInputs(90), replace(
            self.cash, cash_income=90, required_spending=spending, taxable_available=capacity,
            aca_actual=100, aca_benchmark=150, aca_fpl=100,
        ))

    def test_unfunded_choice_uses_best_margin_and_smallest_draw_within_tie(self) -> None:
        """Compare the inclusive ACA cutoff with full capacity on the fresh grid."""
        for improvement, expected in ((0, 310), (5e-8, 310), (2e-7, 410 + 2e-7)):
            with self.subTest(improvement=improvement):
                capacity = 410 + improvement
                search = self._aca_cutoff_search(spending=560, capacity=capacity)
                grid = search._grid(False)
                cutoff_draws = {draw + offset for draw in (10, 43, 310) for offset in (-.00001, 0, .00001)}
                self.assertTrue(cutoff_draws.issubset(grid))
                # At a $310 draw, MAGI is exactly 400% of FPL and the credit
                # is $100. Just above that point the entire credit is lost.
                self.assertEqual(search._quote(310, False), (-60, 400))
                self.assertAlmostEqual(grid[309.99999], -60.00001)
                self.assertAlmostEqual(grid[310.00001], -159.99999)
                self.assertAlmostEqual(grid[capacity] - grid[310], improvement, places=12)
                self.assertEqual(max(grid.values()), grid[capacity])
                self.assertTrue(all(margin < 0 for margin in grid.values()))
                result = search.solve(exact=False)
                self.assertEqual((result.status, result.requested_draw, result.lower_draw), ('Unfunded', expected, 0))
                self.assertAlmostEqual(result.margin, grid[expected])

    def test_grid_requires_nonnegative_cash_even_when_final_label_is_funded(self) -> None:
        """Keep feasibility at zero separate from the final display threshold."""
        for spending, feasible, label in ((500, True, 'Funded'), (500.004, False, 'Funded'),
                                          (500.010001, False, 'Unfunded')):
            with self.subTest(spending=spending):
                search = self._aca_cutoff_search(spending)
                grid = search._grid(False)
                self.assertEqual(any(margin >= 0 for margin in grid.values()), feasible)
                self.assertAlmostEqual(max(grid.values()), 500 - spending)
                result = search.solve(exact=False)
                self.assertEqual((result.status, result.requested_draw), (label, 310))
                self.assertEqual(result.lower_draw > 0, feasible)

        for spending, label in ((.009999, 'Funded'), (.0099990001, 'Unfunded')):
            with self.subTest(spending=spending):
                search = _FundingSearch(self.year, TaxInputs(0), replace(self.cash, required_spending=spending))
                self.assertEqual(search._grid(False), {0: -spending})
                result = search.solve(exact=False)
                self.assertEqual((result.status, result.requested_draw, result.margin), (label, 0, -spending))

    def test_witness_requires_more_than_half_a_cent_of_cash(self) -> None:
        """A nonnegative lower quote may still be below the proof tolerance."""
        for spending, expected in ((.005001, None), (.005, None), (.004999, (.01, .005001))):
            with self.subTest(spending=spending):
                cash = replace(self.cash, taxable_available=1, basis_ratio=1, required_spending=spending)
                search = _FundingSearch(self.year, TaxInputs(0), cash)
                self.assertGreater(search._quote(.01, False)[0], 0)
                self.assertEqual(earlier_feasible_draw(.02, lambda draw: search._quote(draw, False)[0]), expected)

    def test_funded_grid_bracket_must_still_converge_within_half_a_cent(self) -> None:
        """Isolate the bracket-width check after the fixed bisection budget."""
        # For these power-of-two capacities, 49 bisections leave exact widths
        # of 1/256 and 1/128 dollars. Both upper draws can pay all costs.
        for capacity, width, label in ((2**41, 1/256, 'Funded'),
                                       (2**42, 1/128, 'Check funding convergence')):
            with self.subTest(capacity=capacity):
                cash = replace(self.cash, taxable_available=capacity, basis_ratio=1, required_spending=1_000.123)
                result = solve_funding(self.year, TaxInputs(0), cash)
                self.assertEqual((result.status, result.requested_draw), (label, 1_000.125))
                self.assertEqual(result.requested_draw - result.lower_draw, width)
                self.assertAlmostEqual(result.margin, .002)

    def test_unfunded_regular_projection_uses_no_conversion_grid_margin(self) -> None:
        """Preserve policy priority when both modes have a negative best margin."""
        cash = replace(self.cash, cash_income=90, required_spending=300, taxable_available=50,
                       basis_ratio=1, opening_traditional=20, requested_conversion=20, conversion_allowed=True,
                       aca_actual=100, aca_benchmark=150, aca_fpl=100)
        search = _FundingSearch(joint_2027(primary_age=59), TaxInputs(90), cash)
        self.assertEqual(max(search._grid(True).values()), -60)
        self.assertEqual(max(search._grid(False).values()), -160)
        regular = search.solve(exact=False)
        exact = search.solve(exact=True)
        self.assertEqual((regular.status, regular.requested_draw, regular.conversion, regular.margin),
                         ('Unfunded', 50, 0, -160))
        self.assertEqual((exact.status, exact.requested_draw, exact.conversion, exact.margin),
                         ('Unfunded', 50, 20, -60))

    def test_projection_can_skip_unaffordable_conversion_comparison_cannot(self):
        cash = replace(self.cash, opening_traditional=10000, requested_conversion=5000, conversion_allowed=True)
        search = _FundingSearch(joint_2027(primary_age=59), TaxInputs(33000), cash)
        self.assertEqual(max(search._grid(True).values()), -480)
        self.assertEqual(max(search._grid(False).values()), 0)
        projection = search.solve(exact=False)
        comparison = search.solve(exact=True)
        self.assertEqual(projection.conversion, 0)
        self.assertEqual(projection.status, "Funded")
        self.assertEqual(comparison.conversion, 5000)
        self.assertEqual(comparison.status, "Unfunded")
        self.assertAlmostEqual(comparison.margin, -480)
