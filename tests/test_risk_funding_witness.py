"""Reprice the fictional age-121 draw that survived a closing-balance check."""

from dataclasses import replace
import unittest

from funding_selection_witness import earlier_feasible_draw
from risk_engine.funding import FundingInputs, _FundingSearch, solve_funding
from risk_engine.tax import TaxInputs, TaxTables, YearTax


PRIOR_REQUEST = 3_991_289.26141221
CURRENT_REQUEST = 3_761_432.87464338
CURRENT_RETURN = .228424543484169
PRIOR_INCOME = TaxInputs(0, qualified_dividends=3_544_583.8795466814,
                         investment_interest=6_288_777.850808628, social_security=793_684.1431848888)
CURRENT_INCOME = TaxInputs(0, qualified_dividends=3_425_146.5728798574,
                           investment_interest=6_076_872.951883618, social_security=793_684.1431848888)
CURRENT_CASH = FundingInputs(
    cash_income=793_684.1431848888, required_spending=1_244_353.8733681329,
    hsa_paid=0, opening_tax_payable=0, taxable_available=315_681_711.78616196,
    traditional_available=14_935.104862580854, roth_available=94_323_005.41748668,
    mandatory=7_467.552431290427, headroom=0, opening_traditional=14_935.104862580854,
    basis_ratio=.5602285604961862, requested_conversion=0, order=1, conversion_allowed=False,
)
PRIOR_CASH = replace(
    CURRENT_CASH, taxable_available=326_689_758.4835651, traditional_available=13_618.680722342619,
    roth_available=87_815_041.54822695, mandatory=6_809.340361171309,
    opening_traditional=13_618.680722342619, basis_ratio=.48684212438049773,
)


def fixture_tax() -> YearTax:
    """Use the frozen tables and age/year index from the independent reference."""
    cells = {
        'B44': .01, 'B45': 1_000_000, 'B50': .062, 'B52': .0145, 'B53': .009, 'B54': 250_000,
        'B74': .038, 'B75': 250_000, 'B82': .5, 'B83': .26, 'B84': .28, 'B135': 32_000,
        'B136': 44_000, 'B137': .5, 'B138': .85, 'B190': 3_000, 'B192': 504_411,
        'B194': 2_500, 'B195': 6, 'B219': 1, 'B231': 33_200, 'B233': 1_700,
        'B235': 144_700, 'B237': 1_031_900, 'B239': 252_500, 'B240': 192_300,
        'B241': 11_800, 'B243': 316, 'B244': 158, 'B251': .014,
    }
    brackets = (
        (6, (0, 25_600, 104_050, 218_250, 416_650, 529_100, 793_650),
         (.10, .02, .10, .02, .08, .03, .02)),
        (17, (0, 102_100, 633_600), (0, .15, .05)),
        (30, (0, 22_912, 54_314, 85_722, 118_996, 150_394, 768_218, 921_854, 1_536_426),
         (.01, .01, .02, .02, .02, .013, .01, .01, .01)),
    )
    for start, floors, slopes in brackets:
        for row, (floor, slope) in enumerate(zip(floors, slopes, strict=True), start):
            cells[f'E{row}'], cells[f'D{row}'] = floor, slope
    bands = (
        (1, 1.33, .0215, .0215), (1.33, 1.5, .0323, .043), (1.5, 2, .043, .0678),
        (2, 2.5, .0678, .0866), (2.5, 3, .0866, .1022), (3, 4, .1022, .1022),
    )
    for row, band in enumerate(bands, 221):
        cells.update({f'{col}{row}': amount for col, amount in zip('ABCD', band, strict=True)})
    return TaxTables.from_cells(cells).for_year(
        2147, .025, 19.8421035796222, single=False, primary_age=121, spouse_age=121,
    )


class FundingWitnessTests(unittest.TestCase):
    def test_stale_draw_can_balance_and_still_have_a_smaller_funded_quote(self) -> None:
        """Price both observed requests using each path's own independent inputs."""
        tax = fixture_tax()
        prior = _FundingSearch(tax, PRIOR_INCOME, PRIOR_CASH)
        current = _FundingSearch(tax, CURRENT_INCOME, CURRENT_CASH)
        self.assertFalse(prior.aca or prior.conversion or current.aca or current.conversion)
        for search, observed, brokerage in ((prior, PRIOR_REQUEST, 3_984_479.921051042),
                                            (current, CURRENT_REQUEST, 3_753_965.32221208)):
            result = search.solve(exact=False)
            self.assertEqual(result.status, 'Funded')
            self.assertAlmostEqual(result.requested_draw, observed, places=6)
            self.assertAlmostEqual(result.taxable_draw, brokerage, places=6)

        brokerage, traditional, roth, conversion = current._draws(PRIOR_REQUEST, False)
        self.assertAlmostEqual(brokerage, 3_983_821.70898092, places=6)
        self.assertAlmostEqual(traditional, 7_467.55243129, places=6)
        self.assertEqual((roth, conversion), (0, 0))
        # The Mac trace's projection check passes with its cached deposit.
        native_close = (315_681_711.786163 - brokerage + 3.47848981618881e-7) * (1 + CURRENT_RETURN)
        self.assertAlmostEqual(native_close, 382_897_338.323041, delta=.01)

        self.assertAlmostEqual(prior._quote(PRIOR_REQUEST, False)[0], 0, places=6)
        self.assertAlmostEqual(prior._quote(CURRENT_REQUEST, False)[0], -195_532.17576512, places=6)
        self.assertAlmostEqual(current._quote(PRIOR_REQUEST, False)[0], 200_440.86300932, places=6)
        self.assertAlmostEqual(current._quote(CURRENT_REQUEST, False)[0], 0, places=6)
        self.assertAlmostEqual(current._quote(CURRENT_REQUEST - .01, False)[0], -.00872003, places=6)

        def margin(draw: float) -> float:
            """Reprice from trial 7 and keep its no-conversion branch fixed."""
            return current._quote(draw, False)[0]

        witness = earlier_feasible_draw(PRIOR_REQUEST, margin)
        self.assertIsNotNone(witness)
        assert witness is not None
        self.assertAlmostEqual(witness[0], PRIOR_REQUEST - .01, places=8)
        self.assertAlmostEqual(witness[1], 200_440.85428906, places=6)
        self.assertIsNone(earlier_feasible_draw(CURRENT_REQUEST, margin))

    def test_witness_keeps_the_constrained_conversion_policy(self) -> None:
        """An easier quote in a different mode is not a valid selection witness."""
        cash = replace(CURRENT_CASH, requested_conversion=5_000, conversion_allowed=True)
        tax = fixture_tax()
        expected = solve_funding(tax, CURRENT_INCOME, cash)
        search = _FundingSearch(tax, CURRENT_INCOME, cash)
        self.assertEqual(expected.status, 'Funded')
        self.assertEqual(expected.conversion, 5_000)
        self.assertIsNone(earlier_feasible_draw(expected.requested_draw, lambda draw: search._quote(draw, True)[0]))
        self.assertIsNotNone(earlier_feasible_draw(expected.requested_draw, lambda draw: search._quote(draw, False)[0]))

    def test_witness_preserves_zero_near_cent_cliff_and_unfunded_cases(self) -> None:
        """Absence of a witness does not resolve a financially different case."""
        tax = fixture_tax()
        zero = _FundingSearch(tax, CURRENT_INCOME, replace(CURRENT_CASH, cash_income=10_000_000))
        funded = zero.solve(exact=False)
        self.assertEqual((funded.status, funded.requested_draw), ('Funded', 0))
        self.assertEqual(funded.traditional_draw, CURRENT_CASH.mandatory)
        self.assertGreater(funded.margin, 0)
        self.assertIsNone(earlier_feasible_draw(funded.requested_draw, lambda draw: zero._quote(draw, False)[0]))
        self.assertIsNone(earlier_feasible_draw(.004, lambda _: 250))
        self.assertIsNone(earlier_feasible_draw(10, lambda _: .004))
        self.assertIsNone(earlier_feasible_draw(10, lambda _: float('nan')))
        self.assertIsNone(earlier_feasible_draw(100, lambda draw: 20 if draw >= 100 else -3))
        search = _FundingSearch(tax, CURRENT_INCOME, replace(CURRENT_CASH, required_spending=1e9))
        result = search.solve(exact=False)
        self.assertEqual(result.status, 'Unfunded')
        self.assertIsNone(earlier_feasible_draw(result.requested_draw, lambda draw: search._quote(draw, False)[0]))
