"""
test_calculations.py — Sanity tests for the costing business rules.

Plain-stdlib unittest so it runs with `python test_calculations.py` (no pytest
dependency). Focuses on the non-obvious parts: the meters-per-kg formula, the
wastage cost-spreading, the reconciliation check, and graceful handling of
zero/blank/invalid input.
"""

import math
import unittest

from calculations import LotInfo, Product, compute, kg_from_meters, parse_number


class ParseNumberTests(unittest.TestCase):
    def test_blank_and_garbage_return_none(self):
        for bad in ["", "   ", None, "abc", "1.2.3"]:
            self.assertIsNone(parse_number(bad))

    def test_valid_numbers(self):
        self.assertEqual(parse_number("1,234.50"), 1234.5)
        self.assertEqual(parse_number("  42 "), 42.0)
        self.assertEqual(parse_number(7), 7.0)


class LotFormulaTests(unittest.TestCase):
    def _lot(self, **kw):
        base = dict(gsm=180, width_in=90, total_kg=100,
                    rate_per_meter=50, transport_cost=1000)
        base.update(kw)
        return LotInfo(**base)

    def test_meters_per_kg_and_chain(self):
        r = compute(self._lot(), [])
        width_m = 90 * 0.0254
        expected_mpk = 1000 / (width_m * 180)
        self.assertAlmostEqual(r.meters_per_kg, expected_mpk, places=6)
        self.assertAlmostEqual(r.total_meters, expected_mpk * 100, places=6)
        self.assertAlmostEqual(r.fabric_cost, expected_mpk * 100 * 50, places=4)
        self.assertAlmostEqual(r.total_cost, r.fabric_cost + 1000, places=4)
        self.assertAlmostEqual(r.base_cost_per_kg, r.total_cost / 100, places=6)

    def test_missing_inputs_do_not_crash(self):
        r = compute(LotInfo(), [])  # everything blank
        self.assertIsNone(r.meters_per_kg)
        self.assertIsNone(r.total_cost)
        self.assertIsNone(r.profit)

    def test_zero_gsm_is_safe(self):
        r = compute(self._lot(gsm=0), [])
        self.assertIsNone(r.meters_per_kg)  # no divide-by-zero explosion


class WastageSpreadingTests(unittest.TestCase):
    def test_wastage_cost_is_spread_into_adjusted_rate(self):
        lot = LotInfo(gsm=180, width_in=90, total_kg=100,
                      rate_per_meter=50, transport_cost=1000)
        products = [
            Product(name="Sheets", weight_kg=90, pieces=45),
            Product(name="Wastage", weight_kg=10, pieces=0, is_wastage=True),
        ]
        r = compute(lot, products)

        self.assertEqual(r.wastage_weight, 10)
        # Waste cost is spread over the SELLABLE weight (100 - 10 = 90):
        #   wastage_cost_per_kg = (10 * base) / 90
        self.assertAlmostEqual(
            r.wastage_cost_per_kg, (10 * r.base_cost_per_kg) / 90, places=6)
        # adjusted = base + wastage spread, and is strictly higher than base.
        self.assertAlmostEqual(
            r.adjusted_cost_per_kg, r.base_cost_per_kg + r.wastage_cost_per_kg, places=6)
        self.assertGreater(r.adjusted_cost_per_kg, r.base_cost_per_kg)

        # Products are priced at the ADJUSTED rate.
        sheets = r.product_results[0]
        self.assertAlmostEqual(sheets.weight_per_piece, 90 / 45, places=6)
        self.assertAlmostEqual(
            sheets.cost_per_piece, (90 / 45) * r.adjusted_cost_per_kg, places=6)
        self.assertAlmostEqual(sheets.revenue, 45 * sheets.cost_per_piece, places=6)

    def test_full_recovery_when_weights_reconcile(self):
        # Cost is spread over the sellable weight, so ALL of the lot's cost is
        # allocated back onto the finished pieces. When the entered weights
        # reconcile, the allocated cost (total_receipt) equals the total cost
        # paid, i.e. profit is ~0 — every rupee is accounted for.
        lot = LotInfo(gsm=180, width_in=90, total_kg=100,
                      rate_per_meter=50, transport_cost=1000)
        products = [
            Product(name="Sheets", weight_kg=90, pieces=45),
            Product(name="Wastage", weight_kg=10, pieces=0, is_wastage=True),
        ]
        r = compute(lot, products)
        self.assertFalse(r.recon_mismatch)
        self.assertAlmostEqual(r.total_receipt, r.total_cost, places=2)
        self.assertAlmostEqual(r.profit, 0.0, places=2)

    def test_no_wastage_row(self):
        lot = LotInfo(gsm=180, width_in=90, total_kg=100,
                      rate_per_meter=50, transport_cost=0)
        r = compute(lot, [Product(name="Sheets", weight_kg=100, pieces=50)])
        self.assertEqual(r.wastage_weight, 0.0)
        self.assertAlmostEqual(r.wastage_cost_per_kg, 0.0, places=6)
        self.assertAlmostEqual(r.adjusted_cost_per_kg, r.base_cost_per_kg, places=6)


class EdgeCaseTests(unittest.TestCase):
    def test_zero_pieces_gives_na_not_crash(self):
        lot = LotInfo(gsm=180, width_in=90, total_kg=100, rate_per_meter=50)
        r = compute(lot, [Product(name="Odd", weight_kg=10, pieces=0)])
        pr = r.product_results[0]
        self.assertIsNone(pr.weight_per_piece)  # -> "N/A" in the UI
        self.assertIsNone(pr.cost_per_piece)
        self.assertIsNone(pr.revenue)

    def test_reconciliation_mismatch_flagged(self):
        lot = LotInfo(gsm=180, width_in=90, total_kg=116.02, rate_per_meter=50)
        products = [
            Product(name="A", weight_kg=100, pieces=50),
            Product(name="Wastage", weight_kg=18, pieces=0, is_wastage=True),
        ]
        r = compute(lot, products)  # sum = 118.00 vs 116.02
        self.assertTrue(r.recon_mismatch)
        self.assertIn("116.02", r.recon_message)
        self.assertIn("118.00", r.recon_message)

    def test_reconciliation_within_tolerance_ok(self):
        lot = LotInfo(gsm=180, width_in=90, total_kg=100.0, rate_per_meter=50)
        r = compute(lot, [Product(name="A", weight_kg=100.005, pieces=50)])
        self.assertFalse(r.recon_mismatch)  # diff 0.005 < 0.01 tolerance


class KgFromMetersTests(unittest.TestCase):
    def test_round_trips_with_compute(self):
        # meters -> kg -> compute() must return the same meters back.
        kg = kg_from_meters(gsm="254", width_in="90", total_meters="128.74")
        r = compute(LotInfo(gsm=254, width_in=90, total_kg=kg), [])
        self.assertAlmostEqual(r.total_meters, 128.74, places=6)
        # And the derived weight matches the user's real TEST-A lot (~74.75).
        self.assertAlmostEqual(kg, 128.74 * (90 * 0.0254) * 254 / 1000, places=6)

    def test_missing_or_zero_inputs_return_none(self):
        for gsm, width, meters in [
            (None, 90, 100), (254, "", 100), (254, 90, None),
            (0, 90, 100), (254, 0, 100), ("abc", 90, 100),
        ]:
            self.assertIsNone(kg_from_meters(gsm, width, meters),
                              msg=f"({gsm}, {width}, {meters})")


if __name__ == "__main__":
    unittest.main(verbosity=2)
