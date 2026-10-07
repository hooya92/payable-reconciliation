import unittest
from decimal import Decimal
from domain.models import DataQuality
from domain.normalization import parse_amount
from domain.period import prior_period_for


class PeriodTests(unittest.TestCase):
    def test_august_uses_july(self):
        self.assertEqual((prior_period_for(2026, 8).year, prior_period_for(2026, 8).month), (2026, 7))

    def test_january_rolls_back_year(self):
        self.assertEqual((prior_period_for(2027, 1).year, prior_period_for(2027, 1).month), (2026, 12))


class AmountTests(unittest.TestCase):
    def test_commas(self):
        x = parse_amount("86,000,000")
        self.assertEqual(x.value, Decimal("86000000"))
        self.assertEqual(x.quality, DataQuality.CLEAN)

    def test_won_suffix(self):
        x = parse_amount("86,000,000원")
        self.assertEqual(x.value, Decimal("86000000"))
        self.assertEqual(x.quality, DataQuality.NORMALIZED)

    def test_mixed_separator_is_suspicious(self):
        self.assertEqual(parse_amount("86.000,000원").quality, DataQuality.SUSPICIOUS)

    def test_manwon_is_suspicious_even_if_guessable(self):
        x = parse_amount("8600만원")
        self.assertEqual(x.value, Decimal("86000000"))
        self.assertEqual(x.quality, DataQuality.SUSPICIOUS)


if __name__ == "__main__":
    unittest.main()
