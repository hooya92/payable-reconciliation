import unittest
from decimal import Decimal
from domain.models import DataQuality
from domain.normalization import normalize_text, parse_amount
from domain.period import AccountingPeriod, prior_period_for


class PeriodTests(unittest.TestCase):
    def test_august_uses_july(self):
        self.assertEqual((prior_period_for(2026, 8).year, prior_period_for(2026, 8).month), (2026, 7))

    def test_december_to_january_rollover(self):
        self.assertEqual(AccountingPeriod(2027,1).previous(), AccountingPeriod(2026,12))

    def test_period_contains_only_selected_month(self):
        p = AccountingPeriod(2026, 8)
        self.assertTrue(p.contains("2026-08-31"))
        self.assertFalse(p.contains("2026-07-31"))
        self.assertFalse(p.contains("2026-09-01"))

    def test_december_rolls_forward_year(self):
        self.assertEqual(AccountingPeriod(2026,12).next(),AccountingPeriod(2027,1))

    def test_january_rolls_back_year(self):
        self.assertEqual((prior_period_for(2027, 1).year, prior_period_for(2027, 1).month), (2026, 12))


class AmountTests(unittest.TestCase):
    def test_accounting_sign_spacing_is_not_an_amount_error(self):
        for value in ("-500,000","-    500,000","-\u00a0\u00a0500,000","− 500,000"):
            with self.subTest(value=value):
                result=parse_amount(value)
                self.assertEqual(result.value,Decimal("-500000"))
                self.assertEqual(result.quality,DataQuality.CLEAN)

    def test_spaces_inside_digits_are_not_silently_joined(self):
        for value in ("5 00,000","500,000 700,000","- 5 00,000"):
            self.assertEqual(parse_amount(value).quality,DataQuality.SUSPICIOUS)

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


class TextNormalizationTests(unittest.TestCase):
    def test_whitespace_only_differences_are_safe_to_normalize(self):
        self.assertEqual(normalize_text("  유류비   주류  "), normalize_text("유류비 주류"))

    def test_punctuation_is_preserved(self):
        self.assertNotEqual(normalize_text("유류비/주류"), normalize_text("유류비 주류"))


if __name__ == "__main__":
    unittest.main()
