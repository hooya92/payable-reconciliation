import unittest
from decimal import Decimal

from domain.models import JournalLine, PayableItem, Status
from domain.reconciliation import reconcile


def payable(code, desc, amount, name="거래처"):
    return PayableItem(code, name, desc, Decimal(str(amount)))


def debit(code, desc, amount, name="거래처"):
    return JournalLine(code, name, "25301", "미지급금-일반", desc, Decimal(str(amount)), Decimal("0"))


class AdversarialReconciliationTests(unittest.TestCase):
    """Fail-closed cases: uncertainty must never be promoted to an automatic match."""

    def test_unique_same_description_partial_payment_is_review_only(self):
        result = reconcile([payable("001001", "운송비", 100)], [debit("001001", "운송비", 40)])
        self.assertEqual(result[0].status, Status.PARTIAL)
        self.assertEqual(result[0].rule, "POSSIBLE_PARTIAL")

    def test_larger_same_vendor_debit_requires_combined_payment_review(self):
        result = reconcile([payable("001001", "운송비", 100)], [debit("001001", "통합지급", 150)])
        self.assertEqual(result[0].status, Status.AMBIGUOUS)
        self.assertEqual(result[0].rule, "POSSIBLE_COMBINED")

    def test_duplicate_debits_never_choose_one_arbitrarily(self):
        result = reconcile(
            [payable("001001", "운송비", 100)],
            [debit("001001", "운송비", 100), debit("001001", "운송비", 100)],
        )
        self.assertEqual(result[0].status, Status.AMBIGUOUS)

    def test_split_debits_are_not_summed_automatically(self):
        result = reconcile(
            [payable("001001", "운송비", 100)],
            [debit("001001", "운송비", 30), debit("001001", "운송비", 70)],
        )
        self.assertNotEqual(result[0].status, Status.MATCHED)

    def test_same_amount_at_other_vendor_is_not_a_match(self):
        result = reconcile([payable("001001", "운송비", 100)], [debit("009999", "운송비", 100)])
        self.assertEqual(result[0].status, Status.VENDOR_MISMATCH)

    def test_unique_code_amount_with_changed_description_passes_with_note(self):
        result = reconcile([payable("001001", "운송비", 100)], [debit("001001", "보관비", 100)])
        self.assertEqual(result[0].status, Status.MATCHED)
        self.assertEqual(result[0].rule, "CODE_AMOUNT_UNIQUE_WITH_NOTE")
        self.assertIn("적요 차이",result[0].reason)

    def test_vendor_name_change_passes_with_note_when_code_amount_are_unique(self):
        result = reconcile(
            [payable("001001", "운송비", 100, "이전상호")],
            [debit("001001", "운송비", 100, "변경상호")],
        )
        self.assertEqual(result[0].status, Status.MATCHED)
        self.assertIn("거래처명 차이",result[0].reason)

    def test_punctuation_difference_is_preserved_as_a_reference_note(self):
        result = reconcile(
            [payable("001001", "유류비/주류", 100)],
            [debit("001001", "유류비 주류", 100)],
        )
        self.assertEqual(result[0].status, Status.MATCHED)
        self.assertIn("적요 차이",result[0].reason)

    def test_long_carryover_date_is_never_used_as_match_key(self):
        old = PayableItem("001001", "거래처", "운송비", Decimal("100"), "2024-01-31")
        current = JournalLine("001001", "거래처", "25301", "미지급금-일반", "운송비", Decimal("100"), Decimal("0"), "2026-08-31")
        self.assertEqual(reconcile([old], [current])[0].status, Status.MATCHED)


if __name__ == "__main__":
    unittest.main()
