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

    def test_possible_partial_payment_requires_review(self):
        result = reconcile([payable("001001", "운송비", 100)], [debit("001001", "운송비", 40)])
        self.assertEqual(result[0].status, Status.AMBIGUOUS)
        self.assertEqual(result[0].rule, "POSSIBLE_PARTIAL")

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

    def test_same_vendor_amount_but_changed_description_requires_review(self):
        result = reconcile([payable("001001", "운송비", 100)], [debit("001001", "보관비", 100)])
        self.assertEqual(result[0].status, Status.DESCRIPTION_MISMATCH)

    def test_vendor_name_change_does_not_override_code_identity(self):
        result = reconcile(
            [payable("001001", "운송비", 100, "이전상호")],
            [debit("001001", "운송비", 100, "변경상호")],
        )
        self.assertEqual(result[0].status, Status.MATCHED)

    def test_long_carryover_date_is_never_used_as_match_key(self):
        old = PayableItem("001001", "거래처", "운송비", Decimal("100"), "2024-01-31")
        current = JournalLine("001001", "거래처", "25301", "미지급금-일반", "운송비", Decimal("100"), Decimal("0"), "2026-08-31")
        self.assertEqual(reconcile([old], [current])[0].status, Status.MATCHED)


if __name__ == "__main__":
    unittest.main()
