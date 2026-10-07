import unittest
from decimal import Decimal

from domain.models import JournalLine, PayableItem, Status
from domain.reconciliation import reconcile


def p(code, desc, amount):
    return PayableItem(code, "거래처", desc, Decimal(str(amount)))


def j(code, desc, debit):
    return JournalLine(code, "거래처", "25301", "미지급금-일반", desc, Decimal(str(debit)), Decimal("0"))


class ReconcileTests(unittest.TestCase):
    def test_exact_match(self):
        r = reconcile([p("051330", "7월 FMC사업부(주류) 용차료/유류비", 100)],
                      [j("051330", "7월 FMC사업부(주류) 용차료/유류비", 100)])
        self.assertEqual(r[0].status, Status.MATCHED)

    def test_same_amount_different_description_is_exception(self):
        r = reconcile([p("051330", "주류", 100)], [j("051330", "일반", 100)])
        self.assertEqual(r[0].status, Status.DESCRIPTION_MISMATCH)

    def test_same_amount_other_vendor_is_exception(self):
        r = reconcile([p("051330", "주류", 100)], [j("999999", "주류", 100)])
        self.assertEqual(r[0].status, Status.VENDOR_MISMATCH)

    def test_no_debit_is_unpaid(self):
        r = reconcile([p("051330", "주류", 100)], [])
        self.assertEqual(r[0].status, Status.UNPAID)

    def test_description_mismatch_candidate_is_not_reused(self):
        r = reconcile([p("051330", "A", 100), p("051330", "B", 100)], [j("051330", "X", 100)])
        self.assertTrue(all(x.status == Status.AMBIGUOUS for x in r))

    def test_date_is_never_a_match_key(self):
        prior=p("051330","same",100); prior.date="2020-01-01"
        line=j("051330","same",100); line.date="2026-08-31"
        self.assertEqual(reconcile([prior],[line])[0].status, Status.MATCHED)

    def test_partial_payment_is_not_auto_matched(self):
        r=reconcile([p("051330","A",100)],[j("051330","A",40)])
        self.assertEqual(r[0].status, Status.UNPAID)

    def test_duplicate_prior_with_insufficient_debits_is_ambiguous(self):
        r=reconcile([p("051330","A",100),p("051330","A",100)],[j("051330","A",100)])
        self.assertTrue(all(x.status == Status.AMBIGUOUS for x in r))

    def test_duplicate_candidates_are_ambiguous(self):
        r = reconcile([p("051330", "주류", 100)], [j("051330", "주류", 100), j("051330", "주류", 100)])
        self.assertEqual(r[0].status, Status.AMBIGUOUS)


if __name__ == "__main__":
    unittest.main()
