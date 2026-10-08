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

    def test_multiline_description_is_normalized_as_whitespace(self):
        r = reconcile(
            [p("051330", "7월 용차료\n7월 시스템비", 100)],
            [j("051330", "7월 용차료 7월 시스템비", 100)],
        )
        self.assertEqual(r[0].status, Status.MATCHED)
        self.assertEqual(r[0].rule, "EXACT")

    def test_unique_same_vendor_code_and_amount_passes_with_description_note(self):
        r = reconcile([p("051330", "주류", 100)], [j("051330", "일반", 100)])
        self.assertEqual(r[0].status, Status.MATCHED)
        self.assertEqual(r[0].rule, "CODE_AMOUNT_UNIQUE_WITH_NOTE")
        self.assertIn("적요 차이", r[0].reason)

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
        prior=PayableItem("051330","거래처","same",Decimal("100"),"2020-01-01")
        line=JournalLine("051330","거래처","25301","미지급금-일반","same",Decimal("100"),Decimal("0"),"2026-08-31")
        self.assertEqual(reconcile([prior],[line])[0].status, Status.MATCHED)

    def test_unique_partial_payment_is_carried(self):
        r=reconcile([p("051330","A",100)],[j("051330","A",40)])
        self.assertEqual(r[0].status, Status.PARTIAL)
        self.assertEqual(r[0].rule, "PARTIAL_UNIQUE")

    def test_partial_with_duplicate_opening_descriptions_requires_review(self):
        r=reconcile([p("051330","A",100),p("051330","A",200)],[j("051330","A",40)])
        self.assertTrue(all(x.status == Status.AMBIGUOUS for x in r))

    def test_same_amount_debit_not_stolen_from_other_partial_payment(self):
        r=reconcile([p("051330","B",40),p("051330","A",100)],[j("051330","A",40)])
        self.assertTrue(all(x.status == Status.AMBIGUOUS for x in r))

    def test_duplicate_prior_with_insufficient_debits_is_ambiguous(self):
        r=reconcile([p("051330","A",100),p("051330","A",100)],[j("051330","A",100)])
        self.assertTrue(all(x.status == Status.AMBIGUOUS for x in r))

    def test_duplicate_candidates_are_ambiguous(self):
        r = reconcile([p("051330", "주류", 100)], [j("051330", "주류", 100), j("051330", "주류", 100)])
        self.assertEqual(r[0].status, Status.AMBIGUOUS)


if __name__ == "__main__":
    unittest.main()
