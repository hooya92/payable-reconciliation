import unittest
from decimal import Decimal

from domain.models import JournalLine, PayableItem, Status
from domain.reconciliation import reconcile


def prior(code, name, desc, amount):
    return PayableItem(code,name,desc,Decimal(str(amount)))


def line(code, name, desc, debit):
    return JournalLine(code,name,"25301","미지급금-일반",desc,Decimal(str(debit)),Decimal("0"))


class TypoDetectionTests(unittest.TestCase):
    def test_vendor_code_one_digit_typo_is_review_only(self):
        result=reconcile(
            [prior("051330","가상물류","9월 유류비",860000)],
            [line("051380","가상물류","9월 유류비",860000)],
        )[0]
        self.assertEqual(result.status,Status.INPUT_TYPO_SUSPECT)
        self.assertEqual(result.rule,"VENDOR_CODE_TYPO")

    def test_vendor_name_one_character_typo_passes_with_note(self):
        result=reconcile(
            [prior("051330","가상로지스","9월 유류비",860000)],
            [line("051330","가상로지스스","9월 유류비",860000)],
        )[0]
        self.assertEqual(result.status,Status.MATCHED)
        self.assertEqual(result.rule,"CODE_AMOUNT_UNIQUE_WITH_NOTE")
        self.assertIn("거래처명 오타 의심",result.reason)

    def test_description_one_character_typo_passes_with_note(self):
        result=reconcile(
            [prior("051330","가상물류","9월 유류비",860000)],
            [line("051330","가상물류","9월 유류바",860000)],
        )[0]
        self.assertEqual(result.status,Status.MATCHED)
        self.assertEqual(result.rule,"CODE_AMOUNT_UNIQUE_WITH_NOTE")
        self.assertIn("적요 오타 의심",result.reason)

    def test_amount_missing_digit_is_review_only(self):
        result=reconcile(
            [prior("051330","가상물류","9월 유류비",860000)],
            [line("051330","가상물류","9월 유류비",86000)],
        )[0]
        self.assertEqual(result.status,Status.INPUT_TYPO_SUSPECT)
        self.assertEqual(result.rule,"AMOUNT_TYPO")

    def test_large_description_difference_passes_only_when_code_and_amount_are_unique(self):
        result=reconcile(
            [prior("051330","가상물류","9월 유류비",860000)],
            [line("051330","가상물류","9월 창고 보관료",860000)],
        )[0]
        self.assertEqual(result.status,Status.MATCHED)
        self.assertEqual(result.rule,"CODE_AMOUNT_UNIQUE_WITH_NOTE")
        self.assertIn("적요 차이",result.reason)

    def test_code_or_amount_typo_suspicions_are_never_auto_matched(self):
        cases=[
            reconcile([prior("051330","가상물류","9월 유류비",860000)],[line("051380","가상물류","9월 유류비",860000)])[0],
            reconcile([prior("051330","가상물류","9월 유류비",860000)],[line("051330","가상물류","9월 유류비",86000)])[0],
        ]
        self.assertTrue(all(x.status != Status.MATCHED for x in cases))


if __name__=="__main__":
    unittest.main()
