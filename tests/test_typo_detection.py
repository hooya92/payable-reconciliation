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

    def test_vendor_name_one_character_typo_is_review_only(self):
        result=reconcile(
            [prior("051330","가상로지스","9월 유류비",860000)],
            [line("051330","가상로지스스","9월 유류비",860000)],
        )[0]
        self.assertEqual(result.status,Status.INPUT_TYPO_SUSPECT)
        self.assertEqual(result.rule,"VENDOR_NAME_TYPO")

    def test_description_one_character_typo_is_review_only(self):
        result=reconcile(
            [prior("051330","가상물류","9월 유류비",860000)],
            [line("051330","가상물류","9월 유류바",860000)],
        )[0]
        self.assertEqual(result.status,Status.INPUT_TYPO_SUSPECT)
        self.assertEqual(result.rule,"DESCRIPTION_TYPO")

    def test_amount_missing_digit_is_review_only(self):
        result=reconcile(
            [prior("051330","가상물류","9월 유류비",860000)],
            [line("051330","가상물류","9월 유류비",86000)],
        )[0]
        self.assertEqual(result.status,Status.INPUT_TYPO_SUSPECT)
        self.assertEqual(result.rule,"AMOUNT_TYPO")

    def test_large_description_difference_remains_normal_mismatch(self):
        result=reconcile(
            [prior("051330","가상물류","9월 유류비",860000)],
            [line("051330","가상물류","9월 창고 보관료",860000)],
        )[0]
        self.assertEqual(result.status,Status.DESCRIPTION_MISMATCH)

    def test_typo_suspicions_are_never_auto_matched(self):
        cases=[
            reconcile([prior("051330","가상물류","9월 유류비",860000)],[line("051380","가상물류","9월 유류비",860000)])[0],
            reconcile([prior("051330","가상물류","9월 유류비",860000)],[line("051330","가상물류","9월 유류바",860000)])[0],
            reconcile([prior("051330","가상물류","9월 유류비",860000)],[line("051330","가상물류","9월 유류비",86000)])[0],
        ]
        self.assertTrue(all(x.status != Status.MATCHED for x in cases))


if __name__=="__main__":
    unittest.main()
