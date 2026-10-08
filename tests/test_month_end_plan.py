"""Pure month-end plan tests: no Excel files or UI required."""
import unittest
from decimal import Decimal

from application.month_end_plan import build_month_end_plan
from domain.models import JournalLine, NewPayable, PayableItem, ReconcileResult, SourceRef, Status


def opening(amount, row=4):
    return PayableItem("001111","가상물류","전월 운송비",Decimal(amount),"2026-07-31",row,
                       SourceRef("전월.xlsx","26.07",row,"담당"))


def journal(code,desc,debit=0,credit=0,row=5):
    return JournalLine(code,"가상물류" if code=="001111" else "신규업체",
                       "25301","미지급금-일반",desc,
                       Decimal(debit),Decimal(credit),"2026-08-10",row)


class MonthlyPlannerTests(unittest.TestCase):
    def test_original_rows_and_raw_debit_credit_are_all_traceable_once(self):
        p=opening("100")
        j=journal("001111","전월 운송비",debit=100)
        c=journal("001111","당월 운송비",credit=75,row=6)
        r=ReconcileResult(Status.MATCHED,p,j,"정확히 일치","EXACT")
        fresh=NewPayable(c,Decimal(75),"당월 발생",True)
        plan=build_month_end_plan(
            [("001111","가상물류",p.description,p.amount,p.date,4,{})],
            [r],[fresh],[],"전월.xlsx","26.07",journal_items=[j,c,j]
        )
        self.assertEqual([x[3] for x in plan.records],[Decimal(100),Decimal(-100),Decimal(75)])
        self.assertEqual(plan.raw_count,2)
        self.assertEqual(len(plan.offsets),1)
        self.assertEqual(plan.reviews,[])
        self.assertEqual([r[5] for r in plan.records],["대사 일치","대사 일치","당월 발생"])

    def test_unverified_offset_is_not_claimed_matched(self):
        p=opening("100")
        wrong=journal("001111","전월 운송비",debit=90)
        unverified=journal("002222","검토 대변",credit=45,row=7)
        result=ReconcileResult(Status.PARTIAL,p,wrong,"부분지급 가능성 확인 필요","POSSIBLE_PARTIAL")
        new=NewPayable(unverified,Decimal(45),"대변 근거 확인 필요",False)
        plan=build_month_end_plan(
            [("001111","가상물류",p.description,p.amount,p.date,4,{})],
            [result],[new],[],"전월.xlsx","26.07",journal_items=[wrong,unverified]
        )
        self.assertEqual([r[3] for r in plan.records],[Decimal(100),Decimal(-90),Decimal(45)])
        self.assertEqual(len(plan.offsets),0)
        self.assertEqual(plan.raw_count,2)
        self.assertEqual([r[5] for r in plan.records],["확인 필요","확인 필요","확인 필요"])
        self.assertEqual(sum(r[3] for r in plan.records),Decimal(55))
        self.assertEqual(len(plan.reviews),3)


if __name__=="__main__":
    unittest.main()
