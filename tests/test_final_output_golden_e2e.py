import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from adapters.excel.writer import write_result
from application.service import run_reconciliation
from domain.period import AccountingPeriod
from tests.synthetic_adversarial_factory import build


class FinalOutputGoldenE2ETests(unittest.TestCase):
    """Golden test for the complete path: input Excel -> service -> final result Excel."""

    def test_adversarial_inputs_produce_exact_conservative_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            data=build(tmp)
            period=AccountingPeriod(2026,8)

            run=run_reconciliation(
                [data["prior_path"]],
                [data["raw_path"]],
                {"25301"},
                period,
            )

            out=Path(tmp)/"명세서_대사결과.xlsx"
            write_result(
                out,
                run.results,
                run.new_items,
                run.issues,
                [data["prior_path"],data["raw_path"]],
                period.label,
            )

            wb=load_workbook(out,data_only=True)
            self.assertEqual(
                wb.sheetnames,
                ["요약","확인필요","입력데이터확인","차월명세서 초안","당월신규명세","자동대사완료"],
            )

            # Summary is a fixed golden answer for this adversarial fixture.
            summary=dict(wb["요약"].iter_rows(min_row=1,max_col=2,values_only=True))
            self.assertEqual(summary["대상 회계월"],"2026년 8월")
            self.assertEqual(summary["확인 필요"],6)
            self.assertEqual(summary["입력 형식 확인"],3)
            self.assertEqual(summary["차월 초안"],1)
            self.assertEqual(summary["차월 초안 금액"],660000)
            self.assertEqual(summary["당월 신규 명세 검토"],1)
            self.assertEqual(summary["초안 제외 검토건"],6)

            # Only the exact match may appear as automatically completed.
            completed=list(wb["자동대사완료"].iter_rows(min_row=2,values_only=True))
            self.assertEqual(len(completed),1)
            self.assertEqual(completed[0][1],"001001")
            self.assertEqual(completed[0][4],110000)

            # Every non-exact prior item remains visible for human review.
            review=list(wb["확인필요"].iter_rows(min_row=2,values_only=True))
            self.assertEqual(len(review),6)
            review_codes={row[6] for row in review}
            self.assertEqual(review_codes,{"001002","001003","001004","001005","001006","001007"})

            # Dirty source rows are quarantined and never disappear silently.
            input_issues=list(wb["입력데이터확인"].iter_rows(min_row=2,values_only=True))
            self.assertEqual(len(input_issues),3)
            self.assertEqual(sum(row[0]=="전월명세" for row in input_issues),1)
            self.assertEqual(sum(row[0]=="더존" for row in input_issues),2)

            # Draft contains only the definitely unpaid carry-forward.
            draft_rows=list(wb["차월명세서 초안"].iter_rows(min_row=2,values_only=True))
            detail=[row for row in draft_rows if row[0]=="전월이월"]
            self.assertEqual(len(detail),1)
            self.assertEqual(detail[0][1],"001006")
            self.assertEqual(detail[0][4],"장기 이월 임차료")
            self.assertEqual(detail[0][5],660000)
            self.assertFalse(any(row[0]=="당월신규" for row in draft_rows))
            self.assertFalse(any(row[1]=="002001" for row in draft_rows))

            subtotal=[row for row in draft_rows if row[0]=="소계"]
            total=[row for row in draft_rows if row[0]=="전체합계"]
            self.assertEqual(len(subtotal),1)
            self.assertEqual(subtotal[0][5],660000)
            self.assertEqual(len(total),1)
            self.assertEqual(total[0][5],660000)

            # Current-month new item is visible but explicitly withheld from the draft.
            current=list(wb["당월신규명세"].iter_rows(min_row=2,values_only=True))
            self.assertEqual(len(current),1)
            self.assertEqual(current[0][0],"사람 확인 필요")
            self.assertEqual(current[0][3],"002001")
            self.assertEqual(current[0][5],"8월 신규 유류비")
            self.assertEqual(current[0][6],880000)
            self.assertIn("차월 초안 자동포함 안 함",current[0][7])

            wb.close()


if __name__=="__main__":
    unittest.main()
