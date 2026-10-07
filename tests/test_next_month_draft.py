import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from openpyxl import load_workbook
from adapters.excel.writer import write_result
from domain.models import JournalLine, PayableItem, ReconcileResult, SourceRef, Status

class NextMonthDraftTests(unittest.TestCase):
    def test_draft_contains_only_definite_carryover_and_new_credits(self):
        src=SourceRef("담당자A.xlsx","명세",7,"담당자A")
        unpaid=PayableItem("001001","가상A","7월 운송비",Decimal("1000"),"2026-07-31",7,src)
        matched=PayableItem("001002","가상B","7월 보관비",Decimal("2000"),"2026-07-31",8,src)
        review=PayableItem("001003","가상C","7월 용역비",Decimal("3000"),"2026-07-31",9,src)
        j=JournalLine("001002","가상B","25301","미지급금-일반","7월 보관비",Decimal("2000"),Decimal("0"),"2026-08-10",20)
        new=JournalLine("001004","가상D","25301","미지급금-일반","8월 신규비",Decimal("0"),Decimal("4000"),"2026-08-31",30)
        results=[
            ReconcileResult(Status.UNPAID,unpaid,None,"대응 차변 없음","NO_DEBIT"),
            ReconcileResult(Status.MATCHED,matched,j,"일치","EXACT"),
            ReconcileResult(Status.DESCRIPTION_MISMATCH,review,j,"적요 불일치","VENDOR_AMOUNT"),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            out=Path(tmp)/"result.xlsx"
            write_result(out,results,[new],[],[], "2026년 8월")
            wb=load_workbook(out,data_only=True); ws=wb["차월명세서 초안"]
            rows=list(ws.iter_rows(min_row=2,values_only=True))
            detail=[r for r in rows if r[0] in ("전월이월","당월신규")]
            subtotals=[r for r in rows if r[0]=="소계"]
            totals=[r for r in rows if r[0]=="전체합계"]
            self.assertEqual(len(detail),2)
            self.assertEqual({r[0] for r in detail},{"전월이월","당월신규"})
            self.assertEqual({r[5] for r in detail},{1000,4000})
            self.assertNotIn(3000,{r[5] for r in detail})
            self.assertEqual(len(subtotals),2)
            self.assertEqual(len(totals),1)
            self.assertEqual(totals[0][5],5000)
            wb.close()

if __name__=="__main__": unittest.main()
