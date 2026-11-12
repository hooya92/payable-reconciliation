import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

from openpyxl import Workbook, load_workbook

from adapters.excel.month_end_statement import write_month_end_statement
from application.service import run_reconciliation
from tests.test_offset_notes import PERIOD, raw_file

from adapters.excel.month_end_status import link_note_text
from domain.offset_hints import (LinkNote, MAX_MATCH_ROWS, offset_sets, significant_digits)


def kinds(rows):
    return sorted((o.kind,tuple(sorted(o.keys))) for o in offset_sets(rows))


class MixedOffsetTests(unittest.TestCase):
    def test_two_claims_against_one_payment_are_a_candidate(self):
        rows=[("A","GS팀 청구",D(100_000_001)),("B","CU팀 청구",D(200_000_003)),
              ("C","합산 지급",D(-300_000_004))]
        self.assertEqual(kinds(rows),[("mixed",("A","B","C"))])

    def test_one_claim_against_two_payments_is_a_candidate(self):
        rows=[("A","청구",D(123_456_789)),("B","지급1",D(-23_456_789)),("C","지급2",D(-100_000_000))]
        self.assertEqual(kinds(rows),[("mixed",("A","B","C"))])

    def test_ambiguous_sums_are_left_alone(self):
        # 300_000_004 can be paid by A+B or by A2+B2: no guessing
        rows=[("A",  "x",D(100_000_001)),("B", "y",D(200_000_003)),
              ("A2", "z",D(100_000_002)),("B2","w",D(200_000_002)),
              ("C",  "p",D(-300_000_004))]
        self.assertEqual(kinds(rows),[])

    def test_round_or_short_amounts_are_not_matched_by_amount_alone(self):
        self.assertEqual(significant_digits(100_000_000),1)
        self.assertEqual(significant_digits(807_567_651),9)
        rows=[("A","x",D(100)),("B","y",D(200)),("C","z",D(-300))]
        self.assertEqual(kinds(rows),[])
        rows=[("A","x",D(1_000_000)),("B","y",D(2_000_000)),("C","z",D(-3_000_000))]
        self.assertEqual(kinds(rows),[])

    def test_too_many_rows_or_too_large_a_set_are_skipped(self):
        many=[(f"R{i}","x",D(1_000_003+i*7919)) for i in range(MAX_MATCH_ROWS)]
        many.append(("P","y",-(many[0][2]+many[1][2])))
        self.assertEqual(kinds(many),[])
        five=[(f"R{i}","x",D(10_000_019+i*104_729)) for i in range(5)]
        five.append(("P","y",-sum(r[2] for r in five)))
        self.assertEqual(kinds(five),[])

    def test_stricter_stages_keep_priority(self):
        rows=[("A","7월 위탁물류비",D(807_567_651)),("B","8월 위탁물류비",D(-807_567_651))]
        self.assertEqual(kinds(rows),[("pair",("A","B"))])

    def test_wording_says_candidate_and_amount_only(self):
        note=LinkNote("순잔액 대사","",(4,5),True,together=True,candidate=True)
        self.assertEqual(link_note_text(note,[4,5]),
                         "순잔액 대사 · 4~5행과 함께 대차 후보 (합계 0 · 금액만 일치)")
        note=LinkNote("","",(4,),True,together=True,append=True,candidate=True)
        self.assertEqual(link_note_text(note,[4]),"4행과 함께 대차 후보 (합계 0 · 금액만 일치)")


class MixedOffsetStatementTests(unittest.TestCase):
    def test_statement_names_the_rows_as_a_candidate_and_keeps_the_engine_verdict(self):
        with tempfile.TemporaryDirectory() as folder:
            folder=Path(folder); prior=folder/"prior.xlsx"; raw=folder/"raw.xlsx"
            wb=Workbook(); ws=wb.active; ws.title="26.08"
            ws.append(["미지급금 명세"]); ws.append(["코드","거래처","날짜","적요","금액"])
            ws.append([None,None,"2026-08-01","GS팀 위탁물류비",100000001])
            ws.append([None,None,"2026-08-01","CU팀 위탁물류비",200000003])
            ws.append(["001111","가상물류","소계",None,"=SUM(E3:E4)"])
            ws.append(["합 계",None,None,None,'=SUMIF(C3:C5,"소계",E3:E5)'])
            wb.save(prior); wb.close()
            raw_file(raw,[("001111","가상물류","물류비 합산 지급",300000004,0,"2026-09-03")])
            run=run_reconciliation([prior],[raw],{"25301"},PERIOD)
            out=folder/"out.xlsx"
            write_month_end_statement(out,[prior],run.results,run.new_items,run.issues,PERIOD,
                standalone_debits=run.standalone_debits,journal_items=run.journal_items,
                journal_sources=run.journal_sources,source_paths=[prior,raw])
            wb=load_workbook(out); ws=wb["26.09"]
            try:
                self.assertEqual([ws.cell(r,5).value for r in (3,4,5)],[100000001,200000003,-300000004])
                for row,partners in ((3,"4~5"),(4,"3·5"),(5,"3~4")):
                    text=ws.cell(row,6).value
                    # the engine's own "확인 필요" line stays; the hint sits under it and is only a candidate
                    self.assertTrue(text.startswith("확인 필요"),text)
                    self.assertTrue(text.endswith(f"{partners}행과 함께 대차 후보 (합계 0 · 금액만 일치)"),text)
                self.assertEqual(ws.max_row,7)          # nothing below the template body
            finally:
                wb.close()


if __name__=="__main__":
    unittest.main()
