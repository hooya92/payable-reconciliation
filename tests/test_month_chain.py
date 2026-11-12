"""The statement we generate becomes next month's prior statement: our wording must never disturb it."""
import re
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from adapters.excel.month_end_statement import write_month_end_statement
from application.service import run_reconciliation
from domain.offset_hints import differs_only_by_month, month_token
from domain.period import AccountingPeriod
from tests import test_journal_corrections as fixtures
from tests.test_offset_notes import generate, raw_file

PERIOD=fixtures.PERIOD

_LINK=re.compile(r"((?:\d+(?:~\d+)?)(?:·\d+(?:~\d+)?)*)행과 대차")


def next_month(prior, folder, raw_rows, year, month):
    folder=Path(folder)
    raw=folder/f"raw_{year}{month:02d}.xlsx"; raw_file(raw,raw_rows)
    period=AccountingPeriod(year,month)
    run=run_reconciliation([prior],[raw],{"25301"},period)
    out=folder/f"out_{year}{month:02d}.xlsx"
    write_month_end_statement(out,[prior],run.results,run.new_items,run.issues,period,
        standalone_debits=run.standalone_debits,journal_items=run.journal_items,
        journal_sources=run.journal_sources,source_paths=[prior,raw])
    return out


def keep_statement_only(path, drop_status_column=False, blank_status_header=False):
    """What a person may do before next month: delete the review sheet (required), helper sheets
    and, optionally, the status column or just its header."""
    wb=load_workbook(path)
    ws=wb[wb.sheetnames[0]]
    for name in list(wb.sheetnames[1:]):
        del wb[name]
    if drop_status_column:
        ws.delete_cols(6)
    if blank_status_header:
        ws.cell(2,6).value=None
    out=Path(path).with_name("for_next_"+Path(path).name)
    wb.save(out); wb.close()
    return out


def body(ws):
    return [(ws.cell(r,3).value,ws.cell(r,4).value,ws.cell(r,5).value)
            for r in range(3,ws.max_row+1) if ws.cell(r,3).value not in (None,"소계")
            and ws.cell(r,1).value!="합 계" and ws.cell(r,4).value is not None]


def assert_links_point_at_real_partners(testcase, ws):
    """Every "N행과 대차" must name rows that really offset this row."""
    for r in range(3,ws.max_row+1):
        text=str(ws.cell(r,6).value or "")
        match=_LINK.search(text)
        if not match:
            continue
        rows=[]
        for part in match.group(1).split("·"):
            low,_,high=part.partition("~")
            rows+=range(int(low),int(high or low)+1)
        testcase.assertTrue(rows,(r,text))
        testcase.assertNotIn(r,rows,(r,text))
        if "후보" not in text:
            # a one-to-one pair: one partner, equal and opposite amount, same item (month ignored)
            testcase.assertEqual(len(rows),1,(r,text))
            partner=rows[0]
            testcase.assertEqual(ws.cell(r,5).value,-ws.cell(partner,5).value,(r,text))
            mine,theirs=ws.cell(r,4).value,ws.cell(partner,4).value
            testcase.assertTrue(mine==theirs or differs_only_by_month(mine,theirs),(r,text,mine,theirs))
            testcase.assertEqual("(월만 다름)" in text,mine!=theirs,(r,text,mine,theirs))
        else:
            head=text.splitlines()[0]
            month=int(re.search(r"(\d+)월 묶음",head).group(1))
            for partner in rows:
                testcase.assertEqual(month_token(ws.cell(partner,4).value),month,(r,partner,text))


class MonthChainTests(unittest.TestCase):
    OCT=[("001111","가상물류","9월 청구분",50,0,"2026-10-05")]

    def test_next_month_reads_the_generated_statement_unchanged(self):
        with tempfile.TemporaryDirectory() as folder:
            run,sept=generate(folder)
            sept_ws=load_workbook(sept)["26.09"]
            expected=body(sept_ws)
            nxt=keep_statement_only(sept)
            octo=next_month(nxt,folder,self.OCT,2026,10)
            ws=load_workbook(octo)["26.10"]
            # every row of last month's statement is carried exactly (description + amount)
            self.assertEqual(body(ws)[:len(expected)],expected)
            assert_links_point_at_real_partners(self,ws)

    def test_last_months_wording_is_never_carried_into_the_next_month(self):
        with tempfile.TemporaryDirectory() as folder:
            run,sept=generate(folder)
            octo=next_month(keep_statement_only(sept),folder,self.OCT,2026,10)
            ws=load_workbook(octo)["26.10"]
            statuses="\n".join(str(ws.cell(r,6).value or "") for r in range(3,ws.max_row+1))
            # October has no 7월분/8월분 payments, so no 7월/8월 group may appear
            self.assertNotIn("7월 묶음",statuses)
            self.assertNotIn("8월 묶음",statuses)
            # September's subtotal text (prior 590) is replaced by October's own arithmetic
            # October's own arithmetic (the 9월 claim and its payment are now an exact pair,
            # so the subtotal is the short one-line form); never September's 590
            self.assertIn("전월 250 − 차변 50 = 200",statuses)
            self.assertIn("순잔액 대사 · 10행과 대차",statuses)
            self.assertNotIn("590",statuses)
            self.assertEqual(len(ws._hyperlinks),0)

    def test_still_works_when_the_status_column_or_its_header_was_removed(self):
        for label,kwargs in (("status column deleted",{"drop_status_column":True}),
                             ("status header blanked",{"blank_status_header":True})):
            with self.subTest(label), tempfile.TemporaryDirectory() as folder:
                run,sept=generate(folder)
                expected=body(load_workbook(sept)["26.09"])
                octo=next_month(keep_statement_only(sept,**kwargs),folder,self.OCT,2026,10)
                ws=load_workbook(octo)["26.10"]
                self.assertEqual(body(ws)[:len(expected)],expected)
                texts=[str(c.value) for row in ws.iter_rows(min_row=3) for c in row[5:] if c.value is not None]
                # nothing of September's own wording survives in any column
                self.assertFalse([t for t in texts if "7월 묶음" in t or "590" in t],texts)

    def test_three_months_in_a_row(self):
        with tempfile.TemporaryDirectory() as folder:
            run,current=generate(folder)
            sheet="26.09"
            rows=[
                [("001111","가상물류","9월 청구분",50,0,"2026-10-05")],
                [("001111","가상물류","10월 청구분",0,70,"2026-11-03")],
            ]
            for index,raw_rows in enumerate(rows):
                expected=body(load_workbook(current)[sheet])
                year,month=(2026,10) if index==0 else (2026,11)
                current=next_month(keep_statement_only(current),folder,raw_rows,year,month)
                sheet=f"26.{month:02d}"
                ws=load_workbook(current)[sheet]
                self.assertEqual(body(ws)[:len(expected)],expected)
                assert_links_point_at_real_partners(self,ws)

    def test_confirmed_pairs_name_their_partner_and_survive_the_next_month(self):
        # 벨트 교체 700,000 is paid by an equal debit: a confirmed one-to-one pair
        with tempfile.TemporaryDirectory() as folder:
            prior,raw=fixtures.JournalCorrectionTests().make_files(folder,[
                ("001111","가상공영","벨트 교체",700000,0,"2026-09-03")])
            run=run_reconciliation([prior],[raw],{"25301"},PERIOD)
            sept=Path(folder)/"sept.xlsx"
            write_month_end_statement(sept,[prior],run.results,run.new_items,run.issues,PERIOD,
                standalone_debits=run.standalone_debits,journal_items=run.journal_items,
                journal_sources=run.journal_sources,source_paths=[prior,raw])
            ws=load_workbook(sept)["26.09"]
            # rows: 3 벨트(+700,000)  4 도어(+200,000)  5 지급(-700,000)  6 subtotal
            self.assertEqual(ws["F3"].value,"자동 대사 · 5행과 대차")
            self.assertEqual(ws["F5"].value,"자동 대사 · 3행과 대차")
            self.assertEqual(ws["F4"].value,"전월 이월")
            assert_links_point_at_real_partners(self,ws)
            expected=body(ws)
            # next month: the pair rows are still in the statement (the person removes them
            # by hand); the new sheet recomputes every status and never reuses the old ones
            octo=next_month(keep_statement_only(sept),folder,
                            [("000002","정상 업체","정상 비용",100,0,"2026-10-05")],2026,10)
            ws10=load_workbook(octo)["26.10"]
            self.assertEqual(body(ws10)[:len(expected)],expected)
            assert_links_point_at_real_partners(self,ws10)
            self.assertEqual(len(ws10._hyperlinks),0)


if __name__=="__main__":
    unittest.main()
