"""Rows that are the same item and differ only by the month offset each other."""
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

from openpyxl import Workbook, load_workbook

from adapters.excel.month_end_statement import write_month_end_statement
from application.service import run_reconciliation
from domain.models import Status
from domain.offset_hints import differs_only_by_month, offset_sets
from tests import test_journal_corrections as fixtures
from tests.test_offset_notes import raw_file

PERIOD=fixtures.PERIOD


def statement(path, claims):
    wb=Workbook(); ws=wb.active; ws.title="26.08"
    ws.append(["미지급금 명세"]); ws.append(["코드","거래처","날짜","적요","금액"])
    for desc,amount in claims:
        ws.append([None,None,"2026-08-31",desc,amount])
    last=2+len(claims)
    ws.append(["001111","가상물류","소계",None,f"=SUM(E3:E{last})"])
    ws.append(["합 계",None,None,None,f'=SUMIF(C3:C{last+1},"소계",E3:E{last+1})'])
    wb.save(path); wb.close()


def generate(claims, raw_rows, with_subtotal=False):
    with tempfile.TemporaryDirectory() as folder:
        folder=Path(folder)
        prior=folder/"prior.xlsx"; statement(prior,claims)
        raw=folder/"raw.xlsx"; raw_file(raw,raw_rows)
        run=run_reconciliation([prior],[raw],{"25301"},PERIOD)
        out=folder/"out.xlsx"
        write_month_end_statement(out,[prior],run.results,run.new_items,run.issues,PERIOD,
            standalone_debits=run.standalone_debits,journal_items=run.journal_items,
            journal_sources=run.journal_sources,source_paths=[prior,raw])
        ws=load_workbook(out)["26.09"]
        rows={r:(ws.cell(r,4).value,ws.cell(r,5).value,str(ws.cell(r,6).value or ""))
              for r in range(3,ws.max_row+1) if ws.cell(r,4).value}
        if with_subtotal:
            subtotal=next(str(ws.cell(r,6).value) for r in range(3,ws.max_row+1)
                          if ws.cell(r,3).value=="소계")
            return run,rows,subtotal
        return run,rows


def pay(desc, amount, when="2026-09-03"):
    return ("001111","가상물류",desc,amount,0,when)


class MonthOnlyDifferenceTests(unittest.TestCase):
    def test_only_the_month_may_differ(self):
        yes=[("7월 위탁물류비용_GS팀","8월 위탁물류비용_GS팀"),
             ("9월 본사 임차료 및 8월 관리비","10월 본사 임차료 및 9월 관리비"),
             ("7 월 위탁물류_GS","8월  위탁물류_gs")]
        no=[("7월 위탁물류비용_GS팀","8월 위탁물류비용_CU팀"),
            ("7월 위탁물류비용_GS팀","7월 위탁물류비용_GS팀"),
            ("위탁물류","위탁물류비"),("",""),("월말 정산","월말 정리")]
        for a,b in yes: self.assertTrue(differs_only_by_month(a,b),(a,b))
        for a,b in no: self.assertFalse(differs_only_by_month(a,b),(a,b))

    def test_offset_sets_prefer_the_strictest_evidence_and_never_guess(self):
        rows=[("a","7월 X",D(100)),("b","8월 X",D(-100)),                 # month-only pair
              ("c","807",D(807)),("d","807",D(-500)),("e","807",D(-307)),  # 807 = 500 + 307
              ("f","7월 임차료",D(1000)),("g","8월 임차료",D(1000)),
              ("h","7월 임차료",D(-1000)),("i","8월 임차료",D(-1000)),     # exact-text pairs
              ("j","Q",D(10)),("k","Q",D(-4)),                              # does not add up to 0
              ("l","W",D(0)),("m","W",D(0))]                                # zero amounts ignored
        found={(o.kind,tuple(sorted(o.keys))) for o in offset_sets(rows)}
        self.assertEqual(found,{("pair",("a","b")),("group",("c","d","e")),
                                ("pair",("f","h")),("pair",("g","i"))})

    def test_a_group_must_add_up_to_exactly_zero_and_have_both_signs(self):
        self.assertEqual(offset_sets([("a","X",D(100)),("b","X",D(100))]),[])
        self.assertEqual(offset_sets([("a","X",D(100)),("b","X",D(-99))]),[])
        sets=offset_sets([("a","X",D(100)),("b","X",D(200)),("c","X",D(-300))])
        self.assertEqual([(o.kind,tuple(sorted(o.keys))) for o in sets],[("group",("a","b","c"))])

    def test_several_equal_amounts_without_a_unique_pair_still_net_to_zero_as_a_group(self):
        sets=offset_sets([("a","X",D(5)),("b","X",D(5)),("c","X",D(-5)),("d","X",D(-5))])
        self.assertEqual([(o.kind,tuple(sorted(o.keys))) for o in sets],[("group",("a","b","c","d"))])


class EngineTests(unittest.TestCase):
    def test_month_only_difference_is_a_clean_automatic_match(self):
        run,_=generate([("7월 위탁물류비용(기흥센터)_GS팀",807_567_651)],
                       [pay("8월 위탁물류비용(기흥센터)_GS팀",807_567_651)])
        result=run.results[0]
        self.assertEqual(result.status,Status.MATCHED)
        self.assertEqual(result.rule,"CODE_AMOUNT_UNIQUE_MONTH_ONLY")
        self.assertIn("월만 다름",result.reason)

    def test_any_other_description_difference_still_needs_a_look(self):
        run,_=generate([("7월 위탁물류비용(기흥센터)_GS팀",807_567_651)],
                       [pay("8월 위탁물류비용(기흥센터)_CU팀",807_567_651)])
        self.assertEqual(run.results[0].rule,"CODE_AMOUNT_UNIQUE_WITH_NOTE")

    def test_a_different_vendor_name_still_needs_a_look_even_if_only_the_month_differs(self):
        run,_=generate([("7월 위탁물류비용_GS팀",1000)],
                       [("001111","다른 이름","8월 위탁물류비용_GS팀",1000,0,"2026-09-03")])
        self.assertEqual(run.results[0].rule,"CODE_AMOUNT_UNIQUE_WITH_NOTE")


class StatementWordingTests(unittest.TestCase):
    def test_pair_rows_say_they_offset_each_other_and_that_only_the_month_differs(self):
        run,rows=generate([("7월 위탁물류비용(기흥센터)_GS팀",807_567_651)],
                          [pay("8월 위탁물류비용(기흥센터)_GS팀",807_567_651)])
        self.assertEqual(rows[3][2],"자동 대사 · 4행과 대차 (월만 다름)")
        self.assertEqual(rows[4][2],"자동 대사 · 3행과 대차 (월만 다름)")
        self.assertEqual(rows[3][1],-rows[4][1])

    def test_identical_descriptions_carry_no_month_remark(self):
        run,rows=generate([("7월 위탁물류비용_GS팀",1000)],[pay("7월 위탁물류비용_GS팀",1000)])
        self.assertEqual(rows[3][2],"자동 대사 · 4행과 대차")

    def test_net_vendor_rows_that_offset_each_other_name_each_other(self):
        # a negative adjustment makes this a signed-net vendor (no per-item matching)
        run,rows=generate(
            [("7월 위탁물류비용_GS팀",807),("계정정리)7월 패널티공제",-12),("8월 위탁물류비용_CU팀",500)],
            [pay("8월 위탁물류비용_GS팀",807)])
        self.assertEqual(run.results[0].status,Status.SIGNED_NET_AUTO)
        status={r:v[2].splitlines()[0] for r,v in rows.items()}
        self.assertEqual(status[3],"순잔액 대사 · 6행과 대차 (월만 다름)")
        self.assertEqual(status[6],"당월 지급 · 3행과 대차 (월만 다름)")
        # rows with no exact partner keep the generic wording
        self.assertNotIn("대차",status[4]+status[5])

    def test_a_paired_row_is_never_also_a_month_group_candidate(self):
        # 807 (7월 GS) and its payment ("8월 …_GS팀") are exactly paired, so the 8월 CU claim
        # must not claim that same payment as a "candidate", and the subtotal must not
        # report a month group built from it.
        run,rows,subtotal=generate(
            [("7월 위탁물류비용_GS팀",807),("계정정리)7월 패널티공제",-12),("8월 위탁물류비용_CU팀",500)],
            [pay("8월 위탁물류비용_GS팀",807)],with_subtotal=True)
        status={r:v[2] for r,v in rows.items()}
        self.assertEqual(status[5],"순잔액 대사")
        self.assertNotIn("묶음",subtotal)
        self.assertIn("✓ 검산 일치",subtotal)

    def test_an_ambiguous_partner_is_never_picked(self):
        run,rows=generate(
            [("7월 위탁물류비용_GS팀",807),("계정정리)7월 패널티공제",-12)],
            [pay("8월 위탁물류비용_GS팀",807),pay("9월 위탁물류비용_GS팀",807,"2026-09-04")])
        status={r:v[2] for r,v in rows.items()}
        self.assertFalse([t for t in status.values() if "행과 대차" in t and "후보" not in t],status)


class NotOneToOneTests(unittest.TestCase):
    """Payments split or merged, same amounts repeated: rows still say who offsets whom."""

    def last_lines(self, rows):
        return {r:v[2].splitlines()[-1] for r,v in rows.items()}

    def test_a_claim_paid_in_two_parts_names_all_three_rows(self):
        run,rows=generate([("7월 위탁물류비용_GS팀",807)],
                          [pay("7월 위탁물류비용_GS팀",500),pay("7월 위탁물류비용_GS팀",307,"2026-09-04")])
        last=self.last_lines(rows)
        self.assertEqual(last[3],"4~5행과 함께 대차 후보 (합계 0)")
        self.assertEqual(last[4],"3·5행과 함께 대차 후보 (합계 0)")
        self.assertEqual(last[5],"3~4행과 함께 대차 후보 (합계 0)")
        # the engine's own judgement stays on the first line, untouched
        self.assertIn("분할",rows[3][2].splitlines()[0])

    def test_equal_amounts_repeated_each_month_pair_up_month_by_month(self):
        run,rows=generate([("7월 임차료",1000),("8월 임차료",1000)],
                          [pay("7월 임차료",1000),pay("8월 임차료",1000,"2026-09-04")])
        last=self.last_lines(rows)
        self.assertEqual([last[r] for r in (3,4,5,6)],
                         ["5행과 대차 후보","6행과 대차 후보","3행과 대차 후보","4행과 대차 후보"])

    def test_a_verified_net_vendor_gets_the_group_wording_in_place_of_the_generic_one(self):
        # 807 - 12 + 500 - 500 - 307 = 488 > 0 : a verified net balance
        run,rows=generate(
            [("7월 위탁물류비용_GS팀",807),("계정정리)7월 패널티공제",-12),("8월 위탁물류비용_CU팀",500)],
            [pay("7월 위탁물류비용_GS팀",500),pay("7월 위탁물류비용_GS팀",307,"2026-09-04")])
        self.assertEqual(run.results[0].status,Status.SIGNED_NET_AUTO)
        first={r:v[2].splitlines()[0] for r,v in rows.items()}
        self.assertEqual(first[3],"순잔액 대사 · 6~7행과 함께 대차 (합계 0)")
        self.assertEqual(first[6],"당월 지급 · 3·7행과 함께 대차 (합계 0)")
        self.assertEqual(first[7],"당월 지급 · 3·6행과 함께 대차 (합계 0)")

    def test_nothing_is_said_when_the_rows_do_not_add_up_to_zero(self):
        run,rows=generate([("7월 위탁물류비용_GS팀",807)],[pay("7월 위탁물류비용_GS팀",500)])
        self.assertFalse([t for _,_,t in rows.values() if "대차" in t],rows)

    def test_different_items_paid_together_are_not_guessed(self):
        run,rows=generate([("7월 위탁물류비용_GS팀",100),("7월 위탁물류비용_CU팀",200)],
                          [pay("7월 위탁물류비용 합산",300)])
        self.assertFalse([t for _,_,t in rows.values() if "대차" in t],rows)


if __name__=="__main__":
    unittest.main()
