"""Real Excel cell-style and date-type preservation when inserting month-end rows."""
import tempfile
import unittest
from datetime import datetime, date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from adapters.excel.month_end_statement import write_month_end_statement
from domain.models import JournalLine, NewPayable, PayableItem, ReconcileResult, SourceRef, Status
from domain.period import AccountingPeriod


class FullTemplateStyleTests(unittest.TestCase):
    def test_accounting_negative_formats_and_text_sign_gaps_keep_correct_totals(self):
        from adapters.excel.reader import read_prior
        with tempfile.TemporaryDirectory() as folder:
            prior=Path(folder)/"original.xlsx"
            output=Path(folder)/"generated.xlsx"
            wb=Workbook();ws=wb.active;ws.title="26.07"
            ws.append(["7월 명세서"])
            ws.append(["코드","거래처","날짜","적요","금액"])
            ws.append([None,None,"2026-07-31","발생",1000])
            ws.append([None,None,"2026-07-31","지급1",-100])
            ws.append([None,None,"2026-07-31","지급2","-\u00a0 200"])
            ws.append(["001111","가상물류","소계",None,"=SUM(E3:E5)"])
            ws["E4"].number_format='_(* #,##0_);_(* -#,##0_);_(* "-"_);_(@_)'
            ws["E5"].number_format='#,##0;-#,##0;-'
            styles={r:ws.cell(r,5)._style for r in (4,5)}
            wb.save(prior);wb.close()
            before=prior.read_bytes()
            read=read_prior(prior,period=AccountingPeriod(2026,7))
            self.assertEqual(read.issues,[])
            self.assertEqual(read.review_items[0][0].amount,Decimal(700))
            write_month_end_statement(output,[prior],[],[],[],AccountingPeriod(2026,8))
            wb=load_workbook(output);ws=wb["26.08"]
            try:
                self.assertEqual([ws.cell(r,5).value for r in (3,4,5)],[1000,-100,-200])
                self.assertEqual(ws["E6"].value,"=SUM(E3:E5)")
                for row,style in styles.items():
                    self.assertEqual(ws.cell(row,5)._style,style)
                self.assertEqual(prior.read_bytes(),before)
            finally:
                wb.close()

    def test_signed_literal_formula_is_preserved_without_evaluating_references(self):
        from tests.test_month_end_template import write_grouped
        from adapters.excel.reader import read_prior
        with tempfile.TemporaryDirectory() as folder:
            prior=Path(folder)/"original.xlsx"
            output=Path(folder)/"generated.xlsx"
            write_grouped(prior)
            wb=load_workbook(prior);ws=wb["26.07"]
            ws["E3"]=1120000
            ws["E4"]="=-280000"
            ws["E5"]="=SUM(E3:E4)"
            wb.save(prior);wb.close()
            before=prior.read_bytes()
            read=read_prior(prior,period=AccountingPeriod(2026,7))
            self.assertEqual(read.issues,[])
            self.assertEqual(read.review_items[0][0].amount,Decimal("840000"))
            write_month_end_statement(output,[prior],[],[],[],AccountingPeriod(2026,8))
            wb=load_workbook(output)
            self.assertEqual(wb["26.08"]["E4"].value,"=-280000")
            wb.close()
            self.assertEqual(prior.read_bytes(),before)
            wb=load_workbook(prior);wb["26.07"]["E4"]="=-E3"
            wb.save(prior);wb.close()
            with self.assertRaisesRegex(ValueError,"4행.*지원하지 않는 수식"):
                write_month_end_statement(output,[prior],[],[],[],AccountingPeriod(2026,8))
            self.assertTrue(read_prior(prior,period=AccountingPeriod(2026,7)).issues)

    def test_retained_rows_keep_individual_styles_and_original_dates(self):
        from tests.test_month_end_template import write_grouped
        with tempfile.TemporaryDirectory() as folder:
            prior=Path(folder)/"original.xlsx"
            output=Path(folder)/"generated.xlsx"
            write_grouped(prior)
            wb=load_workbook(prior)
            ws=wb["26.07"]
            ws["C3"]="202-07-31"
            ws["C4"]=datetime(2026,7,31)
            ws["C4"].number_format="mm-dd-yy"
            ws["D4"].fill=PatternFill("solid",fgColor="4D009A")
            ws.row_dimensions[4].height=42
            ws["E5"].fill=PatternFill("solid",fgColor="FFFF00")
            ws["E5"].border=Border(bottom=Side(style="double"))
            expected={pos:(ws[pos].value,ws[pos]._style) for pos in ("C3","C4","D4")}
            subtotal_style=ws["E5"]._style
            wb.save(prior);wb.close()
            before=prior.read_bytes()
            write_month_end_statement(output,[prior],[],[],[],AccountingPeriod(2026,8))
            wb=load_workbook(output)
            try:
                ws=wb["26.08"]
                for pos,(value,style) in expected.items():
                    self.assertEqual(ws[pos].value,value)
                    self.assertEqual(ws[pos]._style,style)
                self.assertEqual(ws.row_dimensions[4].height,42)
                self.assertEqual(ws["E5"]._style,subtotal_style)
                self.assertEqual(prior.read_bytes(),before)
            finally:
                wb.close()

    def test_copies_font_borders_alignment_number_formats_and_date_type(self):
        with tempfile.TemporaryDirectory() as folder:
            prior=Path(folder)/"original.xlsx"
            output=Path(folder)/"generated.xlsx"
            wb=Workbook()
            ws=wb.active
            ws.title="26.07"
            ws["A1"]="2026년 7월 가상 미지급금"
            ws.append(["코드","거래처","날짜","적요","금액"])
            ws.append([None,None,datetime(2026,7,31),"7월 운임",100])
            ws.append(["001111","가상물류","소계",None,100])
            border=Border(left=Side(style="thin",color="111111"),
                          right=Side(style="medium",color="333333"),
                          bottom=Side(style="dotted",color="555555"))
            for c in ws[3][:5]:
                c.font=Font(name="굴림",size=10,color="223344")
                c.border=border
                c.alignment=Alignment(horizontal="center",vertical="center")
            ws["C3"].number_format="yyyy/mm/dd"
            ws["E3"].number_format="#,##0;[Red]-#,##0"
            ws.row_dimensions[3].height=29
            for c in ws[4][:5]:
                c.font=Font(name="굴림",size=10,bold=True)
                c.fill=PatternFill("solid",fgColor="FFF2A6")
                c.border=Border(top=Side(style="double",color="001122"))
            ws.row_dimensions[4].height=25
            wb.save(prior)
            wb.close()
            original_bytes=prior.read_bytes()
            payable=PayableItem("001111","가상물류","7월 운임",Decimal(100),"2026-07-31",3,
                                SourceRef(prior.name,"26.07",3,"담당자"))
            payment=JournalLine("001111","가상물류","25301","미지급금","7월 운임",
                                Decimal(100),Decimal(0),"2026-08-05",22)
            new=NewPayable(
                JournalLine("002222","가상신규","25301","미지급금","8월 신규",
                            Decimal(0),Decimal(250),"2026-08-23",33),
                Decimal(250),"당월 신규",True
            )
            write_month_end_statement(
                output,[prior],[ReconcileResult(Status.MATCHED,payable,payment,"일치","EXACT")],
                [new],[],AccountingPeriod(2026,8)
            )
            self.assertEqual(prior.read_bytes(),original_bytes)
            wb=load_workbook(output)
            try:
                ws=wb["26.08"]
                self.assertEqual(ws.max_row,7)
                self.assertEqual([ws.cell(r,5).value for r in (3,4,6)],[100,-100,250])
                self.assertEqual([r[0] for r in list(wb["대사내역"].values)[1:]],["대사 일치 후보"])
                for i in (3,4,6):
                    cell=ws.cell(i,5)
                    self.assertEqual(cell.font.name,"굴림")
                    self.assertEqual(cell.font.sz,10)
                    self.assertEqual(cell.border.left.style,"thin")
                    self.assertEqual(cell.border.right.style,"medium")
                    self.assertEqual(cell.border.bottom.style,"dotted")
                    self.assertEqual(cell.alignment.horizontal,"center")
                    self.assertEqual(cell.number_format,"#,##0;[Red]-#,##0")
                    self.assertEqual(ws.row_dimensions[i].height,29)
                    self.assertEqual(ws.cell(i,3).number_format,"yyyy/mm/dd")
                    self.assertIsInstance(ws.cell(i,3).value,(datetime,date))
                self.assertEqual(ws["C4"].value.date() if isinstance(ws["C4"].value,datetime)
                                 else ws["C4"].value,date(2026,8,5))
                for i in (5,7):
                    subtotal=ws.cell(i,5)
                    self.assertEqual(subtotal.fill.fgColor.rgb,"00FFF2A6")
                    self.assertEqual(subtotal.border.top.style,"double")
                    self.assertTrue(subtotal.font.bold)
                    self.assertEqual(ws.row_dimensions[i].height,25)
                self.assertEqual(ws["E5"].value,"=SUM(E3:E4)")
                self.assertEqual(ws["E7"].value,"=SUM(E6:E6)")
                self.assertEqual(ws["B5"].value,"가상물류")
                self.assertEqual(ws["B7"].value,"가상신규")
                self.assertEqual(ws["F2"].value,"처리상태")  # Extra column leaves source A:E untouched
                self.assertTrue(wb.calculation.fullCalcOnLoad)
            finally:
                wb.close()


if __name__=="__main__":
    unittest.main()
