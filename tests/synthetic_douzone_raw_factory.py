"""Generate a clean, realistic-looking synthetic Douzone Raw workbook.

Run locally:
    python -m tests.synthetic_douzone_raw_factory

The generated .xlsx contains only synthetic data and is ignored by git.
"""
from pathlib import Path
from openpyxl import Workbook


def build(out_dir="synthetic_raw"):
    root=Path(out_dir)
    root.mkdir(parents=True,exist_ok=True)
    path=root/"더존_Raw_정상_2026_1년.xlsx"

    wb=Workbook()
    ws=wb.active
    ws.title="전표출력"

    # Actual Douzone exports may use two identical '코드' headers:
    # account code before 계정과목명, vendor code before 거래처명.
    ws.append(["결의일","결의No","순번","기표일자","기표번호","구분",
               "코드","계정과목명","코드","거래처명","적요","차변","대변"])

    rows=[
        # Prior periods / noise
        ("2026-01-15",1,1,"2026-01-15",1,"일반","25301","미지급금-일반","010001","가상물류A","1월 운송비",0,1250000),
        ("2026-03-20",2,1,"2026-03-20",2,"일반","25301","미지급금-일반","010002","가상주유B","3월 유류비",0,830000),
        ("2026-07-31",3,1,"2026-07-31",3,"일반","25301","미지급금-일반","051330","가상신일물류","7월 FMC사업부(주류) 용차료/유류비",0,64004897),
        ("2026-07-31",4,1,"2026-07-31",4,"일반","25301","미지급금-일반","051330","가상신일물류","7월 FMC사업부(일반) 용차료/유류비",0,19325426),

        # Target month: August payments clearing July lines
        ("2026-08-05",5,1,"2026-08-05",5,"일반","25301","미지급금-일반","051330","가상신일물류","7월 FMC사업부(주류) 용차료/유류비",64004897,0),
        ("2026-08-05",6,1,"2026-08-05",6,"일반","25301","미지급금-일반","051330","가상신일물류","7월 FMC사업부(일반) 용차료/유류비",19325426,0),

        # Target month: new statement items
        ("2026-08-20",7,1,"2026-08-20",7,"일반","25301","미지급금-일반","051330","가상신일물류","8월 FMC사업부(주류) 용차료/유류비",0,68190989),
        ("2026-08-20",8,1,"2026-08-20",8,"일반","25301","미지급금-일반","051330","가상신일물류","8월 FMC사업부(일반) 용차료/유류비",0,18977659),
        ("2026-08-25",9,1,"2026-08-25",9,"일반","25301","미지급금-일반","012345","가상에너지","8월 유류비",0,2450000),

        # Other account: should be ignored when account code 25301 is selected
        ("2026-08-26",10,1,"2026-08-26",10,"일반","11100","보통예금","020001","가상은행","계좌이체",500000,0),

        # Later month / noise
        ("2026-09-10",11,1,"2026-09-10",11,"일반","25301","미지급금-일반","010003","가상서비스C","9월 시스템 유지보수",0,990000),
        ("2026-12-15",12,1,"2026-12-15",12,"일반","25301","미지급금-일반","010004","가상물류D","12월 운송비",0,1550000),
    ]

    for row in rows:
        ws.append(row)

    # Mimic numeric vendor codes with display formats preserving leading zeroes.
    for r in range(2, ws.max_row+1):
        ws.cell(r,9).number_format="000000"

    wb.save(path)
    return path


if __name__=="__main__":
    path=build()
    print(f"생성 완료: {path.resolve()}")
