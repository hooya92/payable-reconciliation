"""Small adversarial Excel dataset for human inspection and E2E safety tests.

Run locally:
    python -m tests.synthetic_adversarial_factory

Generated .xlsx files are ignored by git and contain only synthetic data.
"""
from pathlib import Path
from openpyxl import Workbook


def build(out_dir="synthetic_adversarial"):
    root=Path(out_dir); root.mkdir(parents=True,exist_ok=True)

    prior=root/"악성_전월명세.xlsx"
    wb=Workbook(); ws=wb.active; ws.title="명세"
    ws.append(["거래처코드","거래처명","날짜","적요","금액"])
    ws.append(["001001","가상정상","2026-07-31","7월 유류비",110000])               # exact
    ws.append(["001002","가상적요","2026-07-31","7월 용차료",220000])               # desc mismatch
    ws.append(["001003","가상부분","2026-07-31","7월 보관비",330000])               # partial debit
    ws.append(["001004","가상이전상호","2026-07-31","7월 시스템비",440000])          # vendor name changed
    ws.append(["001005","가상타거래처","2026-07-31","7월 수수료",550000])             # same amount other vendor
    ws.append(["001006","가상미지급","2025-12-31","장기 이월 임차료",660000])          # unpaid / long carry
    ws.append(["001007","가상중복후보","2026-07-31","7월 운송비",770000])             # duplicate debit candidates
    ws.append(["001008","가상이상금액","2026-07-31","7월 기타비","8600만원"])          # suspicious amount
    wb.save(prior)

    raw=root/"악성_더존Raw.xlsx"
    wb=Workbook(); ws=wb.active; ws.title="전표출력"
    ws.append(["기표일자","계정코드","계정과목명","거래처코드","거래처명","적요","차변","대변"])
    ws.append(["2026-08-03","25301","미지급금-일반","001001","가상정상","7월 유류비",110000,0])
    ws.append(["2026-08-04","25301","미지급금-일반","001002","가상적요","7월 용차료 수정",220000,0])
    ws.append(["2026-08-05","25301","미지급금-일반","001003","가상부분","7월 보관비",100000,0])
    ws.append(["2026-08-06","25301","미지급금-일반","001004","가상변경상호","7월 시스템비",440000,0])
    ws.append(["2026-08-07","25301","미지급금-일반","009999","가상다른곳","7월 수수료",550000,0])
    ws.append(["2026-08-08","25301","미지급금-일반","001007","가상중복후보","7월 운송비",770000,0])
    ws.append(["2026-08-09","25301","미지급금-일반","001007","가상중복후보","7월 운송비",770000,0])

    # Current-month new statement item: review only, never auto-added to next-month draft.
    ws.append(["2026-08-20","25301","미지급금-일반","002001","가상신규","8월 신규 유류비",0,880000])

    # Dirty rows: reader must quarantine them, not reinterpret them as normal.
    ws.append(["2026-08-21","25301","미지급금-일반","002002","가상역분개","수정 전표",-990000,0])
    ws.append(["2026-08-22","25301","미지급금-일반","002003","가상이상행","차대 동시",120000,120000])

    # Outside selected month: must be ignored by the period filter.
    ws.append(["2026-07-31","25301","미지급금-일반","001006","가상미지급","장기 이월 임차료",660000,0])
    wb.save(raw)

    return {
        "prior_path": prior,
        "raw_path": raw,
        "expected": {
            "AUTO_MATCH": 1,
            "DESCRIPTION_MISMATCH": 1,
            "POSSIBLE_PARTIAL": 1,
            "VENDOR_NAME_MISMATCH": 1,
            "VENDOR_MISMATCH": 1,
            "UNPAID": 1,
            "DUPLICATE": 1,
            "PRIOR_INPUT_ISSUES": 1,
            "RAW_INPUT_ISSUES": 2,
            "CURRENT_MONTH_NEW_REVIEW": 1,
        },
    }


if __name__=="__main__":
    result=build()
    print(f"생성 완료: {Path(result['prior_path']).parent.resolve()}")
    print(result["expected"])
