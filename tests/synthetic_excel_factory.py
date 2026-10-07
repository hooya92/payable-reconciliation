"""Deterministic synthetic E2E dataset. Contains no company data."""
from pathlib import Path
from random import Random
from openpyxl import Workbook

SEED=20261007

def build(out_dir="synthetic_data", prior_count=2400, raw_noise=7200):
    rng=Random(SEED); root=Path(out_dir); root.mkdir(parents=True,exist_ok=True)
    owners=["담당자A","담당자B","담당자C","담당자D","담당자E","담당자F"]
    priors=[[] for _ in owners]; raw=[]; expected={"MATCHED":0,"MATCHED_WITH_NOTE":0,"DESCRIPTION_MISMATCH":0,"UNPAID":0,"VENDOR_MISMATCH":0}
    for i in range(prior_count):
        code=f"{1000+i:06d}"; vendor=f"가상거래처{i:04d}"; desc=f"7월 가상운송비 {i:04d}"; amount=100000+(i*7919)%90000000
        priors[i%len(owners)].append((code,vendor,"2026-07-31",desc,amount))
        kind=i%40
        if kind==0:
            raw.append(("2026-07-10","25301","미지급금-일반","999999","가상오류거래처",desc,amount,0)); expected["VENDOR_MISMATCH"]+=1
        elif kind==1:
            expected["UNPAID"]+=1
        elif kind==2:
            raw.append(("2026-07-10","25301","미지급금-일반",code,vendor,desc+" 적요변경",amount,0)); expected["MATCHED"]+=1; expected["MATCHED_WITH_NOTE"]+=1
        else:
            raw.append(("2026-07-10","25301","미지급금-일반",code,vendor,desc,amount,0)); expected["MATCHED"]+=1
    # Current-month new credits.
    for i in range(600):
        code=f"{700000+i:06d}"; raw.append(("2026-07-20","25301","미지급금-일반",code,f"가상신규{i:04d}",f"7월 신규비용 {i:04d}",0,200000+(i*3571)%50000000))
    # Long-range noise: same account but outside target month.
    for i in range(raw_noise):
        month=(i%12)+1
        if month==7: month=6
        code=f"{800000+(i%1000):06d}"; raw.append((f"2026-{month:02d}-15","25301","미지급금-일반",code,f"기간외가상{i%1000:04d}",f"기간외 전표 {i}",0,1000+i))
    rng.shuffle(raw)
    for owner,rows in zip(owners,priors):
        wb=Workbook(); ws=wb.active; ws.title="26.07"
        ws.append(["거래처코드","거래처명","날짜","적요","금액"])
        for row in rows: ws.append(row)
        wb.save(root/f"{owner}_2026-07_미지급.xlsx")
    wb=Workbook(); ws=wb.active; ws.title="전표출력"
    ws.append(["결의일","결의No","순번","기표일자","기표번호","구분","코드","계정과목명","코드","거래처명","적요","차변","대변"])
    for n,row in enumerate(raw,1):
        dt,ac,an,vc,vn,desc,debit,credit=row
        ws.append([dt,n,1,dt,n,"일반",ac,an,vc,vn,desc,debit,credit])
    wb.save(root/"더존_Raw_2026_1년.xlsx")
    return {"prior_count":prior_count,"raw_count":len(raw),"new_payables":600,**expected}

if __name__=="__main__":
    print(build())
