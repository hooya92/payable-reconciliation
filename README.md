# Payable Reconciliation

더존 iCUBE의 **전표출력 Excel**과 전월 **미지급(거래처) 세부명세 Excel**을 읽어 사람이 확인해야 할 예외만 추리는 Windows용 로컬 대사 도구입니다.

## 안전 원칙
- 더존 ERP를 직접 조작하거나 입력하지 않습니다.
- 입력 Excel은 읽기 전용으로 사용하며 덮어쓰지 않습니다.
- 결과는 항상 별도 Excel 파일로 저장합니다.
- 확실한 일치만 자동 대사하고, 중복·적요 불일치·미지급 등 애매한 건은 예외로 남깁니다.
- 실제 회사 데이터는 GitHub에 올리지 않습니다.

## 1차 버전
- 전월 명세 + 당월 더존 전표출력 선택
- 거래처코드 + 금액 + 적요 기준 보수적 자동 대사
- 예외 중심 화면
- 당월 대변 신규 미지급 추출
- 별도 `대사결과.xlsx` 생성

## 실행
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

## 테스트
```bash
python -m unittest discover -s tests -v
```

> 현재는 실제 회사 파일이 아닌 합성 데이터 기준으로 개발합니다. 첫 실파일 검증에서 더존/명세서 헤더 구조와 업무 예외를 확인한 뒤 규칙을 보강합니다.
