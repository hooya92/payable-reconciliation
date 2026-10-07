# 명세서 대사 프로그램

더존 iCUBE의 **미지급금 계정 전표출력 Excel**과 전월 **거래처별 명세서 Excel**을 대사해 차월 명세서 초안을 만들고, 사람이 확인해야 할 예외만 추리는 Windows용 로컬 도구입니다.

## 안전 원칙
- 더존 ERP를 직접 조작하거나 입력하지 않습니다.
- 입력 Excel은 읽기 전용으로 사용하며 덮어쓰지 않습니다.
- 결과는 항상 별도 Excel 파일로 저장합니다.
- 확실한 일치만 자동 대사하고, 중복·적요 불일치·미지급 등 애매한 건은 예외로 남깁니다.
- 실제 회사 데이터는 GitHub에 올리지 않습니다.

## 1차 버전
- 전월 명세서 + 당월 더존 전표출력 선택
- 거래처코드 + 금액 + 적요 기준 보수적 자동 대사
- 예외 중심 화면
- 당월 대변 신규 명세 항목 추출
- 별도 `명세서_대사결과.xlsx` 생성

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


## 악성 가상 Excel 직접 생성
실제 회사 자료 없이 지저분한 입력/애매한 대사 케이스를 눈으로 확인하려면:

```bash
python -m tests.synthetic_adversarial_factory
```

프로젝트 아래 `synthetic_adversarial/` 폴더에 가상 `.xlsx` 두 개가 생성됩니다. 생성 파일은 Git에서 무시되며 회사 데이터는 포함하지 않습니다.


## 정상 가상 더존 Raw 생성
실제 더존 Raw 파일이 없어도 UI/대사 흐름을 확인할 수 있도록 정상 가상 Raw를 만들 수 있습니다.

```bash
python -m tests.synthetic_douzone_raw_factory
```

프로젝트 아래 `synthetic_raw/더존_Raw_정상_2026_1년.xlsx`가 생성됩니다. 더존 전표출력 형태처럼 두 개의 `코드` 열을 사용하고, 여러 월의 전표와 25301 외 계정도 섞어 대상월/계정 필터를 확인할 수 있게 구성되어 있습니다.
