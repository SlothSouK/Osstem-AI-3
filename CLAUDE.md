# CLAUDE.md — Osstem-AI-3 프로젝트

## 프로젝트 개요

여섯 가지 독립적인 모듈로 구성된 프로젝트:

1. **챗봇 웹앱** (`backend/` + `frontend/`) — Claude API 기반 풀스택 채팅 앱
2. **충당금 자동화** (`automation/`) — 해외관리2팀 ERP 업무 자동화 스크립트
3. **웹 스크래퍼** (`ostconfin/`) — confinas.osstem.com 데이터 수집 → 엑셀 저장
4. **SAP 채권 자동화** (`sapost/`) — FBL5N 미결항목 ALV 직독 → 엑셀 저장
5. **영업수금 집계** (`collection/`) — 자금일보 엑셀에서 Category=Collection 합산 → 검증 파일 생성
6. **Excel 버전 비교** (`vercheck/`) — 두 버전 Excel 파일을 헤더+행키 기반으로 비교 → 변경 리포트 생성

---

## 디렉토리 구조

```
Osstem-AI-3/
├── backend/              FastAPI + Anthropic SDK (Python)
│   ├── main.py           앱 진입점, CORS 설정
│   ├── routes/chat.py    /api/chat, /api/chat/stream 엔드포인트
│   └── requirements.txt
├── frontend/             React 18 + TypeScript + Vite + Tailwind
│   └── src/
│       ├── App.tsx
│       ├── api/chat.ts   백엔드 API 호출
│       └── components/   ChatWindow, MessageBubble, InputBar
├── automation/           충당금 자동화 (Python)
│   ├── main.py           실행 진입점 (python automation/main.py --month YYYYMM)
│   ├── config/
│   │   ├── config.ini    ERP 컨트롤명, 경로, 셀 위치 설정
│   │   └── .env          로그인 정보 (git 제외)
│   └── src/
│       ├── erp_controller.py   ERP 창 자동 조작
│       ├── downloader.py       엑셀 다운로드
│       ├── data_processor.py   데이터 정제 (pandas)
│       ├── template_writer.py  양식 붙여넣기 (openpyxl)
│       └── utils.py            로거, retry 데코레이터
├── .claude/
│   └── commands/
│       └── sync.md       /sync 커스텀 슬래시 커맨드
├── ostconfin/            웹 데이터 수집 → 엑셀 저장 (Playwright + openpyxl)
│   ├── config/
│   │   ├── config.ini    셀렉터, 시트명, 옵션 설정
│   │   └── .env          URL, 로그인 정보, 엑셀 경로 (git 제외)
│   ├── scraper.py        실행 파일
│   └── requirements.txt
├── sapost/               SAP FBL5N 채권 자동화 (win32com + pandas)
│   ├── main.py           전체 파이프라인 진입점
│   ├── fbl5n_download.py FBL5N 전용 다운로드 스크립트
│   ├── config/
│   │   ├── config.ini    트랜잭션, 필드 ID, 경로 설정
│   │   └── .env          SAP 로그인 정보 (git 제외)
│   ├── src/
│   │   ├── sap_controller.py   SAP GUI 연결/조작 (win32com)
│   │   ├── data_processor.py   데이터 정제 (pandas)
│   │   ├── template_writer.py  엑셀 양식 기입 (openpyxl)
│   │   └── utils.py            로거, retry 데코레이터
│   └── requirements.txt
├── collection/           영업수금 집계 (openpyxl)
│   ├── main.py           실행 진입점 (python collection/main.py)
│   ├── config/
│   │   └── config.ini    자금일보 경로, 집계 대상 Category 설정
│   ├── src/
│   │   └── collector.py  시트 탐색·집계·검증 파일 생성
│   └── requirements.txt
├── vercheck/             Excel 버전 비교 (pandas + openpyxl)
│   ├── main.py           실행 진입점 (python vercheck/main.py --old a.xlsx --new b.xlsx)
│   ├── config/
│   │   └── config.ini    헤더 탐지·키 컬럼·리포트 설정
│   ├── src/
│   │   ├── utils.py          로거, 설정 로드
│   │   ├── excel_reader.py   ExcelReader — 헤더 자동 탐지, SheetData 생성
│   │   ├── comparator.py     SheetComparator, WorkbookComparator — diff 로직
│   │   └── report_writer.py  ReportWriter — 요약·변경내역·비교_* 시트 작성
│   └── requirements.txt
└── CLAUDE.md
```

---

## 챗봇 웹앱 실행

```bash
# 백엔드 (포트 8000)
cd backend
pip install -r requirements.txt
uvicorn main:app --reload

# 프론트엔드 (포트 5173)
cd frontend
npm install
npm run dev
```

환경변수: `backend/.env` 에 `ANTHROPIC_API_KEY=...` 설정

---

## 충당금 자동화 실행

```bash
# 의존성 설치
pip install -r automation/requirements.txt

# 전체 실행 (ERP 자동 조작 포함)
python automation/main.py --month 202503

# ERP 없이 엑셀 처리만 테스트
python automation/main.py --month 202503 --skip-erp
```

실행 전 필수 설정:
- `automation/config/.env` — ERP_USER_ID, ERP_PASSWORD
- `automation/config/config.ini` — window_title, 컨트롤 이름, start_cell

---

## SAP 채권 자동화 실행 (sapost)

"채권명세서 업데이트"라고 부름. 실행 전 법인·기간·유형·경로 조건 확인.

```bash
# 의존성 설치
pip install -r sapost/requirements.txt

# 단일 월 전체 법인
python sapost/fbl5n_download.py --keydate 202603

# 날짜 범위 직접 지정 (여러 달에 걸친 기간)
python sapost/fbl5n_download.py --budat_low 20260101 --budat_high 20260331

# 특정 계정만 재실행
python sapost/fbl5n_download.py --budat_low 20260101 --budat_high 20260331 --accounts 1700006 1700029

# 경로 직접 지정 (config.ini 수정 없이 조건문 경로 override)
python sapost/fbl5n_download.py --budat_low 20260101 --budat_high 20260331 --source_dir "D:\경로\채권명세서"
```

동작:
- `--source_dir` 지정 시 config.ini 수정 없이 해당 경로 사용, raw_dir은 하위 `raw/` 자동 설정
- 파일명 앞 7자리 숫자 = 고객계정 (source_dir 내 파일 자동 탐지)
- 원본 파일은 수정하지 않고 **복사본** 생성 후 작업 (`[고객코드] [법인명]법인채권명세서_YYYYMMDD.xlsx`)
- **미결항목**: FBL5N → 회사코드 1000, 모든 항목, 전기일 기간 조회 → `raw/{계정}-{YYYYMM}.xlsx`
- **반제항목**: 반제일 기간으로 재조회(반제항목 모드) → `raw/{계정}-{YYYYMM}_offset.xlsx`
- SG=M → 미수금(잔액)/미수금 시트 append / 그 외 → 외화외상매출금(잔액)/외화외상매출금 시트 append
- offset 매칭: **지정값 + 증빙일 둘 다 일치**해야 반제로 인식 (지정값은 앞 0 제거 변환도 병행)
- 부분 수금: 잔액 시트 금액 차감 + 경고 알림 / 총액 시트 기상환액 기입 / 완전 수금은 행 삭제
- offset 반영: 기상환액 기입, 상환일·반제전표 기입 (총액 시트만); 반제전표 헤더 없어도 상환일 옆 셀(+1)에 자동 기입, 이미 값 있으면 우측으로 이동하여 빈 셀 쌍에 기입
- 경과기간 수식: append·offset 완료 후 마지막에 일괄 기입, D6 셀에 조회 종료일 기입
- 헤더 행 자동 탐지 + 한국어·영어·스페인어 컬럼명 alias 매핑 (법인별 양식 차이 대응)
- 법인명은 CLAUDE.md 법인코드 목록에서 동적으로 읽어옴
- 만기일: raw 파일 '순 만기일' 컬럼 → 잔액 시트 만기일 헤더 자동 기입 (공백 포함 컬럼명 대응)

채권명세서 파일 위치: `--source_dir` 인수 또는 config.ini `[PATHS] source_dir` — 파일명 앞 7자리 = 고객코드

인터랙티브 모드 법인 입력 시 **'일체' 키워드** 지원:
- `일체` 단독 → 경로 내 전체 (엔터와 동일)
- `해외법인 일체` → CLAUDE.md 법인코드 목록 전체
- `유럽 일체` / `중국 일체` 등 → 해당 키워드 부분 일치 법인 전체 자동 선택

실행 전 필수 설정:
- `sapost/config/.env` — SAP_USER_ID, SAP_PASSWORD, SAP_CLIENT=100
- SAP GUI 실행 중 + 스크립팅 활성화 필요 (Alt+F12 → Options → Scripting)
- `sapost/config/config.ini` — `source_dir`, `raw_dir` 경로 확인 (--source_dir 인수로 override 가능)

---

## 영업수금 집계 실행 (collection)

자금일보 엑셀 파일에서 inflow 테이블의 Category = Collection 행을 합산해 영업수금을 산출한다.

```bash
# 의존성 설치
pip install -r collection/requirements.txt

# 기본 실행 (config.ini source_file 사용)
python collection/main.py

# 파일 직접 지정
python collection/main.py --file "D:\경로\자금일보.xlsx"

# 출력 파일 경로 지정
python collection/main.py --file "D:\경로\자금일보.xlsx" --output "D:\결과\영업수금.xlsx"
```

동작:
- 날짜 형식 시트(DD.MM.YYYY) 전체 탐색
- 각 시트에서 inflow 테이블 헤더 행 자동 탐지 (Category 컬럼 위치 기준)
- Category = Collection (대소문자 무시) 행 Amount 합산 → 영업수금
- 검증용 엑셀 파일 생성 (시트별 소계 + 전체 합계)

설정: `collection/config/config.ini` — source_file 경로, target_category, sheet_pattern

---

## Excel 버전 비교 실행 (vercheck)

두 버전의 Excel 파일을 비교해 변경 포인트를 파악한다.  
셀 위치가 달라져도 헤더명·행 키 기반으로 동일 항목끼리 비교.

```bash
# 의존성 설치
pip install -r vercheck/requirements.txt

# 기본 실행 (두 파일 직접 지정)
python vercheck/main.py --old old.xlsx --new new.xlsx

# 출력 경로 지정
python vercheck/main.py --old old.xlsx --new new.xlsx --output D:\결과\비교리포트.xlsx
```

동작:
- 시트 단위로 비교, 추가/삭제 시트 자동 감지
- 각 시트 헤더 행 자동 탐지 (상위 15행 스캔, 비공백 비율 기준)
- 행 키 컬럼: `config.ini key_columns` 지정 → 고유도 자동 탐지 → 첫 번째 컬럼 순으로 결정
- 숫자 변경: delta·증감률 계산 / 문자 변경: 구→신 표시
- 결과 리포트 시트 구성:
  - `요약`: 시트별 변경셀/추가행/삭제행/추가열/삭제열 수
  - `변경내역`: 모든 변경사항 플랫 리스트
  - `비교_[시트명]`: 색상 코딩 (노랑=변경셀, 초록=추가행, 빨강=삭제행)

설정: `vercheck/config/config.ini` — `key_columns`, `skip_sheets`, `header_scan_rows`

---

## 자동화 모듈 수정 시 주의사항

- **ERP 컨트롤 이름** (`config.ini [ERP]`): Inspect.exe 로 실제 ERP 창에서 확인
- **pywinauto 우선**, UI 접근 불가 시 pyautogui fallback (좌표 수동 수정 필요)
- **템플릿 원본** (`automation/templates/`) 은 수정 금지 — 복사 후 작업
- **중간 결과** (`data/intermediate/*.pkl`): 재실행 시 체크포인트로 사용됨, 재처리 필요 시 삭제

---

## 기술 스택

| 영역 | 스택 |
|------|------|
| 백엔드 | Python 3.11+, FastAPI 0.115, Anthropic SDK 0.34 |
| 프론트엔드 | React 18, TypeScript 5, Vite 5, Tailwind CSS 3 |
| 자동화 | pywinauto 0.6, pyautogui 0.9, pandas 2.2, openpyxl 3.1 |
| 웹 스크래핑 | Playwright 1.44, pandas 2.2, openpyxl 3.1 |
| SAP 자동화 | pywin32 311, pandas 2.2, openpyxl 3.1, python-dotenv 1.0 |
| Excel 비교 | pandas 2.2, openpyxl 3.1 |
| Claude 모델 | claude-sonnet-4-6 |

---

## 코딩 규칙

- Python: 타입 힌트 사용, 클래스 단위로 모듈 분리
- 로그인 정보/API 키는 절대 코드에 직접 입력 금지 → `.env` 사용
- 자동화 모듈은 `--skip-erp` 로 ERP 없이도 단독 테스트 가능하게 유지

---

## 법인코드 목록

| 법인코드 | 법인명 | 고객코드 |
|---------|--------|---------|
| AE1 | 중동법인 | 1700030 |
| AE2 | 중동내수법인 | 1700057 |
| AU1 | 호주법인 | 1700012 |
| BD1 | 방글라데시법인 | 1700018 |
| BR1 | 브라질법인 | 1700029 |
| BR2 | IMPLACIL | 1700058 |
| CA1 | 캐나다법인 | 1700019 |
| CL1 | 칠레법인 | 1700024 |
| CN1 | 중국법인 | 1700003 |
| CN2 | 광동법인 | 1700020 |
| CN3 | 천진법인 | 1700025 |
| CN4 | 염성법인 | 1700028 |
| CO1 | 콜롬비아법인 | 1700056 |
| CZ1 | 유럽법인 | 1700031 |
| DE1 | 독일법인 | 1700001 |
| ES1 | 스페인법인 | 1700036 |
| FR1 | 프랑스법인 | 1700037 |
| GE1 | 조지아법인 | 1700051 |
| HK1 | 홍콩법인 | 1700007 |
| ID1 | 인도네시아법인 | 1700013 |
| IN1 | 인도법인 | 1700004 |
| IT1 | 이탈리아법인 | 1700041 |
| JP1 | 일본법인 | 1700005 |
| JP2 | 일본디지털센터 | |
| KR1 | 오스템임플란트 | |
| KRA | ㈜대한치과교육개발원 | |
| KRB | 오스템파마㈜ | |
| KRC | 오스템글로벌㈜ | |
| KRD | 오스템바스큘라㈜ | |
| KRE | 오스템올소㈜ | |
| KRF | 탑플란㈜ | |
| KRG | 코잔㈜ | |
| KRH | 오스템인테리어㈜ | |
| KRI | ㈜메디칼소프트 | |
| KRJ | 이베스트-어센트신기술투자조합제1호 | |
| KZ1 | 카자흐스탄법인 | 1700017 |
| MN1 | 몽골법인 | 1700022 |
| MX1 | 멕시코법인 | 1700015 |
| MY1 | 말레이시아법인 | 1700009 |
| NL1 | 네덜란드법인 | 1700047 |
| NZ1 | 뉴질랜드법인 | 1700027 |
| PH1 | 필리핀법인 | 1700016 |
| PT1 | 포르투갈법인 | 1700046 |
| RU1 | 러시아법인 | 1700002 |
| SG1 | 싱가폴법인 | 1700008 |
| TH1 | 태국법인 | 1700011 |
| TR1 | 튀르키예법인 | 1700021 |
| TW1 | 대만법인 | 1700000 |
| UA1 | 우크라이나법인 | 1700023 |
| US1 | 미국법인 | 1700006 |
| UZ1 | 우즈베키스탄법인 | 1700026 |
| VN1 | 베트남법인 | 1700014 |
