# CLAUDE.md — Osstem-AI-3 프로젝트

## 프로젝트 개요

네 가지 독립적인 모듈로 구성된 프로젝트:

1. **챗봇 웹앱** (`backend/` + `frontend/`) — Claude API 기반 풀스택 채팅 앱
2. **충당금 자동화** (`automation/`) — 해외관리2팀 ERP 업무 자동화 스크립트
3. **웹 스크래퍼** (`ostconfin/`) — confinas.osstem.com 데이터 수집 → 엑셀 저장
4. **SAP 채권 자동화** (`sapost/`) — FBL5N 미결항목 ALV 직독 → 엑셀 저장

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

```bash
# 의존성 설치
pip install -r sapost/requirements.txt

# FBL5N 전체 실행 (해당 월 전체 법인)
python sapost/fbl5n_download.py --keydate 202603

# 특정 계정만 재실행
python sapost/fbl5n_download.py --keydate 202603 --accounts 1700006 1700029
```

동작:
- `source_dir` (config.ini) 의 파일명 앞 7자리 숫자 = 고객계정
- **미결항목**: FBL5N → 회사코드 1000, 모든 항목, 전기일 기간 조회 → `raw/{계정}-{YYYYMM}.xlsx`
- **반제항목**: 동일 기간을 반제일 기간으로 재조회 → `raw/{계정}-{YYYYMM}_offset.xlsx`
- SG=M → 미수금(잔액)/미수금 시트 append / 그 외 → 외화외상매출금(잔액)/외화외상매출금 시트 append
- offset 데이터: 지정 열 매칭 → 기상환액 기입, 상환일·반제전표 기입, 잔액=0 행 자동 삭제
- 헤더 행 자동 탐지 + 한국어·영어·스페인어 컬럼명 alias 매핑 (법인별 양식 차이 대응)

채권명세서 파일 위치: `D:\해외관리실\채권명세서\자동화\ar\` (파일명 앞 7자리 = 고객코드)

실행 전 필수 설정:
- `sapost/config/.env` — SAP_USER_ID, SAP_PASSWORD, SAP_CLIENT=100
- SAP GUI 실행 중 + 스크립팅 활성화 필요 (Alt+F12 → Options → Scripting)
- `sapost/config/config.ini` — `source_dir` 경로 확인

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
