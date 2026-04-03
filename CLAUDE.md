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

# FBL5N 미결항목 다운로드 (YYYYMM → 해당 월 말일 기준)
python sapost/fbl5n_download.py --keydate 202603
```

동작:
- `D:\해외관리실\해외법인\1. 임플라시우\채권명세서` 파일명 앞 7자리 숫자 = 고객계정
- FBL5N → 회사코드 1000, 미결항목, 특별G/L거래+임시항목 체크 → ALV 직독
- 전기일자 오름차순 정렬 후 `raw/{계정코드}-{YYYYMM}.xlsx` 저장

실행 전 필수 설정:
- `sapost/config/.env` — SAP_USER_ID, SAP_PASSWORD, SAP_CLIENT=100
- SAP GUI 실행 중 + 스크립팅 활성화 필요 (Alt+F12 → Options → Scripting)

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

| 법인코드 | 법인명 |
|---------|--------|
| AE1 | 중동법인 |
| AE2 | 중동내수법인 |
| AU1 | 호주법인 |
| BD1 | 방글라데시법인 |
| BR1 | 브라질법인 |
| BR2 | IMPLACIL |
| CA1 | 캐나다법인 |
| CL1 | 칠레법인 |
| CN1 | 중국법인 |
| CN2 | 광동법인 |
| CN3 | 천진법인 |
| CN4 | 염성법인 |
| CO1 | 콜롬비아법인 |
| CZ1 | 유럽법인 |
| DE1 | 독일법인 |
| ES1 | 스페인법인 |
| FR1 | 프랑스법인 |
| GE1 | 조지아법인 |
| HK1 | 홍콩법인 |
| ID1 | 인도네시아법인 |
| IN1 | 인도법인 |
| IT1 | 이탈리아법인 |
| JP1 | 일본법인 |
| JP2 | 일본디지털센터 |
| KR1 | 오스템임플란트 |
| KRA | ㈜대한치과교육개발원 |
| KRB | 오스템파마㈜ |
| KRC | 오스템글로벌㈜ |
| KRD | 오스템바스큘라㈜ |
| KRE | 오스템올소㈜ |
| KRF | 탑플란㈜ |
| KRG | 코잔㈜ |
| KRH | 오스템인테리어㈜ |
| KRI | ㈜메디칼소프트 |
| KRJ | 이베스트-어센트신기술투자조합제1호 |
| KZ1 | 카자흐스탄법인 |
| MN1 | 몽골법인 |
| MX1 | 멕시코법인 |
| MY1 | 말레이시아법인 |
| NL1 | 네덜란드법인 |
| NZ1 | 뉴질랜드법인 |
| PH1 | 필리핀법인 |
| PT1 | 포르투갈법인 |
| RU1 | 러시아법인 |
| SG1 | 싱가폴법인 |
| TH1 | 태국법인 |
| TR1 | 튀르키예법인 |
| TW1 | 대만법인 |
| UA1 | 우크라이나법인 |
| US1 | 미국법인 |
| UZ1 | 우즈베키스탄법인 |
| VN1 | 베트남법인 |
