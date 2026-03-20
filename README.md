# Osstem-AI-3

Claude AI 기반 풀스택 챗봇 웹 애플리케이션입니다.

## 기술 스택

| 영역 | 기술 |
|------|------|
| Backend | Python, FastAPI, Anthropic SDK |
| Frontend | React 18, TypeScript, Vite, Tailwind CSS |

## 주요 기능

- **멀티턴 대화** — 대화 히스토리를 유지하며 문맥에 맞는 응답
- **스트리밍 응답** — Claude API의 스트리밍을 활용한 실시간 타이핑 효과
- **다크모드 UI** — 깔끔한 채팅 인터페이스

## 시작하기

### 1. Backend 설정

```bash
cd backend
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt

cp .env.example .env
# .env 파일에 ANTHROPIC_API_KEY 입력

uvicorn main:app --reload
```

### 2. Frontend 설정

```bash
cd frontend
npm install
npm run dev
```

브라우저에서 `http://localhost:5173` 접속

## 프로젝트 구조

```
Osstem-AI-3/
├── backend/
│   ├── main.py          # FastAPI 앱 진입점
│   ├── routes/
│   │   └── chat.py      # 채팅 API 엔드포인트
│   ├── requirements.txt
│   └── .env.example
└── frontend/
    ├── src/
    │   ├── App.tsx
    │   ├── api/chat.ts       # API 통신 함수
    │   └── components/
    │       ├── ChatWindow.tsx
    │       ├── MessageBubble.tsx
    │       └── InputBar.tsx
    └── package.json
```

## API 엔드포인트

| Method | Path | 설명 |
|--------|------|------|
| POST | `/api/chat` | 일반 응답 |
| POST | `/api/chat/stream` | SSE 스트리밍 응답 |
