# Market Report System

Google Spreadsheet(원본 Excel과 같은 날짜별 Report Sheet)에서 값을 읽고, 웹앱에서 Draft를 확인한 뒤 Excel/PDF를 만들고 Gmail로 발송합니다.

웹앱은 데이터 입력 도구가 아닙니다. Review / Report Generation / Email Sending만 담당합니다.

## Workflow

```
Google Spreadsheet의 해당 날짜 Sheet에서 값 입력/수정
→ Market Report System
→ Report Date 선택
→ Refresh Data
→ 해당 날짜 Sheet의 셀을 직접 읽기
→ Draft Preview
→ Create / Update Report
→ PDF Preview
→ Email Preview
→ Send Email
```

같은 날짜 Report는 언제든 Google Sheets 최신값으로 다시 생성할 수 있습니다. Create Report는 메일을 보내지 않습니다.

## 로컬 실행

```bat
run.bat
```

로컬 주소는 `http://localhost:8502` 입니다. (`8501`의 Bunker Document System은 사용하지 않습니다.)

## Render

기존 Render 웹 서비스가 GitHub `main`을 자동 배포합니다. Flask 앱 객체는 `server:app` 입니다.

- Build: `pip install -r requirements.txt && python -m playwright install --with-deps chromium`
- Start: `gunicorn server:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 180`
- Health: `GET /` (Sheets/Gmail/News는 페이지 로드 후 async)

환경변수 (값은 Render Dashboard에만 저장):

- `GOOGLE_SHEET_ID`
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_REFRESH_TOKEN` (`sheets_token.json`의 refresh_token)
- `GMAIL_CLIENT_ID`
- `GMAIL_CLIENT_SECRET`
- `GMAIL_REFRESH_TOKEN` (`token.json`의 Gmail refresh_token, Sheets 토큰과 다름)

## Google Spreadsheet 구조

- 날짜별 Report Sheet: `YY.MM.DD 보고자료` — 원본 Excel과 같은 셀 위치
- `Email Recipients`: Type | Name | Email | Active. Active=TRUE인 주소만 기본 TO/CC

로컬 `.env` 예시: `.env.example`  
Cloud secrets 예시: `.streamlit/secrets.toml.example`

## Gmail

OAuth 2.0, scope는 `https://www.googleapis.com/auth/gmail.send`만 요청합니다. Render에서는 `GMAIL_*` 환경변수를 사용합니다. Send Email을 눌렀을 때만 발송합니다.

## Excel / PDF

- PDF: HTML Preview를 Playwright/Chromium으로 A4 portrait 2페이지 (`Weekly Report_Bunkering_YY.MM.DD.pdf`)
- Excel: openpyxl 워크북 다운로드. Windows COM PDF는 사용하지 않습니다.
