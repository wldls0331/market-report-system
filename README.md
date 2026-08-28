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

## Streamlit Community Cloud

Entrypoint: `app.py`

1. GitHub에 이 저장소를 push (인증 파일은 올리지 않음)
2. [share.streamlit.io](https://share.streamlit.io)에서 앱 생성, Main file path = `app.py`
3. Settings → Secrets에 `.streamlit/secrets.toml.example` 구조를 붙여넣고 값을 채움

Windows Excel COM은 Cloud에서 사용할 수 없습니다. Cloud에서는 Google Sheets 값으로 3-page PDF를 생성합니다.

## Google Spreadsheet 구조

- 날짜별 Report Sheet: `YY.MM.DD 보고자료` — 원본 Excel과 같은 셀 위치
- `Email Recipients`: Type | Name | Email | Active. Active=TRUE인 주소만 기본 TO/CC

로컬 `.env` 예시: `.env.example`  
Cloud secrets 예시: `.streamlit/secrets.toml.example`

## Gmail

OAuth 2.0, scope는 `https://www.googleapis.com/auth/gmail.send`만 요청합니다. Cloud에서는 refresh_token을 Streamlit Secrets에서 읽습니다.

## Excel / PDF

- 로컬 Windows: Microsoft Excel COM으로 차트 포함 3-page PDF
- Linux / Streamlit Cloud: Excel COM 없이 Google Sheets 값으로 3-page PDF

- PDF: `Weekly Report_Bunkering_YY.MM.DD.pdf`
