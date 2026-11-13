# Formly 아티팩트 버전 (서버리스)

`doc-autofill-mvp v0.1.0`(FastAPI + openpyxl + LibreOffice)을 브라우저 단일 페이지로 옮긴 버전입니다.

| 원본 (서버) | 아티팩트 (브라우저) |
|---|---|
| openpyxl로 셀 쓰기 | JSZip + DOMParser로 시트 XML 직접 수정 (서식·병합·인쇄설정 보존) |
| LibreOffice `--convert-to pdf` | 시트를 HTML로 재현 → html2canvas → jsPDF (A4, 여백·방향 반영, 행 경계에서 페이지 분할) |
| SQLite + `data/` 폴더 | 아티팩트 `db` 개인 영역(`data/users/<id>/`)에 매핑 JSON + xlsx(base64 분할) 저장 |
| FileResponse 다운로드 | `downloads` capability |

- 소스: `src/engine.js`(엑셀 엔진), `src/app.js`(UI, 원본 web/app.js 포팅), `src/base.css`(원본 CSS), `src/extra.css`
- 빌드: `python3 build.py <샘플.xlsx>` → `index.html` (샘플 xlsx 내장). 출장비 정산 샘플은 실명·사번이 있어 저장소에 넣지 않습니다.
