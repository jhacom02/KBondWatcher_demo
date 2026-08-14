# KBondWatcher

Windows에서 채권 채팅 호가를 감시하고, Excel PnL이 임계를 만족하면 확정 문장을 UI로 전송한다.

설계·로직: [docs/KBondWatcher_doc.md](docs/KBondWatcher_doc.md)  
에러 표: [docs/error_table.md](docs/error_table.md)

## MODE

| MODE | 소스 | 전송 |
|------|------|------|
| 1 | KBond 분리창 (`KBOND_CHAT_TITLE`) | 같은 창 |
| 2 | 동일 | 메모장 |
| 3 | FORESTBOND (UIA) | 메모장 |

## 설치

```bat
cd <프로젝트>
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

`.env`에서 `MODE`, `EXCEL_WORKBOOK`(이미 열린 통합문서 경로), MODE 1·2면 `KBOND_CHAT_TITLE`, `SENT_AFTER`를 맞춘다. Excel에 `vba/KBondWatcher.bas`를 넣고 경로를 이 PC에 맞게 수정한다.

## 실행

Excel Start, 또는:

```bat
python main.py --config .env
python main.py --config .env --diagnose-source
python main.py --config .env --diagnose-send
python main.py --config .env --test-parser "25-10 23+"
python main.py --config .env --test-send
python main.py --config .env --perf-summary
pytest -q
```
