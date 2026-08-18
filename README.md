# KBondWatcher

Windows에서 채권 채팅 호가를 감시하고, Excel PnL이 임계를 만족하면 확정 문장을 UI로 전송한다.

설계·로직: [docs/KBondWatcher_doc.md](docs/KBondWatcher_doc.md)  
에러 표: [docs/error_table.md](docs/error_table.md)  
Pilot smoke: [docs/pilot_smoke.md](docs/pilot_smoke.md)

## MODE

| MODE | 소스 | 전송 |
|------|------|------|
| 1 | KBond 분리창 (`KBOND_CHAT_TITLE`) | 같은 창 |
| 2 | 동일 | 메모장 |
| 3 | FORESTBOND (UIA) | 메모장 |

## 설치 (개발)

```bat
cd <프로젝트>
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 1차 데모 UI (권장)

Excel·채팅을 연 뒤:

```bat
python main.py --serve
```

브라우저에서 `http://127.0.0.1:8765/` — Profile / Calibration / START / STOP.

Admin (별도, 선택):

```bat
set KBOND_ADMIN_URL=http://127.0.0.1:8770
python main.py --serve-admin
```

Watcher subprocess: `python main.py --run-profile` (Web START가 기동).

레거시 `.env` + VBA 경로도 당분간 동작한다:

```bat
python main.py --config .env
```

## 기타 CLI

```bat
python main.py --config .env --diagnose-source
python main.py --config .env --diagnose-send
python main.py --config .env --test-parser "25-10 23+"
python main.py --config .env --test-send
python main.py --config .env --perf-summary
pytest -q
```

## Trader Pilot 패키징

```bat
powershell -ExecutionPolicy Bypass -File build\build_nuitka.ps1
```

생성 폴더만 배포 (`.py` 없음, 시스템 Python 불필요). 세부 체크는 `docs/pilot_smoke.md`.
