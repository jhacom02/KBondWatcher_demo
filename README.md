# KBondWatcher

Windows에서 채권 채팅 호가를 감시하고, Excel PnL이 임계를 만족하면 확정 문장을 UI로 전송한다.

설계·로직: [docs/KBondWatcher_doc.md](docs/KBondWatcher_doc.md)  
에러 표: [docs/error_table.md](docs/error_table.md)  
Pilot smoke: [docs/pilot_smoke.md](docs/pilot_smoke.md)

## MODE

| MODE | 소스 | 전송 |
|------|------|------|
| 1 | KBond 분리창 (Profile `kbond_chat_title`) | 같은 창 |
| 2 | 동일 | 메모장 |
| 3 | FORESTBOND (UIA) | 메모장 |

## 설치 (개발)

```bat
cd <프로젝트>
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Trader UI (Pilot)

Excel·채팅을 연 뒤:

```bat
python main.py --serve
```

브라우저 `http://127.0.0.1:8765/` — Profile / Calibration / START / STOP.  
Watcher: Web START → `python main.py --run-profile` (frozen이면 `main.exe --run-profile`).

Admin (별도):

```bat
set KBOND_ADMIN_URL=http://127.0.0.1:8770
python main.py --serve-admin
```

```bat
pytest -q
```

## Trader Pilot 패키징

```bat
powershell -ExecutionPolicy Bypass -File build\build_nuitka.ps1
```

생성 폴더만 배포 (`.py` 없음, 시스템 Python 불필요). 체크리스트: `docs/pilot_smoke.md`.
