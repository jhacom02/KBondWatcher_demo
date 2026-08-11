# KBondWatcher

One-shot watcher: FORESTBOND Chrome (UIA) → Excel calc → threshold → KakaoTalk send.

## Install

```bat
cd C:\mycode\KBondWatcher
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
```

Edit `.env` (required keys, no code defaults for business values).

## Flow

1. Excel START → `pythonw main.py --config .env`
2. Read FORESTBOND chat via UIA (`CHROME_TITLE`)
3. Parse TARGET quote → write Excel input cell → auto-calc → read P&L
4. If `pnl >= PNL_THRESHOLD` → open KakaoTalk room → paste `MESSAGE_TEMPLATE` → Enter → exit

## Diagnose

```bat
python main.py --config .env --diagnose-chrome
python main.py --config .env --diagnose-kakao
python main.py --config .env --test-parser "25-10 23+"
python main.py --config .env --test-send
pytest -q
```

## Modules

| File | Role |
|------|------|
| `forestbond_reader.py` / `quote_parser.py` | message source (FORESTBOND Chrome) |
| `excel_bridge.py` | value input / calc |
| `trigger.py` | send trigger rule |
| `message_sender.py` | KakaoTalk send |
| `main.py` | orchestration |

## VBA

Import `vba/KBondWatcher.bas` into `sample/sample.xlsm`. Align Const paths and status cells with `.env`. Connect START/STOP to `StartKBondWatcher` / `StopKBondWatcher`.
