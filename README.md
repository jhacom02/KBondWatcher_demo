# KBondWatcher

One-shot watcher: UIA/Win32 source (`SOURCE_*`, KBond Messenger) → Excel auto-calc → threshold → UI send (`SEND_*`, KBond).

## Install

```bat
cd C:\mycode\KBondWatcher
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Edit `.env`. Keep **K-Bond Messenger** running before send/read tests.

## Flow

1. Excel START → `pythonw main.py --config .env`
2. Source chat via `SOURCE_WINDOW_TITLE` / `SOURCE_PROCESS_NAME` (default: `K-Bond` / `KBondMessenger.exe`)
3. Parse TARGET quote → write Excel input → auto-calc → read P&L
4. If `pnl >= PNL_THRESHOLD` → KBond input click → paste `MESSAGE_TEMPLATE` → Enter → exit

## Diagnose

```bat
python main.py --config .env --diagnose-source
python main.py --config .env --diagnose-send
python main.py --config .env --test-parser "25-10 23+"
python main.py --config .env --test-send
pytest -q
```

## Modules

| File | Role |
|------|------|
| `source_reader.py` / `quote_parser.py` | source window + quote parse |
| `excel_bridge.py` | input / auto-calc read |
| `trigger.py` | threshold + message template |
| `send_ui.py` | generic UI send (`SEND_*`) |
| `main.py` | orchestration |

## VBA

Import `vba/KBondWatcher.bas`. Align Const paths with install dir. Wire START/STOP to `StartKBondWatcher` / `StopKBondWatcher`. Prefer `.venv\Scripts\pythonw.exe` in `PYTHONW_PATH`.
