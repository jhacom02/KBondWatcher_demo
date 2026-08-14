# KBondWatcher

One-shot watcher: source (MODE) → `EXCEL_WATCH_CELL` slot → `EXCEL_PNL_THRESHOLD_CELL` → send (MODE).

## MODE

| MODE | Source | Send |
|------|--------|------|
| 1 | KBondMessenger (`KBOND_CHAT_TITLE` 분리창) | 같은 분리창 |
| 2 | KBondMessenger (`KBOND_CHAT_TITLE` 분리창) | Notepad |
| 3 | FORESTBOND (UIA) | Notepad |

Source process is fixed by `MODE`. MODE 1·2 chat room title is `.env` `KBOND_CHAT_TITLE`. Click ratios come from `.env` (`SEND_INPUT_X_M1`/`Y_M1` relative to the detached chat window, `SEND_INPUT_X_M23`/`Y_M23`).

## Install

```bat
cd C:\mycode\KBondWatcher
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Edit `.env` (`MODE=1|2|3`). Set `EXCEL_WORKBOOK` to the full absolute path of the open workbook.

Operational defaults: `PROCESS_EXISTING_ON_START=false` (skip lines already on screen at start; only watch newer lines). Set `true` only when replaying visible chat for parser tests. After a threshold send the process exits (one-shot).

## Flow

1. Excel START → `pythonw main.py --config .env`
2. `EXCEL_WATCH_CELL` (default D2) selects one allowlisted slot (`EXCEL_SLOT_ROWS`); signed qty on that row (negative→`BID`, positive→`OFFER`; abs=억 size) → G2 as `{instrument} / BID` or `{instrument} / OFFER`. Threshold from `EXCEL_PNL_THRESHOLD_CELL` (default E2).
3. Yield prefix: B6 → rows 19/25; B5 → rows 41/46/56 (cells/rows from `.env`)
4. Chat match → write that row's input col → wait `xlDone` and numeric PnL → read PnL at row+offset. Abort if `|pnl-threshold|` exceeds `EXCEL_PNL_SANITY_BAND`.
5. Threshold cell hit → send flipped `{confirm_token} ㅎㅈ` → exit (one-shot)

## Diagnose

```bat
python main.py --config .env --diagnose-source
python main.py --config .env --diagnose-send
python main.py --config .env --test-parser "25-10 23+"
python main.py --config .env --test-send
pytest -q
```

## Layout

| Path | Role |
|------|------|
| `main.py` | CLI / orchestration |
| `config/` | `.env` loader + MODE presets |
| `source/` | MODE source factory, RichEdit, UIA, quote parser |
| `send/` | click / paste / Enter UI send |
| `excel/` | 5 slots + B5/B6 prefix + F2–J2 status |
| `core/` | models, trigger, logger |
| `docs/` `logs/` `sample/` `tests/` `tools/` `vba/` | unchanged |

## VBA

Import `vba/KBondWatcher.bas`. Status cells F2/G2/H2/I2/J2. Paths point at `C:\mycode\KBondWatcher`.
