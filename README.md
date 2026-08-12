# KBondWatcher

One-shot watcher: source (MODE) → 5 Excel slots → PnL threshold → send (MODE).

## MODE

| MODE | Source | Send |
|------|--------|------|
| 1 | KBondMessenger (TElTree) | KBondMessenger |
| 2 | KBondMessenger (TElTree) | Notepad |
| 3 | FORESTBOND (UIA) | Notepad |

Source/send window identity and click ratios are fixed by `MODE` (not by `.env` `SOURCE_*` / `SEND_*` keys).

## Install

```bat
cd C:\mycode\KBondWatcher_kbond
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Edit `.env` (`MODE=1|2|3`). Set `EXCEL_WORKBOOK` to the full absolute path of the open workbook.

## Flow

1. Excel START → `pythonw main.py --config .env`
2. Load active slots from A/E; Looking For from E (±100) → G2 `OFFER` or `BID` only (all slots same direction)
3. Yield prefix: B6 → rows 19/25; B5 → rows 41/46/56
4. Chat match any slot instrument + BUY/SELL → write D → wait `CalculationState==xlDone` → read F(row+3)
5. Threshold hit → send flipped `{confirm_token} ㅎㅈ` → exit (one-shot)

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
| `source_reader.py` / `source_reader_kbond.py` / `source_reader_uia.py` / `eltree_reader.py` | MODE source factory + TElTree / UIA |
| `excel_bridge.py` | 5 slots + B5/B6 prefix + F2–J2 status |
| `trigger.py` | OFFER/BID threshold + side flip |
| `send_ui.py` | clipboard send (target from MODE) |
| `main.py` | orchestration |

## VBA

Import `vba/KBondWatcher.bas`. Status cells F2/G2/H2/I2/J2.
