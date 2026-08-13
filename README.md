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
cd C:\mycode\KBondWatcher
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Edit `.env` (`MODE=1|2|3`). Set `EXCEL_WORKBOOK` to the full absolute path of the open workbook.

## Flow

1. Excel START → `pythonw main.py --config .env`
2. Load active slots from instrument/qty cols (default A/E); Looking For from qty (`-100`→`BID`/사자, `+100`→`OFFER`/팔자) → G2 only (all slots same direction)
3. Yield prefix: B6 → rows 19/25; B5 → rows 41/46/56 (cells/rows from `.env`)
4. Chat match → write input col → wait `CalculationState==xlDone` → read PnL col at row+offset
5. Threshold hit → send flipped `{confirm_token} ㅎㅈ` → exit (one-shot)

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
| `source/` | MODE source factory, TElTree, UIA, quote parser |
| `send/` | click / paste / Enter UI send |
| `excel/` | 5 slots + B5/B6 prefix + F2–J2 status |
| `core/` | models, trigger, logger |
| `docs/` `logs/` `sample/` `tests/` `tools/` `vba/` | unchanged |

## VBA

Import `vba/KBondWatcher.bas`. Status cells F2/G2/H2/I2/J2. Paths point at `C:\mycode\KBondWatcher`.
