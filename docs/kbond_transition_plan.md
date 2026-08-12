---
name: KBond source and send switch
overview: 현재 SOURCE_*/SEND_* 골격 위에서, KBond PC에서 진단→.env→필요 시 source_reader만 최소 수정해 읽기·전송을 KBond로 전환한다. Excel/파서/트리거/VBA 골격은 유지한다.
todos:
  - id: env-send-kbond
    content: ".env SEND_* to KBond process/title/input 0.825/0.940; diagnose-send + test-send"
    status: pending
  - id: env-source-kbond
    content: ".env SOURCE_WINDOW_TITLE to KBond; diagnose-source; confirm quote lines in dump"
    status: pending
  - id: source-reader-fixup
    content: "Only if diagnose fails: add SOURCE_PROCESS_NAME and/or UIA filter; keep SourceReader API"
    status: pending
  - id: parser-kbond-lines
    content: "Add KBond real-line accept/reject tests; adjust quote_parser only if needed"
    status: pending
  - id: e2e-docs
    content: "E2E watcher SENT/STOPPED; update .env samples and KBondWatcher_doc/README for KBond defaults"
    status: pending
isProject: false
---

# Build plan: switch source + send to KBond

Agent instructions: implement on a Windows PC **with KBond running**. Do not invent Kakao paths. Do not add code comments. Prefer `.env` changes; change Python only when diagnose proves UIA/title matching is insufficient. Keep module names `source_reader.py` / `send_ui.py` and config prefixes `SOURCE_*` / `SEND_*`.

## Current baseline (already in repo)

```text
Excel VBA → main.py
  → SourceReader(SOURCE_WINDOW_TITLE) UIA Text + watermark
  → quote_parser → excel_bridge.write_yield_read_pnl
  → trigger → send_ui.send_text (process+title, click ratio, Ctrl+V, Enter)
```

| Area | File | Status |
|------|------|--------|
| Source | [`source_reader.py`](../source_reader.py) | Title substring UIA; no process filter yet |
| Send | [`send_ui.py`](../send_ui.py) | Generic click-paste-Enter; ready for KBond via `.env` |
| Config | [`config.py`](../config.py), [`.env`](../.env) | `SOURCE_WINDOW_TITLE`, `SEND_*` |
| CLI | [`main.py`](../main.py) | `--diagnose-source`, `--diagnose-send`, `--test-send` |
| Excel / trigger / VBA | unchanged | Keep STATUS 4종 + J2 phrases |

Do **not** reintroduce `KAKAO_*`, `CHROME_TITLE`, `forestbond_reader`, `message_sender`.

## Target behavior

- **Source**: watch KBond chat panel text via UIA (same `SourceReader` API).
- **Send**: click KBond input (red-box region on main window) → paste `MESSAGE_TEMPLATE` → Enter → `SENT` → exit.
- One-shot watcher semantics unchanged.

## KBond UI facts (for matching)

- Multi-panel messenger; use **main top-level window** for send clicks (not child HWND first).
- Send input (red box) ≈ main-window ratios **X 0.72–0.93, Y 0.92–0.96** → click center **`SEND_INPUT_X=0.825`**, **`SEND_INPUT_Y=0.940`**. Keep `SEND_INPUT_X ≤ 0.90` (avoid participant list).
- Source text: message list above input; ignore participant names column if mixed into UIA dump (parser should reject non-quotes).

## Implementation order (mandatory)

### Step 1 — Send (Track B): `.env` only first

Update [`.env`](../.env) (measure real process/title on the KBond PC; placeholders below):

```env
SEND_PROCESS_NAME=KBond.exe
SEND_WINDOW_TITLE=KBond
SEND_INPUT_X=0.825
SEND_INPUT_Y=0.940
```

Keep existing `SEND_*_PAUSE_SECONDS` unless timing fails.

Verify:

```bat
python main.py --config .env --diagnose-send
python main.py --config .env --test-send
```

Pass criteria: diagnose shows hwnd + click_point; test-send pastes into the **input box** (not participant list / wrong panel). If click misses, remeasure ratios from `GetWindowRect` of the matched hwnd and update `.env` only.

Code change for send: **none** unless process name or window matching is wrong (then fix matching in `send_ui.py` minimally).

### Step 2 — Source (Track A): `.env` then diagnose

Set:

```env
SOURCE_WINDOW_TITLE=<substring of KBond main or chat window title>
```

Verify:

```bat
python main.py --config .env --diagnose-source
```

Pass criteria: dump includes chat lines that look like trade quotes for `TARGET`.

### Step 3 — Source fixup only if Step 2 fails

If window not found or Text list empty/wrong:

1. Add optional **`SOURCE_PROCESS_NAME`** (mirror send): require in `config.py`, filter UIA/top-level windows by PID set like `send_ui.find_target_window`, then title substring. Wire `SourceReader` + `.env`.
2. If Text still empty: try other UIA `control_type`s (`ListItem`, etc.) inside `SourceReader._collect_text_controls` without changing `main.py` loop API (`get_new_message_lines` / watermark).
3. If multiple panels pollute lines: filter controls whose rectangle intersects ROI ≈ main-window **X 0.70–0.93, Y 0.08–0.90** (message list). Prefer config ratios `SOURCE_ROI_*` only if hardcoding ratios is unavoidable—prefer env keys consistent with `SOURCE_*`.

Do **not** rename modules to `kbond_*` unless necessary; keep `source` / `send` naming.

### Step 4 — Parser

Capture real KBond lines from diagnose dump into [`tests/test_quote_parser.py`](../tests/test_quote_parser.py).

- If lines already match `TARGET` + `_PRICE_SIDE` fullmatch → no parser change.
- If trailing junk / different side markers → extend parser + tests only as needed.
- Keep `MESSAGE_TEMPLATE` using `{instrument} {raw_token}` unless product asks otherwise.

### Step 5 — E2E + docs

1. Excel workbook open, auto-calc on, VBA paths point at this install.
2. Run watcher (Excel START or `python main.py --config .env`).
3. Confirm G2/J2: `WATCHING` / `Start Successful` → quote path → `Message Sent: ...` + `SENT` on trigger, or `Stopped` on stop flag.
4. Update [`docs/KBondWatcher_doc.md`](KBondWatcher_doc.md) and [`README.md`](../README.md) defaults to KBond `SOURCE_*` / `SEND_*` examples (remove FORESTBOND/Notepad as primary).
5. `pytest -q` must pass.

## Out of scope

- Auto-launch KBond / auto-login
- OCR
- KakaoTalk
- Changing Excel STATUS/J2 contract
- Rewriting `main.py` orchestration

## Definition of Done

- [ ] `--diagnose-send` + `--test-send` hit KBond input
- [ ] `--diagnose-source` shows usable quote lines
- [ ] Watcher one-shot send on threshold; STOP works
- [ ] `.env` documents KBond values; docs match
- [ ] No legacy Kakao/Chrome config keys
- [ ] `pytest -q` green
