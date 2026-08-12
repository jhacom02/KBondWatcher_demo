# KBondWatcher — 프로젝트 흐름 및 로직 문서

프로젝트 경로: `C:\mycode\KBondWatcher`

---

## 1. 목적

Windows에서 다음을 **한 번(one-shot)** 수행한다.

1. `MODE`로 소스/전송 대상을 정하고 채팅을 읽어 신규 라인을 수집한다.
2. Excel 5슬롯(A/E)에서 종목·Looking For를 읽고, B5/B6에서 yield prefix를 정한다.
3. 매칭 슬롯의 D에 수익률을 쓰고 `CalculationState==xlDone` 후 F(합계) PnL을 읽는다.
4. 임계값을 만족하면 side flip 확정 메시지를 보내고 **종료**한다 (one-shot).
5. 미달이면 계속 감시. STOP 플래그면 `STOPPED`.

OCR·Selenium·카카오톡 연동은 사용하지 않는다. 소스는 MODE에 따라 TElTree 또는 FORESTBOND UIA, 전송은 `send_ui` 한 경로(대상은 MODE 프리셋)만 사용한다.

---

## 2. 전체 흐름

```
Excel VBA StartKBondWatcher
        │
        ▼
pythonw main.py --config .env
        │
        ├─ Config.load(.env) — MODE presets for source/send
        ├─ Excel: load 5 slots + B5/B6 prefixes
        ├─ F2=WATCHING, G2=OFFER|BID, J2=(HH:MM:SS) Start Successful
        ├─ create_source_reader(MODE) → watermark
        │
        └─ 폴링
              ├─ stop → STOPPED
              ├─ source lines (TElTree or UIA)
              ├─ match any slot (instrument + BUY/SELL)
              ├─ write D → xlDone → read F
              ├─ skip → Quote Skipped
              └─ trigger → flip send → SENT exit
```

트리거 성공 후 재감시하지 않는다. 전송 성공 시 즉시 종료한다(one-shot SENT exit).

---

## 3. 디렉터리·모듈 역할

| 경로 | 역할 |
|------|------|
| `main.py` | CLI·감시 루프 오케스트레이션 |
| `config.py` | `.env` 로드·검증 (`MODE` 프리셋 포함) |
| `models.py` | `AppStatus`, `Quote`, `TriggerResult`, `WatcherSession` |
| `eltree_reader.py` | KBondMessenger `TElTree` 탐색·원격 항목 텍스트 읽기 |
| `source_common.py` | watermark 공통 베이스 |
| `source_reader_kbond.py` | MODE 1/2 TElTree 소스 |
| `source_reader_uia.py` | MODE 3 FORESTBOND UIA 소스 |
| `source_reader.py` | `create_source_reader(cfg)` 팩토리 |
| `quote_parser.py` | 채팅 라인 → `Quote` (엄격 정규식) |
| `excel_bridge.py` | 실행 중 Excel COM: 입력/자동계산 후 P&L/상태셀 |
| `trigger.py` | Looking For별 PnL 임계·side flip·메시지 템플릿 |
| `send_ui.py` | 범용 UI 클릭·붙여넣기·Enter (TOPMOST로 전송 대상 확보, 대상은 MODE) |
| `logger.py` | `kbond_watcher` 로거 (파일 롤링 + 콘솔) |
| `vba/KBondWatcher.bas` | Excel START/STOP 매크로 |
| `tests/` | 파서·트리거·브릿지·센더·MODE 단위 테스트 |
| `requirements.txt` | 의존성 (`pywinauto` 포함, MODE 3) |
| `.env` | 운영 설정 (`MODE` + Excel/타이밍/템플릿) |

---

## 4. 실행 진입점 (`main.py`)

### 4.1 CLI

```text
python main.py --config .env
python main.py --config .env --diagnose-source
python main.py --config .env --diagnose-send
python main.py --config .env --test-parser "25-10 23+"
python main.py --config .env --test-send
```

| 옵션 | 동작 |
|------|------|
| (기본) | `run_watcher` |
| `--diagnose-source` | 소스 창·Text 컨트롤·메시지 라인 덤프 |
| `--diagnose-send` | MODE send 대상 HWND·클릭 좌표 진단 |
| `--test-parser LINE` | 한 줄 파싱 결과 출력 (실패 시 exit 1) |
| `--test-send` | 샘플 `Quote`로 템플릿 메시지 실제 전송 |

설정 로드 실패 시 exit `2`. 런타임 오류 시 exit `1`.

### 4.2 `run_watcher` 상세

1. `WatcherSession(status=STARTING)` 생성 (내부 상태; Excel에는 쓰지 않음).
2. `STOP_FLAG_PATH` 파일이 있으면 삭제 (`clear_stop_flag`).
3. `ExcelBridge.connect()` → `load_slots()` (B5/B6 prefix, 5 slots, 단일 Looking For) → `update_status(WATCHING, looking_for=OFFER|BID)`.
4. `create_source_reader(cfg)` → `find_source_window` → `ensure_target_window` → watermark. 소스/전송 창 실패 시 즉시 ERROR exit (재시도 없음).
5. 루프: 라인 읽기 실패(`SourceReaderError`)도 즉시 ERROR. 매칭 → `write_yield_read_pnl(D, F)` → evaluate → Quote Skipped 또는 SENT exit.

### 4.3 Excel 상태 행 (F2~J2)

| 셀 | 역할 | 허용값 |
|----|------|--------|
| F2 Status | 감시 상태 | `WATCHING` / `STOPPED` / `SENT` / `ERROR` |
| G2 Looking For | 탐색 호가 방향 | `BID` / `OFFER` |
| H2 Last Quote | 마지막 호가 | `{instrument} {raw_token}` |
| I2 Last PnL | 마지막 P&L | 숫자 |
| J2 Last Action | 마지막 동작 | 모두 `(HH:MM:SS) {action}` 형식. 예: `(16:15:35) Start Successful` / `(16:15:35) Quote Skipped` / `(16:15:35) Message Sent: ...` / `(16:15:35) Stopped` / `(16:15:35) Python Error: ...` |

| E41 | Looking For | 수집 side | 트리거 |
|-----|-------------|-----------|--------|
| `-100` | `OFFER` | BUY (사자) | `pnl >= PNL_THRESHOLD` |
| `+100` | `BID` | SELL (팔자) | `pnl <= -PNL_THRESHOLD` |

G2에는 `OFFER` 또는 `BID`만. 활성 슬롯 방향이 섞이면 시작 에러. 로그/파서는 BUY/SELL.

Quote Skipped는 호가/PnL 없이 `(HH:MM:SS) Quote Skipped`만 쓴다. 호가·PnL은 H2/I2에 남긴다.

---

## 5. 설정 (`config.py` / `.env`)

`Config.load(path)`:

- 파일이 없으면 `ConfigError`.
- `python-dotenv`로 로드한 뒤, 파일을 다시 읽어 key=value를 파싱한다 (주석·빈 줄 무시, 따옴표 trim).
- 값은 파일 우선, 없으면 환경변수. 필수 키 누락·형식 오류 시 `ConfigError`.
- `MODE`가 source/send 창 identity·click ratio를 고정한다. `.env`에 `SOURCE_*` / `SEND_PROCESS_NAME` / `SEND_WINDOW_TITLE` / `SEND_INPUT_*`가 있어도 무시된다.

### 5.0 MODE

| MODE | Source | Send |
|------|--------|------|
| 1 | KBondMessenger (`K-Bond` / TElTree) | KBondMessenger (click 0.825, 0.940) |
| 2 | KBondMessenger (TElTree) | Notepad (`메모장`, click 0.5, 0.5) |
| 3 | FORESTBOND (UIA) | Notepad |

### 5.1 감시

| 키 | 설명 |
|----|------|
| `MODE` | `1` / `2` / `3` (필수) |
| `POLL_INTERVAL_MS` | 폴링 간격 (≥ 100) |
| `PROCESS_EXISTING_ON_START` | 시작 시 기존 라인 처리 여부 |

`TARGET` / `YIELD_PREFIX` / `SOURCE_*` 없음. 종목·prefix는 Excel 슬롯·B5/B6. 소스 창은 MODE 프리셋.

### 5.2 Excel

| 키 | 설명 |
|----|------|
| `EXCEL_WORKBOOK` | 열린 통합문서 절대경로 |
| `EXCEL_SHEET` | 시트명 |
| `EXCEL_SLOT_ROWS` | `19,25,41,46,56` |
| `EXCEL_ROWS_10Y` | `19,25` (prefix B6) |
| `EXCEL_ROWS_3Y` | `41,46,56` (prefix B5) |
| `EXCEL_PREFIX_3Y_CELL` | B5 |
| `EXCEL_PREFIX_10Y_CELL` | B6 |
| `EXCEL_STATUS_CELL` | F2 |
| `EXCEL_LOOKING_FOR_CELL` | G2 (`OFFER`/`BID`만) |
| `EXCEL_LAST_QUOTE_CELL` | H2 |
| `EXCEL_LAST_PNL_CELL` | I2 |
| `EXCEL_LAST_ACTION_CELL` | J2 |

슬롯: A{R} 종목(`국고` 제거), E{R} ±100, D{R} 수익률 입력, F{R+3} PnL.
| `PNL_THRESHOLD` | 트리거 기준 (`pnl >= threshold`) |

### 5.3 전송 (대상은 MODE, 타이밍만 `.env`)

| 키 | 설명 |
|----|------|
| `MESSAGE_TEMPLATE` | `str.format` 템플릿 |
| `SEND_FOREGROUND_RETRY_PAUSE_SECONDS` | 포그라운드 재시도 대기 |
| `SEND_ACTIVATE_SHOW_PAUSE_SECONDS` | Show/Restore 후 대기 |
| `SEND_AFTER_ACTIVATE_PAUSE_SECONDS` | 활성화 후 대기 |
| `SEND_INPUT_CLICK_PAUSE_SECONDS` | 입력란 클릭 후 대기 |
| `SEND_PASTE_PAUSE_SECONDS` | 붙여넣기 후 대기 |
| `SEND_SEND_PAUSE_SECONDS` | Enter 후 대기 |

프로세스/제목/클릭 비율은 MODE 프리셋. `send_ui`는 동일 경로를 사용한다.

### 5.4 기타

| 키 | 설명 |
|----|------|
| `STOP_FLAG_PATH` | STOP용 플래그 파일 경로 |
| `LOG_LEVEL` | 로그 레벨 |
| `LOG_PATH` | 로그 파일 (상대면 `.env` 디렉터리 기준) |

### 5.5 `MESSAGE_TEMPLATE` 플레이스홀더

`str.format`에 넘기는 키는 다음뿐이다.

| 키 | 의미 |
|----|------|
| `instrument` | TARGET |
| `raw_token` | 호가 토큰 (예: `23+`) |
| `yield_value` | Excel에 쓴 수익률 |
| `side` | `BUY` / `SELL` |
| `pnl` | 읽은 P&L (float) |
| `raw_line` | 원문 라인 |

예: `{instrument} {raw_token} 확정`

---

## 6. 데이터 모델 (`models.py`)

### `AppStatus`

문자열 enum. 로깅·세션용으로 여러 값이 있다. Excel 셀 갱신은 `main.py`에서 `WATCHING` / `SENT` / `STOPPED` / `ERROR` 4종만 사용한다.

### `Quote`

| 필드 | 의미 |
|------|------|
| `instrument` | TARGET |
| `raw_line` | 원문 라인 (strip) |
| `raw_token` | TARGET 이후 호가 토큰 (예: `23+`) |
| `yield_value` | Excel에 넣을 수익률 |
| `side` | `BUY` / `SELL` |
| `timestamp`, `sender` | `보낸이 (HH:MM[:SS]) :` 메타가 있으면 파싱 |

`fingerprint`: `sha1(raw_line)` — 세션 내 중복 처리 방지.

### `WatcherSession`

- `processed_fingerprints`: 이미 처리한 quote fingerprint set
- `status`: 현재 내부 상태

### `TriggerResult`

`triggered`, `reason`, `pnl`, `quote`.

---

## 7. 소스 읽기 (`create_source_reader` / ElTree / UIA)

MODE 1·2: KBond Messenger `TElTree` 원격 읽기 (`source_reader_kbond.py` + `eltree_reader.py`).  
MODE 3: FORESTBOND UIA Text 수집 (`source_reader_uia.py`, KBondWatcher 로직 이식).  
공통 watermark는 `source_common.BaseSourceReader`.

### 7.1 TElTree (MODE 1·2)

1. MODE 프리셋 프로세스/제목으로 메인 창 HWND 선택(면적 최대).
2. 자식 중 class가 `TElTree`이고, 부모 대비 중심 X 비율 ≥ 0.55 인 것 중 면적 최대를 채팅 본문으로 선택.
3. `OpenProcess` 후 `TVM_GETCOUNT` / `TVM_GETNEXTITEM` / `TVM_GETITEM(A|W)`로 항목 순회.
4. 메시지 ID·구조 레이아웃은 `eltree_reader.py` 상단 상수.

폴링 중 메신저를 포그라운드로 가져오지 않는다. 클립보드는 `send_ui`에서만 사용.

### 7.2 창 찾기

1. 프로세스 미실행·제목 불일치·TElTree/UIA 미발견 시 `SourceReaderError` → 즉시 ERROR 종료 (폴링 재시도 없음).
2. `create_source_reader(cfg)`가 MODE에 맞는 리더를 반환.

### 7.3 텍스트 수집

- MODE 1·2: `read_eltree_lines` → strip·빈 줄·중복 제거.
- MODE 3: UIA Document/Text descendants → strip·중복 제거.

### 7.4 신규 라인·워터마크

- 라인 fingerprint = SHA1(UTF-8 텍스트).
- `initialize_watermark(process_existing_on_start)`:
  - `True`: watermark를 비움 → 현재 화면에 있는 것도 “신규”로 처리 가능.
  - `False`: 현재 라인을 watermark에 넣어 **이후 새 것만** 반환.
- `get_new_message_lines`: 미초기화면 watermark 규칙 적용 후, 이후에는 새 fingerprint만 반환.

세션 쪽(`WatcherSession.processed_fingerprints`)은 파싱 성공 호가용이며 ElTree watermark와 역할이 다르다.

### 7.5 진단 (`diagnose`)

`--diagnose-source`:

- 메인 창 제목·HWND·process
- tree class / HWND / rect
- `TVM_GETCOUNT` item count
- 메시지 라인 수와 샘플(최대 N줄)
- 읽기 실패 시 `read error:` 한 줄

### 7.6 한계

| 한계 | 영향 |
|------|------|
| ElTree 메시지/구조 상수 불일치 | count 0·빈 텍스트·드물게 대상 프로세스 오류 → 상수 튜닝 필요 |
| 가상화·오프스크린 | 트리에 없는 줄 누락 가능 |
| 창 제목·프로세스명 변경 | 매칭 실패 |
| 좌측 룸 리스트도 TElTree | 중심 X≥0.55·면적으로 본문 선택; 레이아웃 변경 시 오선택 가능 |
| `OpenProcess` 거부 | 권한/백신 → 읽기 실패 |

운영 전 로그인 PC에서 채팅방 연 뒤 `--diagnose-source`로 호가 라인을 확인한다.

---

## 8. 호가 파싱 (`quote_parser.py`)

### TARGET 경계

`build_target_pattern`: `TARGET`를 escape한 뒤, 앞뒤가 숫자/`-`가 아니어야 매칭.  
예: `25-11`은 `125-11`, `25-110`과 매칭되지 않음.

### 메타 (선택)

`보낸이 (HH:MM[:SS]) : body` 형태면 sender/timestamp를 뽑고, 파싱은 원문 전체에서 TARGET을 찾는다.

### TARGET 이후 토큰 (필수, fullmatch)

```text
^\s+(?P<price>\d{2,3})\s*(?P<side>[+-]|사자|팔자)\s*$
```

수락 예: `735+`, `735 +`, `735사자`, `735 사자`, `23-`.  
거부 예: 뒤에 `40억`, `ㅎㅈ`, `자투리` 등 추가 텍스트, 비호가 문장.

### 수익률 변환 (`digits_to_yield`)

- 2자리: `YIELD_PREFIX + n/100` (예: prefix 3, `23` → `3.23`)
- 3자리: `YIELD_PREFIX + n/1000` (예: prefix 3, `735` → `3.735`)

### 사이드

`+` / `사자` → `BUY`, `-` / `팔자` → `SELL`.  
런타임 `required_side`가 `BUY`/`SELL`이면 다른 쪽은 `None` 반환 (E41 Looking For에서 유도).

---

## 9. Excel (`excel_bridge.py`)

### 연결

- 이미 실행 중인 `Excel.Application`에 `GetActiveObject`로 붙는다. Excel을 새로 띄우지 않는다.
- 워크북: `EXCEL_WORKBOOK` 절대경로와 `wb.FullName` 일치 우선, 아니면 `wb.Name`이 파일명과 일치, 아니면 name endswith.
- 시트: `EXCEL_SHEET`가 있으면 해당 시트, 없으면 ActiveSheet.

### 계산 파이프라인

`write_yield_read_pnl(yield_value)`:

1. 입력 셀에 float 기록 (`write_yield`)
2. `Application.CalculationState == xlDone(0)`까지 폴링 (타임아웃 시 `ExcelBridgeError`)
3. P&L 셀을 float로 읽어 반환 (`read_pnl`)

별도 Python PnL 공식 없음. `read_qty()`로 E41 수량을 읽는다.

### 상태 갱신

`update_status`는 STATUS는 항상 쓰고, 인자로 준 looking_for / last_quote / last_pnl / last_action만 선택 갱신.  
상태 쓰기 실패는 경고 로그만 (감시 루프를 깨지 않음).

---

## 10. 트리거·메시지 (`trigger.py`)

| Looking For | 트리거 |
|-------------|--------|
| `OFFER` | `pnl >= PNL_THRESHOLD` |
| `BID` | `pnl <= -PNL_THRESHOLD` |

`format_message`: `{confirm_token}` = `raw_token` side flip (`715+`→`715-`).  
예 템플릿: `{instrument} {confirm_token} ㅎㅈ`.

---

## 11. 범용 UI 전송 (`send_ui.py`)

카카오톡 API·방 검색·채팅 탭 클릭은 **없다**.  
MODE 프리셋의 프로세스/제목 창에 붙여넣기 후 Enter한다 (MODE 1: K-Bond, MODE 2·3: Notepad).

대상 프로세스가 **이미 실행 중**이어야 한다. 없으면 `SendError`로 즉시 실패한다 (자동 실행 없음).

### `send_text(text, cfg)`

1. `ensure_target_window`: 프로세스·제목으로 HWND 선택
2. `activate_window` 후 잠시 `HWND_TOPMOST`로 z-order 확보
3. 입력 비율 클릭 → 포그라운드·클릭 지점이 타깃 앱인지 확인 (아니면 `SendError`, paste 안 함)
4. 클립보드 복사 → Ctrl+V → Enter → TOPMOST 해제
5. 로그 `MESSAGE_SENT`

클릭 좌표는 창 `GetWindowRect` 기준 상대 비율이다. UI 해상도·창 크기가 바뀌면 `.env`의 X/Y·pause를 재조정한다.

### 진단

`--diagnose-send` → 프로세스 실행 여부, PID, HWND, rect, 클릭 좌표를 문자열로 출력한다.

---

## 12. Excel VBA (`vba/KBondWatcher.bas`)

모듈명: `KBondWatcher`.

### `StartKBondWatcher`

1. F2/G2/H2/I2/J2를 빈 값으로 초기화 (STARTING/launched 문자열을 쓰지 않음).
2. `WScript.Shell`로 숨김(`0`) 실행:

```text
pythonw.exe "<MAIN_PATH>" --config "<CONFIG_PATH>"
```

`CurrentDirectory` = `PROJECT_DIR`.

3. VBA 자체 실패 시만 F2=`ERROR`, J2=`VBA Error: ...`.

### `StopKBondWatcher`

1. `STOP_FLAG_PATH` 상위 폴더가 없으면 생성.
2. 플래그 파일에 `stop` 기록.
3. **STATUS 셀은 건드리지 않음** — Python이 플래그를 보고 `STOPPED` / `Stopped`를 기록.
4. VBA 실패 시만 ERROR 셀 기록.

### 상수 (코드에 박혀 있음 — 배포 경로와 반드시 맞출 것)

| Const | 값 |
|-------|-----|
| `PYTHONW_PATH` | `pythonw.exe` |
| `PROJECT_DIR` | `C:\mycode\KBondWatcher_kbond` |
| `MAIN_PATH` | `C:\mycode\KBondWatcher_kbond\main.py` |
| `CONFIG_PATH` | `C:\mycode\KBondWatcher_kbond\.env` |
| `STOP_FLAG_PATH` | `C:\temp\kbond_watcher.stop` |
| 상태 셀 | F2 Status / G2 Looking For / H2 / I2 / J2 |

운영 시 VBA Const와 `.env`의 `STOP_FLAG_PATH`·Excel 셀 키가 일치하는지 확인한 뒤 워크북에 모듈을 재임포트하고, 버튼에 `StartKBondWatcher` / `StopKBondWatcher`를 연결한다.  
가능하면 `PYTHONW_PATH`를 `.venv\Scripts\pythonw.exe` 절대 경로로 두는 것을 권장한다.

---

## 13. 로깅 (`logger.py`)

- 로거 이름: `kbond_watcher`
- `setup_logger`: 부모 디렉터리 생성, RotatingFileHandler(2MB×5) + StreamHandler
- 포맷: `%(asctime)s %(levelname)s %(message)s`

감시 중 주요 로그 키워드: `WATCHING`, `QUOTE_FOUND`, `NO_TRIGGER`, `TRIGGERED`, `SENT`, `STOPPED`, `ERROR`, `EXIT`, `EXCEL_*`, `MESSAGE_SENT`.

---

## 14. 운영 전제 조건

1. Windows + Excel이 대상 워크북을 **연 상태**, 계산 옵션은 Automatic.
2. MODE에 맞는 소스 창이 열려 있음 (1·2: K-Bond, 3: FORESTBOND).
3. MODE에 맞는 전송 대상 앱이 실행 중 (1: K-Bond, 2·3: Notepad).
4. `.env`에 `MODE` 및 Excel/타이밍 키 완비, VBA 경로·스톱 플래그·상태 셀 일치.
5. `pip install -r requirements.txt` (pywin32, psutil, python-dotenv, pyautogui, pyperclip, pywinauto, pytest).

의존성 설치·진단:

```bat
cd C:\mycode\KBondWatcher_kbond
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py --config .env --diagnose-source
python main.py --config .env --diagnose-send
pytest -q
```

---

## 15. 테스트

| 파일 | 내용 |
|------|------|
| `tests/test_quote_parser.py` | 수락/거부 라인, side 필터 |
| `tests/test_trigger.py` | Looking For 트리거, side flip, confirm_token |
| `tests/test_excel_bridge.py` | 브릿지 유틸·워크북 경로 매칭 |
| `tests/test_send_ui.py` | 좌표·창 매칭 등 센더 단위 |
| `tests/test_eltree_reader.py` | TElTree ROI 선택·라인 정규화 |
| `tests/test_config_mode.py` | MODE 프리셋·팩토리·잘못된 MODE |

UI/실기 Excel·KBond 전송은 자동화 테스트 범위 밖이며, CLI `--diagnose-source` / `--diagnose-send` / `--test-send`로 확인한다.

---

## 16. 장애 시 확인 순서

1. Excel F2: `ERROR`면 J2(`Python Error:` / `VBA Error:`) 요약 + `LOG_PATH` 로그.
2. F2가 계속 비어 있음: Python이 안 붙음 → VBA 경로·`pythonw`·`.env` 위치.
3. `WATCHING`인데 호가 미반응: G2 Looking For·E41 수량, `--diagnose-source`, `--test-parser`.
4. 계산만 되고 전송 없음: I2 P&L vs Looking For별 임계값, CalculationState 대기.
5. 전송 실패: 대상 프로세스 실행 여부, `--diagnose-send`, 클릭 비율·타이밍.

---

## 17. 설계상 고정된 동작 요약

- **One-shot**: 전송 성공(`SENT`) 또는 STOP/ERROR로 종료.
- **Fail-fast**: 소스/전송 창 미발견·읽기 실패·전송 실패는 재시도 없이 즉시 `ERROR`.
- **설정 외부화**: 비즈니스·좌표·타이밍은 `.env`만. `REQUIRED_SIDE` 없음(E41 유도).
- **상태셀**: F2 Status 4종, G2 Looking For(`BID`/`OFFER`), H2/I2/J2.
- **소스**: MODE별 TElTree 또는 FORESTBOND UIA.
- **계산**: `write_yield` → `CalculationState==xlDone` 대기 → `read_pnl`.
- **확정 메시지**: side flip (`confirm_token`) 후 MODE send 대상 전송.
- **워크북**: 프로젝트 밖 절대경로 `EXCEL_WORKBOOK` 지원.
