# KBondWatcher — 프로젝트 흐름 및 로직 문서

프로젝트 경로: `C:\mycode\KBondWatcher`

---

## 1. 목적

Windows에서 다음을 **한 번(one-shot)** 수행한다.

1. `SOURCE_WINDOW_TITLE` / `SOURCE_PROCESS_NAME`에 맞는 KBond 창을 읽어 신규 텍스트 라인을 수집한다.
2. 설정된 종목(`TARGET`) 호가만 파싱한다.
3. Excel에 수익률을 쓰고, 워크북 자동계산 결과를 기다린 뒤 P&L을 읽는다 (`Worksheet.Calculate` 호출 없음).
4. P&L이 임계값 이상이면 `SEND_*`로 지정한 KBond 입력란에 메시지를 보내고 **프로세스를 종료**한다.
5. 임계값 미달이면 계속 감시한다. Excel STOP(스톱 플래그)이면 `STOPPED`로 종료한다.

OCR·Selenium·카카오톡 연동은 사용하지 않는다. 메시지 전송은 범용 UI 자동화(`SEND_*`)만 사용한다.

---

## 2. 전체 흐름

```
Excel VBA StartKBondWatcher
        │
        ▼
pythonw main.py --config .env
        │
        ├─ Config.load(.env)
        ├─ Excel COM 연결 → G2=WATCHING, J2=Start Successful
        ├─ SourceReader(source_window_title=...) → find_source_window()
        ├─ initialize_watermark(PROCESS_EXISTING_ON_START)
        │
        └─ 폴링 루프 (POLL_INTERVAL_MS)
              │
              ├─ stop 파일 있으면 → G2=STOPPED, J2=Stopped → exit 0
              ├─ UIA로 신규 채팅 라인 수집
              ├─ quote_parser로 TARGET 호가 파싱 (실패/중복면 skip)
              ├─ Excel: write_yield_read_pnl (입력 쓰기 → 자동계산 → P&L 읽기)
              ├─ pnl < threshold → G2=WATCHING, H2/I2/J2 갱신 → 계속
              └─ pnl >= threshold → send_ui.send_text → G2=SENT → exit 0
```

트리거 성공 후 재감시하지 않는다. 전송 성공 시 즉시 종료한다(one-shot SENT exit).

---

## 3. 디렉터리·모듈 역할

| 경로 | 역할 |
|------|------|
| `main.py` | CLI·감시 루프 오케스트레이션 |
| `config.py` | `.env` 로드·검증 (`Config` dataclass) |
| `models.py` | `AppStatus`, `Quote`, `TriggerResult`, `WatcherSession` |
| `source_reader.py` | UIA 텍스트 수집·신규 라인 감지 (`SOURCE_WINDOW_TITLE`) |
| `quote_parser.py` | 채팅 라인 → `Quote` (엄격 정규식) |
| `excel_bridge.py` | 실행 중 Excel COM: 입력/자동계산 후 P&L/상태셀 |
| `trigger.py` | `pnl >= threshold` 판정, 메시지 템플릿 포맷 |
| `send_ui.py` | 범용 UI 창 활성화·클릭·붙여넣기·Enter (`SEND_*`) |
| `logger.py` | `kbond_watcher` 로거 (파일 롤링 + 콘솔) |
| `vba/KBondWatcher.bas` | Excel START/STOP 매크로 |
| `tests/` | 파서·트리거·브릿지·센더 단위 테스트 |
| `requirements.txt` | 의존성 |
| `.env` | 운영 설정 (비즈니스 기본값 코드에 없음) |

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
| `--diagnose-send` | `SEND_PROCESS_NAME` / 창 HWND·클릭 좌표 진단 |
| `--test-parser LINE` | 한 줄 파싱 결과 출력 (실패 시 exit 1) |
| `--test-send` | 샘플 `Quote`로 템플릿 메시지 실제 전송 |

설정 로드 실패 시 exit `2`. 런타임 오류 시 exit `1`.

### 4.2 `run_watcher` 상세

1. `WatcherSession(status=STARTING)` 생성 (내부 상태; Excel에는 쓰지 않음).
2. `STOP_FLAG_PATH` 파일이 있으면 삭제 (`clear_stop_flag`).
3. `ExcelBridge` 생성 후 `connect()` → `update_status(WATCHING, last_action="Start Successful")`.
4. `SourceReader(source_window_title=cfg.source_window_title)` → `find_source_window()` → `initialize_watermark(PROCESS_EXISTING_ON_START)`.
5. 무한 루프:
   - stop 파일이 있으면 `STOPPED` / `Stopped` 기록, 플래그 삭제, return `0`.
   - `get_new_message_lines(...)` 실패(`SourceReaderError`) 시 경고 로그만 남기고 sleep 후 continue (프로세스는 유지).
   - 각 신규 라인에 대해 `parse_quote_line` → `None`이면 skip.
   - `quote.fingerprint`(raw_line SHA1)가 세션 set에 있으면 skip, 없으면 추가.
   - 로그 `QUOTE_FOUND` → `excel.write_yield_read_pnl` → `evaluate`.
   - **미트리거**: 로그 `NO_TRIGGER`, Excel은 `WATCHING` + last_quote / last_pnl / `Quote Passed: {instrument} {raw_token} (pnl={#,##0})`.
   - **트리거**: `format_message` → `send_ui.send_text(text, cfg)` → Excel `SENT` + `Message Sent: {text}` → return `0`.
6. 예외(`ConfigError`, `ExcelBridgeError`, `SourceReaderError`, `SendError` 및 기타): Excel이 연결되어 있으면 `ERROR` + `Python Error: ...`(최대 200자), return `1`.
7. `finally`: `excel.close()` (COM `CoUninitialize`).

### 4.3 Excel에 쓰는 STATUS (트레이더용 4종)

내부 `AppStatus` enum에는 `QUOTE_FOUND`, `CALCULATING`, `SENDING` 등이 있으나, **셀에는 아래만 기록**한다.

| G2 (STATUS) | 의미 | J2 (LAST ACTION) |
|-------------|------|------------------|
| `WATCHING` | 시작 직후 / 미트리거 후 재개 | `Start Successful` 또는 `Quote Passed: {instrument} {raw_token} (pnl={#,##0})` |
| `SENT` | UI 전송 성공 (종료 직전) | `Message Sent: {실제 전송 문구}` |
| `STOPPED` | STOP 플래그 | `Stopped` |
| `ERROR` | 실패 | `Python Error: ...` 또는 VBA의 `VBA Error: ...` |

H2 = 마지막 호가 문자열(`{instrument} {raw_token}`), I2 = 마지막 P&L (갱신 시).

P&L을 J2에 넣을 때는 `int(round(pnl))`를 천단위 콤마(`#,##0`)로 포맷한다. 예: `pnl=1,234,567`.

---

## 5. 설정 (`config.py` / `.env`)

`Config.load(path)`:

- 파일이 없으면 `ConfigError`.
- `python-dotenv`로 로드한 뒤, 파일을 다시 읽어 key=value를 파싱한다 (주석·빈 줄 무시, 따옴표 trim).
- 값은 파일 우선, 없으면 환경변수. 필수 키 누락·형식 오류 시 `ConfigError`.
- 비즈니스 수치/셀/창 제목 등의 **코드 기본값은 없다**.

### 5.1 감시·파싱

| 키 | 설명 |
|----|------|
| `TARGET` | 감시 종목 토큰 (예: `25-10`) |
| `SOURCE_WINDOW_TITLE` | 소스 창 제목 **부분 문자열** (예: `K-Bond`). `CHROME_TITLE` 키는 없음 |
| `SOURCE_PROCESS_NAME` | 선택. 프로세스명으로 창 필터 (예: `KBondMessenger.exe`) |
| `YIELD_PREFIX` | 호가 자리수 → 수익률 변환 시 정수부 |
| `REQUIRED_SIDE` | `ANY` / `BUY` / `SELL` |
| `POLL_INTERVAL_MS` | 폴링 간격 (≥ 100) |
| `PROCESS_EXISTING_ON_START` | 시작 시 이미 화면에 있는 라인도 처리할지 |

### 5.2 Excel

| 키 | 설명 |
|----|------|
| `EXCEL_WORKBOOK` | 열린 통합문서 이름 (정확 일치 또는 접미사) |
| `EXCEL_SHEET` | 시트명 (비우면 ActiveSheet) |
| `EXCEL_INPUT_CELL` | 수익률 입력 셀 |
| `EXCEL_PNL_CELL` | P&L 읽기 셀 |
| `EXCEL_STATUS_CELL` | STATUS (예: G2) |
| `EXCEL_LAST_QUOTE_CELL` | LAST QUOTE (예: H2) |
| `EXCEL_LAST_PNL_CELL` | LAST PNL (예: I2) |
| `EXCEL_LAST_ACTION_CELL` | LAST ACTION (예: J2) |
| `PNL_THRESHOLD` | 트리거 기준 (`pnl >= threshold`) |

### 5.3 전송 (`SEND_*`, 카카오 키 없음)

| 키 | 설명 |
|----|------|
| `SEND_PROCESS_NAME` | 대상 프로세스 (예: `KBondMessenger.exe`) |
| `SEND_WINDOW_TITLE` | 창 제목 부분 문자열 (예: `K-Bond`) |
| `MESSAGE_TEMPLATE` | `str.format` 템플릿 |
| `SEND_INPUT_X` / `SEND_INPUT_Y` | 창 `GetWindowRect` 대비 입력란 클릭 비율 (0~1) |
| `SEND_FOREGROUND_RETRY_PAUSE_SECONDS` | 포그라운드 재시도 대기 |
| `SEND_ACTIVATE_SHOW_PAUSE_SECONDS` | Show/Restore 후 대기 |
| `SEND_AFTER_ACTIVATE_PAUSE_SECONDS` | 활성화 후 대기 |
| `SEND_INPUT_CLICK_PAUSE_SECONDS` | 입력란 클릭 후 대기 |
| `SEND_PASTE_PAUSE_SECONDS` | 붙여넣기 후 대기 |
| `SEND_SEND_PAUSE_SECONDS` | Enter 후 대기 |

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

## 7. UIA 소스 읽기 (`source_reader.py`)

이 모듈은 특정 앱 API 이름에 묶이지 않는다. Win32로 창(제목·선택적 프로세스)을 찾은 뒤 UIA `Text`/`ListItem` 등을 모으고, 비어 있으면 Win32 텍스트로 보완한다.

### 7.1 원리

Windows UI Automation(UIA)은 접근성 트리에서 컨트롤 타입·이름·텍스트를 제공한다.  
`pywinauto.Desktop(backend="uia")`로 최상위 창을 나열한 뒤, 선택한 창의 descendants에서 `Document` / `Text`를 순회해 `window_text()`를 수집한다.

채팅 UI가 가상화되어 있으면 **화면에 보이지 않는(또는 트리에 없는) 줄은 UIA에 나타나지 않을 수 있다.**  
Ctrl+F 검색·스크롤·클립보드 복사·OCR은 하지 않는다.

### 7.2 창 찾기 규칙 (`find_source_window`)

1. `Desktop(backend="uia").windows()`로 창 나열.
2. `SOURCE_WINDOW_TITLE`이 창 제목에 **부분 포함**(대소문자 무시)되면 후보.
3. 면적(가로×세로)이 가장 큰 창을 선택.
4. 후보가 없으면 `SourceReaderError`.

생성자: `SourceReader(source_window_title=...)`.  
감시 시작 시 `main.py`는 `find_source_window()`를 한 번 호출해 창 존재 여부를 확인한다.

### 7.3 텍스트 수집 규칙

1. Document 컨트롤 descendants를 우선 root로 삼고, 없으면 창 자체.
2. 각 root에서 `control_type="Text"` descendants의 `window_text()`를 줄 단위로 split.
3. Document에서 Text가 비면 창 전체 Text descendants로 재시도.
4. strip 후 빈 줄·동일 문자열 중복 제거 후 리스트로 반환 (`get_visible_message_lines`).

### 7.4 신규 라인·워터마크

- 라인 fingerprint = SHA1(UTF-8 텍스트).
- `initialize_watermark(process_existing_on_start)`:
  - `True`: watermark를 비움 → 이후(또는 미초기화 첫 호출 시) 현재 화면에 있는 것도 “신규”로 처리 가능.
  - `False`: 현재 화면 라인을 전부 watermark에 넣어 **이후 새로 나타난 것만** 반환.
- `get_new_message_lines`:
  - 미초기화면 `initialize_watermark`와 동일 규칙으로 초기화한 뒤, `process_existing_on_start=True`면 현재 라인을 반환·watermark에 추가, `False`면 빈 리스트.
  - 초기화 이후에는 이미 본 fingerprint는 skip, 새 것만 리스트로 반환하며 watermark에 추가.

세션 쪽(`WatcherSession.processed_fingerprints`)은 파싱에 성공한 호가의 `Quote.fingerprint`를 별도로 관리한다. UIA watermark와는 역할이 다르다.

### 7.5 진단 (`diagnose`)

`--diagnose-source` → `SourceReader(...).diagnose(max_messages=200)`:

- 창 제목, HWND
- Text 컨트롤 수
- 메시지 라인 수와 샘플(최대 N줄)

### 7.6 한계

| 한계 | 영향 |
|------|------|
| 가상화·오프스크린 메시지 | UIA에 안 잡힘 → 감시 누락 가능 |
| 창 제목 변경 | `SOURCE_WINDOW_TITLE` 부분 일치 실패 |
| 동일 제목 다중 창 | 면적 최대 창만 선택 |
| Document/Text 구조 변경 | 수집 실패 또는 빈 라인 |
| 비텍스트 렌더링(이미지 전용) | UIA로 읽을 수 없음 |

운영 전 `--diagnose-source`로 실제 호가 라인이 보이는지 확인한다.

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
`REQUIRED_SIDE`가 `BUY`/`SELL`이면 다른 쪽은 `None` 반환.

---

## 9. Excel (`excel_bridge.py`)

### 연결

- 이미 실행 중인 `Excel.Application`에 `GetActiveObject`로 붙는다. Excel을 새로 띄우지 않는다.
- 워크북: 이름이 `EXCEL_WORKBOOK`과 대소문자 무시 일치, 또는 그 이름으로 끝나면 매칭.
- 시트: `EXCEL_SHEET`가 있으면 해당 시트, 없으면 ActiveSheet.

### 계산 파이프라인

`write_yield_read_pnl(yield_value)`:

1. 입력 셀에 float 기록 (`write_yield`)
2. P&L 셀을 float로 읽어 반환 (`read_pnl`)
3. **`Worksheet.Calculate` / `Application.Calculate`를 호출하지 않는다.** Excel 자동계산(Automatic)에 의존한다.

빈값·변환 실패 시 `ExcelBridgeError`.

### 상태 갱신

`update_status`는 STATUS는 항상 쓰고, 인자로 준 last_quote / last_pnl / last_action만 선택 갱신.  
상태 쓰기 실패는 경고 로그만 (감시 루프를 깨지 않음).

---

## 10. 트리거·메시지 (`trigger.py`)

```text
triggered  ⇔  pnl >= PNL_THRESHOLD
```

`format_message(template, quote, pnl)` → `template.format(instrument=..., raw_token=..., yield_value=..., side=..., pnl=..., raw_line=...)`.

---

## 11. 범용 UI 전송 (`send_ui.py`)

카카오톡 API·방 검색·채팅 탭 클릭은 **없다**.  
`SEND_PROCESS_NAME` / `SEND_WINDOW_TITLE`으로 지정한 창에 붙여넣기 후 Enter한다. 운영 기본은 KBond Messenger(`KBondMessenger.exe`, 제목 `K-Bond`).

대상 프로세스가 **이미 실행 중**이어야 한다. 없으면 `SendError`로 즉시 실패한다 (자동 실행 없음).

### `send_text(text, cfg)`

1. `ensure_target_window`:
   - 프로세스 실행 여부 확인
   - 제목에 `SEND_WINDOW_TITLE` 부분 포함 + 해당 PID인 창 중 rank(가시성·면적) 최고 HWND 선택
2. `activate_window`: Restore/Show → Alt+SetForeground / AttachThreadInput으로 포그라운드
3. `SEND_INPUT_X/Y` 비율 클릭
4. 클립보드에 `text` 복사 → `Ctrl+V` → Enter
5. 로그 `MESSAGE_SENT`

클릭 좌표는 창 `GetWindowRect` 기준 상대 비율이다. UI 해상도·창 크기가 바뀌면 `.env`의 X/Y·pause를 재조정한다.

### 진단

`--diagnose-send` → 프로세스 실행 여부, PID, HWND, rect, 클릭 좌표를 문자열로 출력한다.

---

## 12. Excel VBA (`vba/KBondWatcher.bas`)

모듈명: `KBondWatcher`.

### `StartKBondWatcher`

1. G2/H2/I2/J2를 빈 값으로 초기화 (STARTING/launched 문자열을 쓰지 않음).
2. `WScript.Shell`로 숨김(`0`) 실행:

```text
pythonw.exe "<MAIN_PATH>" --config "<CONFIG_PATH>"
```

`CurrentDirectory` = `PROJECT_DIR`.

3. VBA 자체 실패 시만 G2=`ERROR`, J2=`VBA Error: ...`.

### `StopKBondWatcher`

1. `STOP_FLAG_PATH` 상위 폴더가 없으면 생성.
2. 플래그 파일에 `stop` 기록.
3. **STATUS 셀은 건드리지 않음** — Python이 플래그를 보고 `STOPPED` / `Stopped`를 기록.
4. VBA 실패 시만 ERROR 셀 기록.

### 상수 (코드에 박혀 있음 — 배포 경로와 반드시 맞출 것)

| Const | 값 |
|-------|-----|
| `PYTHONW_PATH` | `pythonw.exe` |
| `PROJECT_DIR` | `C:\mycode\KBondWatcher` |
| `MAIN_PATH` | `C:\mycode\KBondWatcher\main.py` |
| `CONFIG_PATH` | `C:\mycode\KBondWatcher\.env` |
| `STOP_FLAG_PATH` | `C:\temp\kbond_watcher.stop` |
| 상태 셀 | G2 / H2 / I2 / J2 |

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
2. 소스 창이 열려 있고, 창 제목에 `SOURCE_WINDOW_TITLE`이 포함됨.
3. 전송 대상 앱이 실행 중이고, 창 제목에 `SEND_WINDOW_TITLE`이 포함됨 (기본: K-Bond Messenger).
4. `.env` 완비, VBA 경로·스톱 플래그·상태 셀이 `.env`와 일치.
5. `pip install -r requirements.txt` (pywin32, pywinauto, psutil, python-dotenv, pyautogui, pyperclip, pytest).

의존성 설치·진단:

```bat
cd C:\mycode\KBondWatcher
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
| `tests/test_quote_parser.py` | 수락/거부 라인, `REQUIRED_SIDE` 필터 |
| `tests/test_trigger.py` | threshold 비교, 템플릿 포맷 |
| `tests/test_excel_bridge.py` | 브릿지 유틸/동작 |
| `tests/test_send_ui.py` | 좌표·창 매칭 등 센더 단위 |

UI/실기 Excel·KBond 전송은 자동화 테스트 범위 밖이며, CLI `--diagnose-source` / `--diagnose-send` / `--test-send`로 확인한다.

---

## 16. 장애 시 확인 순서

1. Excel G2: `ERROR`면 J2(`Python Error:` / `VBA Error:`) 요약 + `LOG_PATH` 로그.
2. G2가 계속 비어 있음: Python이 안 붙음 → VBA 경로·`pythonw`·`.env` 위치.
3. `WATCHING`인데 호가 미반응: `--diagnose-source`으로 라인이 UIA에 보이는지, `--test-parser`로 형식 일치 여부.
4. 계산만 되고 전송 없음: I2 P&L vs `PNL_THRESHOLD`, Excel 자동계산 여부.
5. 전송 실패: 대상 프로세스 실행 여부, `--diagnose-send`, 클릭 비율·타이밍.

---

## 17. 설계상 고정된 동작 요약

- **One-shot**: 전송 성공(`SENT`) 또는 STOP/ERROR로 종료.
- **설정 외부화**: 비즈니스·좌표·타이밍은 `.env`만. `KAKAO_*` / `CHROME_TITLE` 키 없음.
- **상태셀 단순화**: 트레이더는 G2의 `WATCHING` / `SENT` / `STOPPED` / `ERROR`만 보면 됨.
- **소스**: `SOURCE_WINDOW_TITLE` 창의 UIA Text만.
- **계산**: `write_yield_read_pnl` + Excel 자동계산 (`Worksheet.Calculate` 없음).
- **액션**: `SEND_*` 창 활성화 → 비율 클릭 → 붙여넣기 → Enter.
