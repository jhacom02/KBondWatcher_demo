# KBondWatcher — 설계·로직 명세

Windows에서 채권 채팅 호가를 감시하고, Excel PnL 임계값을 만족하면 확정 메시지를 UI로 전송하는 워처이다.  
이 문서는 **현재 저장소 구현**만 기술한다. 코드와 어긋나면 코드를 진실로 삼고 문서를 갱신한다.

운영 원칙: **폴백 없음.** 애매하면 보내지 않고 ERROR. 재시도는 Excel COM busy·PnL 대기·전송 foreground/focus(최대 2회)·UIA Text 열거(최대 2회)만.

에러 메시지·exit 코드: [`error_table.md`](error_table.md).

---

## 1. 목적과 범위

### 1.1 하는 일

1. `.env`의 `MODE`에 따라 채팅 소스·전송 대상을 정한다.
2. 이미 실행 중인 Excel 워크북에서 **`EXCEL_WATCH_CELL`이 가리키는 한 종목 슬롯**과 `EXCEL_PNL_THRESHOLD_CELL` PnL 임계점을 읽는다.
3. 채팅에서 **신규 라인**만 골라 호가 파싱한다.
4. 해당 슬롯의 수익률 입력 셀에 yield를 쓰고, PnL을 읽는다.
5. Looking For·임계값 조건이 맞으면 side-flip 확정 문자열을 전송 대상 창에 붙여 넣고 Enter한다.
6. 중지 플래그 또는 오류 시 상태를 Excel에 남기고 종료한다.

### 1.2 하지 않는 일

- OCR, Selenium, 카카오톡 전용 연동은 사용하지 않는다.
- Excel을 새로 기동하지 않는다 (`GetActiveObject`로 이미 열린 인스턴스에 붙는다).
- 소스/전송 프로세스 identity는 `MODE` 프리셋이 고정한다. MODE 1·2 채팅방 제목만 `.env`의 `KBOND_CHAT_TITLE`.
- 전송 클릭 비율은 `.env`의 `SEND_INPUT_*_M1` / `SEND_INPUT_*_M23`. MODE 1은 **분리 채팅창** 기준이다.
- 여러 종목을 동시에 감시하지 않는다. 한 폴에서 호가가 2건 이상 매칭되면 보내지 않는다.

### 1.3 런타임 전제

| 전제 | 설명 |
|------|------|
| OS | Windows |
| Excel | 대상 워크북이 **열린 상태**, 계산 옵션 Automatic 권장 |
| MODE 1·2 소스 | `KBondMessenger.exe`. 제목이 `KBOND_CHAT_TITLE`을 포함하는 분리 채팅창의 `TJvRichEdit` |
| MODE 3 소스 | 창 제목에 `FORESTBOND`, UIA `Text` |
| MODE 1 전송 | 같은 `KBOND_CHAT_TITLE` 분리창 입력란 (`SEND_INPUT_*_M1`) |
| MODE 2·3 전송 | `notepad.exe`, 제목에 `메모장` |
| Python | `requirements.txt` 의존성 (venv 권장) |

---

## 2. 아키텍처

```text
Excel VBA Start
    → pythonw main.py --config .env
        → Config.load(.env)
        → ExcelBridge.connect / load_slots
        → create_source_reader(MODE)
        → send.ensure_target_window
        → watermark 초기화
        → loop:
              stop flag?
              get_new_message_lines
              배치 매칭 (0 skip / 1 proceed / 2+ ERROR)
              write yield → wait PnL
              load_slots 재확인 → evaluate
              skip 또는 send_text
              SENT_AFTER=exit 종료 / loop 이면 reseed 후 WATCHING
```

| 경로 | 책임 |
|------|------|
| `main.py` | CLI, 감시 루프, 배치 매칭, 상태 전이 |
| `config/` | `.env` 로드·검증, MODE 프리셋 |
| `source/` | 채팅 라인 수집, watermark, 호가 파서 |
| `source/win32mem.py` | PID 열거, RichEdit 원격 메모리 상수 |
| `excel/` | COM 브리지, 슬롯·Looking For, yield/PnL, 상태 셀 |
| `core/` | Quote·세션 모델, 트리거·메시지 포맷, 로거 |
| `send/` | 대상 창 활성화, 클릭·클립보드 붙여넣기·Enter |
| `vba/` | Excel Start/Stop |
| `tests/` | pytest |
| `docs/` | 이 명세와 에러 표 |
| `logs/` | 런타임 로그 (`LOG_PATH`, gitignore) |

의존성: `pywin32`, `psutil`, `python-dotenv`, `pyautogui`, `pyperclip`, `pywinauto`, `pytest`.

---

## 3. MODE와 입출력

| MODE | 소스 리더 | 소스 창 | 전송 대상 | 클릭 비율 키 |
|------|-----------|---------|-----------|--------------|
| 1 | `KbondSourceReader` | `KBondMessenger.exe` / `KBOND_CHAT_TITLE` / `TJvRichEdit` | 동일 분리창 | `SEND_INPUT_X_M1`, `SEND_INPUT_Y_M1` |
| 2 | `KbondSourceReader` | 동일 | `notepad.exe` / `메모장` | `SEND_INPUT_X_M23`, `SEND_INPUT_Y_M23` |
| 3 | `UiaSourceReader` | 제목 `FORESTBOND` (프로세스명 없음) | Notepad | `SEND_INPUT_X_M23`, `SEND_INPUT_Y_M23` |

`create_source_reader` (`source/reader.py`): MODE 1·2 → KBond, MODE 3 → UIA, 그 외 `SourceReaderError`.

KBond 채팅 본문은 UIA `Text`가 비어 있어 RichEdit Win32로 읽는다. FORESTBOND는 UIA `Text`가 열려 MODE 3을 둔다.

---

## 4. 설정 (`config/loader.py`)

### 4.1 로드 규칙

1. `--config` 파일이 없으면 `ConfigError`.
2. `python-dotenv` 로드 후 파일을 다시 읽어 `key=value` 파싱 (빈 줄·`#` 무시, 따옴표 trim).
3. 값 우선순위: **파일 키 > 환경변수**. 필수 키 누락·형식 오류 → `ConfigError`, exit 2.

### 4.2 MODE 프리셋

코드 상수: `KBondMessenger.exe`, `FORESTBOND`, `notepad.exe`, `메모장`.

MODE 1·2는 `.env` **`KBOND_CHAT_TITLE`** (비면 안 됨)으로 분리창을 고른다. 대소문자 무시 부분일치. MODE 1 전송 제목도 이 값. MODE 3은 이 키를 쓰지 않는다.

`.env`에 `SOURCE_WINDOW_TITLE` / `SOURCE_PROCESS_NAME` / `SEND_PROCESS_NAME` / `SEND_WINDOW_TITLE`가 있어도 **무시**된다.

### 4.3 클릭 비율

| 키 | MODE | 의미 |
|----|------|------|
| `SEND_INPUT_X_M1` / `SEND_INPUT_Y_M1` | 1 | 분리 채팅창 대비 클릭 비율 (0~1) |
| `SEND_INPUT_X_M23` / `SEND_INPUT_Y_M23` | 2, 3 | 동일 |

로드 시 MODE에 맞는 쌍만 `Config.send_input_x/y`에 넣는다.

### 4.4 키 요약

**감시**

| 키 | 제약 |
|----|------|
| `MODE` | 1 / 2 / 3 |
| `KBOND_CHAT_TITLE` | MODE 1·2 필수 |
| `POLL_INTERVAL_MS` | 정수 ≥ 100 |
| `PROCESS_EXISTING_ON_START` | bool |
| `SENT_AFTER` | `exit` (전송 후 종료) 또는 `loop` (reseed 후 계속) |

**Excel**

| 키 | 설명 |
|----|------|
| `EXCEL_WORKBOOK` | 열린 통합문서 경로 또는 이름 매칭 |
| `EXCEL_SHEET` | 선택. 비면 ActiveSheet |
| `EXCEL_SLOT_ROWS` | 슬롯 허용 행. 워치는 이 중 **한 행**만 |
| `EXCEL_ROWS_10Y` / `EXCEL_ROWS_3Y` | prefix 매핑. 모든 slot row가 합집합에 속해야 함 |
| `EXCEL_INSTRUMENT_COL` 등 | 열 문자 |
| `EXCEL_PNL_ROW_OFFSET` | PnL 행 = 슬롯행 + offset (≥ 0) |
| `EXCEL_PREFIX_3Y_CELL` / `EXCEL_PREFIX_10Y_CELL` | 예: B5 / B6 |
| `EXCEL_WATCH_CELL` | 감시 종목 (관례 D2) |
| `EXCEL_PNL_THRESHOLD_CELL` | PnL 임계 (관례 E2) |
| `EXCEL_PNL_SANITY_BAND` | `\|pnl - threshold\|` 초과 시 ERROR |
| `EXCEL_STATUS_CELL` … `EXCEL_LAST_ACTION_CELL` | F2~J2 |

**전송·운영**

| 키 | 설명 |
|----|------|
| `MESSAGE_TEMPLATE` | `str.format` |
| `SEND_*_PAUSE_SECONDS` | 포그라운드·클릭·붙여넣기·Enter |
| `STOP_FLAG_PATH` | 파일이 있으면 STOP |
| `LOG_LEVEL` / `LOG_PATH` | 상대 경로는 `.env` 부모 기준 |

`MESSAGE_TEMPLATE` 치환: `instrument`, `raw_token`, `confirm_token`, `yield_value`, `side`, `pnl`, `raw_line`, `quantity`.

---

## 5. Excel (`excel/bridge.py`)

### 5.1 연결

`GetActiveObject("Excel.Application")`. 워크북은 설정 경로와 `FullName` 일치, 또는 `Name`/파일명 매칭. RPC busy는 같은 호출만 50×0.1s 재시도.

### 5.2 슬롯

`EXCEL_SLOT_ROWS`는 허용 행이다. 워처는 `EXCEL_WATCH_CELL`로 고른 **한 행만** 로드한다.

| 필드 | 출처 |
|------|------|
| `instrument` | `{INSTRUMENT_COL}{row}`, 선행 `국고` 제거 |
| `looking_for` / `required_side` / `qty_abs` | `{QTY_COL}{row}`. 0이 아닌 정수. 부호=방향, 절댓값=억 수량 |
| `yield_prefix` | 10Y 행 → PREFIX_10Y 셀, 3Y 행 → PREFIX_3Y. `floor(abs(값))` |
| `input_cell` | `{INPUT_COL}{row}` |
| `pnl_cell` | `{PNL_COL}{row + PNL_ROW_OFFSET}` |

감시 셀: Formula가 `=A41` / `=$A$41` / `=현재시트!A41`이면 그 행(허용 목록만). `=`가 아니면 표시 문자열을 허용 행 A열과 **정확히 1건** 매칭.

임계셀은 float. 기동 시와 **새 채팅 줄이 있을 때마다** 슬롯·임계를 다시 읽는다. 유휴 폴링에서는 Excel을 읽지 않는다.

### 5.3 Looking For · 트리거

| E(qty) | G2 | 파서 `required_side` | 트리거 |
|--------|-----|----------------------|--------|
| 음수 | `{종목} / BID` | `BUY` | `pnl <=` 임계 |
| 양수 | `{종목} / OFFER` | `SELL` | `pnl >=` 임계 |

evaluate에는 G2 문자열이 아니라 슬롯의 `BID`/`OFFER`를 넘긴다.

### 5.4 yield와 PnL

`write_yield_read_pnl`:

1. 입력 셀에 yield 기록  
2. `CalculationState == xlDone(0)` 이고 PnL이 float일 때까지 대기 (30s, 50ms). CVErr·빈칸은 숫자로 쓰지 않음.  
3. `|pnl - threshold| > EXCEL_PNL_SANITY_BAND`이면 `main`이 ERROR.

### 5.5 상태 셀

| 셀 | 역할 |
|----|------|
| F2 | `WATCHING` / `SENT` / `STOPPED` / `ERROR` |
| G2 | `{instrument} / BID` 또는 `… / OFFER` |
| H2 | `{instrument} {raw_token}` |
| I2 | PnL |
| J2 | `(HH:MM:SS) …` 또는 `(HH:MM:SS) Error: …` |

세션 enum에는 `QUOTE_FOUND` 등이 있으나 F2에는 위 운영 값 위주다. 셀 쓰기 실패는 즉시 종료.

---

## 6. 소스와 watermark

공통 API (`source/common.py`): `find_source_window`, `get_visible_message_lines`, `get_new_message_lines`, `initialize_watermark`, `reseed_watermark_from_visible`, `diagnose`.

라인 식별: `SHA1(UTF-8 watermark_key)` (`message_fingerprint`). 창 크기 `WATERMARK_WINDOW` = 2000줄 (꼬리만 비교).

### 6.1 MODE 1·2 (RichEdit)

의도: 분리 채팅창 문서 전체를 스크롤과 무관하게 읽는다.

1. `KBondMessenger.exe` 최상위 창 중 제목에 `KBOND_CHAT_TITLE` 포함. 0건 또는 **서로 다른 창 2개 이상**이면 ERROR.  
2. 자식 `TJvRichEdit` 중 가시·면적/높이 최소 이상인 것 중 최대. 입력란 `TRxRichEdit`는 제외.  
3. 잡은 HWND는 `IsWindow`인 동안 재사용한다. 컨트롤이 잠깐 숨어도 재탐색하지 않는다. 핸들이 파괴되면 다시 찾는다.  
4. `WM_GETTEXTLENGTH`가 직전과 같으면 본문 API를 생략하고 캐시.  
5. 길이 ≤ 2,000,000 wchar이면 `WM_GETTEXT`. 초과면 **끝 200만**만 `EM_GETTEXTRANGE` (종료하지 않음). OpenProcess 실패면 ERROR.  
6. `watermark_key` = 원문 줄.

메신저를 포그라운드로 강제하지 않는다.

### 6.2 MODE 3 (UIA)

의도: FORESTBOND가 접근성 트리에 노출하는 `Text`만 읽는다. 뷰포트 조각이므로 시각 토큰이 없는 호가 줄은 매칭하지 않는다.

1. 제목 `FORESTBOND` 창 중 면적 최대. 폴마다 데스크톱을 다시 훑지 않고 핸들을 든다. 무효·비가시면 `find_source_window`.  
2. `Text` descendants. 열거 예외는 최대 2회 재시도 후 ERROR. 0건이면 즉시 ERROR.  
3. 바로 앞 줄이 시간 토큰이면 다음 줄 `watermark_key` = `(시각) : {호가}`. 파서 `text`는 호가 조각.  
4. 한 Text에 시각+호가가 이미 붙어 있으면 재결합하지 않는다.  
5. 매칭 시 `watermark_key`에 `(HH:MM`이 없으면 skip (ERROR 아님).

위로 스크롤하면 옛 조각이 신규로 보이거나 최신이 트리에서 빠질 수 있다. 맨 아래 유지가 전제다.

### 6.3 Watermark

| API | 동작 |
|-----|------|
| `initialize_watermark(false)` | 꼬리 창 fp로 set을 채움. 이후 새로 들어온 줄만 통과 |
| `initialize_watermark(true)` | set을 비움. 꼬리 창도 신규 가능 |
| `get_new_message_lines` | 꼬리 창에서 set에 없는 줄만 반환·FIFO add |
| `reseed_watermark_from_visible` | 꼬리 창을 set에 union |

`PROCESS_EXISTING_ON_START=false`: 기동 시점 꼬리 창은 스킵. `true`: 기동 시 보이는 줄도 처리 (재현 테스트용).

워터마크는 프로세스 메모리뿐이며 디스크에 저장하지 않는다.

### 6.4 세션 fingerprint

`WatcherSession.processed_fingerprints`는 **파싱에 성공한** `watermark_key` SHA1이다.

- 소스 watermark: 이 채팅 식별 키를 이미 봤는가  
- session set: 이 식별 키로 이미 매칭·처리했는가  

MODE 3에서 같은 호가 문구라도 시각이 다르면 키가 달라 재기회다. `SENT_AFTER=loop`여도 session set은 비우지 않는다. `Quote.fingerprint`는 `raw_line` SHA1이며 워처 매칭에는 쓰지 않는다.

---

## 7. 호가 파서 (`source/quote_parser.py`)

입력: 한 줄 + `instrument`, `yield_prefix`, `required_side`, `required_qty`.  
실패는 예외가 아니라 `None` (전송하지 않고 다음 줄).

종목은 앞뒤가 숫자/`-`가 아닌 경계에서만 매칭 (`25-11` ≠ `125-11`).

종목 직후:

```text
^\s+(?P<price>\d{2,3})\s*(?P<side>[+-]|사자|팔자)\s*
```

- 수량 생략 또는 `(` / `*` trailing만 → 100억. `required_qty == 100`일 때만 통과.  
- 그 외 `^\d+\s*억?`만 인정하고 `required_qty`와 같아야 함.  
- `있나요`, `ㅎㅈ`, 추가 숫자는 거부.

2자리 가격 → `prefix + n/100`, 3자리 → `prefix + n/1000`. `+`/`사자` = BUY, `-`/`팔자` = SELL.

---

## 8. 트리거와 메시지 (`core/trigger.py`)

- BID: `pnl <= threshold`  
- OFFER: `pnl >= threshold`  

`flip_side_token`은 전송문에만 쓴다 (`+`↔`-`, `사자`↔`팔자`).  
`quantity != 100`이면 confirm_token 뒤에 ` {n}억`.

---

## 9. 전송 (`send/ui.py`)

1. 전송 프로세스·제목 부분일치 창.  
2. `activate_window`: 아이콘이면 Restore. 이미 전경이면 Show/sleep 생략.  
3. TOPMOST → 비율 클릭 → 포그라운드·커서 하위 창이 대상 앱인지 확인. `failed to foreground` / `send focus not on target`만 최대 2회 재시도.  
4. `pyperclip.copy` 후 클립보드가 전송문과 같아야 함. 아니면 `SendError`, Ctrl+V 전 중단(재시도 없음).  
5. Ctrl+V → Enter. finally TOPMOST off.

`--diagnose-send`로 HWND·제목·클릭 좌표를 본다. MODE 1은 분리창에서 M1 비율을 재측정한다.

---

## 10. 메인 루프 (`main.py`)

### 10.1 CLI

| 플래그 | 동작 |
|--------|------|
| (기본) | `run_watcher` |
| `--config PATH` | dotenv (기본 `.env`) |
| `--diagnose-source` / `--diagnose-send` | 창·본문 또는 전송 좌표 |
| `--test-send` | 샘플 Quote 전송 |
| `--test-parser LINE` | Excel 슬롯 기준 파서 |
| `--perf-summary` | `logs/sent_perf.csv` mean/median |

Config 실패 exit **2**. 그 외 운영 오류 exit **1**. Stop은 exit **0**.

### 10.2 시작

PID 파일 기록 · stop 플래그 삭제 → Excel connect → `load_slots` → F2 `WATCHING` → 소스·전송 창 → watermark 초기화 → 폴링.

### 10.3 한 사이클

```text
stop? → STOPPED, return 0
lines = get_new_message_lines
없으면 sleep
있으면 load_slots, LINE 로그, collect_batch_matches
  0건 → sleep
  2건+ → ERROR ambiguous quotes in one poll
  1건 → yield/PnL → load_slots 재확인
        종목·looking_for·qty·threshold 변경 → ERROR
        sanity band 초과 → ERROR
        evaluate skip → Quote Skipped
        trigger → send_text → SENT → SENT_AFTER
sleep(POLL_INTERVAL_MS)
```

### 10.4 `SENT_AFTER`

| 값 | 동작 |
|----|------|
| `exit` | SENT 기록 후 exit 0 |
| `loop` | `reseed_watermark_from_visible` 후 WATCHING 유지 |

### 10.5 SENT 지연 로그

확정(`send_text` 성공)마다 [`logs/sent_perf.csv`](../logs/sent_perf.csv)에 한 줄 append. Quote Skipped·ERROR는 기록하지 않는다.

`total_ms`는 **호가를 매칭한 직후부터 확정 메시지를 다 보내기까지** (Excel PnL 대기 + 클릭·붙여넣기·Enter). 채팅 게시 시각은 초 단위라 시작점으로 쓰지 않는다. 폴링으로 줄을 보기 전 지연은 포함하지 않는다.

| 열 | 의미 |
|----|------|
| `ts` | 전송 완료 ISO |
| `total_ms` | 매칭 → `send_text` return |
| `excel_ms` | `write_yield_read_pnl` |
| `send_ms` | `send_text` |
| `mode` | 1 / 2 / 3 |
| `looking_for` | G2 라벨 |
| `raw_line` | MODE 1·2 `text`, MODE 3 `watermark_key` |
| `sent_message` | 실제로 보낸 확정 문장 |

기록 실패는 전송을 되돌리지 않는다. 통계: `python main.py --config .env --perf-summary`.

---

## 11. VBA (`vba/KBondWatcher.bas`)

상수의 `PROJECT_DIR` / `PYTHONW_PATH` / `STOP_FLAG_PATH`를 설치 PC에 맞게 수정한다. STOP 경로는 `.env`와 같아야 한다. F2~J2 주소는 `.env` `EXCEL_*_CELL`과 같아야 한다.

| Sub | 동작 |
|-----|------|
| `StartKBondWatcher` | 기존 워처 종료 → 플래그 삭제 → F2~J2 클리어 → `pythonw main.py --config .env` |
| `StopKBondWatcher` | stop 플래그 → taskkill → F2 `STOPPED` |
| Fail | F2 `ERROR`, J2 `(HH:nn:ss) Error: …` |

`KillWatcher`는 PID `taskkill`(0·128 정상) 후 **python/pythonw**의 `main.py`만 PowerShell로 정리한다.

---

## 12. 로깅

로거 `kbond_watcher`. RotatingFileHandler 2MB × 5 + 콘솔.  
`%(asctime)s %(levelname)s %(message)s`. 운영 경로에 warning/폴백 없음.

| 키워드 | 때 |
|--------|-----|
| `WATCHING` | 기동·스킵 후 복귀 |
| `LINE` | watermark를 통과한 신규 줄 (`mode`/`looking_for`/`threshold` 정수/`raw_line`). 폴당 최대 20, 160자 truncate. MODE 3는 `watermark_key` |
| `LINE_OMITTED` | 20줄 초과분 |
| `QUOTE_FOUND` | 파싱 성공 1건 |
| `NO_TRIGGER` / `TRIGGERED` / `SENT` / `STOPPED` / `ERROR` / `EXIT` | 상태 |
| `EXCEL_CONNECTED` / `EXCEL_WRITE` / `PNL` | COM |
| `SLOTS_LOADED` | DEBUG |
| `MESSAGE_SENT` | 전송 완료 |
| `source watermark` / `reseed` | 게이트 |

---

## 13. 테스트

```bat
pytest -q
```

| 파일 | 범위 |
|------|------|
| `tests/test_config_mode.py` | MODE, 클릭 비율, `SENT_AFTER` |
| `tests/test_quote_parser.py` | 파서 accept/reject |
| `tests/test_trigger.py` | BID/OFFER, sanity band |
| `tests/test_excel_bridge.py` | 감시 셀, CVErr, busy |
| `tests/test_richedit_reader.py` | 캡, 창 제목 선택, 핸들 유효 |
| `tests/test_send_ui.py` | 클릭 좌표, 클립보드 |
| `tests/test_watcher_guards.py` | 배치 매칭, LINE 로그 |
| `tests/test_perf_log.py` | SENT CSV append, mean/median |
| `tests/test_uia_time.py` | 시각 토큰, UIA 창 캐시 |
| `tests/test_watermark.py` | 2000줄 창, 신규 게이트 |

---

## 14. 운영 체크

1. `.env`의 `MODE`, `KBOND_CHAT_TITLE`(1·2), `SENT_AFTER`, `EXCEL_WORKBOOK`, `STOP_FLAG_PATH`.  
2. VBA 경로·`pythonw`를 이 PC에 맞게.  
3. 소스·전송·Excel을 연 뒤 `--diagnose-source` / `--diagnose-send`.  
4. D2가 허용 행 한 종목, 그 행 E열이 0이 아닌 정수, E2가 숫자.  
5. MODE 1·2는 제목이 맞는 **분리창**이 살아 있어야 한다. 닫거나 메인에 붙이면 ERROR.  
6. MODE 3는 채팅을 맨 아래에 둔다. 시간 토큰 없는 호가 줄은 매칭하지 않는다.

| 증상 | 점검 |
|------|------|
| 시작 즉시 ERROR | 창, Excel, 설정 키, J2 `Error:` |
| WATCHING인데 무반응 | qty/side, `PROCESS_EXISTING_ON_START=false`, diagnose-source |
| PnL만 되고 전송 없음 | I2 vs 임계, Looking For |
| I2가 −2146826273 | `#VALUE!` 타임아웃 |
| 전송이 Excel로 감 | diagnose-send, 포커스, 클릭 비율 |
| MODE 1·2 소스 없음 | 분리창 제목, `TJvRichEdit` |

---

## 15. 데이터 객체

```text
config.loader     → Config
source.*          → list[SourceLine]
quote_parser      → Optional[Quote]
excel.bridge      → InstrumentSlot, pnl, status
core.trigger      → TriggerResult, outbound text
send.ui           → paste + Enter
main              → 배치 매칭 + session fingerprint + stop
```

- `SourceLine` — `text`, `watermark_key`  
- `Quote` — instrument, raw_line, raw_token, yield_value, side, quantity, sender/ts  
- `InstrumentSlot` — row, looking_for, required_side, qty_abs, yield_prefix, cells  
- `TriggerResult` — triggered, reason, pnl, quote  
- `WatcherSession` — processed_fingerprints, status  
