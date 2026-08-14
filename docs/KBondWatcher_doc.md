# KBondWatcher — 설계·로직 명세 (현재 코드 기준)

Windows에서 채권 채팅 호가를 감시하고, Excel PnL 임계값을 만족하면 확정 메시지를 UI로 전송하는 워처이다.  
이 문서는 **현재 저장소 구현**만 기술한다. 다른 개발자가 프로세스·모듈 경계를 파악하고 유지보수할 수 있도록 작성했다.

---

## 1. 목적과 범위

### 1.1 하는 일

1. `.env`의 `MODE`에 따라 채팅 소스·전송 대상을 정한다.
2. 이미 실행 중인 Excel 워크북에서 **D2가 가리키는 한 종목 슬롯**과 E2 PnL 임계점을 읽는다.
3. 채팅에서 **신규 라인**만 골라 호가 파싱한다.
4. 해당 슬롯의 수익률 입력 셀에 yield를 쓰고, PnL을 읽는다.
5. Looking For·임계값 조건이 맞으면 side-flip 확정 문자열을 전송 대상 창에 붙여 넣고 Enter한다.
6. 중지 플래그 또는 오류 시 상태를 Excel에 남기고 종료한다.

### 1.2 하지 않는 일

- OCR, Selenium, 카카오톡 전용 연동은 사용하지 않는다.
- Excel을 새로 기동하지 않는다 (`GetActiveObject`로 이미 열린 인스턴스에 붙는다).
- 소스/전송 창 identity(프로세스·제목)는 `.env`의 `SOURCE_*` / `SEND_PROCESS_NAME` 등으로 바꾸지 않는다. `MODE` 프리셋이 고정한다.
- 전송 클릭 비율만 `.env`의 `SEND_INPUT_*_M1` / `SEND_INPUT_*_M23`으로 조절한다.

### 1.3 런타임 전제

| 전제 | 설명 |
|------|------|
| OS | Windows |
| Excel | 대상 워크북이 **열린 상태**, 계산 옵션 Automatic 권장 |
| MODE 1·2 소스 | `KBondMessenger.exe`. 분리 채팅창 `TJvRichEdit`(`WM_GETTEXT`) 우선, 없으면 제목 `K-Bond`의 `TElTree` |
| MODE 3 소스 | 창 제목에 `FORESTBOND`, UIA로 `Text` 노출 |
| MODE 1 전송 | 동일 KBond 창 입력란 (클릭 비율) |
| MODE 2·3 전송 | `notepad.exe`, 제목에 `메모장` |
| Python | `requirements.txt` 의존성 설치 (venv 권장) |

---

## 2. 아키텍처 개요

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
              parse / match slot
              write yield → read PnL
              evaluate → skip or send
              (현재) SENT 후 reseed → WATCHING 유지
```

| 패키지 | 책임 |
|--------|------|
| `main.py` | CLI, 감시 루프, 상태 전이, 슬롯 매칭 오케스트레이션 |
| `config/` | `.env` 로드·검증, MODE 프리셋(창 identity) |
| `source/` | 채팅 라인 수집(TElTree / UIA), watermark, 호가 파서 |
| `excel/` | COM 브리지, 슬롯·Looking For, yield/PnL, 상태 셀 |
| `core/` | 모델, 트리거·메시지 포맷, 로거 |
| `send/` | 대상 창 활성화, 클릭·클립보드 붙여넣기·Enter |
| `vba/` | Excel에서 Start/Stop |
| `tests/` | pytest |
| `docs/` `sample/` `logs/` `tools/` | 문서·샘플·로그·진단 도구 |

의존성: `pywin32`, `psutil`, `python-dotenv`, `pyautogui`, `pyperclip`, `pywinauto`, `pytest`.

---

## 3. MODE와 입출력 매핑

| MODE | 소스 리더 | 소스 창 | 전송 대상 | 클릭 비율 키 |
|------|-----------|---------|-----------|--------------|
| 1 | `KbondSourceReader` (RichEdit / TElTree) | `KBondMessenger.exe` / 채팅 분리창 또는 제목 `K-Bond` | 동일 KBond | `SEND_INPUT_X_M1`, `SEND_INPUT_Y_M1` |
| 2 | `KbondSourceReader` (RichEdit / TElTree) | 동일 | `notepad.exe` / `메모장` | `SEND_INPUT_X_M23`, `SEND_INPUT_Y_M23` |
| 3 | `UiaSourceReader` | 제목 `FORESTBOND` (프로세스명 없음) | Notepad | `SEND_INPUT_X_M23`, `SEND_INPUT_Y_M23` |

`create_source_reader(cfg)` (`source/reader.py`):

- MODE ∈ {1, 2} → KBond 리더  
- MODE == 3 → UIA 리더  
- 그 외 → `SourceReaderError`

---

## 4. 설정 (`config/loader.py`)

### 4.1 로드 규칙

1. `--config` 경로의 파일이 없으면 `ConfigError`.
2. `python-dotenv`로 로드한 뒤, 파일을 다시 읽어 `key=value`를 파싱한다 (빈 줄·`#` 주석 무시, 따옴표 trim).
3. 값 우선순위: **파일 키 > 환경변수**. 필수 키 누락·형식 오류 시 `ConfigError`.

### 4.2 MODE 프리셋 (창 identity만)

코드 상수:

- `KBOND_PROCESS = KBondMessenger.exe`, `KBOND_TITLE = K-Bond`
- `FORESTBOND_TITLE = FORESTBOND`
- `NOTEPAD_PROCESS = notepad.exe`, `NOTEPAD_TITLE = 메모장`

`.env`의 `SOURCE_WINDOW_TITLE`, `SOURCE_PROCESS_NAME`, `SEND_PROCESS_NAME`, `SEND_WINDOW_TITLE`는 **무시**된다.

### 4.3 클릭 비율 (`.env` 필수)

| 키 | 사용 MODE | 의미 |
|----|-----------|------|
| `SEND_INPUT_X_M1` / `SEND_INPUT_Y_M1` | 1 | 대상 창 client 대비 클릭 비율 (0~1) |
| `SEND_INPUT_X_M23` / `SEND_INPUT_Y_M23` | 2, 3 | 동일 |

로드 시 MODE에 맞는 쌍만 `Config.send_input_x/y`에 넣는다. 범위 밖이면 `ConfigError`.

### 4.4 필수·선택 키 요약

**감시**

| 키 | 제약 |
|----|------|
| `MODE` | 1 / 2 / 3 |
| `POLL_INTERVAL_MS` | 정수 ≥ 100 |
| `PROCESS_EXISTING_ON_START` | bool (`true`/`false` 등) |

**Excel**

| 키 | 설명 |
|----|------|
| `EXCEL_WORKBOOK` | 열린 통합문서 절대경로(또는 이름 매칭 가능한 값) |
| `EXCEL_SHEET` | 선택. 비면 ActiveSheet |
| `EXCEL_SLOT_ROWS` | 슬롯 허용 목록 (예: `19,25,41,46,56`). D2가 이 중 **한 행**만 선택 |
| `EXCEL_ROWS_10Y` / `EXCEL_ROWS_3Y` | prefix 셀 매핑용 행. 모든 slot row가 합집합에 속해야 함 |
| `EXCEL_INSTRUMENT_COL` 등 | 열 문자 (A, E, D, F …) |
| `EXCEL_PNL_ROW_OFFSET` | PnL 행 = 슬롯행 + offset (≥ 0). 41이면 F44 |
| `EXCEL_PREFIX_3Y_CELL` / `EXCEL_PREFIX_10Y_CELL` | 예: B5 / B6 |
| `EXCEL_STATUS_CELL` … `EXCEL_LAST_ACTION_CELL` | 상태 행 (관례상 F2~J2) |

감시 종목은 `.env`가 아니라 시트 **D2** (`=A41` 또는 종목 문자열). 임계점은 **E2** 숫자(부호 그대로). 저장 없이 열린 워크북 COM 값. 잘못된 D2/E2는 즉시 중단.

**전송·운영**

| 키 | 설명 |
|----|------|
| `MESSAGE_TEMPLATE` | `str.format` (아래 필드) |
| `SEND_*_PAUSE_SECONDS` | 포그라운드·클릭·붙여넣기·Enter 타이밍 |
| `STOP_FLAG_PATH` | 존재하면 STOP |
| `LOG_LEVEL` / `LOG_PATH` | 상대 경로는 `.env` 부모 기준 |

`MESSAGE_TEMPLATE` 치환 키: `instrument`, `raw_token`, `confirm_token`, `yield_value`, `side`, `pnl`, `raw_line`, `quantity`.

---

## 5. Excel 계약 (`excel/bridge.py`)

### 5.1 연결

- `win32com` `GetActiveObject("Excel.Application")` — Excel이 꺼져 있으면 실패.
- 워크북: 설정 경로와 `FullName` 일치, 또는 `Name`/파일명 매칭.
- 시트: `EXCEL_SHEET` 지정 또는 ActiveSheet.

### 5.2 슬롯 (`InstrumentSlot`)

`EXCEL_SLOT_ROWS`는 시트에 있는 슬롯 구조(허용 행)다. 워처는 D2로 고른 **한 행만** 로드한다. 예: D2 `=A41` → row 41.

| 필드 | 출처 |
|------|------|
| `instrument` | `{INSTRUMENT_COL}{row}` (A41), 선행 `국고` 제거. 빈 값이면 오류 |
| `looking_for` / `required_side` / `qty_abs` | `{QTY_COL}{row}` (E41). **0이 아닌 정수**. 부호 → 방향, 절댓값 → 호가 수량(억). 소수·0이면 오류. 다른 슬롯 E열은 읽지 않음 |
| `yield_prefix` | 행이 10Y 목록이면 `PREFIX_10Y` 셀, 3Y면 `PREFIX_3Y` 셀. `floor(abs(값))` |
| `input_cell` | `{INPUT_COL}{row}` (D41) |
| `pnl_cell` | `{PNL_COL}{row + PNL_ROW_OFFSET}` (F44) |

D2 해석: Formula가 `=A41` / `=$A$41` / `=현재시트!A41`이면 그 행(허용 목록만). `=`가 아니면 표시 문자열을 허용 행 A열과 **정확히 1건** 매칭. 그 외 수식·빈칸·0/2건 매칭·다른 시트 참조는 즉시 오류.

E2는 PnL 임계점 float. 빈칸·비숫자는 즉시 오류. 기동 시와 **새 채팅 줄이 있을 때마다** D2/E2와 선택 슬롯을 다시 읽는다 (저장·재실행 불필요). 유휴 폴링에서는 Excel을 읽지 않는다.

### 5.3 Looking For · 수집 side · 트리거

| E(qty) | Looking For (G2) | 파서 `required_side` | 호가 수량 | 트리거 조건 |
|--------|------------------|----------------------|-----------|-------------|
| 음수 (예 `-80`, `-100`) | `BID` | `BUY` (`+` / `사자`) | `abs(E)`억 | `pnl <= E2` → 확정 시 팔자 토큰 |
| 양수 (예 `+80`, `+100`) | `OFFER` | `SELL` (`-` / `팔자`) | `abs(E)`억 | `pnl >= E2` → 확정 시 사자 토큰 |

E2는 부호 있는 숫자 그대로다. BID인데 E2가 큰 양수면 거의 모든 호가가 확정된다.

### 5.4 yield 기록과 PnL 읽기

`write_yield_read_pnl(input_cell, pnl_cell, yield_value)`:

1. 입력 셀에 yield 기록  
2. `Application.CalculationState == xlDone(0)` 될 때까지 폴링 (타임아웃 30s, 간격 50ms)  
3. PnL 셀을 float로 읽어 반환  

### 5.5 상태 셀 (`update_status`)

| 셀 (관례) | 역할 | Excel에 쓰는 값 |
|-----------|------|-----------------|
| F2 Status | 감시 상태 | 주로 `WATCHING` / `SENT` / `STOPPED` / `ERROR` |
| G2 Looking For | 방향 | `BID` / `OFFER` |
| H2 Last Quote | 마지막 호가 | `{instrument} {raw_token}` |
| I2 Last PnL | 마지막 PnL | 숫자 |
| J2 Last Action | 마지막 동작 | `(HH:MM:SS) …` |

`update_status`는 status를 항상 쓰고, 인자로 준 looking_for / last_quote / last_pnl / last_action만 선택 갱신한다. 셀 쓰기 실패는 경고 로그만 (예외로 루프를 끊지 않음).

세션 내부 enum에는 `QUOTE_FOUND`, `CALCULATING`, `TRIGGERED`, `SENDING` 등도 있으나, Excel F2에는 위 운영 4종(+ WATCHING 복귀) 위주로 반영한다.

---

## 6. 소스 읽기와 watermark

소스 리더는 모두 `BaseSourceReader` (`source/common.py`)를 구현한다.  
**공통 API:** `find_source_window`, `get_visible_message_lines`, `get_new_message_lines`, `initialize_watermark`, `reseed_watermark_from_visible`, `diagnose`.

### 6.1 왜 방식이 둘인가

| | UIA (MODE 3) | KBond RichEdit / TElTree (MODE 1·2) |
|--|--------------|--------------------------------------|
| 메커니즘 | Windows UI Automation 접근성 트리의 `Text` | 분리 채팅창 `TJvRichEdit`에 `WM_GETTEXT`. 없으면 메인 창 `TElTree` 메모리 읽기 |
| 전제 | 앱이 텍스트를 UIA에 노출 | 채팅 본문이 `TJvRichEdit` 또는 `TElTree` |
| 대상 예 | FORESTBOND | KBond Messenger |
| 줄 형태 | 컨트롤 단위라 **조각화** 가능 | RichEdit는 `\r\n` 줄, TElTree는 노드 ≈ 한 줄 |
| 범용성 | 노출된 앱에만 | KBond(Delphi) 레이아웃에 맞춤 |

KBond는 UIA `Text`가 비어 있고, 채팅을 분리하면 제목에 `K-Bond`가 없는 `TfrmDetach` + `TJvRichEdit`가 된다. 그래서 MODE 1·2는 **RichEdit를 먼저** 읽고, 없을 때만 TElTree로 떨어진다. FORESTBOND는 UIA가 열려 MODE 3을 둔다. OCR은 사용하지 않는다.

### 6.2 KBond 소스 흐름 (MODE 1·2)

**RichEdit (우선):**

1. `KBondMessenger.exe` PID의 **가시** 최상위 창을 열거 (제목 `K-Bond` 불필요 — 분리창 `[채팅] …` 포함).  
2. 자식 중 class `TJvRichEdit`이면서 가시·면적/높이 최소값 이상인 컨트롤 중 **가장 큰 것**을 채팅 본문으로 선택 (입력란 `TRxRichEdit`는 제외).  
3. `WM_GETTEXTLENGTH`를 먼저 본다. 직전 길이와 같으면 `WM_GETTEXT`를 생략하고 캐시한 `list[str]`을 쓴다. 길이가 늘거나 처음이면 `WM_GETTEXT`로 본문을 읽는다.  
4. 덤프는 스크롤과 무관한 **문서 전체**다. 위로 스크롤해도 내용이 바뀌지 않는다. strip · 빈 줄 제거 · 문자열 중복 제거 → `list[str]`.  
5. 버퍼 상한은 약 2,000,000 wchar다. `GETTEXTLENGTH`가 이보다 크면 **앞에서 자른 본문을 쓰지 않는다** (신규는 맨 아래에 있어 절단 시 유실됨). 경고를 남기고 직전 정상 캐시를 유지한다. `--diagnose-source` 헤더의 `gettext_len` / `clipped`로 절단 여부를 확인한다.

**TElTree (폴백):**

1. 제목에 `K-Bond` 포함 최상위 창 선택.  
2. 자식 `TElTree`, 부모 대비 중심 X 비율 ≥ 0.55.  
3. `OpenProcess` 후 `TVM_GETCOUNT` / `TVM_GETITEM`으로 아이템 텍스트 수집.

폴링 중 메신저를 포그라운드로 강제하지 않는다.

### 6.3 UIA 흐름 (MODE 3)

1. `Desktop(backend="uia").windows()`에서 제목에 `FORESTBOND` 포함 창 → 면적 최대. **프로세스명 필터 없음.**  
2. `Document` descendants의 `Text` 우선, 없으면 창 전체 `Text`.  
3. `window_text`를 줄 단위로 분해 · strip. 바로 앞 줄이 시간 토큰(`권** (17:48:01) :` 또는 `(17:48:01) :`)이면 다음 줄의 `watermark_key`는 `(17:48:01) : {호가}`. 파서에 넘기는 `text`는 호가 조각 그대로.  
4. dedupe는 quote 문자열이 아니라 `watermark_key`. 같은 호가가 초만 다르면 둘 다 남는다. 시간 줄이 없으면 키 = 호가 조각 (기존과 동일).

한 Text에 시각+호가가 이미 붙어 있으면 재결합하지 않는다.

### 6.4 Watermark (신규 라인 게이트)

라인 fingerprint = `SHA1(UTF-8 watermark_key)` hex (`message_fingerprint`). MODE 1·2는 `watermark_key =` 원문 줄. MODE 3은 가능하면 `(시각) : 호가`.

게이트는 **맨 뒤 `WATERMARK_WINDOW`(2000)줄만** 본다. 앞쪽 히스토리는 비교하지 않는다. watermark는 삽입 순서 deque+set이며 창을 넘으면 가장 오래된 fp부터 삭제한다. 전체 덤프를 계속 스캔하면서 set만 FIFO로 지우면 과거 호가가 신규로 부활하므로 그렇게 하지 않는다.

`TJvRichEdit`+`WM_GETTEXT`는 스크롤과 무관하게 문서 전체를 주므로, 위로 스크롤해도 창 밖 옛줄은 검사하지 않으면 재오탐이 나지 않는다. (스크롤 재오탐은 UIA MODE 3 쪽 이슈다.)

| API | 동작 |
|-----|------|
| `initialize_watermark(false)` | 현재 덤프의 **꼬리 창** fp로 set을 **채움**. 이후 그 창에 새로 들어온 문자열만 통과. |
| `initialize_watermark(true)` | set을 **비움**. (이어지는 `get_new`가 꼬리 창을 신규로 반환 가능) |
| `get_new_message_lines` | 미초기화면 위 규칙 적용. 이후: 꼬리 창에서 set에 없는 줄만 반환하며 FIFO add. |
| `reseed_watermark_from_visible` | 꼬리 창 줄을 set에 **union**. 창 밖 fp는 FIFO로만 제거. |

`PROCESS_EXISTING_ON_START`:

- **`false` (운영 권장):** 기동 시점 **꼬리 창**은 스킵, **그 다음부터** 창에 들어온 줄만 처리. 중지 중 쌓인 줄은 재기동 시에도 그때 창 안이면 패스.  
- **`true`:** 기동 시 꼬리 창에 있는 줄도 신규로 처리 (파서·재현 테스트용). 잔여 체결분 재전송 위험.

워터마크는 **프로세스 메모리 set**이며 디스크에 저장하지 않는다. SENT/STOPPED로 프로세스가 끝나면 소멸한다.

### 6.5 세션 fingerprint (두 번째 층)

`WatcherSession.processed_fingerprints`는 파싱에 **성공한** 줄의 `SourceLine.watermark_key` SHA1이다.  
소스 watermark와 알고리즘은 같으나 역할이 다르다.

- watermark: “이 채팅 **식별 키**를 이미 소스에서 봤는가”  
- session set: “이 **식별 키**로 이미 매칭·처리했는가”  

MODE 3에서 파서 입력(`Quote.raw_line`)은 호가 조각이고, 세션·QUOTE_FOUND 로그는 `watermark_key`(시각+호가)를 쓴다. 그래서 같은 호가라도 초가 다르면 재기회다.

현재 test-loop에서 SENT 후 reseed해도 session set은 비우지 않으므로, **동일 watermark_key는 같은 프로세스에서 재트리거되지 않는다.**

---

## 7. 호가 파서 (`source/quote_parser.py`)

입력: 소스에서 온 한 줄 + 슬롯의 `instrument`, `yield_prefix`, `required_side`, `required_qty`.

### 7.1 종목 매칭

`build_target_pattern`: escape한 종목 문자열이 앞뒤로 숫자/`-`가 아닌 경계에서만 매칭.  
예: `25-11`은 `125-11`, `25-110`과 매칭되지 않음.

### 7.2 메타 (선택)

`보낸이 (HH:MM[:SS]) : …` 형이면 sender/timestamp를 뽑는다. 종목 검색은 **원문 전체**에서 수행한다.

### 7.3 호가 토큰

종목 직후부터:

```text
^\s+(?P<price>\d{2,3})\s*(?P<side>[+-]|사자|팔자)\s*
```

- `raw_token` = 가격+side 구간만 (strip). 수량은 `Quote.quantity`에 둔다 (`flip_side_token`이 토큰 끝을 뒤집기 위함).  
- side 뒤 수량: 없거나 `(` / `*` trailing만이면 **생략 = 100억**. `required_qty == 100`일 때만 통과.  
- 그 외는 `^\d+\s*억?`만 수량으로 인정 (`80`, `80억`, `100`, `100억`). `억` 정규화 후 숫자가 `required_qty`와 같아야 함.  
- `required_qty != 100`이면 수량 생략 호가는 거부.  
- 수량 뒤 trailing: 비거나 `(` / `*`로 시작. `있나요`, 추가 숫자, `ㅎㅈ` 등은 거부.  
- side 뒤가 숫자로 시작하지 않는 문자열(`있으신가요`)은 거부.

### 7.4 yield · side

| 자릿수 | 변환 |
|--------|------|
| 2 | `prefix + n/100` |
| 3 | `prefix + n/1000` |

`+`/`사자` → `BUY`, `-`/`팔자` → `SELL`.  
`required_side`가 BUY/SELL이면 다른 쪽은 `None`.

`Quote.raw_line` = 입력 라인 strip 전체 (소스가 준 문자열 그대로).

---

## 8. 트리거와 확정 메시지 (`core/trigger.py`)

### 8.1 `evaluate`

- `BID` (현물 매도): `pnl <= threshold` 이면 트리거. 임계값 부호는 강제하지 않음.  
- `OFFER` (현물 매수): `pnl >= threshold` 이면 트리거.  
- 그 외 looking_for → `ValueError`.

### 8.2 `flip_side_token` → `confirm_token`

토큰 끝의 `+`↔`-`, `사자`↔`팔자`.  
매칭용 side와 별개로, **전송 문장**에만 사용한다 (상대방에게 확정 호가 제시).

### 8.3 `format_message`

`MESSAGE_TEMPLATE.format(...)`. 기본 예: `{instrument} {confirm_token} ㅎㅈ`.  
`quantity != 100`이면 confirm_token 뒤에 ` {n}억`을 붙인다. 예: `25-10 695- 80억 ㅎㅈ`. 100억(생략·명시)은 기존 문장.

---

## 9. 전송 (`send/ui.py`)

1. 전송 프로세스 실행 여부 확인.  
2. 프로세스 PID + 제목 부분일치 최상위 창 중 점수 최대 선택.  
3. 복원·Show → 포그라운드 강제 (Alt / AttachThreadInput 등, `.env` 타이밍).  
4. `send_text`:  
   - TOPMOST on → 재포그라운드  
   - 창 사각형 × `(send_input_x, send_input_y)` 클릭  
   - 포그라운드·커서 하위 창이 대상 앱인지 검증 (아니면 `SendError` — Excel 오입력 방지)  
   - 클립보드 → Ctrl+V → Enter  
   - finally TOPMOST off  

진단: `--diagnose-send`로 HWND·비율·클릭 좌표 출력.

---

## 10. 메인 루프 (`main.py`)

### 10.1 CLI

| 플래그 | 동작 |
|--------|------|
| (기본) | `run_watcher` |
| `--config PATH` | dotenv 경로 (기본 `.env`) |
| `--diagnose-source` | 소스 diagnose 출력 후 종료 |
| `--diagnose-send` | 전송 diagnose 출력 후 종료 |
| `--test-send` | 샘플 Quote로 전송만 |
| `--test-parser LINE` | Excel 슬롯 기준 파서 시험 |

설정 로드 실패 exit **2**. 그 외 운영 오류 exit **1**.

### 10.2 시작 시퀀스

1. stop 플래그 파일 삭제 시도.  
2. Excel connect → `load_slots`(D2 한 슬롯 + E2) → F2 `WATCHING`, J2 Start Successful.  
3. 소스 창·전송 창 resolve.  
4. `initialize_watermark(PROCESS_EXISTING_ON_START)`.  
5. 폴링 루프.

소스/전송 창을 못 찾거나 읽기/전송/Excel 오류 시 즉시 `ERROR` (폴링 재시도 없음).

### 10.3 폴링 한 사이클

```text
stop flag? → STOPPED, return 0
lines = get_new_message_lines(...)
if lines: load_slots again (D2 row + E2 threshold)
for line in lines:
    parse_quote_line against the single loaded slot
    session fingerprint 중복이면 skip
    QUOTE_FOUND 로그 (raw_token + raw_line)
    write_yield_read_pnl (D{row} / F{row+offset})
    evaluate vs E2
    if not triggered:
        H2/I2/J2 Quote Skipped, WATCHING, continue
    format_message → send_text
    SENT 상태 기록
    → (현재 구현) reseed_watermark_from_visible → WATCHING → continue
sleep(POLL_INTERVAL_MS)
```

### 10.4 SENT 이후 동작 (현재 코드)

**Test-loop (활성):**

1. Excel에 `SENT` 및 last_quote / pnl / Message Sent 기록.  
2. `reseed_watermark_from_visible()` — 지금 화면 전체를 “이미 본 줄”로 표시 → **이후 새로 올라온 줄만** 다시 탐지.  
3. F2를 `WATCHING`으로 되돌리고 루프 계속.

**One-shot (코드에 주석으로 존재):**

- `log.info("EXIT"); return 0`  
- 주석을 해제하고 test-loop 블록을 비활성화하면 전송 후 프로세스 종료.

유지보수 시 **어느 쪽이 활성인지 `main.py` SENT 직후를 확인**할 것.

### 10.5 매칭 순서

`_match_quote`는 D2가 선택한 **슬롯 하나**만 파서에 넘긴다. 다른 `EXCEL_SLOT_ROWS` 종목은 같은 줄에 있어도 트리거되지 않는다.

---

## 11. VBA (`vba/KBondWatcher.bas`)

| 상수 | 의미 |
|------|------|
| `PROJECT_DIR` / `MAIN_PATH` / `CONFIG_PATH` | 설치 경로 (배포 PC에 맞게 수정) |
| `PYTHONW_PATH` | `pythonw.exe` — **venv의 pythonw 절대경로 권장** |
| `STOP_FLAG_PATH` | `.env`의 `STOP_FLAG_PATH`와 **반드시 동일** |
| `PID_PATH` | stop 경로와 같은 폴더의 `kbond_watcher.pid` (Python이 PID 기록) |

| Sub | 동작 |
|-----|------|
| `StartKBondWatcher` | 기존 워처 프로세스 종료 → stop 플래그 삭제 → F2~J2 클리어 → `pythonw main.py --config .env` 숨김 실행 |
| `StopKBondWatcher` | stop 플래그 생성 → PID/`main.py` 프로세스 `taskkill` → F2=`STOPPED`, J2=`(HH:MM:SS) Stopped` |

Python은 COM busy(`RPC_E_SERVERCALL_RETRYLATER`)를 수 초간 재시도하고, 계산 대기·라인 처리 중에도 stop 플래그를 본다. VBA Stop은 프로세스를 직접 죽이므로 워처가 이미 죽었거나 COM에 막혀 있어도 셀이 갱신된다.

VBA 셀 주소 상수(F2~J2)는 `.env`의 `EXCEL_*_CELL`과 일치해야 한다.

---

## 12. 로깅

- 로거 이름: `kbond_watcher`  
- RotatingFileHandler (2MB × 5) + 콘솔  
- 포맷: `%(asctime)s %(levelname)s %(message)s`

주요 키워드: `WATCHING`, `QUOTE_FOUND`, `NO_TRIGGER`, `TRIGGERED`, `SENT`, `STOPPED`, `ERROR`, `EXIT`(one-shot 시), `EXCEL_*`, `MESSAGE_SENT`, `source watermark`, `source watermark reseed`.

`QUOTE_FOUND`는 `raw_token`과 소스 식별 키(`watermark_key`)를 함께 남긴다. H2는 `{instrument} {raw_token}`만 사용한다.

---

## 13. 진단·테스트

```bat
cd <프로젝트>
.venv\Scripts\activate
python main.py --config .env --diagnose-source
python main.py --config .env --diagnose-send
python main.py --config .env --test-parser "25-10 23+"
python main.py --config .env --test-send
pytest -q
```

| 테스트 영역 | 파일 |
|-------------|------|
| MODE·클릭 비율 로드 | `tests/test_config_mode.py` |
| 파서 accept/reject | `tests/test_quote_parser.py` |
| Looking For·flip | `tests/test_trigger.py` |
| Excel 헬퍼 | `tests/test_excel_*.py` 등 |
| ElTree / send | 해당 test_* |

---

## 14. 운영·유지보수 체크리스트

1. `.env`의 `MODE`, `EXCEL_WORKBOOK`, `PROCESS_EXISTING_ON_START`, 클릭 비율, `STOP_FLAG_PATH` 확인.  
2. VBA 경로·`pythonw`·STOP 경로를 설치 PC에 맞게 수정.  
3. KBond/FORESTBOND·전송 대상·Excel을 연 뒤 diagnose.  
4. D2가 허용 목록의 한 종목을 가리키고, 그 행 E열이 0이 아닌 정수인지(부호=방향, 절댓값=억 수량), E2 임계점이 숫자인지 확인.  
5. SENT 후 동작이 test-loop인지 one-shot인지 `main.py` 확인.  
6. MODE 3 QUOTE_FOUND의 `raw_line=`은 `(시각) : 호가` 키. 파서 입력은 호가 조각.  
7. 동일 `watermark_key`는 세션 내 재처리되지 않음. 호가 문구만 같고 시각이 다르면 MODE 3에서 재기회.  
8. 감시 중 채팅을 위로 스크롤하면(특히 UIA) 예상 밖 줄이 신규로 보이거나 최신을 놓칠 수 있음 — 맨 아래 유지 권장.

### 일반적인 장애

| 증상 | 점검 |
|------|------|
| 시작 즉시 ERROR | 소스/전송 창, Excel 미실행, 설정 키 누락 |
| WATCHING인데 무반응 | Looking For·qty, `PROCESS_EXISTING=false`로 기존 화면만 있음, diagnose-source |
| PnL만 되고 전송 없음 | I2 vs 임계값·Looking For 방향 |
| 전송이 Excel로 감 | diagnose-send 좌표, TOPMOST/포커스, 클릭 비율 |
| MODE 1 트리 없음 | 채팅 분리창 `TJvRichEdit`, 또는 `TElTree`·채팅 창 위치 |

---

## 15. 모듈 책임 한눈에

```text
config.loader     → Config (불변)
source.*          → list[str] 신규 라인
quote_parser      → Optional[Quote]
excel.bridge      → slots, pnl, status cells
core.trigger      → TriggerResult, outbound text
send.ui           → UI paste+Enter
main              → 위 조립 + 세션 fingerprint + stop
```

데이터 객체:

- `Quote` — instrument, raw_line, raw_token, yield_value, side, optional sender/ts, fingerprint  
- `InstrumentSlot` — row, looking_for, required_side, qty_abs, yield_prefix, cells  
- `TriggerResult` — triggered, reason, pnl, quote  
- `WatcherSession` — processed_fingerprints, status  

이 명세와 코드가 어긋나면 **코드를 진실로** 삼고 문서를 갱신한다.
