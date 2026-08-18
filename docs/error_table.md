# KBondWatcher 에러 표

운영 실패는 F2 `ERROR`, J2 `(HH:MM:SS) Error: {원인}` (최대 200자), `pythonw` **exit 1**. Config는 같은 셀 기록을 시도한 뒤 **exit 2**. VBA Fail은 Python보다 앞이며 J2 시각 형식이 `HH:nn:ss`다.

메시지 템플릿 단위다. Config 키 하나하나를 행으로 쪼개지 않았다 (`missing required config key`는 키가 달라도 같은 유형). 파서 실패는 예외 없이 `None`을 돌려 아래 “에러 아님”과 같다.

## Config — exit 2

| 조건 | 메시지 | 의미 |
|------|--------|------|
| `.env` 파일 없음 | `config file not found: {path}` | 설정 파일을 열 수 없음. stderr + Excel ERROR 시도 후 종료 |
| 필수 키 없음 | `missing required config key: {key}` | MODE, EXCEL_*, SEND_*_PAUSE, SENT_AFTER, LOG_* 등이 파일/환경에 없음 |
| 형식 오류 | `{key} must be a boolean/number/integer` / `column letter` / `between 0 and 1` / `must not be empty` | 값 타입·범위가 잘못됨 |
| MODE 범위 밖 | `MODE must be 1, 2, or 3, got {mode}` | 1·2·3만 허용 |
| SENT_AFTER 불량 | `SENT_AFTER must be exit or loop` | `exit`=전송 후 종료, `loop`=reseed 후 계속 |
| MODE 1·2인데 채팅 제목 없음 | `KBOND_CHAT_TITLE must not be empty` | 분리창 제목 필터가 비어 있음. MODE 3은 이 키 불필요 |
| 폴링 너무 짧음 | `POLL_INTERVAL_MS must be >= 100` | CPU·UI 부하 하한 |
| 빈 경로/셀/템플릿 | `EXCEL_WORKBOOK` / `MESSAGE_TEMPLATE` / `STOP_FLAG_PATH` / `LOG_PATH` / `{cell_key} must not be empty` | 워크북·상태셀·템플릿·로그 경로가 빈 문자열 |
| 오프셋·밴드 | `EXCEL_PNL_ROW_OFFSET must be >= 0` / `EXCEL_PNL_SANITY_BAND must be > 0` | PnL 행 오프셋과 sanity band 제약 |
| 슬롯 행이 만기 목록 밖 | `EXCEL_SLOT_ROWS contains {row} not in EXCEL_ROWS_10Y/3Y` | 허용 슬롯이 3Y/10Y prefix 매핑에 없음 |

## Excel · PnL — exit 1

| 조건 | 메시지 | 의미 |
|------|--------|------|
| Excel 미실행·COM 없음 | `pywin32 is required` / `Failed to connect to running Excel.Application: {exc}` | pywin32 미설치이거나 GetObject/ROT가 아닌 이유로 COM 연결 실패. 워크북이 닫힌 경우는 아래 EXCEL_WAIT |
| 워크북 FullName 읽기 실패 | `Failed to read FullName for workbook {name}` | 열린 객체의 FullName COM 실패. 종료 |
| 시트 없음 | `Worksheet '{sheet}' not found` / `No ActiveSheet available` | `EXCEL_SHEET` 이름 오류 또는 ActiveSheet 없음 |
| Excel busy 5초 초과 | `Excel busy after 50 retries: {last}` | RPC busy를 50×0.1s 재시도한 뒤 포기 |
| 감시 셀 수식 오류 | `{cell} formula is not a single A{row} ref` / `must reference column A` / `must reference the current sheet` / `row N is not in EXCEL_SLOT_ROWS` / `is empty` / `instrument matches N slot rows` | D2가 `=A41` 형태가 아니거나 허용 행이 아님. 종목 문자열이 0건·2건 매칭 |
| 슬롯 값 불량 | `{A}{row} is empty` / `{E}{row} must be a non-zero integer` / `row N is not mapped to 3Y or 10Y` / `failed to read yield prefixes` / `slot row allowlist is empty` | 종목 빈칸, E열 0·소수, prefix 셀 비숫자 |
| 셀이 숫자가 아님 | `Excel cell is #VALUE!` (`#NULL!` `#DIV/0!` `#REF!` `#NAME?` `#NUM!` `#N/A`) / `empty/None` / `blank` / `unexpected boolean` / `cannot convert…` | CVErr·빈칸을 PnL로 쓰지 않음. 대기 루프 안에서는 재시도 |
| 30초 동안 숫자 PnL 없음 | `{pnl_cell} not numeric after 30.0s ({detail}, CalculationState={state}, value=…)` | xlDone이 안 되거나 `#VALUE!`/빈칸이 유지됨. `Calculate` 강제 없음 |
| CalculationState 읽기 실패 | `Failed to read Excel CalculationState: {exc}` | COM이 계산 상태를 반환하지 못함 |
| 상태 셀 쓰기 실패 | `Excel status update failed: {exc}` | F2~J2 쓰기 실패는 즉시 종료. ERROR 기록 자체 실패만 로그 |
| \|pnl − threshold\| > band | `PnL {pnl} outside sanity band {band} of threshold {threshold}` | E2 기준 ±band. I2에 해당 pnl 기록 후 종료 |
| 임계값 미로드 | `PnL threshold is not loaded` | `load_slots`가 threshold 없이 진행된 내부 불일치 |
| 감시 슬롯이 PnL 중 변경 | `watch slot changed during PnL ({before} -> {after})` | evaluate 직전 재로드한 종목·방향·수량·임계가 호가 시점과 다름 |
| 한 폴에 호가 2건 이상 | `ambiguous quotes in one poll: {n}` | 어느 줄이 대상인지 모호. 전송하지 않음 |
| 확정 토큰을 뒤집을 수 없음 | `cannot flip side token: {raw_token}` | 트리거 후 `format_message`. 파서가 통과한 토큰이면 드묾 |
| evaluate looking_for 이상 | `looking_for must be BID or OFFER, got {value}` | 슬롯 내부값은 BID/OFFER. G2 표시용 `25-11 / BID`는 evaluate에 안 넘김 |

## 소스 — exit 1

| 조건 | 메시지 | 의미 |
|------|--------|------|
| MODE 비정상 | `unsupported MODE: {mode}` | loader를 우회한 경우에만 |
| 프로세스/제목 설정 빠짐 | `source_process_name is required` / `source_window_title is required` | MODE 1·2 생성자 가드 |
| 메신저 미실행 | `process not running: 'KBondMessenger.exe'` | 프로세스가 없음 |
| 방 제목 0건 | `no visible TJvRichEdit whose window title contains '{needle}'` | `KBOND_CHAT_TITLE`과 맞는 가시 분리창이 없음 |
| 방 제목 2창 이상 | `ambiguous chat windows matching '{needle}': [titles]` | 같은 문자열이 서로 다른 창에 걸림 |
| 창은 맞는데 본문 컨트롤 없음 | `no visible TJvRichEdit chat pane in window matching '{needle}'` | 제목은 맞지만 면적/높이 조건의 채팅 RichEdit가 없음 |
| 핸들 무효 | `TJvRichEdit hwnd not resolved` | 캐시된 채팅 컨트롤을 읽기 전에 창을 못 잡음 |
| 문서 >200만자, 프로세스 메모리 읽기 실패 | `OpenProcess` / `IsWow64Process` / `VirtualAllocEx` / `WriteProcessMemory` / `ReadProcessMemory failed` | `EM_GETTEXTRANGE`용 원격 버퍼 실패. 앞에서 자른 GETTEXT로 폴백하지 않음 |
| 본문 읽기 API 실패 | `WM_GETTEXTLENGTH failed` / `WM_GETTEXT failed: {exc}` | 길이·본문 SendMessage 실패 |
| FORESTBOND 창 없음 (MODE 3) | `window containing title 'FORESTBOND' not found` | 제목 부분일치 창이 데스크톱에 없음 |
| UIA 열거 실패 | `Failed to enumerate UIA desktop windows` / `Failed to read FORESTBOND window size` / `Failed to enumerate Text controls` / `Failed to read UIA Text: {exc}` | Text 열거 예외는 최대 2회 재시도 후 ERROR. 그 외 폴백 없음 |
| Text 컨트롤 0건 | `no UIA Text controls` | 창은 있으나 Text 역할 컨트롤이 없음 |

## 전송 — exit 1

| 조건 | 메시지 | 의미 |
|------|--------|------|
| 전송 프로세스 없음 | `{process} is not running` | MODE 1은 KBondMessenger, 2·3은 notepad.exe |
| 전송 창 제목 없음 | `window containing title '{title}' not found` | MODE 1은 `KBOND_CHAT_TITLE`, 2·3은 `메모장` |
| 포그라운드 실패 | `failed to foreground hwnd={hwnd}` / `invalid window handle` | Alt/AttachThreadInput으로도 전경이 안 됨. foreground 실패는 최대 2회 재시도 후 ERROR (`invalid window handle`는 재시도 없음) |
| 클릭 후 포커스가 다른 앱 | `send focus not on target fg='…' under='…'` | Excel 등 다른 창에 붙여넣기 방지. foreground/focus 오류는 최대 2회 재시도 후 ERROR |
| 클립보드 불일치 | `clipboard mismatch after copy` | `copy` 직후 클립보드가 전송문과 다름. Ctrl+V 전에 중단 |
| 창 크기·비교 실패 | `invalid window size` / `failed to compare send target windows: {exc}` | rect가 0이거나 HWND 비교 실패 |

## VBA (Python보다 앞)

| 조건 | 메시지 | 의미 |
|------|--------|------|
| Start/Stop Fail | `(HH:nn:ss) Error: {Err.Description}` | pythonw가 안 뜬 상태일 수 있음. `.bas`를 엑셀에 다시 넣어야 최신 로직 적용 |
| taskkill 비정상 코드 | `taskkill failed: {rc}` | 소프트 스톱(최대 8s) 뒤 최후 `/F`. 0=성공, 128=이미 없음(정상). 그 외는 Fail |
| PowerShell sweep 종료코드 ≠ 0 | `Stop-Process failed: {rc}` | python/pythonw의 `main.py`만 대상. Excel은 죽이지 않음 |

## 에러가 아닌 것 (중단하지 않음)

| 조건 | 보이는 것 | 의미 |
|------|-----------|------|
| 호가 미매칭 | 없음 | 감시 종목·side·수량이 안 맞으면 다음 줄 |
| fingerprint 중복 | 없음 | MODE 1·2: 같은 `watermark_key`는 세션 내 재처리 안 함. MODE 3는 세션 set을 쓰지 않고, 리더가 개수가 늘어난 줄만 검토 |
| 임계 미달 | J2 `Quote Skipped` | PnL이 BID/OFFER 조건을 못 채움. F2는 WATCHING |
| 정상 중지 | J2 `Stopped`, F2 `STOPPED`, exit 0 | stop 플래그 또는 `StopRequested` |
| Excel 닫힘 대기 | J2 `Excel closed; waiting to reopen {파일명}`, F2 `EXCEL_WAIT` | 워크북/RPC 소멸. 파일이 다시 열리면 WATCHING. 전송 없음. exit 하지 않음 |
| `SENT_AFTER=exit` 전송 성공 | J2 `Message Sent`, F2 `SENT`, exit 0 | 한 번 보내고 종료 |
| 문서 > 200만 wchar | 종료 없음 | 끝 200만만 읽음. 읽기 실패만 ERROR |

PID/stop 파일 정리 실패, ERROR 셀 쓰기 실패, `CoUninitialize` 실패는 이미 종료 중이라 `log.error`만 남긴다.
