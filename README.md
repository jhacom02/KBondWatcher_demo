# FORESTBOND → Excel → K-Bond Watcher (POC)

Windows 전용 one-shot 자동화 POC입니다.

흐름: Excel START → Python watcher → Chrome FORESTBOND(UIA) 메시지 감시
→ TARGET 호가 파싱 → Excel 입력/Calculate/P&L 읽기 → threshold 통과 시
axis.exe Client 상대좌표 클릭 후 `ㅎㅈ` 붙여넣기 → `READY_TO_SUBMIT` 후 종료.
**Enter는 보내지 않습니다.** 사용자가 K-Bond에서 직접 확정합니다.

## 요구 사항

- Windows 10/11
- Python 3.11+
- Excel (이미 인포맥스 연동된 계산식 포함 워크북)
- Chrome에서 FORESTBOND 화면 오픈
- K-Bond (`axis.exe`) 실행

## 1. 설치

```bat
cd C:\mycode\KBondHandler
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 2. 설정

```bat
copy config.env.example config.env
```

`config.env`에서 최소한 다음을 확인/수정합니다.

| 키 | 설명 |
|---|---|
| `TARGET` | 감시 종목 (예: `25-11`) |
| `YIELD_PREFIX` | 수익률 앞자리 (예: `4` → `23+` = 4.23) |
| `EXCEL_WORKBOOK` | 기본 `sample.xlsx` |
| `EXCEL_INPUT_CELL` / `EXCEL_PNL_CELL` | 노란 입력셀 / P&L 결과셀 |
| `PNL_THRESHOLD` | 트리거 기준 (기본 1,000,000) |
| `KBOND_PID` | `axis.exe` PID (`0`이면 프로세스명 검색) |
| `WIN_X` / `WIN_Y` | Client 영역 상대 클릭 좌표 (0~1) |
| `SEND_TEXT` | 기본 `ㅎㅈ` |
| `PROCESS_EXISTING_ON_START` | POC는 `true`, 실운용은 `false` 권장 |

## 3. Excel 워크북

실사용 파일: `sample/sample.xlsx`

1. Excel에서 `sample/sample.xlsx`를 연다.
2. 노란색 입력 셀 / P&L / 상태 셀(J16~J19)이 config와 맞는지 확인한다.
3. VBA 모듈 `vba/ForestBondWatcher.bas`를 임포트한다.
4. START / STOP 버튼을 시트에 추가하고 각각 `StartForestBondWatcher` / `StopForestBondWatcher`에 연결한다.
5. `.bas` 상단 `PYTHONW_PATH`, `PROJECT_DIR`, `MAIN_PATH`, `CONFIG_PATH`를 환경에 맞게 수정한다.

## 4. axis.exe PID 확인

PowerShell:

```powershell
Get-Process axis | Select-Object Id, ProcessName
```

값을 `config.env`의 `KBOND_PID`에 넣습니다. 프로세스가 하나뿐이면 `KBOND_PID=0`도 가능합니다.

## 5. 진단 순서 (권장)

Chrome FORESTBOND와 Excel(`sample.xlsx`), K-Bond를 켠 뒤:

### 5-1. K-Bond HWND / 클릭 좌표

```bat
python main.py --config config.env --diagnose-kbond
```

PID, HWND, WindowRect, ClientRect, 계산된 screen click 좌표가 출력됩니다.

### 5-2. K-Bond prefill (Enter 없음)

```bat
python main.py --config config.env --prefill-kbond
```

지정 상대좌표 클릭 후 `ㅎㅈ`만 붙여넣습니다. **Enter는 자동으로 누르지 않습니다.**
입력창에 정확히 들어갔는지 확인한 뒤 사용자가 직접 Enter로 확정/취소합니다.

### 5-3. FORESTBOND Chrome UIA

```bat
python main.py --config config.env --diagnose-chrome
```

창 제목, HWND, Text control 수, 인식된 메시지 목록을 확인합니다.

### 5-4. Parser

```bat
python main.py --config config.env --test-parser "25-11 23+"
```

예상 출력:

```
instrument = 25-11
yield = 4.23
side = BUY
raw_token = 23+
```

### 5-5. 단위 테스트

```bat
pytest -q
```

## 6. 전체 one-shot 실행

1. Chrome FORESTBOND 오픈
2. `sample/sample.xlsx` 오픈 (인포맥스 연동·수식 준비)
3. `axis.exe` 실행
4. Excel START 버튼 클릭 (또는 `python main.py --config config.env`)

동작:

1. STATUS=`WATCHING`
2. TARGET 호가 감지 → Excel 입력 → Calculate → P&L 읽기
3. threshold 미달이면 계속 감시
4. threshold 통과 시 K-Bond 클릭 + `ㅎㅈ` paste → STATUS=`READY_TO_SUBMIT` → Python 종료
5. 사용자가 K-Bond에서 Enter로 최종 확정
6. 다시 감시하려면 START를 다시 누른다 (자동 재시작 없음)

STOP 버튼은 `STOP_FLAG_PATH` 파일을 만들어 watcher를 `STOPPED`로 종료합니다.

## 7. 로그

`logs/watcher.log` (RotatingFileHandler)

예시:

```
2026-08-10 16:20:30 INFO WATCHING
2026-08-10 16:20:34 INFO QUOTE_FOUND | 25-11 | 23+ | 4.230 | BUY
2026-08-10 16:20:34 INFO EXCEL_WRITE | D19=4.23
2026-08-10 16:20:34 INFO PNL | F22=1532000
2026-08-10 16:20:34 INFO TRIGGERED
2026-08-10 16:20:35 INFO READY_TO_SUBMIT
2026-08-10 16:20:35 INFO EXIT
```

## 8. 안전 규칙

- `axis.exe`에 대해 Enter / WM_KEYDOWN / SendMessage submit 자동화를 하지 않습니다.
- OCR / Selenium / Chrome DevTools API를 사용하지 않습니다.
- Excel COM은 watcher 단일 스레드에서만 조작합니다.

## 9. 프로젝트 구조

```
KBondHandler/
├─ main.py
├─ config.py
├─ forestbond_reader.py
├─ quote_parser.py
├─ excel_bridge.py
├─ kbond_controller.py
├─ models.py
├─ logger.py
├─ config.env / config.env.example
├─ sample/sample.xlsx
├─ vba/ForestBondWatcher.bas
├─ logs/
└─ tests/
```
