# 플랜: 화면 읽기·메시지 전송을 KBond로 전환

이 문서는 현재 `KBondWatcher`(UIA 읽기 `READ_WINDOW_TITLE` → Excel → 범용 UI 전송 `SEND_*`)를  
**KBond 메신저에서 읽고 / KBond 메신저로 보내는** 구조로 바꿀 때 개발자가 따라갈 구현 플랜이다.

구현 코드는 포함하지 않는다. 변경 범위·좌표·설정·검증 순서만 정의한다.

참고 화면: 사용자가 제공한 KBond 스크린샷  
(우측 세로 채팅방 하단 **빨간 네모** = 확정 메시지 입력란)

---

## 0. 목표 요약

| 구분 | 현재 | 전환 후 |
|------|------|---------|
| 메시지 소스 | UIA + `READ_WINDOW_TITLE` (예: FORESTBOND), `forestbond_reader.py` | 동일 `READ_*` 키로 KBond 창 제목·(추후) 프로세스 |
| 메시지 액션 | `message_sender.py` + `SEND_*` (기본 메모장, 클릭→paste→Enter) | `.env`의 `SEND_*`만 KBond로 교체 |
| Excel / 파서 / 트리거 | 유지 | **유지** |

두 전환은 독립적으로 가능하다.

- **Track A**: 소스만 KBond (`READ_WINDOW_TITLE` 값·reader 구현)
- **Track B**: 전송만 KBond (`SEND_*` 값) — **시퀀스는 이미 구현됨**
- **Track A+B**: 둘 다 KBond (최종 권장)

권장 구현 순서: **창/컨트롤 진단 → Track A → Track B(`.env` SEND_* 교체)**.

---

## 1. 화면 구조 (기준 스크린샷)

KBond는 단일 메인 프레임 안에 **여러 채팅 패널(MDI/도킹)** 이 배치된다.

```
┌─────────────────────────────────────────────────────────────┐
│  KBond 메인 (메뉴/툴바)                                      │
├──────────┬──────────┬──────────┬────────────────────────────┤
│ 패널1    │ 패널2    │ 패널3    │                            │
├──────────┼──────────┤          │  우측 세로 채팅방 (타깃)    │
│ 패널4    │ 패널5    │          │  ┌────────────────┬─────┐  │
│          │          │          │  │ 메시지 목록     │참가자│  │
│          │          │          │  │ (읽기 영역)     │목록  │  │
│          │          │          │  ├────────────────┴─────┤  │
│          │          │          │  │ [입력란 ← 빨간 네모]  │  │
└──────────┴──────────┴──────────┴────────────────────────────┘
```

전환 시 핵심은 **우측 세로 채팅방 하나**다.

| 영역 | 역할 |
|------|------|
| 메시지 목록 (입력란 위) | Track A: 신규 호가 라인 수집 |
| 참가자 목록 (우측 세로) | 읽기/클릭 대상 아님 (오탐 방지) |
| 하단 입력란 (빨간 네모) | Track B: 포커스 후 확정 메시지 전송 |

메시지 라인은 `[시:분:초]` + 본문 형태가 보이며, 배경색(노랑/녹/청 등)은 UIA/텍스트 수집과 무관하다.  
`quote_parser`의 TARGET+호가 규칙은 **라인 문자열만** 보면 되므로 파서 자체는 유지한다.  
다만 KBond 라인이 FORESTBOND와 형식이 다르면 파서 수락/거부 케이스만 추가한다 (섹션 8).

---

## 2. 좌표 (메인 창 client/window 비율, 스크린샷 기준)

좌표는 **KBond 메인 최상위 창의 `GetWindowRect` (left, top, width, height)** 기준 상대비율이다.  
자식 패널 HWND를 따로 쓰지 않고 메인 창에 클릭하는 방식을 1차 기본으로 한다.

### 2.1 전송 입력란 (빨간 네모) — Track B

스크린샷에서 입력란 대략 범위:

| | 비율 (메인 창 대비) |
|--|-------------------|
| 왼쪽 `X0` | **0.72** |
| 오른쪽 `X1` | **0.93** |
| 위 `Y0` | **0.92** |
| 아래 `Y1` | **0.96** |

**클릭 기본값 (중심점)** — `.env` 초안:

```env
KBOND_INPUT_X=0.825
KBOND_INPUT_Y=0.940
```

계산: `((0.72+0.93)/2, (0.92+0.96)/2) ≈ (0.825, 0.940)`

절대 픽셀:

```text
x = left + int(width  * KBOND_INPUT_X)
y = top  + int(height * KBOND_INPUT_Y)
```

주의:

- 참가자 목록이 입력란 오른쪽에 있으므로 **X를 0.95 이상으로 올리면 목록을 클릭**할 수 있다. 상한 권장 `≤ 0.90`.
- 입력란보다 위(`Y < 0.90`)면 메시지 본문을 클릭한다.
- 창 최대화/복원, DPI, 해상도가 바뀌면 비율 재측정이 필요하다. diagnose에서 클릭 전 좌표를 로그로 남긴다.

### 2.2 읽기 영역 (우측 패널 메시지 목록) — Track A 참고용

클릭용이 아니라, UIA/텍스트가 안 잡힐 때 ROI·진단·자식 창 필터에 쓴다.

| | 대략 비율 |
|--|-----------|
| 메시지 목록 X | **0.70 ~ 0.93** (참가자 목록 제외) |
| 메시지 목록 Y | **0.08 ~ 0.90** (헤더 아래 ~ 입력란 위) |

UIA로 전체 Text를 긁은 뒤 `quote_parser`로 거르는 방식이 1차다.  
ROI는 “우측 패널만” 좁힐 때 또는 OCR 폴백 시에만 사용한다.

### 2.3 좌표 재측정 절차 (운영 PC에서 필수)

1. KBond를 **운영과 동일한 배치**(우측 세로방 고정)로 연다.
2. 진단 스크립트로 메인 HWND rect를 출력한다.
3. 입력란 중앙을 마우스로 찍고 화면 좌표 `(sx, sy)`를 구한다.
4. `KBOND_INPUT_X = (sx - left) / width`, `KBOND_INPUT_Y = (sy - top) / height` 로 `.env`를 갱신한다.
5. `--test-send`로 클릭→붙여넣기→Enter만 검증한다 (실제 방 전송 주의).

스크린샷 비율은 **초깃값**이다. 첫 배포 PC에서 한 번 재보정하는 것을 완료 조건에 넣는다.

---

## 3. 아키텍처 변경

### 3.1 유지

- `main.py` 감시 루프 골격 (폴링, stop flag, Excel STATUS 4종, one-shot SENT 종료)
- `excel_bridge.py`, `trigger.py`, `quote_parser.py`(규칙 동일 시), `models.py`, `logger.py`, VBA START/STOP
- Excel 셀·`PNL_THRESHOLD`·`MESSAGE_TEMPLATE` 의미

### 3.2 교체·추가

| 현재 | 전환 후 |
|------|---------|
| `forestbond_reader.py` | `kbond_reader.py` (또는 동일 파일 개조) |
| `message_sender.py` + `SEND_*` | 동일 유지 — `.env` 값만 KBond로 |
| `READ_WINDOW_TITLE` | 값은 KBond 창 제목; 필요 시 `READ_PROCESS_NAME` 추가 |
| `--diagnose-read` / `--diagnose-send` | CLI 이름 유지, 대상만 KBond |

### 3.3 `main.py` 연결 지점

현재:

```text
ForestBondReader → parse_quote_line → Excel → evaluate → message_sender.send_text(cfg)
```

목표:

```text
KBondReader → parse_quote_line → Excel → evaluate → message_sender.send_text(cfg)
```

루프·STATUS·에러 처리 패턴은 그대로 두고 **reader 구현 / `READ_*`·`SEND_*` 값**을 바꾼다.

---

## 4. Track A — 화면 읽기를 KBond로

### 4.1 선행 진단 (코딩 전 필수)

KBond가 떠 있는 PC에서 확인:

1. 프로세스명 (예: `KBond.exe` — **실측 후 `.env`에 기입**)
2. 메인 창 클래스명 / 제목
3. 채팅 패널이 **별도 HWND 자식**인지, 메인 클라이언트 위 **커스텀 그리기**인지
4. UIA `Text` / `Document` / `ListItem` 로 메시지 문자열이 보이는지
5. 가상화로 스크롤 밖 라인이 안 잡히는지

진단 CLI 요구사항:

- 메인 HWND, class, title, rect
- (가능하면) 자식 창 트리 요약
- UIA로 수집한 텍스트 라인 N개 덤프
- TARGET이 포함된 라인만 필터한 미리보기

UIA에 텍스트가 전혀 없으면 Win32 접근성/클립보드 복사 등 대안을 플랜 B로 문서화하고, OCR은 최후 수단으로만 검토한다.

### 4.2 `KBondReader` 책임 (현 `ForestBondReader`와 동일 계약)

구현이 맞추어야 할 공개 API (이름만 예시):

| 메서드 | 동작 |
|--------|------|
| `find_kbond_window()` | 프로세스+제목/클래스로 메인(또는 타깃 패널) 창 확보 |
| `get_visible_message_lines()` | 채팅 텍스트 라인 리스트 |
| `initialize_watermark(process_existing)` | 시작 시 기존 라인 처리 여부 |
| `get_new_message_lines(...)` | watermark 기반 신규만 반환 |
| `diagnose(...)` | 운영 점검용 덤프 |

watermark·fingerprint(SHA1) 로직은 현 `forestbond_reader.py`와 동일하게 가져가도 된다.

### 4.3 텍스트 수집 전략 (우선순위)

1. **UIA Text/Document descendants** (현 FORESTBOND와 동일 패턴) — KBond 메인 또는 우측 패널 wrapper부터.
2. 패널이 여러 개면:
   - 설정 `KBOND_CHAT_TITLE` (방 제목 부분문자열)로 자식/컨트롤 필터, 또는
   - ROI(섹션 2.2) 안에 bounding rect가 들어오는 Text만 채택.
3. 참가자 목록 텍스트(이름만 있는 짧은 줄)는 파서가 자연 거부하도록 두거나, ROI로 제외.

### 4.4 파서 연동

- 1차는 기존 `parse_quote_line` 재사용.
- KBond 라인에 타임스탬프·접두가 달라도 TARGET 경계 + `_PRICE_SIDE` fullmatch만 만족하면 통과한다.
- 실패 라인이 많으면 `tests/test_quote_parser.py`에 KBond 실라인을 accept/reject로 추가한다.
- `REQUIRED_SIDE` / `YIELD_PREFIX` / `TARGET` 의미 변경 없음.

### 4.5 설정 키 (Track A)

추가 예:

```env
KBOND_PROCESS_NAME=KBond.exe
KBOND_WINDOW_CLASS=
KBOND_WINDOW_TITLE=KBond
KBOND_CHAT_TITLE=
KBOND_READ_BACKEND=uia
```

제거(Track A 완료 후): `READ_WINDOW_TITLE` 및 Chrome 전용 diagnose.

`config.py`에 필수 검증 추가. 코드 하드코딩 금지(현 프로젝트 규칙).

### 4.6 `main.py` 변경 포인트 (Track A)

- `ForestBondReader` import/생성 → `KBondReader`
- `ForestBondReaderError` → `KBondReaderError` (또는 공통 `SourceReaderError`)
- `--diagnose-read` → KBond read diagnose
- 시작 시 `find_*` / watermark 호출은 동일 순서

Excel·트리거·`SEND_*` 전송 블록은 Track A만 할 때 건드리지 않는다.

---

## 5. Track B — 메시지 전송을 KBond로

### 5.1 현황

전송 시퀀스는 이미 범용 `SEND_*`로 구현되어 있다.

```text
1. SEND_PROCESS_NAME 실행 여부 확인
2. SEND_WINDOW_TITLE 포함 창 찾기 → activate
3. SEND_INPUT_X/Y 클릭
4. Ctrl+V → Enter
```

Track B는 `.env` 교체 + 좌표 재측이다.

```env
SEND_PROCESS_NAME=KBond.exe
SEND_WINDOW_TITLE=KBond
SEND_INPUT_X=0.825
SEND_INPUT_Y=0.940
```

스크린샷 기준 입력란 초깃값은 §2.1. `--diagnose-send` / `--test-send`로 검증.

비율 키는 **0~1** 검증 (`config.py`의 `_require_ratio` 재사용).

`MESSAGE_TEMPLATE`은 유지 (예: `{instrument} {raw_token} ㅎㅈ`).

### 5.2 포그라운드·포커스 리스크

- KBond가 다른 창에 가려지면 클릭이 빗나간다 → 전송 직전 반드시 activate.
- 입력란 클릭 후 잘못된 창에 paste되지 않도록 전경·제목 확인.
- 여러 모니터/DPI: `--diagnose-send`의 click_point 확인.

### 5.3 안전장치

- `--test-send`는 실제 창에 문구가 들어간다. KBond 전환 전 테스트 환경에서만 검증.

---

## 6. 설정·문서·VBA 정리

### 6.1 `.env` 최종 형태 (A+B)

유지: `TARGET`, `YIELD_*`, `REQUIRED_SIDE`, `POLL_*`, `PROCESS_EXISTING_*`, `EXCEL_*`, `PNL_*`, `MESSAGE_TEMPLATE`, `STOP_*`, `LOG_*`

읽기: `READ_WINDOW_TITLE` 값/구현. 전송: `SEND_*` 값 (섹션 5).

### 6.2 `config.py`

- 읽기 모듈을 KBond UIA에 맞게 교체; `SEND_*`/`READ_WINDOW_TITLE` 키 이름은 유지 가능
- 누락 키는 `ConfigError` (기본값 금지)

### 6.3 VBA

- START/STOP·상태셀 초기화 로직 변경 없음
- `PROJECT_DIR` / `MAIN_PATH` / `CONFIG_PATH`만 실제 `KBondWatcher` 경로와 일치시킬 것
- `STOP_FLAG_PATH`는 `.env`와 동일

### 6.4 문서

- `KBondWatcher_doc.md` / `README.md`의 소스 설명을 KBond 읽기 기준으로 갱신
- 본 플랜의 좌표·시퀀스를 운영 문서로 남기거나 요약 이관

---

## 7. 구현 작업 분해 (체크리스트)

### Phase 0 — 진단

- [ ] KBond 프로세스명·창 class/title 실측
- [ ] UIA로 우측 방 메시지 텍스트 수집 가능 여부 확인
- [ ] 입력란 클릭 비율을 운영 해상도에서 재측정 (`KBOND_INPUT_X/Y`)
- [ ] `--diagnose-kbond-*` 초안 CLI 추가

### Phase 1 — Track A (읽기)

- [ ] `kbond_reader.py` 구현 (watermark API 호환)
- [ ] `config` / `.env`에 read용 `KBOND_*` 추가
- [ ] `main.py` reader 교체
- [ ] 실라인으로 `quote_parser` 테스트 보강
- [ ] 폴링 중 신규 호가 → Excel 반영까지 확인 (전송은 기존 유지 가능)

### Phase 2 — Track B (전송)

- [ ] KBond `send_text` 구현 (클릭 → paste → Enter)
- [ ] 스크린샷 초깃값 `0.825 / 0.940` 적용 후 실측 보정
- [ ] `.env` `SEND_*`를 KBond로 변경 (코드 경로 유지)
- [ ] `--test-send`가 KBond로 나가게 변경
- [ ] 트리거 시 G2=`SENT`, 프로세스 종료 확인

### Phase 3 — 정리

- [ ] `forestbond_reader.py`를 KBond reader로 교체 또는 이름 정리
- [ ] pytest 전부 통과
- [ ] `KBondWatcher_doc.md` 갱신

---

## 8. 테스트 계획

| 항목 | 방법 |
|------|------|
| 창 찾기 | diagnose: 프로세스 미실행 / 실행 / 최소화 복원 |
| 읽기 | diagnose 덤프에 기대 호가 라인 포함 |
| 파서 | KBond 실문장 accept/reject 단위 테스트 |
| 클릭 | test-send 전 좌표 로그; 입력란에만 포커스되는지 육안 |
| 전송 | 테스트 문구 1회 → 방에 표시 → Enter 중복 전송 없는지 |
| E2E | 감시 → Excel P&L ≥ threshold → KBond 전송 → `SENT` → 종료 |
| STOP | 플래그 파일 → `STOPPED` |
| 회귀 | Excel STATUS 4종, 미트리거 시 `WATCHING`+J2 pnl 문구 |

---

## 9. 리스크와 대응

| 리스크 | 대응 |
|--------|------|
| UIA에 채팅 텍스트 없음 | 자식 HWND·다른 control_type 탐색; 안 되면 별도 접근 방식 재설계 |
| 패널 레이아웃 변경 시 좌표 이탈 | `.env` 비율만 재측정; 코드 수정 최소화 |
| 여러 채팅방 텍스트 혼입 | `KBOND_CHAT_TITLE` 또는 ROI 필터 |
| 숨김/최소화 KBond | activate 실패 시 `ERROR` |
| 잘못된 창에 붙여넣기 | 전송 직전 foreground가 KBond인지 확인 |
| 입력란에 잔여 텍스트 | `KBOND_CLEAR_INPUT_BEFORE_SEND` |
| DPI/멀티모니터 | 실측 좌표와 Win32 rect 불일치 시 보정 |

---

## 10. 완료 정의 (Definition of Done)

1. FORESTBOND Chrome 없이도 신규 호가를 KBond에서 읽어 Excel 계산까지 수행한다.
2. 트리거 시 우측 방 입력란(빨간 네모)에 확정 메시지가 전송된다 (`SEND_*` 교체).
3. 입력 클릭 비율이 `.env`로만 조정 가능하고, 초깃값은 본 문서의 **0.825 / 0.940**이다.
4. Excel STATUS는 계속 `WATCHING` / `SENT` / `STOPPED` / `ERROR`만 사용한다.
5. 운영 PC에서 diagnose + test-send + 1회 E2E가 문서화·통과된다.

---

## 11. 비범위

- KBond 자동 로그인·자동 실행
- 여러 채팅방에 동시에 보내기
- OCR 기본 경로
- 파서 비즈니스 규칙 전면 개편 (실라인 차이만 보정)
- Excel 수식/시트 구조 변경
