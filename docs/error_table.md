# KBondWatcher 에러 표 (Trader Web)

트레이더 Local Web에서 보이는 메시지 기준.  
운영 실패는 Status `ERROR` + Last Action/`last_error`, watcher **exit 1**. Profile·START 거부는 formFlash(빨간 글씨).

애매하면 **전송하지 않고 즉시 ERROR**. 폴백 없음. VBA·`.env`·Excel F2–J2 UI 없음.

**대응**은 트레이더가 할 일을 짧게 적는다. 코드/배포 문제는 `개발자 패치 요청`.

---

## Profile / Save·Submit (formFlash profile)

| 조건 | 메시지 | 의미 | 대응 |
|------|--------|------|------|
| Name 빈칸 | `profile_name is required` | 필수 | Name 입력 후 Save |
| 종목 불허 | `instrument must be one of …` | 허용 종목만 | 목록에서 선택 |
| Looking For 이상 | `looking_for must be BID or OFFER` | | BID/OFFER 선택 |
| 수량 ≤0 | `required_qty must be > 0` | | 양의 정수 |
| 임계 비숫자 | `threshold must be numeric` | | 숫자 입력 |
| 임계 연산자 | `threshold_op must be <= or >=` | | ≤ / ≥ 선택 |
| 워크북 없음 | `excel_workbook FullName is required` | | Find로 열린 워크북 선택 |
| 워크북이 파일명만 | `excel_workbook must be a FullName path, not a bare Name` | 전체 경로 필요 | Find 다시 |
| 시트 없음 | `excel_sheet is required` | | Sheet 선택 |
| 셀 형식 | `yield_input_cell` / `pnl_cell` `must be a cell like D19` | | 예: D41, F44 |
| Yield Prefix | `yield_prefix must be 3 or 4` | | 3 또는 4 |
| Mode | `mode must be 1, 2, or 3` | | Mode 선택 |
| MODE 1·2 채팅제목 | `kbond_chat_title is required for MODE 1/2` | | Chat Title 입력 |
| Loop | `sent_after must be exit or loop` | | exit/loop |
| MODE 1 + loop | `sent_after must be exit when mode is 1` | Mode 1은 exit만 | exit로 변경 |
| 감시 중 프로필 변경 | `stop watcher before changing profile` | START 중 Save/Submit/Calibrate 금지 | STOP 후 재시도 |
| Runtime Save | `Runtime saved (no re-approve) · …` | 잠금 필드 동일 → `profile.json`만 갱신, 서명·version 유지 | START 가능; Admin 재승인 불필요 |
| Draft Save | `Draft saved — Submit & Admin Approve…` | 잠금 필드 변경 또는 미승인 | Submit→Admin Approve→적용 |
| 엔진 업그레이드 후 서명 거부 | `profile signature invalid` | policy 전용 서명으로 전환됨 | 한 번 재 Submit→Approve |
| Admin URL 없음 | `KBOND_ADMIN_URL required` | Submit에 Admin 필요 | start.bat/환경에 URL 설정, 개발자 확인 |
| draft 없음 | `no draft: …` | Submit 전 Save 필요 | Save 후 Submit |
| Admin HTTP/네트워크 | `admin HTTP {code}: …` / `admin request failed: …` | Admin 거부·끊김 | Admin 상태·URL·HTTPS 확인; 지속 시 개발자 |
| Pilot HTTP URL | `pilot mode requires HTTPS KBOND_ADMIN_URL` | Pilot는 HTTPS만 | URL을 HTTPS로 |
| 열린 워크북 0개 | `No open Excel workbook.` | Find 결과 | Excel에서 대상 xlsm 연 뒤 Find |
| schema 불일치 | `unsupported profile_schema_version …` | 엔진·프로필 포맷 불일치 | 개발자 패치/재배포 |
| Not Authorized | `profile not authorized` | 서명 미적용 | Profile Submit→Approve→자동 apply 대기 |
| Settings lease 만료 | `license lease expired` | lease 만료 | Admin 켜고 lease 갱신 후 Save |
| Settings 잠금 변경 | `locked fields cannot change in Settings` | Settings는 runtime만 | Profile 탭에서 Submit |
| Profile Revert | `↺ Revert` (Not Authorized·수정 시) | 마지막 서명본 또는 기본값 | 클릭 시 폼만 복원 (Save 전) |
| Settings Revert | `↺ Revert` (저장 후 수정 시) | 마지막 저장 runtime | 클릭 시 Settings 폼 복원 |

---

## Watcher / Coordinate

| 조건 | 메시지 | 의미 | 대응 |
|------|--------|------|------|
| Test Click 성공 | `Test click ok` | 좌표 클릭만 (붙여넣기 없음) | |
| Calibrate/Test while START | `stop watcher before changing profile` | 감시 중 좌표 변경 금지 | STOP 후 |
| Not Authorized + START/STOP/좌표 | `profile not authorized` | 승인 전 Watcher 조작 불가 | Profile 승인 |

---

## START / STOP (formFlash status)


| 조건 | 메시지 | 의미 | 대응 |
|------|--------|------|------|
| Profile 검증 실패 | (위 Profile 메시지와 동일) | START 전 validate | 필드 수정 후 START |
| 이미 실행 중 | `watcher already running pid=…` | 이중 START | STOP 후 START; 안 되면 작업관리자에서 watcher/`main` 종료 |
| START 성공 | `started pid=…` | 정상 | — |
| Job 배정 실패(경고) | `started pid=… · console-close kill unavailable` | 콘솔 강제종료 시 자식이 남을 수 있음 | STOP/탭닫기로 중단; 가능하면 콘솔를 직접 실행 |
| STOP 성공 | `stopped` | | — |
| STOP 강제종료 실패 | `failed to terminate pid=…` | 프로세스 안 죽음 | 작업관리자에서 해당 PID 종료 후 재실행 |
| 기기 미활성화 | `device is not activated` | Admin 미승인 | Admin에서 device activate |
| 원격 disable | `device is remotely disabled` / `remotely disabled` | Admin 차단 | Admin Enable 요청 |
| Lease 없음 | `license lease missing` | | Admin URL·네트워크 확인 후 재START; 개발자 |
| Lease 서명 불량 | `license lease signature invalid` | 키/위조 | `KBOND_SIGNING_PUBLIC_KEY`·Admin 키 확인, 개발자 |
| Lease device 불일치 | `license lease device mismatch` | | 재등록/개발자 |
| Lease disabled | `license lease disabled` | | Admin 확인 |
| Lease 만료 | `license lease expired` | | Admin lease 재발급·재START |
| profile_version 불일치 | `lease profile_version … != local …` | 승인 프로필과 로컬 불일치 | Admin 승인 프로필 동기화(약 60초) 후 START |
| 엔진 버전 부족 | `engine … below minimum …` | min_engine | 최신 패키지 설치, 개발자 |
| 프로필 서명 없음/불량 | `profile signature required` / `profile signature invalid` | 미승인·키 없음·위조 | Submit→Approve→자동적용 후 START; public key 설정 확인 |
| Demo 만료 | `demo expired on …` | 파일 날짜 지남 | 개발자(새 빌드/expiry) |
| Demo 파일 없음/불량 | `demo_expiry.txt missing…` / `invalid date` / `empty` | Pilot fail-closed | 개발자 패치 |
| credential | `device credential missing` / `failed to unprotect…` | DPAPI/자격 | 같은 Windows 계정으로 재실행; 안 되면 개발자 |
| Pilot 로컬 lease 금지 | `local lease issuance forbidden in pilot mode` | Admin lease만 | Admin·URL 확인 |
| Excel preflight | `excel preflight failed: …` / `workbook FullName mismatch …` | 워크북 미오픈·경로 불일치 | Excel에서 Profile과 같은 파일 연 뒤 START |
| Source preflight | `source preflight failed: …` | 채팅 창 없음 | KBond/FORESTBOND·제목 확인 후 재실행 |
| Send preflight | `send target preflight failed: …` | 전송 창 없음 | Notepad/KBond 전송창 연 뒤 재실행 |
| 탭 닫기/새로고침 | (자동 STOP) | 의도적 중단 | 다시 보려면 START |

---

## Coordinate (formFlash coord)

| 조건 | 메시지 | 의미 | 대응 |
|------|--------|------|------|
| 전송 창 없음 | `send target window not found` | | Mode에 맞는 창 연 뒤 Calibrate |
| 창 크기 이상 | `invalid send window size` | | 창 최대화/복원 후 재시도 |
| 클릭 타임아웃 | `timed out waiting for click` | | 안내대로 대상 창을 1회 클릭 |

---

## Watcher ERROR — Excel · PnL (Status Last Action / last_error)

| 조건 | 메시지 | 의미 | 대응 |
|------|--------|------|------|
| Excel/COM 없음·끊김 | `Workbook '…' is not open in Excel` / `Failed to connect…` / `pywin32 is required…` | **재연결 없음·ERROR** | Excel에서 대상 파일 연 뒤 **재START**; COM 반복 시 작업관리자에서 Excel 종료 후 재오픈 |
| 시트 없음 | `Worksheet '{sheet}' not found` | Profile sheet 오류 | Sheet Name 수정·Save/Approve 후 START |
| Excel busy 초과 | `Excel busy after … retries` | RPC busy | Excel 응답 대기·다른 매크로 중단 후 재START; 반복 시 Excel 재시작 |
| PnL 비숫자/CVErr | `Excel cell is #VALUE!` 등 / `blank` / `cannot convert…` | | 시트 수식·입력셀 확인; 개발자/시트 담당 |
| PnL 대기 타임아웃 | `{pnl_cell} not numeric after 30.0s …` | 계산 미완료 | 자동계산·수식 확인 후 재START |
| Sanity band | `PnL {pnl} outside sanity band {band} of threshold {threshold}` | 비정상 PnL·**전송 없음** | threshold/시트 확인; 이상치면 개발자 |
| 확정 토큰 flip | `cannot flip side token: …` | | 해당 호가 형식 이슈 → 개발자 패치 요청 |

---

## Watcher ERROR — 소스 (채팅)

| 조건 | 메시지 | 의미 | 대응 |
|------|--------|------|------|
| 메신저 미실행 | `process not running: 'KBondMessenger.exe'` | MODE 1·2 | KBond 실행 후 재START |
| 방 제목 0건 | `no visible … title contains '…'` | Chat Title 불일치 | 분리창 제목·Profile Title 맞춘 뒤 재START |
| 방 제목 2건+ | `ambiguous chat windows matching '…'` | 동일 제목 창 복수 | 다른 창 닫거나 Title을 더  uniquely |
| 채팅 pane 없음 | `no visible … chat pane…` / `TJvRichEdit hwnd not resolved` | | 분리 채팅창 상태 확인·재START |
| FORESTBOND 없음 | `window containing title 'FORESTBOND' not found` | MODE 3 | FORESTBOND 실행 |
| UIA/메모리 읽기 실패 | `Failed to enumerate…` / `OpenProcess failed…` / `WM_GETTEXT…` / `no UIA Text controls` | | 창 다시 열고 재START; 반복 시 개발자 |
| 한 폴 호가 2건+ | `ambiguous quotes in one poll: {n}` | **전송 없음** | 조건/시장 확인 후 재START; 반복 시 개발자 |

---

## Watcher ERROR — 전송

| 조건 | 메시지 | 의미 | 대응 |
|------|--------|------|------|
| 전송 프로세스 없음 | `{process} is not running` | notepad/KBond | 대상 앱 실행 후 재START |
| 전송 창 없음 | `window containing title '…' not found` | | 메모장/채팅 입력창 연 뒤 재START |
| 포그라운드 실패 | `failed to foreground hwnd=…` | | 창을 앞으로·다른 전체화면 닫고 재START |
| 포커스 이탈 | `send focus not on target …` | 다른 창에 붙여넣기 방지 | 전송 창만 두고 재START; Coordinate 재설정 |
| 클립보드 불일치 | `clipboard mismatch after copy` | | 재START; 반복 시 개발자 |
| 창 크기/핸들 | `invalid window size` / `invalid window handle` | | 창 복원 후 재START |

---

## 에러가 아닌 것 (중단·전송 없음 또는 정상)

| 조건 | 보이는 것 | 의미 | 대응 |
|------|-----------|------|------|
| 호가 미매칭 | (변화 없음) | 종목·side·수량 불일치 | Profile/Looking For 확인 |
| 임계 미달 | Last Action `Quote Skipped`, WATCHING | NO_TRIGGER | 정상; threshold 조정은 필요 시 |
| 정상 STOP | `stopped` / STOPPED | | — |
| 전송 성공 exit | SENT 후 종료 | `sent_after=exit` | 다시 보려면 START |
| 전송 성공 loop | SENT → WATCHING | 데모 loop | STOP으로만 완전 중단 |
| 프로필 동기화 대기 | status에 승인/적용 안내 | Admin Approve 대기 | Approve 후 약 60초 대기 |
| MODE 3 UIA 고점 | (변화 없음) | 세션 고점보다 문구 개수가 늘어난 줄만 검토. 칸 깜빡임 복원은 재전송 없음 | 정상 |
