# KBondWatcher — 설계·로직 명세

Windows에서 채권 채팅 호가를 감시하고, Excel PnL 임계값을 만족하면 확정 메시지를 UI로 전송하는 워처이다.  
코드와 어긋나면 코드를 진실로 삼는다.

**Pilot-only:** 제어 UI는 Local Web (`--serve`). VBA·`.env`·Excel F2–J2 상태 UI는 제거됨. Excel은 **PnL COM**만 사용.

운영 원칙: **폴백 없음.** 애매하면 보내지 않고 ERROR.  
데모 예외: 전송 **성공 후** `sent_after=loop`로 재감시 가능.

| 문서 | 용도 |
|------|------|
| [`error_table.md`](error_table.md) | **운영 우선** — 트레이더 Web 에러·대응 |
| [`pilot_smoke.md`](pilot_smoke.md) | **배포 검증** 체크리스트 |
| 이 문서 | 아키텍처·MODE·루프 요약 (레거시 `.env`/VBA 서술 없음) |

---

## 1. 목적과 범위

### 1.1 하는 일

1. Trader Profile(+ machine 좌표)과 MODE 프리셋으로 소스·전송 대상을 정한다.
2. **Policy(서명·승인):** Name, Chat Title, Mode, Loop, Excel workbook/sheet, message template. **Runtime(STOP 후 Save, 재승인 없음):** instrument / looking_for / qty / threshold / yield_prefix / Input·Output cell.
3. 채팅에서 **신규 라인**만 골라 호가 파싱한다.
4. yield 셀에 쓰고 PnL을 읽은 뒤 sanity band·evaluate한다.
5. 조건이 맞으면 side-flip 확정 문자열을 전송 창에 붙여 넣고 Enter한다.
6. 상태는 `%LOCALAPPDATA%\KBondWatcher\runtime_status.json` + 웹 Status에 반영한다. Audit에 prefs 스냅샷이 포함된다.

### 1.2 하지 않는 일

- OCR, Selenium, 카카오톡 전용 연동 없음.
- Excel을 새로 기동하지 않는다. 이미 열린 워크북에만 붙는다.
- 한 폴에서 호가 2건 이상 매칭 시 보내지 않는다 (ERROR).
- Excel 끊김 시 재연결 대기 없이 ERROR.
- 잘못된 PnL(sanity band 밖)·전송 실패·창 ambiguous는 폴백 없이 ERROR.

### 1.3 런타임 전제

| 전제 | 설명 |
|------|------|
| OS | Windows |
| Excel | 대상 워크북 **열린 상태**, Automatic 계산 권장 |
| UI | `python main.py --serve` (또는 frozen `main.exe --serve`) |
| Watcher | START → `--run-profile` |
| MODE 1·2 소스 | `KBondMessenger.exe` + Profile 채팅 제목 |
| MODE 3 소스 | FORESTBOND UIA Text |
| 전송 좌표 | `machine.json` send_input_x/y |

---

## 2. 아키텍처

```text
main.py --serve
    → Local Web START
        → preflight (signed profile, lease, Excel bind)
        → spawn --run-profile
            → config_from_profile + slot_from_profile
            → ExcelBridge.connect / write_yield_read_pnl
            → create_source_reader(MODE)
            → loop:
                  stop flag? → STOPPED
                  ExcelDisconnected → ERROR (no reconnect)
                  get_new_message_lines
                  배치 매칭 (0 skip / 1 proceed / 2+ ERROR)
                  write yield → wait PnL → sanity → evaluate
                  skip 또는 send_text
                  sent_after=exit 종료 / loop 이면 reseed 후 WATCHING
```

| 경로 | 책임 |
|------|------|
| `main.py` | CLI (`--serve` / `--serve-admin` / `--run-profile`), 배치 매칭 헬퍼 |
| `app/watcher_profile.py` | Pilot 감시 루프 |
| `app/web/` | Trader UI |
| `app/adapter.py` | Profile → Config |
| `config/` | MODE 프리셋, Config dataclass |
| `source/` | 채팅 라인·호가 파서 |
| `excel/` | COM bind, yield/PnL |
| `core/` | Quote·트리거·로거 |
| `send/` | 붙여넣기·Enter |
| `admin/` | 서명·lease·정책 |
| `tests/` | pytest |

의존성: `requirements.txt`.

---

## 3. MODE와 입출력

| MODE | 소스 | 전송 |
|------|------|------|
| 1 | KBond 분리창 (`kbond_chat_title`) | 같은 창 |
| 2 | 동일 | Notepad (`메모장`) |
| 3 | FORESTBOND (UIA Text). 세션 고점보다 개수가 늘어난 줄만 신규. 깜빡임 복원은 재전송 없음 | Notepad |

설정은 Profile + `machine.json` + `DeveloperDefaults` (poll, sanity band, send pauses). `.env` / VBA는 사용하지 않는다.

---

## 4. 상태·중단

| 상태 | 의미 |
|------|------|
| WATCHING | 감시 중 |
| SENT | 전송 성공 |
| STOPPED | STOP / 탭 닫기 / stop 플래그 |
| ERROR | Excel 끊김, ambiguous, sanity, send/source 실패 등 — **전송 없음·즉시 종료** |

상세 메시지·대응: [`error_table.md`](error_table.md).  
배포 전 확인: [`pilot_smoke.md`](pilot_smoke.md).
