변경된 .env (실측)
키	값
SOURCE_WINDOW_TITLE
K-Bond
SOURCE_PROCESS_NAME
KBondMessenger.exe
SEND_PROCESS_NAME
KBondMessenger.exe
SEND_WINDOW_TITLE
K-Bond
SEND_INPUT_X / Y
0.825 / 0.940
계획서의 KBond.exe가 아니라 실제 프로세스는 KBondMessenger.exe, 창 제목은 K-Bond Messenger 입니다.

코드
UIA Desktop에 KBond가 안 잡혀 source_reader에 Win32 창 찾기 + SOURCE_PROCESS_NAME 필터를 추가했습니다. README/문서 기본값도 KBond로 갱신했습니다.

검증
--diagnose-send / --test-send: 성공 (25-10 00+ 확정 전송)
pytest: 35 passed
--diagnose-source: 창은 찾음, 호가 라인은 0건
KBond가 Delphi(TElTree)라 채팅 텍스트가 UIA로 노출되지 않습니다. 호가 감시가 되려면 채팅방이 열린 상태에서 텍스트가 읽히는 경로를 추가 확인하거나, ElTree 전용 읽기가 필요합니다. 채팅방을 연 뒤 --diagnose-source를 다시 돌려 보시면 됩니다.
