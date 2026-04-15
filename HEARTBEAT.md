# HEARTBEAT.md - 돌쇠의 순찰 (Periodic Awareness)

## 실행 주기
Windows 작업 스케줄러("OpenClaw Heartbeat")가 매 30분마다 heartbeat_check.ps1을 자동 실행하고,
결과를 memory/heartbeat-result.txt에 저장한다.

## ⛔ 금지 사항 (NON-NEGOTIABLE)
1. **PowerShell 직접 실행 절대 금지** — 돌쇠가 powershell, pwsh, .ps1 파일을 직접 실행하지 않는다. Windows 스케줄러가 알아서 실행한다.
2. **새 스크립트 파일 생성 절대 금지** — heartbeat_XXXX.ps1, check_vN.ps1 등 새 파일을 만들지 않는다.
3. **승인 창/팝업 절대 금지** — MessageBox, confirm, alert 등 어떤 형태의 팝업도 띄우지 않는다.
4. **중복 보고 금지** — 이전 순찰과 동일한 내용이면 보고하지 않는다.
5. **장황한 보고 금지** — 사과, 해명, 경과 설명 없이 결과만 보고한다.

## 순찰 방법
1. memory/heartbeat-result.txt 파일을 읽는다 (PowerShell 실행하지 않는다!)
2. 첫 줄의 CHECKED 타임스탬프가 이전 보고 시점보다 새로운지 확인한다
3. VIP_MAIL, MAIL, CAL 항목이 있으면 보고한다
4. HEARTBEAT_OK이면 아무 말도 하지 않는다

## 결과 파일 형식 (memory/heartbeat-result.txt)
- CHECKED|날짜시간 — 점검 시각
- VIP_MAIL|시간|발신자|제목 — VIP 메일 (우선 보고)
- MAIL|시간|발신자|제목 — 일반 새 메일
- CAL|시간|일정명 — 2시간 내 일정
- HEARTBEAT_OK — 특이사항 없음
- ERROR|메시지 — 점검 오류

## 보고 규칙
- **보고할 것이 있을 때만** 채팅으로 간결하게 보고한다
- HEARTBEAT_OK이면 완전 침묵
- 보고 형식: 표 또는 한두 줄 요약
- 밤 23시~아침 8시: 긴급 건(VIP 메일) 외에는 침묵
