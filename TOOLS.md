# TOOLS.md - 돌쇠의 연장통 (Capability Map)

## 도구별 권한 및 제약

### Outlook (메일/일정)
- **권한:** 메일 스캔, 초안 생성, 일정 조회/초안 등록
- **제약:** 직접 발송 불가 — 초안 → 결재 → 전송 플로우
- **인코딩:** UTF-16 LE BOM 필수
- **메일 본문:** HTML 형식 (merged table, grouped rows, 간결 bullet)
- **일정:** Outlook에 직접 등록 → Google Calendar는 별도 설정

### Google Calendar
- **권한:** 일정 조회, 초안 등록
- **제약:** 직접 API 등록 금지 → Outlook 경유
- **참조:** [LESSON 2026-03-13]

### Excel
- **권한:** 읽기/쓰기, 데이터 정리, 보고서 생성
- **주의:** 큰 숫자 소수점/단위 변환 시 이중 검증
- **출력:** 바로 붙여넣기 가능한 형태

### Google Drive
- **권한:** 문서 읽기/쓰기, 공유 문서 접근
- **제약:** 권한 설정 변경은 주인님 확인 후
- **주의:** 공유 문서 내 개인정보 노출 금지

### Windows 파일탐색기
- **권한:** 워크스페이스 내 읽기/쓰기
- **제약:** 시스템 핵심 폴더 접근 불가
- **원칙:** 삭제는 항상 휴지통(trash) 우선, 파일명 영문 권장

### PowerShell
- **권한:** 자동화 스크립트 실행
- **제약:** exec 도구에서 `$` 특수문자 손실 [LESSON 2026-03-09]
- **원칙:** 파일 기반 실행 권장, VBS→PS1 체인 최소화
- **인코딩:** UTF-16 LE BOM + [System.IO.File]::WriteAllText()

### Agent Browser (Web Tool)
- **권한:** 정보 검색, 웹 기반 자동화
- **제약:** 사내 보안 영역 접근 불가

## OpenClaw 인프라
| 항목 | 값 |
|------|-----|
| Gateway | localhost:18792 |
| Copilot Relay | localhost:18795 |
| WebSocket | ws://127.0.0.1:18792 |
| 브라우저 프로필 | Comet → "openclaw", Chrome → "chrome" |
| 동기화 스크립트 | C:\Users\307984\.openclaw\workspace\sync_memory.py |