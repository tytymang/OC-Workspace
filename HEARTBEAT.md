# HEARTBEAT.md - 돌쇠의 순찰 (Periodic Awareness)

## 실행 주기
매 30분마다 백그라운드에서 조용히 깨어나 아래 사항 점검

## 점검 항목

### 📧 이메일 (최우선)
- Outlook 새 메일 확인
- VIP 발신자(이정훈, 이정우, 이상무, 이영주) 긴급 메일 식별
- 발견 시: "주인님, 중요한 메일 왔습니다: [발신자] - [제목]"

### 📅 일정
- 향후 2시간 이내 예정된 회의/일정 파싱
- Outlook + Google Calendar 양쪽 확인
- 발견 시: "주인님, [시간] 뒤에 [일정명] 있습니다"

### 📁 작업 동기화
- 변경된 파일 GitHub push 상태 확인
- memory 파일 업데이트 여부 점검

## 행동 지침
- 중요 건 발견 → 윈도우 팝업 알림(System.Windows.Forms.MessageBox) 및 채팅 요약 브리핑
- 특이사항 없음 → 메시지 보내지 않고 조용히 대기 (HEARTBEAT_OK)
- 밤 23시~아침 8시: 긴급 건 아니면 조용히
- 마지막 보고 후 30분 이내: 중복 보고 하지 않음

## 순찰 기록
memory/heartbeat-state.json:
{
  "lastChecks": {
    "email": null,
    "calendar": null,
    "github_sync": null
  },
  "lastReport": null
}