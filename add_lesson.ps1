$memory = Get-Content "C:\Users\307984\.openclaw\workspace\MEMORY.md" -Encoding UTF8
$lesson = "[LESSON] 2026-03-13 | 캘린더 | 사용자가 'Google Calendar 일정 등록'을 명시했을 때, Outlook 동기화를 이용하지 말고 브라우저를 통해 구글 캘린더(calendar.google.com) 웹에 직접 일정을 등록할 것."
$memory += $lesson
$memory | Out-File "C:\Users\307984\.openclaw\workspace\MEMORY.md" -Encoding UTF8
