Add-Type -AssemblyName System.Windows.Forms
$msg = "주인님, 8:30에 [외주 P/O 분할 입고 현황] 일정이 있습니다.`n`n최근 중요 메일:`n1. 이예진 주임: 임원 KPI 평가 요청`n2. 손민혁 상무: G2(8GR) 회의 내용 공유"
[System.Windows.Forms.MessageBox]::Show($msg, "돌쇠 순찰 알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
