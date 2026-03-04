
$o = New-Object -ComObject Outlook.Application
$a = $o.CreateItem(1)
$a.Subject = "[회의] 재고상태 구분 신규추가 건 확인 및 운영 적용"
$a.Start = "2026-02-26 10:30:00"
$a.Duration = 60
$a.Location = "106 접견실 (https://sskv.webex.com/sskv/j.php?MTID=mb9e45aa44dbb35ccfc41197aabb63e1b)"
$a.Body = "김종원 선임님, 김승현 대리님, 이동석 팀장님, 육하나님`n`n어제 공유해주신 '재고상태 구분 신규추가' 건과 관련하여, 금일 오전 확인 및 운영 적용을 위한 미팅을 제안드립니다.`n`n[회의 링크]`nhttps://sskv.webex.com/sskv/j.php?MTID=mb9e45aa44dbb35ccfc41197aabb63e1b`n`n[주요 확인 사항]`n- 기존 1월 재고 꼬리표와 차이분석`n- 1개월 미만 연령 재고 생성 정상 여부`n- 재고 배분 금액 확인 (Detail)`n- 기존 등록된 재고 소진계획 유지 여부`n`n문제 없으면 기존 1월 재고 꼬리표에 Update 진행 예정입니다.`n`n감사합니다."
$a.MeetingStatus = 1
$null = $a.Recipients.Add("김종원")
$null = $a.Recipients.Add("김승현")
$null = $a.Recipients.Add("이동석")
$null = $a.Recipients.Add("육하나")
$a.Display()
