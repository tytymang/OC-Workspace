[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001

$outlook = New-Object -ComObject Outlook.Application
$appointment = $outlook.CreateItem(1) # olAppointmentItem

$appointment.Subject = "SCM 援ъ텞 ?묒쓽"
$appointment.Start = "2026-02-27 15:00"
$appointment.End = "2026-02-27 16:00"
$appointment.Location = "1痢?108 ?뚯쓽??(https://sskv.webex.com/sskv/j.php?MTID=m35ed789b69ac6c98c2496b844b12571f)"
$appointment.Body = @"
5. ?묒쓽 ?댁슜
- ?④퀎蹂?SCM 援ъ텞 ?꾨왂
- 1?④퀎 怨쇱젣 ?곸꽭 ?묒쓽 > T3 Smart SCM Platform Migration (湲곕뒫 媛쒖꽑, ?닿? ?꾨왂 嶺?
"@

$attendees = @("理쒗쁽援?, "?대룞??, "源?쒕┝", "源醫낆썝", "媛뺣룞誘?, "?≫븯??, "源?뱁쁽", "?뺥쁺怨?)
foreach ($person in $attendees) {
    $recipient = $appointment.Recipients.Add($person)
    $recipient.Type = 1 # olRequired
}

$appointment.Save()
Write-Output "Outlook ?쇱젙 ?깅줉 ?꾨즺: $($appointment.Subject)"

