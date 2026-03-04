Set objOutlook = CreateObject("Outlook.Application")
' "주말 스탭 임원회의"를 유니코드 문자로 조합
' 주(44514) 말(47568) (32) 스(49828) 탭(53409) (32) 임(51076) 원(50896) 회(54940) 의(51032)
meetingTitle = ChrW(51452) & ChrW(47568) & ChrW(32) & ChrW(49828) & ChrW(53409) & ChrW(32) & ChrW(51076) & ChrW(50896) & ChrW(54940) & ChrW(51032)

Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(9)
Set colItems = objFolder.Items

found = False
For Each item In colItems
    ' 2/28 08:00 일정을 찾아 제목 교체
    If DateDiff("n", item.Start, "2026-02-28 08:00") = 0 Then
        item.Subject = meetingTitle
        item.Save
        WScript.Echo "Success"
        found = True
        Exit For
    End If
Next

If Not found Then
    ' 만약 못 찾으면 새로 생성
    Set objAppt = objOutlook.CreateItem(1)
    objAppt.MeetingStatus = 1
    objAppt.Subject = meetingTitle
    objAppt.Start = "2026-02-28 08:00"
    objAppt.End = "2026-02-28 14:00"
    objAppt.Location = "본관접견실 108"
    objAppt.Save
    WScript.Echo "New Success"
End If
