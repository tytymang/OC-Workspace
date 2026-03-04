Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(9) ' olFolderCalendar
Set colItems = objFolder.Items
Set objMeeting = colItems.Find("[Subject] = 'ָ  ӿȸ'")

If Not objMeeting Is Nothing Then
    objMeeting.Subject = "주말 스탭 임원회의"
    objMeeting.Save
    WScript.Echo "Outlook Meeting Fixed: 주말 스탭 임원회의"
Else
    ' 만약 위 제목으로 못 찾으면 오늘 날짜의 08시 일정을 찾음
    For Each item In colItems
        If item.Start = #02/28/2026 08:00:00 AM# Then
            item.Subject = "주말 스탭 임원회의"
            item.Save
            WScript.Echo "Outlook Meeting Fixed by Time: 주말 스탭 임원회의"
            Exit For
        End If
    Next
End If
