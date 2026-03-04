
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Calendar = $Namespace.GetDefaultFolder(9) # olFolderCalendar
$Items = $Calendar.Items

# 오늘 생성된 일정 중 제목이 깨졌거나 해당 패턴인 것 삭제
$Items | Where-Object { $_.Subject -like "*[출근]*" -or $_.Subject -like "*?*" } | ForEach-Object {
    if ($_.Start -gt [DateTime]::Now) {
        $_.Delete()
        Write-Host "DELETED: $($_.Subject) ($($_.Start.ToString('yyyy-MM-dd')))"
    }
}
