
try {
    $outlook = New-Object -ComObject Outlook.Application
    "Outlook COM Success" | Out-File -FilePath "test.txt"
} catch {
    $_.Exception.Message | Out-File -FilePath "test.txt"
}
