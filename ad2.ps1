$ids = "405977", "307420", "306732", "308406", "402440"
$results = @()
foreach ($id in $ids) {
    $user = Get-ADUser -Filter "SamAccountName -eq '$id'" -Properties DisplayName, Department, Title
    if ($user) {
        $results += @{
            ID = $id
            Name = $user.DisplayName
            Department = $user.Department
            Title = $user.Title
        }
    }
}
$results | ConvertTo-Json -Depth 3 | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\ad.json" -Encoding UTF8
