[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ids = "405977", "307420", "306732", "308406", "402440"
foreach ($id in $ids) {
    Get-ADUser -Filter "SamAccountName -eq '$id'" -Properties DisplayName, Department, Title | Select-Object SamAccountName, DisplayName, Department, Title
}
