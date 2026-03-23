import os

ps1_content = """
$ErrorActionPreference = "Stop"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open("C:\\Users\\307984\\.openclaw\\document\\최종_전사_AI_과제_마스터_리스트.xlsx")
    $ws = $wb.Worksheets.Item("Sheet1")
    
    $data = @()
    for ($r=2; $r -le 4; $r++) {
        $row = @{}
        for ($c=1; $c -le 5; $c++) {
            $row[$c.ToString()] = $ws.Cells.Item($r, $c).Text
        }
        $data += $row
    }
    
    $data | ConvertTo-Json -Depth 5 | Out-File "C:\\Users\\307984\\.openclaw\\workspace\\target_sample.json" -Encoding UTF8

} finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
}
"""

with open(r"C:\Users\307984\.openclaw\workspace\get_target_data.py", "w", encoding="utf-16le") as f:
    f.write("\ufeff" + ps1_content)
