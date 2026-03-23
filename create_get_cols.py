import os

ps1_content = """
$ErrorActionPreference = "Stop"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open("C:\\Users\\307984\\.openclaw\\document\\최종_전사_AI_과제_마스터_리스트.xlsx")
    $ws = $wb.Worksheets.Item("Sheet1")
    $cols = $ws.UsedRange.Columns.Count
    for ($i=1; $i -le $cols; $i++) {
        Write-Output "$($i): $($ws.Cells.Item(1, $i).Text)"
    }
} finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
}
"""

with open(r"C:\Users\307984\.openclaw\workspace\get_target_cols_u16.ps1", "w", encoding="utf-16le") as f:
    f.write("\ufeff" + ps1_content)
