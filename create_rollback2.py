import os

ps1_content = """
$ErrorActionPreference = "Stop"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $targetPath = "C:\\Users\\307984\\.openclaw\\document\\최종_전사_AI_과제_마스터_리스트.xlsx"
    $wbTarget = $excel.Workbooks.Open($targetPath)
    $wsTarget = $wbTarget.Worksheets.Item("Sheet1")
    
    $lr = $wsTarget.UsedRange.Rows.Count
    if ($lr -gt 22) {
        $rangeToDelete = $wsTarget.Range($wsTarget.Cells.Item(23, 1), $wsTarget.Cells.Item($lr, $wsTarget.UsedRange.Columns.Count))
        $rangeToDelete.EntireRow.Delete()
        $wbTarget.Save()
        Write-Output "Reverted target file back to 22 rows. (Deleted $($lr - 22) rows)"
    } else {
        Write-Output "Target file already at $lr rows. No rollback needed."
    }
} catch {
    Write-Output "Error: $_"
} finally {
    if ($wbTarget) { $wbTarget.Close($false) }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
    }
}
"""

with open(r"C:\Users\307984\.openclaw\workspace\rollback2_u16.ps1", "w", encoding="utf-16le") as f:
    f.write("\ufeff" + ps1_content)
