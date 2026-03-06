
$ErrorActionPreference = "Stop"

try {
    # 1. Find File
    $files = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*20260305_02*.xlsx"
    $targetFile = $files[0].FullName
    
    # 2. Open Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($targetFile)
    $sheet = $workbook.Sheets.Item(1)
    
    # Dump 11-20 rows, 1-10 cols
    for ($r=11; $r -le 20; $r++) {
        $rowStr = ""
        for ($c=1; $c -le 10; $c++) {
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val) {
                $rowStr += "[$r,$c]: $val | "
            }
        }
        Write-Host $rowStr
    }

} catch {
    Write-Error $_
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { 
        $excel.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
