
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
    
    # 3. Find Name
    $searchName = [char]0xC120 + [char]0xBD09 + [char]0xADDC
    
    # Search entire used range for name
    $found = $sheet.UsedRange.Find($searchName)
    
    if ($found) {
        Write-Host "Name found at Row $($found.Row), Col $($found.Column)"
        
        # Dump that row
        $r = $found.Row
        $rowStr = ""
        for ($c=1; $c -le 20; $c++) {
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val) {
                $rowStr += "[$r,$c]: $val | "
            }
        }
        Write-Host "Row Dump: $rowStr"
        
        # Dump header row (assumed to be row 1 or row above name)
        # But maybe header is far above. Let's dump rows around name - 10
        
    } else {
        Write-Host "Name not found in UsedRange"
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
