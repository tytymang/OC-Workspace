$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $folder = "C:\Users\307984\.openclaw\workspace\temp_attachments"
    $files = Get-ChildItem $folder -Filter "*.xlsx"
    
    $workbook = $excel.Workbooks.Open($files[0].FullName)
    $sheet = $workbook.Sheets.Item(1)
    
    $results = @()
    # 상단 30행만 추출 (보통 여기에 주요 내용이 있음)
    for ($r = 1; $r -le 30; $r++) {
        $rowStr = ""
        for ($c = 1; $c -le 15; $c++) {
            $val = $sheet.Cells.Item($r, $c).Text
            if (![string]::IsNullOrWhiteSpace($val)) {
                $rowStr += "[$val] "
            }
        }
        if (![string]::IsNullOrWhiteSpace($rowStr)) {
            $results += $rowStr
        }
    }
    
    $workbook.Close($false)
    $excel.Quit()
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}