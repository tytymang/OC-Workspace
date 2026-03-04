
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$word = New-Object -ComObject Word.Application
try {
    $pdfPath = "C:\Users\307984\.openclaw\workspace\menu.pdf"
    $doc = $word.Documents.Open($pdfPath, $false, $true)
    
    # 2월 26일 메뉴가 있는 테이블을 찾아보겠습니다.
    foreach ($table in $doc.Tables) {
        $tableText = ""
        for ($r = 1; $r -le $table.Rows.Count; $r++) {
            for ($c = 1; $c -le $table.Columns.Count; $c++) {
                $cellText = $table.Cell($r, $c).Range.Text -replace "[\r\n\x07]", ""
                $tableText += $cellText + "`t"
            }
            $tableText += "`n"
        }
        if ($tableText -match "2/26" -or $tableText -match "목요일") {
            Write-Output "--- Table Found ---"
            Write-Output $tableText
        }
    }
    $doc.Close($false)
} catch {
    Write-Error $_.Exception.Message
} finally {
    $word.Quit()
}
