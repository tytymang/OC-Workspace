
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
$filePath = Join-Path $downloadsPath "서울반도체 중석식_2.23 (1).pdf"

if (Test-Path $filePath) {
    # .txt로 변환 시도 (Windows의 기본 기능을 이용해 텍스트 추출 시도)
    $word = New-Object -ComObject Word.Application
    try {
        $doc = $word.Documents.Open($filePath, $false, $true)
        $text = $doc.Content.Text
        # 2월 26일 또는 2/26 주변 텍스트 추출
        $index = $text.IndexOf("2/26")
        if ($index -lt 0) { $index = $text.IndexOf("2월 26일") }
        
        if ($index -ge 0) {
            $start = [Math]::Max(0, $index - 100)
            $length = [Math]::Min($text.Length - $start, 1000)
            Write-Output "--- EXTRACTED CONTENT ---"
            Write-Output $text.Substring($start, $length)
        } else {
            Write-Output "DATE_NOT_FOUND_IN_TEXT"
            # 전체 텍스트 일부 출력
            Write-Output $text.Substring(0, [Math]::Min($text.Length, 500))
        }
        $doc.Close($false)
    } finally {
        $word.Quit()
    }
} else {
    Write-Output "FILE_NOT_FOUND"
}
