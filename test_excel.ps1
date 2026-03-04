
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$excel = New-Object -ComObject Excel.Application
try {
    $pdfPath = "C:\Users\307984\.openclaw\workspace\menu.pdf"
    # 엑셀은 PDF를 직접 열 수 없지만, 워드 실패 시 시도해볼 수 있는 대안
    # 하지만 보통 식단표는 엑셀에서 PDF로 변환되는 경우가 많으므로 파일명을 다시 확인
    Write-Output "Excel object created successfully"
} catch {
    Write-Error $_.Exception.Message
} finally {
    $excel.Quit()
}
