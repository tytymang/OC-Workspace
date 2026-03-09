# 1. 파일 이름 변경 (xlsx -> zip)
$zipPath = "C:\Users\307984\.openclaw\workspace\temp_exchange.zip"
Copy-Item "C:\Users\307984\.openclaw\workspace\temp_exchange.xlsx" $zipPath -Force

# 2. 압축 해제
$extractPath = "C:\Users\307984\.openclaw\workspace\temp_extract"
if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

# 3. 데이터 검색 (sharedStrings.xml + sheet*.xml)
# sharedStrings.xml에는 모든 텍스트 데이터가 들어있음
$sharedStringsPath = "$extractPath\xl\sharedStrings.xml"

if (Test-Path $sharedStringsPath) {
    $xml = [xml](Get-Content $sharedStringsPath)
    $texts = $xml.sst.si.t
    
    # 2025.12, 2026.01, 2026.02 등 검색
    $targets = @("2025.12", "2026.01", "2026.02", "2025-12", "2026-01", "2026-02", "Dec-25", "Jan-26", "Feb-26")
    
    foreach ($t in $texts) {
        foreach ($target in $targets) {
            if ($t -like "*$target*") {
                Write-Host "FOUND TEXT: $t"
            }
        }
    }
} else {
    Write-Host "No shared strings found. Scanning sheets directly..."
}

# 4. 시트 데이터 스캔 (숫자값 확인)
# 2025.12 근처에 있는 숫자(환율)를 찾아야 함.
# 하지만 XML 구조가 복잡하므로, 텍스트로 덤프해서 패턴 매칭

$sheets = Get-ChildItem "$extractPath\xl\worksheets\sheet*.xml"
foreach ($s in $sheets) {
    $content = Get-Content $s.FullName
    # 정규식으로 '1300.50' 같은 환율 패턴 검색
    # 2025.12와 가까운 위치에 있는 숫자를 찾아야 함.
    
    # 간단히 1000~2000 사이 숫자 검색
    $matches = [regex]::Matches($content, ">(1[0-4][0-9][0-9]\.?[0-9]*)<")
    foreach ($m in $matches) {
        Write-Host "CANDIDATE: $($m.Groups[1].Value)"
    }
}
