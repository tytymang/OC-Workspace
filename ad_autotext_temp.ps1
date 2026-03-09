Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$name = [Microsoft.VisualBasic.Interaction]::InputBox("상용구(단축어)로 등록할 임직원 이름(예: 최현구)을 입력하세요.", "마당쇠 - AD 자동 상용구 등록기", "")

if ([string]::IsNullOrWhiteSpace($name)) { exit }

$searcher = [adsisearcher]""
$searcher.Filter = "(&(objectCategory=person)(objectClass=user)(name=$name*))"
$searcher.PropertiesToLoad.Add("name") | Out-Null
$searcher.PropertiesToLoad.Add("mailnickname") | Out-Null
$results = $searcher.FindAll()

if ($results.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("'$name' 님을 AD에서 찾을 수 없습니다.", "마당쇠 알림", 0, [System.Windows.Forms.MessageBoxIcon]::Warning)
    exit
}

$user = $results[0].Properties
$fullName = $user["name"][0]
$empId = $user["mailnickname"][0]

$textToInsert = "$fullName 사번: $empId"

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $templatePath = "$env:APPDATA\Microsoft\Templates\NormalEmail.dotm"
    
    if (-not (Test-Path $templatePath)) {
        [System.Windows.Forms.MessageBox]::Show("NormalEmail.dotm 파일을 찾을 수 없습니다. 아웃룩을 먼저 실행해주세요.", "오류", 0, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
    
    $doc = $word.Documents.Open($templatePath)
    $range = $doc.Range()
    $range.Text = $textToInsert
    
    try {
        $doc.AttachedTemplate.AutoTextEntries.Item($name).Delete()
    } catch {}
    
    $doc.AttachedTemplate.AutoTextEntries.Add($name, $range) | Out-Null
    $doc.Save()
    $doc.Close()
    
    [System.Windows.Forms.MessageBox]::Show("[$name] 님의 상용구가 성공적으로 등록되었습니다!`n`n[아웃룩 입력 내용]`n$textToInsert", "마당쇠 보고", 0, [System.Windows.Forms.MessageBoxIcon]::Information)
} catch {
    [System.Windows.Forms.MessageBox]::Show("상용구 등록 중 오류가 발생했습니다:`n$_", "에러", 0, [System.Windows.Forms.MessageBoxIcon]::Error)
} finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}