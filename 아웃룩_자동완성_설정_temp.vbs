Set objWord = CreateObject("Word.Application")
objWord.Visible = False
Set objFSO = CreateObject("Scripting.FileSystemObject")
strAppData = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%APPDATA%")
strTemplatePath = strAppData & "\Microsoft\Templates\NormalEmail.dotm"

If Not objFSO.FileExists(strTemplatePath) Then
    MsgBox "NormalEmail.dotm 파일을 찾을 수 없습니다. 아웃룩을 먼저 실행해주세요.", 48, "오류"
    objWord.Quit
    WScript.Quit
End If

Set objDoc = objWord.Documents.Open(strTemplatePath)
Set objRange = objDoc.Range()
objRange.Text = "최현구(Choi, HyunGoo/상무) 사번: 123456"

On Error Resume Next
objDoc.AttachedTemplate.AutoTextEntries.Item("최현구").Delete
On Error GoTo 0

objDoc.AttachedTemplate.AutoTextEntries.Add "최현구", objRange
objDoc.Save
objDoc.Close
objWord.Quit

MsgBox "설정이 완료되었습니다!" & vbCrLf & "아웃룩 새 메일에서 '최현구'를 치고 Enter를 누르면 자동 완성됩니다.", 64, "마당쇠 보고"